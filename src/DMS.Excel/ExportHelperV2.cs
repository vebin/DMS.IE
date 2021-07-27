using DMS.Excel.Attributes.Export;
using DMS.Excel.Extension;
using DMS.Excel.Models;
using DMSN.Common.Extensions;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace DMS.Excel
{
    /// <summary>
    /// 导出辅助类
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ExportHelperV2<T> where T : class, new()
    {
        private ExcelPackage _excelPackage;
        private ExcelWorksheet _excelWorksheet;
        private ExcelExporterAttribute _excelExporterAttribute;
        private List<ExporterHeaderInfo> _exporterHeaderList;
        /// <summary>
        /// 当前工作
        /// </summary>
        protected List<ExcelWorksheet> ExcelWorksheets { get; set; } = new List<ExcelWorksheet>();
        /// <summary>
        /// 当前Sheet索引
        /// </summary>
        protected int SheetIndex = 0;
        public ExportHelperV2()
        {

        }


        /// <summary>
        /// 导出类设置
        /// </summary>
        public ExcelExporterAttribute ExcelExporterSettings
        {
            get
            {
                if (_excelExporterAttribute == null)
                {
                    var type = typeof(T);
                    if (typeof(DataTable) == type)
                    {
                        _excelExporterAttribute = new ExcelExporterAttribute();
                    }
                    else
                        _excelExporterAttribute = type.GetAttribute<ExcelExporterAttribute>(true);

                    if (_excelExporterAttribute == null)
                    {
                        var exporterAttribute = type.GetAttribute<ExporterAttribute>(true);
                        if (exporterAttribute != null)
                        {
                            _excelExporterAttribute = new ExcelExporterAttribute()
                            {
                                Author = exporterAttribute.Author,
                                AutoFitAllColumn = exporterAttribute.AutoFitAllColumn,
                                AutoFitMaxRows = exporterAttribute.AutoFitMaxRows,
                                FontSize = exporterAttribute.FontSize,
                                HeaderFontSize = exporterAttribute.HeaderFontSize,
                                MaxRowNumberOnASheet = exporterAttribute.MaxRowNumberOnASheet,
                                Name = exporterAttribute.Name,
                                TableStyle = _excelExporterAttribute?.TableStyle ?? TableStyles.None,
                                AutoCenter = _excelExporterAttribute != null && _excelExporterAttribute.AutoCenter,
                            };
                        }
                        else
                            _excelExporterAttribute = new ExcelExporterAttribute();
                    }
                }

                return _excelExporterAttribute;
            }
            set => _excelExporterAttribute = value;
        }
        /// <summary>
        /// 排序的属性
        /// </summary>
        protected virtual List<PropertyInfo> SortedProperties
        {
            get
            {
                var type = typeof(T);
                var objProperties = type.GetProperties()
                    .OrderBy(p => p.GetAttribute<ExporterHeaderAttribute>()?.ColumnIndex ?? 10000).ToList();
                return objProperties;
            }
        }
        /// <summary>
        /// 当前Excel包
        /// </summary>
        public ExcelPackage CurrentExcelPackage
        {
            get
            {
                if (_excelPackage == null)
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    _excelPackage = new ExcelPackage();

                    if (ExcelExporterSettings?.Author != null)
                        _excelPackage.Workbook.Properties.Author = ExcelExporterSettings?.Author;
                }

                return _excelPackage;
            }
            set => _excelPackage = value;
        }
        /// <summary>
        /// 当前工作Sheet
        /// </summary>
        protected ExcelWorksheet CurrentExcelWorksheet
        {
            get
            {
                if (_excelWorksheet == null)
                {
                    var name = ExcelExporterSettings?.Name ?? "导出结果";
                    if (SheetIndex != 0)
                    {
                        name += "-" + SheetIndex;
                    }

                    _excelWorksheet = CurrentExcelPackage.Workbook.Worksheets.Add(name);
                    _excelWorksheet.OutLineApplyStyle = true;
                    ExcelWorksheets.Add(_excelWorksheet);
                }
                return _excelWorksheet;
            }
            set => _excelWorksheet = value;
        }
        /// <summary>
        /// 表头列表
        /// </summary>
        protected List<ExporterHeaderInfo> ExporterHeaderList
        {
            get
            {
                if (_exporterHeaderList == null)
                {
                    GetExporterHeaderInfoList();
                }

                return _exporterHeaderList;
            }
            set => _exporterHeaderList = value;
        }
        /// <summary>
        /// 获取头部定义
        /// </summary>
        /// <returns></returns>
        protected virtual void GetExporterHeaderInfoList(DataTable dt = null, ICollection<T> dataItems = null)
        {
            _exporterHeaderList = new List<ExporterHeaderInfo>();

            //var type = _type ?? typeof(T);
            //#179 GetProperties方法不按特定顺序（如字母顺序或声明顺序）返回属性，因此此处支持按ColumnIndex排序返回
            //var objProperties = type.GetProperties().OrderBy(p => p.GetAttribute<ExporterHeaderAttribute>()?.ColumnIndex ?? 10000).ToArray();
            var objProperties = SortedProperties;
            if (objProperties.Count == 0)
                return;
            for (var i = 0; i < objProperties.Count; i++)
            {

                var item = new ExporterHeaderInfo
                {
                    Index = i + 1,
                    PropertyName = objProperties[i].Name,
                    ExporterHeaderAttribute =
                        (objProperties[i].GetCustomAttributes(typeof(ExporterHeaderAttribute), true) as
                            ExporterHeaderAttribute[])?.FirstOrDefault() ??
                        new ExporterHeaderAttribute(objProperties[i].GetDisplayName() ?? objProperties[i].Name),
                    CsTypeName = objProperties[i].PropertyType.GetCSharpTypeName(),
                    ExportImageFieldAttribute = objProperties[i].GetAttribute<ExportImageFieldAttribute>(true),

                };

                //设置列显示名
                item.DisplayName = item.ExporterHeaderAttribute == null ||
                                   item.ExporterHeaderAttribute.DisplayName == null ||
                                   item.ExporterHeaderAttribute.DisplayName.IsNullOrEmpty()
                    ? item.PropertyName
                    : item.ExporterHeaderAttribute.DisplayName;
                //设置Format
                item.ExporterHeaderAttribute.Format = item.ExporterHeaderAttribute.Format.IsNullOrEmpty()
                    ? objProperties[i].GetDisplayFormat()
                    : item.ExporterHeaderAttribute.Format;

                _exporterHeaderList.Add(item);
            }
        }


        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <returns>文件</returns>
        public virtual ExcelPackage Export(ICollection<T> dataItems)
        {
            var data = ParseData(dataItems);

            AddDataItems(data);


            //// 为了传入dataItems，在这里提前调用一下
            //if (_exporterHeaderList == null) GetExporterHeaderInfoList(null, dataItems);


            //DisableAutoFitWhenDataRowsIsLarge(dataItems.Count);
            return AddHeaderAndStyles();
        }
        /// <summary>
        /// 解析数据
        /// </summary>
        /// <param name="dataItems"></param>
        /// <returns></returns>
        protected virtual IEnumerable<ExpandoObject> ParseData(ICollection<T> dataItems)
        {
            var type = typeof(T);
            var properties = SortedProperties;
            List<ExpandoObject> list = new List<ExpandoObject>();
            foreach (var dataItem in dataItems)
            {
                dynamic obj = new ExpandoObject();
                foreach (var propertyInfo in properties)
                {
                    if (propertyInfo.PropertyType.GetCSharpTypeName() == "Boolean")
                    {
                        var col = ExporterHeaderList.First(a => a.PropertyName == propertyInfo.Name);
                        var val = type.GetProperty(propertyInfo.Name)?.GetValue(dataItem).ToString();
                        bool value = Convert.ToBoolean(val);

                        ((IDictionary<string, object>)obj)[propertyInfo.Name] = value;
                    }
                    else if (propertyInfo.PropertyType.GetCSharpTypeName() == "Nullable<Boolean>")
                    {
                        var col = ExporterHeaderList.First(a => a.PropertyName == propertyInfo.Name);
                        var value = Convert.ToBoolean(type.GetProperty(propertyInfo.Name)?.GetValue(dataItem));

                        ((IDictionary<string, object>)obj)[propertyInfo.Name] = value;
                    }
                    else
                    {
                        ((IDictionary<string, object>)obj)[propertyInfo.Name] = type.GetProperty(propertyInfo.Name)?.GetValue(dataItem)?.ToString();
                    }
                }

                //yield return obj;
                list.Add(obj);
            }
            //list.Add(obj);
            return list;
        }
        /// <summary>
        /// 添加导出数据
        /// </summary>
        /// <param name="dataItems"></param>
        /// <param name="excelRange"></param>
        protected void AddDataItems(IEnumerable<ExpandoObject> dataItems)
        {
            var excelRange = CurrentExcelWorksheet.Cells["A1"];
            if (dataItems == null || !dataItems.Any())
            {
                return;
            }

            excelRange.LoadFromDictionaries(dataItems, true, ExcelExporterSettings.TableStyle);
        }

        /// <summary>
        /// 添加表头、样式以及忽略列、格式处理
        /// </summary>
        /// <returns></returns>
        private ExcelPackage AddHeaderAndStyles()
        {
            AddHeader();

            if (ExcelExporterSettings.AutoFitAllColumn)
            {
                CurrentExcelWorksheet.Cells[CurrentExcelWorksheet.Dimension.Address].AutoFitColumns();
            }

            AddStyle();
            //DeleteIgnoreColumns();
            //以便支持导出多Sheet
            SheetIndex++;
            SetSkipRows();
            return CurrentExcelPackage;
        }
        /// <summary>
        /// 创建表头
        /// </summary>
        protected void AddHeader()
        {
            foreach (var exporterHeaderDto in ExporterHeaderList)
            {
                var exporterHeaderAttribute = exporterHeaderDto.ExporterHeaderAttribute;
                if (exporterHeaderAttribute != null)
                {
                    var colCell = CurrentExcelWorksheet.Cells[1, exporterHeaderDto.Index];
                    colCell.Style.Font.Bold = exporterHeaderAttribute.IsBold;


                    colCell.Value = exporterHeaderDto.DisplayName;


                    var size = ExcelExporterSettings?.HeaderFontSize ?? exporterHeaderAttribute.FontSize;
                    if (size.HasValue)
                        colCell.Style.Font.Size = size.Value;
                }
            }
        }
        /// <summary>
        /// 添加样式
        /// </summary>
        protected virtual void AddStyle()
        {
            foreach (var exporterHeader in ExporterHeaderList)
            {
                var col = CurrentExcelWorksheet.Column(exporterHeader.Index);
                if (exporterHeader.ExporterHeaderAttribute != null && !string.IsNullOrWhiteSpace(exporterHeader.ExporterHeaderAttribute.Format))
                {
                    col.Style.Numberformat.Format = exporterHeader.ExporterHeaderAttribute.Format;

                }
                else
                {
                    //处理日期格式
                    switch (exporterHeader.CsTypeName)
                    {
                        case "DateTime":
                        case "DateTimeOffset":
                        //case "DateTime?":
                        case "Nullable<DateTime>":
                        case "Nullable<DateTimeOffset>":
                            //设置本地化时间格式
                            col.Style.Numberformat.Format = CultureInfo.CurrentUICulture.DateTimeFormat.FullDateTimePattern;
                            break;
                        default:
                            break;
                    }
                }

                if (!ExcelExporterSettings.AutoFitAllColumn && exporterHeader.ExporterHeaderAttribute != null && exporterHeader.ExporterHeaderAttribute.IsAutoFit)
                    col.AutoFit();

                if (exporterHeader.ExportImageFieldAttribute != null)
                {
                    col.Width = exporterHeader.ExportImageFieldAttribute.Width;
                }

                if (exporterHeader.ExporterHeaderAttribute != null)
                {
                    //设置单元格宽度
                    var width = exporterHeader.ExporterHeaderAttribute.Width;
                    if (width > 0)
                    {
                        col.Width = width;
                    }

                    if (exporterHeader.ExporterHeaderAttribute.AutoCenterColumn)
                    {
                        col.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }

                    if (exporterHeader.ExporterHeaderAttribute.WrapText)
                    {
                        col.Style.WrapText = exporterHeader.ExporterHeaderAttribute.WrapText;
                    }
                    col.Hidden = exporterHeader.ExporterHeaderAttribute.Hidden;
                }
            }
        }
        /// <summary>
        ///设置x行开始追加内容
        /// </summary>
        private void SetSkipRows()
        {
            if (ExcelExporterSettings.HeaderRowIndex > 1)
            {
                CurrentExcelWorksheet.InsertRow(1, ExcelExporterSettings.HeaderRowIndex);
            }
        }
    }
}
