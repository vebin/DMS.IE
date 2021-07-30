using DMS.Excel.Attributes.Export;
using DMS.Excel.Models;
using DMSN.Common.Extensions;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace DMS.Excel
{
    public class ExportHelper<T> where T : class, new()
    {
        private List<ExporterHeaderInfo> _exporterHeaderInfoList;
        private ExcelExporterAttribute _excelExporterAttribute;

        private ExcelWorksheet _excelWorksheet;
        private ExcelPackage _excelPackage;
        /// <summary>
        /// 当前Sheet索引
        /// </summary>
        protected int SheetIndex = 0;
        /// <summary>
        /// 导出类全局配置
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
                                //AutoFitMaxRows = exporterAttribute.AutoFitMaxRows,
                                FontSize = exporterAttribute.FontSize,
                                HeaderFontSize = exporterAttribute.HeaderFontSize,
                                MaxRowNumberOnASheet = exporterAttribute.MaxRowNumberOnASheet,
                                Name = exporterAttribute.Name,
                                TableStyle = _excelExporterAttribute?.TableStyle ?? TableStyles.None,
                                AutoCenter = _excelExporterAttribute != null && _excelExporterAttribute.AutoCenter,
                                IsBold = exporterAttribute.IsBold,
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
                    //.OrderBy(p => p.GetAttribute<ExporterHeaderAttribute>()?.ColumnIndex ?? 10000)
                    .ToList();
                return objProperties;
            }
        }

        /// <summary>
        /// 表头列表
        /// </summary>
        protected List<ExporterHeaderInfo> ExporterHeaderInfoList
        {
            get
            {
                if (_exporterHeaderInfoList == null || _exporterHeaderInfoList.Count <= 0)
                {
                    _exporterHeaderInfoList = new List<ExporterHeaderInfo>();
                    var objProperties = SortedProperties;
                    if (objProperties.Count == 0)
                        return _exporterHeaderInfoList;

                    for (var i = 0; i < objProperties.Count; i++)
                    {
                        var item = new ExporterHeaderInfo
                        {
                            Index = i + 1,
                            PropertyName = objProperties[i].Name,
                            ExporterHeaderAttribute = (objProperties[i].GetCustomAttributes(typeof(ExporterHeaderAttribute), true) as ExporterHeaderAttribute[])?.FirstOrDefault(),
                            CsTypeName = objProperties[i].PropertyType.GetTypeName(),
                            ExportImageFieldAttribute = objProperties[i].GetAttribute<ExportImageFieldAttribute>(true),

                        };

                        //设置列显示名
                        item.DisplayName = item.ExporterHeaderAttribute == null ||
                                           item.ExporterHeaderAttribute.DisplayName.IsNullOrEmpty()
                            ? item.PropertyName
                            : item.ExporterHeaderAttribute.DisplayName;

                        ////设置Format
                        //if (item.ExporterHeaderAttribute != null && !item.ExporterHeaderAttribute.Format.IsNullOrEmpty())
                        //{
                        //    item.ExporterHeaderAttribute.Format = item.ExporterHeaderAttribute.Format;
                        //}

                        _exporterHeaderInfoList.Add(item);
                    }
                }

                return _exporterHeaderInfoList;
            }
            set => _exporterHeaderInfoList = value;
        }
        /// <summary>
        /// 
        /// </summary>
        public ExcelPackage CurrentExcelPackage
        {
            get
            {
                if (_excelPackage == null)
                {
                    _excelPackage = new ExcelPackage();

                    if (ExcelExporterSettings?.Author != null)
                        _excelPackage.Workbook.Properties.Author = ExcelExporterSettings?.Author;
                }

                return _excelPackage;
            }
            set => _excelPackage = value;
        }
        protected ExcelWorksheet CurrentExcelWorksheet
        {
            get
            {
                if (_excelWorksheet == null)
                {
                    AddExcelWorksheet();
                }

                return _excelWorksheet;
            }
            set => _excelWorksheet = value;
        }

        public ExcelWorksheet AddExcelWorksheet(string name = null)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                name = ExcelExporterSettings?.Name ?? "导出结果";
            }

            if (SheetIndex != 0)
            {
                name += "-" + SheetIndex;
            }

            _excelWorksheet = CurrentExcelPackage.Workbook.Worksheets.Add(name);
            _excelWorksheet.OutLineApplyStyle = true;
            return _excelWorksheet;
        }

        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <returns>文件</returns>
        public virtual ExcelPackage Export(ICollection<T> dataItems)
        {
            AddDataItems(dataItems);
            AddHeader();
            AddColumnStyle();
            SheetIndex++;
            return CurrentExcelPackage;
        }

        protected void AddDataItems(ICollection<T> dataItems)
        {
            ExcelRangeBase excelRange = CurrentExcelWorksheet.Cells["A1"];
            excelRange.LoadFromCollection(dataItems, true, ExcelExporterSettings.TableStyle);
        }

        /// <summary>
        /// 设置头部样式
        /// </summary>
        protected void AddHeader()
        {
            //全局剧中
            if (ExcelExporterSettings.AutoCenter)
            {
                CurrentExcelWorksheet.Cells[1, 1, CurrentExcelWorksheet.Dimension?.End.Row ?? 10, ExporterHeaderInfoList.Count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            //头部全部加粗
            if (ExcelExporterSettings.IsBold)
            {
                CurrentExcelWorksheet.Cells[1, 1, 1, ExporterHeaderInfoList.Count].Style.Font.Bold = ExcelExporterSettings.IsBold;
            }

            foreach (var exporterHeader in ExporterHeaderInfoList)
            {
                var colCell = CurrentExcelWorksheet.Cells[1, exporterHeader.Index];
                colCell.Value = exporterHeader.DisplayName;

                var exporterHeaderAttribute = exporterHeader.ExporterHeaderAttribute;
                if (exporterHeaderAttribute != null)
                {
                    //colCell.Style.Font.Bold = exporterHeaderAttribute.IsBold;//当前字段加粗
                    var size = ExcelExporterSettings?.HeaderFontSize ?? exporterHeaderAttribute.FontSize;
                    if (size.HasValue)
                        colCell.Style.Font.Size = size.Value;
                }
            }
        }
        /// <summary>
        /// 添加列的样式
        /// </summary>
        protected void AddColumnStyle()
        {
            foreach (var exporterHeader in ExporterHeaderInfoList)
            {
                var colColumn = CurrentExcelWorksheet.Column(exporterHeader.Index);
                if (exporterHeader.ExporterHeaderAttribute != null && !string.IsNullOrWhiteSpace(exporterHeader.ExporterHeaderAttribute.Format))
                {
                    colColumn.Style.Numberformat.Format = exporterHeader.ExporterHeaderAttribute.Format;

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
                            colColumn.Style.Numberformat.Format = CultureInfo.CurrentUICulture.DateTimeFormat.FullDateTimePattern;
                            break;
                        default:
                            break;
                    }
                }
                if (ExcelExporterSettings.AutoFitAllColumn)
                {
                    colColumn.AutoFit();
                }

                var exporterHeaderAttribute = exporterHeader.ExporterHeaderAttribute;
                if (exporterHeaderAttribute != null)
                {
                    //设置单元格宽度
                    var width = exporterHeader.ExporterHeaderAttribute.Width;
                    if (width > 0)
                    {
                        colColumn.Width = width;
                    }

                    //自动换行未起效果
                    //if (exporterHeader.ExporterHeaderAttribute.WrapText)
                    //{
                    //    colColumn.Style.WrapText = exporterHeader.ExporterHeaderAttribute.WrapText;
                    //}
                    colColumn.Hidden = exporterHeader.ExporterHeaderAttribute.Hidden;
                }


            }
        }
    }
}
