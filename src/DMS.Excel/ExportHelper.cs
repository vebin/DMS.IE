using DMS.Common.Extensions;
using DMS.Excel.Attributes.Export;
using DMS.Excel.Extension;
using DMS.Excel.Models;
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
    public class ExportHelper<T> where T : class, new()
    {
        private ExcelPackage _excelPackage;
        private ExcelWorksheet _excelWorksheet;
        private ExcelExporterAttribute _excelExporterAttribute;
        private List<ExporterHeaderInfo> _exporterHeaderList;
        /// <summary>
        /// 当前Sheet索引
        /// </summary>
        protected int SheetIndex = 0;
        /// <summary>
        /// 当前工作
        /// </summary>
        protected List<ExcelWorksheet> ExcelWorksheets { get; set; } = new List<ExcelWorksheet>();
        /// <summary>
        /// 是否为动态DataTable导出
        /// </summary>
        protected bool IsDynamicDatableExport { get; set; }
        /// <summary>
        /// 是否为ExpandoObject类型，调用LoadFromDictionaries以支持动态导出
        /// </summary>
        protected bool IsExpandoObjectType { get; set; }
        public ExportHelper()
        {
            if (typeof(DataTable).Equals(typeof(T)))
            {
                IsDynamicDatableExport = true;
            }
            if (typeof(ExpandoObject).Equals(typeof(T)))
            {
                IsExpandoObjectType = true;
            }
        }


        /// <summary>
        /// 导出设置
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
                    AddExcelWorksheet();
                }


                return _excelWorksheet;
            }
            set => _excelWorksheet = value;
        }
        /// <summary>
        /// 添加Sheet
        /// 支持同一个数据拆成多个Sheet
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public ExcelWorksheet AddExcelWorksheet()
        {
            var name = ExcelExporterSettings?.Name ?? "导出结果";
            if (SheetIndex != 0)
            {
                name += "-" + SheetIndex;
            }

            _excelWorksheet = CurrentExcelPackage.Workbook.Worksheets.Add(name);
            _excelWorksheet.OutLineApplyStyle = true;
            ExcelWorksheets.Add(_excelWorksheet);
            return _excelWorksheet;
        }


        /// <summary>
        /// Excel数据表
        /// </summary>
        protected ExcelTable CurrentExcelTable { get; set; }
        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <returns>文件</returns>
        public virtual ExcelPackage Export(ICollection<T> dataItems)
        {
            if (!IsExpandoObjectType)
            {
                var data = ParseData(dataItems);
                AddDataItems(data);
            }
            else
            {
                //ExpandoObject类型
                AddDataItems(dataItems);
            }

            // 为了传入dataItems，在这里提前调用一下
            if (_exporterHeaderList == null) GetExporterHeaderInfoList(null, dataItems);
            //仅当存在图片表头才渲染图片
            if (ExporterHeaderList.Any(p => p.ExportImageFieldAttribute != null))
            {
                //AddPictures(dataItems.Count);
            }

            DisableAutoFitWhenDataRowsIsLarge(dataItems.Count);
            return AddHeaderAndStyles();
        }
        /// <summary>
        /// 解析数据dto
        /// </summary>
        /// <param name="dataItems"></param>
        /// <returns></returns>
        protected virtual IEnumerable<ExpandoObject> ParseData(ICollection<T> dataItems)
        {
            var type = typeof(T);
            var properties = SortedProperties;
            foreach (var dataItem in dataItems)
            {
                dynamic obj = new ExpandoObject();
                foreach (var propertyInfo in properties)
                {
                    if (propertyInfo.PropertyType.IsEnum)
                    {
                        //var col = ExporterHeaderList.First(a => a.PropertyName == propertyInfo.Name);
                        //var value = type.GetProperty(propertyInfo.Name)?.GetValue(dataItem)?.ToString();

                        //if (col.MappingValues.Count > 0 && col.MappingValues.ContainsValue(value ?? string.Empty))
                        //{
                        //    var mapValue = col.MappingValues.FirstOrDefault(f => f.Key == value);
                        //    ((IDictionary<string, object>)obj)[propertyInfo.Name] = mapValue.Value;
                        //}
                        //else
                        //{

                        //}
                        if (
                            propertyInfo.PropertyType.IsEnum ||
                            propertyInfo.PropertyType.GetNullableUnderlyingType() != null &&
                            propertyInfo.PropertyType.GetNullableUnderlyingType().IsEnum)
                        {
                            {
                                var value = (int)type.GetProperty(propertyInfo.Name)?.GetValue(dataItem);
                                {
                                    var col = ExporterHeaderList.First(a => a.PropertyName == propertyInfo.Name);

                                    if (false)// if (col.MappingValues.Count > 0 && col.MappingValues.ContainsValue(value))
                                    {
                                        //var mapValue = col.MappingValues.FirstOrDefault(f => f.Value == value);
                                        //dr[propertyInfo.Name] = mapValue.Value;
                                        //((IDictionary<string, object>)obj)[propertyInfo.Name] = mapValue.Key;
                                    }
                                    else
                                    {
                                        var enumDefinitionList = propertyInfo.PropertyType.GetEnumDefinitionList();
                                        if (enumDefinitionList == null)
                                        {
                                            enumDefinitionList = propertyInfo.PropertyType.GetNullableUnderlyingType()
                                                .GetEnumDefinitionList();
                                        }

                                        var tuple = enumDefinitionList.FirstOrDefault(f => f.Item1 == value.ToString());
                                        if (tuple != null)
                                        {
                                            if (!tuple.Item4.IsNullOrEmpty())
                                            {
                                                //dr[propertyInfo.Name] = tuple.Item4;
                                                ((IDictionary<string, object>)obj)[propertyInfo.Name] = tuple.Item4;
                                            }
                                            else
                                            {
                                                ((IDictionary<string, object>)obj)[propertyInfo.Name] = tuple.Item2;
                                                // dr[propertyInfo.Name] = tuple.Item2;
                                            }
                                        }
                                        else
                                        {
                                            ((IDictionary<string, object>)obj)[propertyInfo.Name] = value;
                                            //dr[propertyInfo.Name] = value;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else if (propertyInfo.PropertyType.GetCSharpTypeName() == "Boolean")
                    {
                        var col = ExporterHeaderList.First(a => a.PropertyName == propertyInfo.Name);
                        var val = type.GetProperty(propertyInfo.Name)?.GetValue(dataItem).ToString();
                        bool value = Convert.ToBoolean(val);
                        if (false)//col.MappingValues.Count > 0 && col.MappingValues.ContainsValue(value)
                        {
                            //var mapValue = col.MappingValues.FirstOrDefault(f => f.Value.ToString() == value.ToString());
                            //((IDictionary<string, object>)obj)[propertyInfo.Name] = mapValue.Key;
                        }
                        else
                        {
                            ((IDictionary<string, object>)obj)[propertyInfo.Name] = value;
                        }
                    }
                    else if (propertyInfo.PropertyType.GetCSharpTypeName() == "Nullable<Boolean>")
                    {
                        var col = ExporterHeaderList.First(a => a.PropertyName == propertyInfo.Name);
                        var value = Convert.ToBoolean(type.GetProperty(propertyInfo.Name)?.GetValue(dataItem));

                        if (false)//col.MappingValues.Count > 0 && col.MappingValues.ContainsValue(value.ToString())
                        {
                            //var mapValue = col.MappingValues.FirstOrDefault(f => f.Value.to == value.ToString());
                            //((IDictionary<string, object>)obj)[propertyInfo.Name] = mapValue.Key;
                        }
                        else
                        {
                            ((IDictionary<string, object>)obj)[propertyInfo.Name] = value;
                        }
                    }
                    else
                    {
                        ((IDictionary<string, object>)obj)[propertyInfo.Name] = type.GetProperty(propertyInfo.Name)?.GetValue(dataItem)?.ToString();
                    }
                }

                yield return obj;
            }
            //list.Add(obj);
            // return list;
        }







        /// <summary>
        /// 添加导出数据
        /// </summary>
        /// <param name="dataItems"></param>
        /// <param name="excelRange"></param>
        protected void AddDataItems(IEnumerable<T> dataItems, ExcelRangeBase excelRange = null)
        {
            if (excelRange == null)
                excelRange = CurrentExcelWorksheet.Cells["A1"];

            if (dataItems == null || !dataItems.Any())
            {
                return;
            }

            if (ExcelExporterSettings.ExcelOutputType == ExcelOutputTypes.DataTable)
            {
                if (IsExpandoObjectType)
                    excelRange.LoadFromDictionaries((IEnumerable<IDictionary<string, object>>)dataItems, true, ExcelExporterSettings.TableStyle);
                else
                {
                    //如果TableStyle=None则Table不为null
                    var er = excelRange.LoadFromCollection(dataItems, true, ExcelExporterSettings.TableStyle);
                    CurrentExcelTable = CurrentExcelWorksheet.Tables.GetFromRange(er);
                }
            }
            else
            {
                excelRange.LoadFromCollection(dataItems, true, ExcelExporterSettings.TableStyle);
            }
        }
        /// <summary>
        /// 添加导出数据
        /// </summary>
        /// <param name="dataItems"></param>
        /// <param name="excelRange"></param>
        protected void AddDataItems(IEnumerable<ExpandoObject> dataItems, ExcelRangeBase excelRange = null)
        {
            if (excelRange == null)
                excelRange = CurrentExcelWorksheet.Cells["A1"];

            if (dataItems == null || !dataItems.Any())
            {
                return;
            }

            if (ExcelExporterSettings.ExcelOutputType == ExcelOutputTypes.DataTable)
            {
                if (IsExpandoObjectType)
                    excelRange.LoadFromDictionaries(dataItems, true, ExcelExporterSettings.TableStyle);
                else
                {
                    //如果TableStyle=None则Table不为null
                    var er = excelRange.LoadFromDictionaries(dataItems, true, ExcelExporterSettings.TableStyle);
                    CurrentExcelTable = CurrentExcelWorksheet.Tables.GetFromRange(er);
                }
            }
            else
            {
                //if (IsExpandoObjectType)
                //  excelRange.LoadFromDictionaries(dataItems, true, ExcelExporterSettings.TableStyle);
                //else
                excelRange.LoadFromDictionaries(dataItems, true, ExcelExporterSettings.TableStyle);
                //CurrentExcelTable = CurrentExcelWorksheet.Tables.GetFromRange(er);
            }
        }
        /// <summary>
        /// 在数据达到设置值时禁用自适应列
        /// </summary>
        /// <param name="count"></param>
        private void DisableAutoFitWhenDataRowsIsLarge(int count)
        {
            //如果已经设置了AutoFitMaxRows并且当前数据超过此设置，则关闭自适应列的配置
            if (ExcelExporterSettings.AutoFitMaxRows != 0 && count > ExcelExporterSettings.AutoFitMaxRows)
            {
                ExcelExporterSettings.AutoFitAllColumn = false;
                foreach (var item in ExporterHeaderList)
                {
                    if (item.ExporterHeaderAttribute != null)
                        item.ExporterHeaderAttribute.IsAutoFit = false;
                }
            }
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
        ///设置x行开始追加内容
        /// </summary>
        private void SetSkipRows()
        {
            if (ExcelExporterSettings.HeaderRowIndex > 1)
            {
                CurrentExcelWorksheet.InsertRow(1, ExcelExporterSettings.HeaderRowIndex);
            }
        }
        /// <summary>
        /// 创建表头
        /// </summary>
        protected void AddHeader()
        {
            //NoneStyle的时候没创建Table
            //https://github.com/JanKallman/EPPlus/blob/4dacf27661b24d92e8ba3d03d51dd5468845e6c1/EPPlus/ExcelRangeBase.cs#L2013
            var isNoneStyle = ExcelExporterSettings.TableStyle == TableStyles.None;

            if (CurrentExcelTable == null && ExcelExporterSettings.ExcelOutputType == ExcelOutputTypes.DataTable && !isNoneStyle)
            {
                var cols = ExporterHeaderList.Count;
                var range = CurrentExcelWorksheet.Cells[1, 1, CurrentExcelWorksheet.Dimension?.End.Row ?? 10, cols];
                //https://github.com/dotnetcore/Magicodes.IE/issues/66
                CurrentExcelTable = CurrentExcelWorksheet.Tables.Add(range, $"Table{CurrentExcelWorksheet.Index}");
                CurrentExcelTable.ShowHeader = true;
                //Enum.TryParse(ExcelExporterSettings.TableStyle, out TableStyles outStyle);
                CurrentExcelTable.TableStyle = ExcelExporterSettings.TableStyle;
            }

            if (ExcelExporterSettings.AutoCenter)
            {
                //自己注释
                //CurrentExcelWorksheet.Cells[1, 1, CurrentExcelWorksheet.Dimension?.End.Row ?? 10, ExporterHeaderList.Count].Style.HorizontalAlignment = ExceHorizontalAlignment.Center;
            }

            foreach (var exporterHeaderDto in ExporterHeaderList)
            {
                var exporterHeaderAttribute = exporterHeaderDto.ExporterHeaderAttribute;
                if (exporterHeaderAttribute != null)
                {
                    var colCell = CurrentExcelWorksheet.Cells[1, exporterHeaderDto.Index];
                    colCell.Style.Font.Bold = exporterHeaderAttribute.IsBold;

                    if (CurrentExcelTable != null)
                    {
                        var col = CurrentExcelTable.Columns[exporterHeaderDto.Index - 1];
                        col.Name = exporterHeaderDto.DisplayName;
                    }
                    else
                    {
                        colCell.Value = exporterHeaderDto.DisplayName;
                    }


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

                if ((_exporterHeaderList == null || _exporterHeaderList.Count == 0) && !IsDynamicDatableExport &&
                    !IsExpandoObjectType) throw new ArgumentException("请定义表头！");
               
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
            if (dt != null)
            {
                for (var i = 0; i < dt.Columns.Count; i++)
                {
                    var item = new ExporterHeaderInfo
                    {
                        Index = i + 1,
                        PropertyName = dt.Columns[i].ColumnName,
                        ExporterHeaderAttribute = new ExporterHeaderAttribute(dt.Columns[i].ColumnName),
                        CsTypeName = dt.Columns[i].DataType.GetCSharpTypeName(),
                        DisplayName = dt.Columns[i].ColumnName
                    };
                    AddExportHeaderInfo(item);
                }
            }
            else if (IsExpandoObjectType)
            {
                var items = dataItems as IEnumerable<IDictionary<string, object>>;
                var keys = new List<string>(items.First().Keys);
                for (int i = 0; i < keys.Count; i++)
                {
                    var item = new ExporterHeaderInfo
                    {
                        Index = i + 1,
                        PropertyName = keys[i],
                        ExporterHeaderAttribute = new ExporterHeaderAttribute(keys[i]),
                        CsTypeName = keys[i].GetType().GetCSharpTypeName(),
                        DisplayName = keys[i]
                    };
                    AddExportHeaderInfo(item);
                }
            }
            else if (!IsDynamicDatableExport)
            {
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
                        ExportImageFieldAttribute = objProperties[i].GetAttribute<ExportImageFieldAttribute>(true)

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
                    //设置Ignore
                    //item.ExporterHeaderAttribute.IsIgnore =
                    //    (objProperties[i].GetAttribute<IEIgnoreAttribute>(true) == null)
                    //        ? item.ExporterHeaderAttribute.IsIgnore
                    //        : objProperties[i].GetAttribute<IEIgnoreAttribute>(true).IsExportIgnore;

                    //var itemMappingValues = item.MappingValues;
                    //objProperties[i].ValueMapping(ref itemMappingValues);
                    //var mappings = objProperties[i].GetAttributes<ValueMappingAttribute>().ToList();
                    //foreach (var mappingAttribute in mappings.Where(mappingAttribute =>
                    //    !item.MappingValues.ContainsKey(mappingAttribute.Value)))
                    //    item.MappingValues.Add(mappingAttribute.Value, mappingAttribute.Text);

                    ////如果存在自定义映射，则不会生成默认映射
                    //if (!mappings.Any())
                    //{
                    //    if (objProperties[i].PropertyType.IsEnum)
                    //    {
                    //        var propType = objProperties[i].PropertyType;
                    //        var isNullable = propType.IsNullable();
                    //        if (isNullable) propType = propType.GetNullableUnderlyingType();
                    //        var values = propType.GetEnumTextAndValues();

                    //        foreach (var value in values.Where(value => !item.MappingValues.ContainsKey(value.Key)))
                    //            item.MappingValues.Add(value.Value, value.Key);

                    //        if (isNullable)
                    //            if (!item.MappingValues.ContainsKey(string.Empty))
                    //                item.MappingValues.Add(string.Empty, null);

                    //    }
                    //}

                    AddExportHeaderInfo(item);
                }
            }
        }
        /// <summary>
        /// 添加列头并执行列头筛选器
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        protected virtual void AddExportHeaderInfo(ExporterHeaderInfo item)
        {
            //执行列头筛选器
            //if (ExporterHeaderFilter != null)
            //{
            //    item = ExporterHeaderFilter.Filter(item);
            //}

            _exporterHeaderList.Add(item);
        }
    }
}
