using DMS.Excel.Attributes.Export;
using DMS.Excel.Models;
using DMSN.Common.Extensions;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace DMS.Excel
{
    public class ExportHelper<T> : IDisposable where T : class, new()
    {
        private List<ExporterHeaderInfo> _exporterHeaderInfoList;
        private ExcelExporterAttribute _excelExporterAttribute;

        private ExcelWorksheet _excelWorksheet;
        private ExcelPackage _excelPackage;
        private string _sheetName;
        /// <summary>
        /// 
        /// </summary>
        public ExportHelper(string sheetName = null)
        {
            _sheetName = sheetName;
        }
        public ExportHelper(ExcelPackage existExcelPackage, string sheetName = null)
        {
            if (existExcelPackage != null)
            {
                this._excelPackage = existExcelPackage;
            }

            _sheetName = sheetName;
        }
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
                                AllFontSize = exporterAttribute.AllFontSize,
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
                    .OrderBy(p => p.GetAttribute<ExporterHeaderAttribute>()?.ColumnIndex ?? 10000)
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
                    AddExcelWorksheet(_sheetName);
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
            SetHeader();
            SetColumn();
            SheetIndex++;

            //int columns = CurrentExcelWorksheet.Dimension.Columns;
            //CurrentExcelWorksheet.Cells[1, columns + 1].Value = "自定义";
            //CurrentExcelWorksheet.Cells[2, columns + 1].Value = "自定义值 ";
            //CurrentExcelWorksheet.InsertRow(1, 1);

            //CurrentExcelWorksheet.Cells[1, 1, 1, 2].Merge = true;
            //CurrentExcelWorksheet.Cells[1, 1, 1, 2].Value = "合并 ";
            //CurrentExcelWorksheet.Cells[1, 1, 1, 2].Style.Font.Bold = true;
            //CurrentExcelWorksheet.Cells[1, 1, 1, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //CurrentExcelWorksheet.DeleteColumn(exporterHeaderDto.Index - deletedCount);
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
        protected void SetHeader()
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
            //全局字体大小
            bool HeaderFontSizeFlag = false;
            if (ExcelExporterSettings.AllFontSize > 0)
            {
                CurrentExcelWorksheet.Cells.Style.Font.Size = ExcelExporterSettings.AllFontSize;
            }
            else
            {
                if (ExcelExporterSettings.HeaderFontSize > 0)
                {
                    HeaderFontSizeFlag = true;

                }
                if (ExcelExporterSettings.FontSize > 0)
                {
                    //正文字体大小
                    //从头部行计算，正文大小
                }
            }

            foreach (var exporterHeader in ExporterHeaderInfoList)
            {
                var colCell = CurrentExcelWorksheet.Cells[1, exporterHeader.Index];
                colCell.Value = exporterHeader.DisplayName;
                if (HeaderFontSizeFlag)
                {
                    //全局优先，局部设置头部字体大小
                    colCell.Style.Font.Size = ExcelExporterSettings.HeaderFontSize;
                }

                var exporterHeaderAttribute = exporterHeader.ExporterHeaderAttribute;
                if (exporterHeaderAttribute != null)
                {
                    //colCell.Style.Font.Bold = exporterHeaderAttribute.IsBold;//当前字段加粗
                }
            }

        }
        /// <summary>
        /// 添加列的样式
        /// </summary>
        protected void SetColumn()
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


        /// <summary>
        /// 添加动态表头
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="headerTexts"></param>
        /// <param name="headerTextsDictionary"></param>
        public static void AddHeader(ExcelWorksheet sheet, string[] headerTexts, string[] headerTextsDictionary)
        {
            for (var i = 0; i < headerTextsDictionary.Length; i++)
            {
                AddHeader(sheet, i + 1 + headerTexts.Length, headerTextsDictionary[i]);
            }
        }
        /// <summary>
        /// 添加表头
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnIndex"></param>
        /// <param name="headerText"></param>
        public static void AddHeader(ExcelWorksheet sheet, int columnIndex, string headerText)
        {
            sheet.Cells[1, columnIndex].Value = headerText;
            sheet.Cells[1, columnIndex].Style.Font.Bold = true;
        }


        /// <summary>
        /// 添加动态数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startRowIndex"></param>
        /// <param name="items"></param>
        /// <param name="propertySelectors"></param>
        /// <param name="dictionaryKeys"></param>

        //public static void AddObjects(ExcelWorksheet sheet, int startRowIndex, IList<Student> items, Func<Student, object>[] propertySelectors, List<string> dictionaryKeys)
        //{
        //    for (var i = 0; i < items.Count; i++)
        //    {
        //        for (var j = 0; j < dictionaryKeys.Count; j++)
        //        {
        //            sheet.Cells[i + startRowIndex, j + 1 + propertySelectors.Length].Value = items[i].Dictionarys[dictionaryKeys[j]];
        //        }
        //    }

        //}

        public void Dispose()
        {
            _excelPackage.Dispose();
        }

    }
}
