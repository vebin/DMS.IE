using DMS.Excel.Attributes.Import;
using DMS.Excel.Models;
using DMS.Excel.Result;
using DMSN.Common.Extensions;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace DMS.Excel
{
    public class ImportHelper<T> : IDisposable where T : class, new()
    {
        /// <summary>
        /// 
        /// </summary>
        private Dictionary<string, dynamic> dicMergePreValues = new Dictionary<string, dynamic>();
        /// <summary>
        /// 导入结果
        /// </summary>
        internal ImportResult<T> ImportResult { get; set; }
        /// <summary>
        /// 列头集合
        /// </summary>
        protected List<ImporterHeaderInfo> ImporterHeaderInfos { get; set; }
        /// <summary>
        /// 
        /// </summary>
        private ExcelImporterAttribute _excelImporterAttribute;
        /// <summary>
        /// 获取自定义属性全局配置
        /// </summary>
        protected ExcelImporterAttribute ExcelImporterSettings
        {
            get
            {
                if (_excelImporterAttribute == null)
                {
                    var type = typeof(T);
                    _excelImporterAttribute = type.GetAttribute<ExcelImporterAttribute>(true);
                    if (_excelImporterAttribute != null)
                        return _excelImporterAttribute;

                    var importerAttribute = type.GetAttribute<ImporterAttribute>(true);
                    if (importerAttribute != null)
                    {
                        _excelImporterAttribute = new ExcelImporterAttribute()
                        {
                            DataRowStartIndex = importerAttribute.DataRowStartIndex,
                            DataRowEndIndex = importerAttribute.DataRowEndIndex,
                        };
                    }
                    else
                    {
                        _excelImporterAttribute = new ExcelImporterAttribute();
                    }
                    return _excelImporterAttribute;
                }

                return _excelImporterAttribute;

            }
            set => _excelImporterAttribute = value;
        }


        /// <summary>
        /// 导入数据
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public Task<ImportResult<T>> Import(Stream stream)
        {
            ImportResult = new ImportResult<T>();
            using (stream)
            {
                using (var excelPackage = new ExcelPackage(stream))
                {
                    ParseExcelTemplate(excelPackage);
                    ParseData(excelPackage);
                }
            }


            return Task.FromResult(ImportResult);
        }

        /// <summary>
        /// 解析Excel与DTO模型列的匹配
        /// </summary>
        /// <returns></returns>
        protected bool ParseExcelTemplate(ExcelPackage excelPackage)
        {
            //根据名称获取Sheet，如果不存在则取第一个
            try
            {
                var worksheet = GetWorksheet(excelPackage, ExcelImporterSettings.SheetIndex);
                var excelHeaders = new Dictionary<string, int>();
                var endColumnCount = worksheet.Dimension.End.Column;
                for (var columnIndex = 1; columnIndex <= endColumnCount; columnIndex++)
                {
                    //起始可能不为1，后期根据模板优化
                    var header = worksheet.Cells[ExcelImporterSettings.DataRowStartIndex, columnIndex].Text;
                    excelHeaders.Add(header, columnIndex);
                }
                ImporterHeaderInfos = new List<ImporterHeaderInfo>();
                var objProperties = typeof(T).GetProperties();
                foreach (var propertyInfo in objProperties)
                {
                    var importerHeaderAttribute = (propertyInfo.GetCustomAttributes(typeof(ImporterHeaderAttribute), true) as ImporterHeaderAttribute[])?.FirstOrDefault() ?? new ImporterHeaderAttribute
                    {
                        //Name = propertyInfo.GetDisplayName() ?? propertyInfo.Name,
                        Name = propertyInfo.Name,//如果不设置，则自动使用默认属性名称
                    };
                    if (importerHeaderAttribute.ColumnIndex == 0)
                        importerHeaderAttribute.ColumnIndex = excelHeaders[importerHeaderAttribute.Name];

                    var colHeader = new ImporterHeaderInfo
                    {
                        PropertyName = propertyInfo.Name,
                        HeaderAttribute = importerHeaderAttribute,
                        PropertyInfo = propertyInfo,
                    };
                    ImporterHeaderInfos.Add(colHeader);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"模板出现未知错误：{ex.Message}", ex);
            }
            return true;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="excelPackage"></param>
        protected virtual void ParseData(ExcelPackage excelPackage)
        {
            var worksheet = GetWorksheet(excelPackage, ExcelImporterSettings.SheetIndex);

            #region 检查导入最大条数限制
            #endregion

            ImportResult.Data = new List<T>();
            var propertyInfos = new List<PropertyInfo>(typeof(T).GetProperties());

            for (var rowIndex = ExcelImporterSettings.DataRowStartIndex + 1;
                rowIndex <= worksheet.Dimension.End.Row; rowIndex++)
            {
                //跳过空行
                if (worksheet.Cells[rowIndex, 1, rowIndex, worksheet.Dimension.End.Column].All(p => p.Text == string.Empty))
                {
                    //EmptyRows.Add(rowIndex);
                    continue;
                }
                else
                {

                    //模板与DTO匹配字段
                    var propertyInfoList = propertyInfos.Where(p => ImporterHeaderInfos.Any(p1 => p1.PropertyName == p.Name));
                    if (propertyInfoList != null && propertyInfoList.Count() > 0)
                    {
                        var dataItem = new T();
                        foreach (var propertyInfo in propertyInfoList)
                        {
                            var col = ImporterHeaderInfos.First(a => a.PropertyName == propertyInfo.Name);

                            var cell = worksheet.Cells[rowIndex, col.HeaderAttribute.ColumnIndex];

                            try
                            {
                                //如果是合并行并且值不为NULL，则暂存值
                                if (cell.Merge && cell.Value == null && dicMergePreValues.ContainsKey(propertyInfo.Name))
                                {
                                    propertyInfo.SetValue(dataItem, dicMergePreValues[propertyInfo.Name]);
                                    continue;
                                }

                                var cellValue = cell.Value?.ToString();
                                switch (propertyInfo.PropertyType.GetTypeName())
                                {
                                    #region 类型
                                    case "Boolean":
                                        SetValue(cell, dataItem, propertyInfo, false);
                                        break;

                                    case "Nullable<Boolean>":
                                        if (string.IsNullOrWhiteSpace(cellValue))
                                            SetValue(cell, dataItem, propertyInfo, null);
                                        else
                                        {
                                            //AddRowDataError(rowIndex, col, $"值 {cellValue} 不合法！");
                                        }
                                        break;

                                    case "String":
                                        SetValue(cell, dataItem, propertyInfo, cellValue);
                                        break;
                                    //long
                                    case "Int64":
                                        {
                                            if (!long.TryParse(cellValue, out var number))
                                            {
                                                //AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                                                break;
                                            }
                                            SetValue(cell, dataItem, propertyInfo, number);
                                        }
                                        break;

                                    case "Nullable<Int64>":
                                        {
                                            if (string.IsNullOrWhiteSpace(cellValue))
                                            {
                                                SetValue(cell, dataItem, propertyInfo, null);
                                                break;
                                            }

                                            if (!long.TryParse(cellValue, out var number))
                                            {
                                                //AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                                                break;
                                            }

                                            SetValue(cell, dataItem, propertyInfo, number);
                                        }
                                        break;

                                    case "Int32":
                                        {
                                            if (!int.TryParse(cellValue, out var number))
                                            {
                                                //AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                                                break;
                                            }

                                            SetValue(cell, dataItem, propertyInfo, number);
                                        }
                                        break;

                                    case "Nullable<Int32>":
                                        {
                                            if (string.IsNullOrWhiteSpace(cellValue))
                                            {
                                                SetValue(cell, dataItem, propertyInfo, null);
                                                break;
                                            }

                                            if (!int.TryParse(cellValue, out var number))
                                            {
                                                //AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                                                break;
                                            }

                                            SetValue(cell, dataItem, propertyInfo, number);
                                        }
                                        break;

                                    case "Int16":
                                        {
                                            if (!short.TryParse(cellValue, out var number))
                                            {
                                                //AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                                                break;
                                            }

                                            SetValue(cell, dataItem, propertyInfo, number);
                                        }
                                        break;

                                    case "Nullable<Int16>":
                                        {
                                            if (string.IsNullOrWhiteSpace(cellValue))
                                            {
                                                SetValue(cell, dataItem, propertyInfo, null);
                                                break;
                                            }

                                            if (!short.TryParse(cellValue, out var number))
                                            {
                                                //AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                                                break;
                                            }

                                            SetValue(cell, dataItem, propertyInfo, number);
                                        }
                                        break;

                                    case "Decimal":
                                        {
                                            if (!decimal.TryParse(cellValue, out var number))
                                            {
                                                //AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                                                break;
                                            }

                                            SetValue(cell, dataItem, propertyInfo, number);
                                        }
                                        break;

                                    case "Nullable<Decimal>":
                                        {
                                            if (string.IsNullOrWhiteSpace(cellValue))
                                            {
                                                SetValue(cell, dataItem, propertyInfo, null);
                                                break;
                                            }

                                            if (!decimal.TryParse(cellValue, out var number))
                                            {
                                                //AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                                                break;
                                            }

                                            SetValue(cell, dataItem, propertyInfo, number);
                                        }
                                        break;

                                    case "Double":
                                        {
                                            if (!double.TryParse(cellValue, out var number))
                                            {
                                                //AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                                                break;
                                            }

                                            SetValue(cell, dataItem, propertyInfo, number);
                                        }
                                        break;

                                    case "Nullable<Double>":
                                        {
                                            if (string.IsNullOrWhiteSpace(cellValue))
                                            {
                                                SetValue(cell, dataItem, propertyInfo, null);
                                                break;
                                            }

                                            if (!double.TryParse(cellValue, out var number))
                                            {
                                                //AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                                                break;
                                            }

                                            SetValue(cell, dataItem, propertyInfo, number);
                                        }
                                        break;
                                    //case "float":
                                    case "Single":
                                        {
                                            if (!float.TryParse(cellValue, out var number))
                                            {
                                                //AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                                                break;
                                            }

                                            SetValue(cell, dataItem, propertyInfo, number);
                                        }
                                        break;

                                    case "Nullable<Single>":
                                        {
                                            if (string.IsNullOrWhiteSpace(cellValue))
                                            {
                                                SetValue(cell, dataItem, propertyInfo, null);
                                                break;
                                            }

                                            if (!float.TryParse(cellValue, out var number))
                                            {
                                                //AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                                                break;
                                            }

                                            SetValue(cell, dataItem, propertyInfo, number);
                                        }
                                        break;

                                    case "DateTime":
                                        {
                                            if (cell.Value == null || cell.Text.IsNullOrEmpty())
                                            {
                                                //AddRowDataError(rowIndex, col, $"值 {cell.Value} 无效，请填写正确的日期时间格式！");
                                                break;
                                            }
                                            try
                                            {
                                                var date = cell.GetValue<DateTime>();
                                                SetValue(cell, dataItem, propertyInfo, date);
                                            }
                                            catch (Exception)
                                            {
                                                //AddRowDataError(rowIndex, col, $"值 {cell.Value} 无效，请填写正确的日期时间格式！");
                                                break;
                                            }
                                        }
                                        break;

                                    case "DateTimeOffset":
                                        {
                                            if (!DateTimeOffset.TryParse(cell.Text, out var date))
                                            {
                                                //AddRowDataError(rowIndex, col, $"值 {cell.Text} 无效，请填写正确的日期时间格式！");
                                                break;
                                            }

                                            SetValue(cell, dataItem, propertyInfo, date);
                                        }
                                        break;

                                    case "Nullable<DateTime>":
                                        {
                                            if (string.IsNullOrWhiteSpace(cell.Text))
                                            {
                                                SetValue(cell, dataItem, propertyInfo, null);
                                                break;
                                            }

                                            if (!DateTime.TryParse(cell.Text, out var date))
                                            {
                                                //AddRowDataError(rowIndex, col, $"值 {cell.Text} 无效，请填写正确的日期时间格式！");
                                                break;
                                            }

                                            SetValue(cell, dataItem, propertyInfo, date);
                                        }
                                        break;

                                    case "Nullable<DateTimeOffset>":
                                        {
                                            if (string.IsNullOrWhiteSpace(cell.Text))
                                            {
                                                SetValue(cell, dataItem, propertyInfo, null);
                                                break;
                                            }

                                            if (!DateTimeOffset.TryParse(cell.Text, out var date))
                                            {
                                                //AddRowDataError(rowIndex, col, $"值 {cell.Text} 无效，请填写正确的日期时间格式！");
                                                break;
                                            }

                                            SetValue(cell, dataItem, propertyInfo, date);
                                        }
                                        break;

                                    case "Guid":
                                        {
                                            if (!Guid.TryParse(cellValue, out var guid))
                                            {
                                                //AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的Guid格式！");
                                                break;
                                            }

                                            SetValue(cell, dataItem, propertyInfo, guid);
                                        }
                                        break;

                                    case "Nullable<Guid>":
                                        {
                                            if (string.IsNullOrWhiteSpace(cellValue))
                                            {
                                                SetValue(cell, dataItem, propertyInfo, null);
                                                break;
                                            }

                                            if (!Guid.TryParse(cellValue, out var guid))
                                            {
                                                //AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的Guid格式！");
                                                break;
                                            }

                                            SetValue(cell, dataItem, propertyInfo, guid);
                                        }
                                        break;

                                    default:
                                        SetValue(cell, dataItem, propertyInfo, cell.Value);
                                        break;
                                        #endregion
                                }
                            }
                            catch (Exception ex)
                            {
                                //AddRowDataError(rowIndex, col, ex.Message);
                            }
                        }

                        ImportResult.Data.Add(dataItem);
                    }
                }
            }




        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="dataItem"></param>
        /// <param name="propertyInfo"></param>
        /// <param name="value"></param>
        private void SetValue(ExcelRange cell, T dataItem, PropertyInfo propertyInfo, dynamic value)
        {
            if (cell.Merge && value != null)
            {
                dicMergePreValues[propertyInfo.Name] = value;
            }
            propertyInfo.SetValue(dataItem, value);
        }



        /// <summary>
        /// 
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <returns></returns>
        protected virtual ExcelWorkbook GetWorkbook(ExcelPackage excelPackage)
        {
            return excelPackage.Workbook;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="sheetIndex"></param>
        /// <returns></returns>
        protected virtual ExcelWorksheet GetWorksheet(ExcelPackage excelPackage, int sheetIndex)
        {
            return excelPackage.Workbook.Worksheets[sheetIndex] ?? excelPackage.Workbook.Worksheets[0];
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheetIndex"></param>
        /// <returns></returns>
        protected virtual ExcelWorksheet GetWorksheet(ExcelWorkbook workbook, int sheetIndex)
        {
            return workbook.Worksheets[sheetIndex] ?? workbook.Worksheets[0];
        }
        /// <summary>
        /// 
        /// </summary>
        public void Dispose()
        {
            ExcelImporterSettings = null;
            ImporterHeaderInfos = null;
            ImportResult = null;
            dicMergePreValues = null;
            GC.Collect();
        }
    }
}
