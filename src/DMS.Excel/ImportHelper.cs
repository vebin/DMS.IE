using DMS.Common.Extensions;
using DMS.Excel.Attributes;
using DMS.Excel.Attributes.Import;
using DMS.Excel.Extension;
using DMS.Excel.Models;
using DMS.Excel.Result;
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
        /// </summary>
        /// <param name="filePath"></param>
        public ImportHelper(string filePath = null)
        {
            FilePath = filePath;
        }

        /// <summary>
        /// </summary>
        /// <param name="stream"></param>
        public ImportHelper(Stream stream)
        {
            Stream = stream;
        }
        /// <summary>
        /// 
        /// </summary>
        private Dictionary<string, dynamic> dicMergePreValues = new Dictionary<string, dynamic>();
        /// <summary>
        /// 导入文件路径
        /// </summary>
        protected string FilePath { get; set; }
        /// <summary>
        /// 文件流
        /// </summary>
        protected Stream Stream { get; set; }
        /// <summary>
        /// 导入结果
        /// </summary>
        internal ImportResult<T> ImportResult { get; set; }
        /// <summary>
        /// 列头定义
        /// </summary>
        protected List<ImporterHeaderInfo> ImporterHeaderInfos { get; set; }
        /// <summary>
        /// 
        /// </summary>
        private ExcelImporterAttribute _excelImporterAttribute;
        /// <summary>
        /// 获取头部属性值
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
                            HeaderRowIndex = importerAttribute.HeaderRowIndex, 
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

        public Task<ImportResult<T>> Import(string filePath = null)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            if (!string.IsNullOrWhiteSpace(filePath)) FilePath = filePath;
            ImportResult = new ImportResult<T>();

            if (Stream == null)
            {
                CheckExcelFilePath(FilePath);
                Stream = new FileStream(FilePath, FileMode.Open);
            }

            using (Stream)
            {
                using (var excelPackage = new ExcelPackage(Stream))
                {
                    //获取导入实体列定义
                    ParseHeader();
                    ParseTemplate(excelPackage);
                    ImportResult.ImporterHeaderInfos = ImporterHeaderInfos;
                    if (ImportResult.HasError) return Task.FromResult(ImportResult);

                    ParseData(excelPackage);

                }
            }

            return Task.FromResult(ImportResult);
        }

        /// <summary>
        /// 解析实体头部
        /// </summary>
        /// <returns></returns>
        protected virtual bool ParseHeader()
        {
            ImporterHeaderInfos = new List<ImporterHeaderInfo>();
            ImportResult.TemplateErrors = new List<TemplateErrorInfo>();
            var objProperties = typeof(T).GetProperties();
            if (objProperties.Length == 0)
            {
                ImportResult.TemplateErrors.Add(new TemplateErrorInfo
                {
                    Message = $"解析实体为空对象"
                });
                return false;
            }
            foreach (var propertyInfo in objProperties)
            {
                //如果不设置，则自动使用默认定义
                var importerHeaderAttribute = (propertyInfo.GetCustomAttributes(typeof(ImporterHeaderAttribute), true) as ImporterHeaderAttribute[])?.FirstOrDefault() ?? new ImporterHeaderAttribute
                {
                    //如果没有设置ImporterHeader，在检查是否有设置Display，都没有则为默认名称
                    Name = propertyInfo.GetDisplayName() ?? propertyInfo.Name,
                };

                var colHeader = new ImporterHeaderInfo
                {
                    IsRequired = propertyInfo.IsRequired(),
                    PropertyName = propertyInfo.Name,
                    Header = importerHeaderAttribute,
                    ImportImageFieldAttribute = propertyInfo.GetAttribute<ImportImageFieldAttribute>(true),
                    PropertyInfo = propertyInfo
                };
                ImporterHeaderInfos.Add(colHeader);
            }
            return true;
        }
        /// <summary>
        /// 解析模板并设置列索引
        /// </summary>
        /// <returns></returns>
        protected virtual void ParseTemplate(ExcelPackage excelPackage)
        {
            try
            {
                //根据名称获取Sheet，如果不存在则取第一个
                var worksheet = GetWorksheet(excelPackage, ExcelImporterSettings.SheetIndex);
                var excelHeaders = new Dictionary<string, int>();
                var endColumnCount = worksheet.Dimension.End.Column;
                for (var columnIndex = 1; columnIndex <= endColumnCount; columnIndex++)
                {
                    var header = worksheet.Cells[ExcelImporterSettings.HeaderRowIndex, columnIndex].Text;
                    excelHeaders.Add(header, columnIndex);
                }

                foreach (var item in ImporterHeaderInfos)
                {
                    //设置列索引
                    if (item.Header.ColumnIndex == 0)
                        item.Header.ColumnIndex = excelHeaders[item.Header.Name];
                }
            }
            catch (Exception ex)
            {
                ImportResult.TemplateErrors.Add(new TemplateErrorInfo
                {
                    Message = $"模板出现未知错误：{ex}"
                });
                throw new Exception($"模板出现未知错误：{ex.Message}", ex);
            }
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

            for (var rowIndex = ExcelImporterSettings.HeaderRowIndex + 1;
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

                            var cell = worksheet.Cells[rowIndex, col.Header.ColumnIndex];

                            try
                            {
                                //如果是合并行并且值不为NULL，则暂存值
                                if (cell.Merge && cell.Value == null && dicMergePreValues.ContainsKey(propertyInfo.Name))
                                {
                                    propertyInfo.SetValue(dataItem, dicMergePreValues[propertyInfo.Name]);
                                    continue;
                                }

                                var cellValue = cell.Value?.ToString();
                                switch (propertyInfo.PropertyType.GetCSharpTypeName())
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
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 不合法！");
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
                                                AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
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
                                                AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                                                break;
                                            }

                                            SetValue(cell, dataItem, propertyInfo, number);
                                        }
                                        break;

                                    case "Int32":
                                        {
                                            if (!int.TryParse(cellValue, out var number))
                                            {
                                                AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
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
                                                AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                                                break;
                                            }

                                            SetValue(cell, dataItem, propertyInfo, number);
                                        }
                                        break;

                                    case "Int16":
                                        {
                                            if (!short.TryParse(cellValue, out var number))
                                            {
                                                AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
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
                                                AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                                                break;
                                            }

                                            SetValue(cell, dataItem, propertyInfo, number);
                                        }
                                        break;

                                    case "Decimal":
                                        {
                                            if (!decimal.TryParse(cellValue, out var number))
                                            {
                                                AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
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
                                                AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                                                break;
                                            }

                                            SetValue(cell, dataItem, propertyInfo, number);
                                        }
                                        break;

                                    case "Double":
                                        {
                                            if (!double.TryParse(cellValue, out var number))
                                            {
                                                AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
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
                                                AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
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
                                                AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
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
                                                AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                                                break;
                                            }

                                            SetValue(cell, dataItem, propertyInfo, number);
                                        }
                                        break;

                                    case "DateTime":
                                        {
                                            if (cell.Value == null || cell.Text.IsNullOrEmpty())
                                            {
                                                AddRowDataError(rowIndex, col, $"值 {cell.Value} 无效，请填写正确的日期时间格式！");
                                                break;
                                            }
                                            try
                                            {
                                                var date = cell.GetValue<DateTime>();
                                                SetValue(cell, dataItem, propertyInfo, date);
                                            }
                                            catch (Exception)
                                            {
                                                AddRowDataError(rowIndex, col, $"值 {cell.Value} 无效，请填写正确的日期时间格式！");
                                                break;
                                            }
                                        }
                                        break;

                                    case "DateTimeOffset":
                                        {
                                            if (!DateTimeOffset.TryParse(cell.Text, out var date))
                                            {
                                                AddRowDataError(rowIndex, col, $"值 {cell.Text} 无效，请填写正确的日期时间格式！");
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
                                                AddRowDataError(rowIndex, col, $"值 {cell.Text} 无效，请填写正确的日期时间格式！");
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
                                                AddRowDataError(rowIndex, col, $"值 {cell.Text} 无效，请填写正确的日期时间格式！");
                                                break;
                                            }

                                            SetValue(cell, dataItem, propertyInfo, date);
                                        }
                                        break;

                                    case "Guid":
                                        {
                                            if (!Guid.TryParse(cellValue, out var guid))
                                            {
                                                AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的Guid格式！");
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
                                                AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的Guid格式！");
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
                                AddRowDataError(rowIndex, col, ex.Message);
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
        /// 验证路径是否为空
        /// 验证后缀是否是Excel文件
        /// 验证文件路径是否存在
        /// </summary>
        /// <param name="fileName"></param>
        public static void CheckExcelFilePath(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentNullException(Resource.FileNameShouldNotBeEmpty, nameof(filePath));
            }
            if (!Path.GetExtension(filePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException(Resource.ExportingIsOnlySupportedXLSX, nameof(filePath));
            }
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("导入文件不存在!");
            }
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
        ///     添加数据行错误
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="importerHeaderInfo"></param>
        /// <param name="errorMessage"></param>
        protected virtual void AddRowDataError(int rowIndex, ImporterHeaderInfo importerHeaderInfo, string errorMessage = "数据格式无效！")
        {

            //if (ImportResult.RowErrors == null) ImportResult.RowErrors = new List<DataRowErrorInfo>();

            //var dataRowError = ImportResult.RowErrors.FirstOrDefault(p => p.RowIndex == rowIndex);
            //if (dataRowError == null)
            //{
            //    dataRowError = new DataRowErrorInfo
            //    {
            //        RowIndex = rowIndex,
            //    };
            //    ImportResult.RowErrors.Add(dataRowError);
            //}
            //dataRowError.FieldErrors.Add(importerHeaderInfo.Header.Name, errorMessage);
        }
        public void Dispose()
        {
        }
    }
}
