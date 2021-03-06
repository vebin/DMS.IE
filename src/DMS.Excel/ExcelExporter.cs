using DMS.Excel.Extensions;
using DMS.Excel.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace DMS.Excel
{
    /// <summary>
    /// 
    /// </summary>
    public class ExcelExporter : IExcelExporter
    {

        private ExcelPackage _excelPackage;
        public ExcelExporter()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }
        #region 最原始导入
        /// <summary>
        /// 最原始导入
        /// 没有样式
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <param name="dataItems"></param>
        public async Task ExportLoadFromCollection<T>(string fileName, ICollection<T> dataItems) where T : class, new()
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("结果导出");
                worksheet.Cells.LoadFromCollection(dataItems, true, TableStyles.None);
                //worksheet.Cells[2,1].Style.Font.Bold = true;
                //worksheet.Cells[2, 1, 10, 5].Style.Font.Bold = true;//范围行与例加粗

                //worksheet.Cells.Style.Font.Bold = true;
                //worksheet.Cells.Style.Font.Size = 14;
                worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; //水平居中
                worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;     //垂直居中

                //worksheet.Cells.Style.Font.Name = "微软雅黑";
                //worksheet.Cells.Style.ShrinkToFit = true;//单元格自动适应大小
                //worksheet.Column(4).AutoFit();

                await package.SaveAsAsync(new FileStream(fileName, FileMode.Create));
            };
        }
        #endregion

        /// <summary>
        /// 导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <param name="dataItems"></param>
        /// <returns></returns>
        public async Task<ExportFileInfo> Export<T>(string fileName, ICollection<T> dataItems) where T : class, new()
        {
            var bytes = await ExportAsByteArray(dataItems);
            return bytes.ToExcelExportFileInfo(fileName);
        }
        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="dataItems">数据</param>
        /// <returns>文件二进制数组</returns>
        public Task<byte[]> ExportAsByteArray<T>(ICollection<T> dataItems) where T : class, new()
        {
            var helper = new ExportHelper<T>();

            if (helper.ExcelExporterSettings.MaxRowNumberOnASheet > 0 &&
                dataItems.Count > helper.ExcelExporterSettings.MaxRowNumberOnASheet)
            {
                using (helper.CurrentExcelPackage)
                {
                    var sheetCount = (int)(dataItems.Count / helper.ExcelExporterSettings.MaxRowNumberOnASheet) +
                                     ((dataItems.Count % helper.ExcelExporterSettings.MaxRowNumberOnASheet) > 0
                                         ? 1
                                         : 0);
                    for (int i = 0; i < sheetCount; i++)
                    {
                        var sheetDataItems = dataItems.Skip(i * helper.ExcelExporterSettings.MaxRowNumberOnASheet)
                            .Take(helper.ExcelExporterSettings.MaxRowNumberOnASheet).ToList();
                        helper.AddExcelWorksheet();
                        helper.Export(sheetDataItems);
                    }

                    return Task.FromResult(helper.CurrentExcelPackage.GetAsByteArray());
                }
            }
            else
            {
                using (var ep = helper.Export(dataItems))
                {
                    return Task.FromResult(ep.GetAsByteArray());
                }
            }

        }






        /// <summary>
        /// 追加集合
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataItems"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public ExcelExporter Append<T>(ICollection<T> dataItems, string sheetName = null) where T : class, new()
        {
            var helper = this._excelPackage == null ? new ExportHelper<T>(sheetName) : new ExportHelper<T>(_excelPackage, sheetName);
            //var name = helper.ExcelExporterSettings?.Name;
            //if (this._excelPackage?.Workbook.Worksheets.Any(x => x.Name == name) ?? false)
            //{
            //    throw new ArgumentNullException($"A sheet with the name already exists:{name}");
            //}
            this._excelPackage = helper.Export(dataItems);
            return this;

        }
        /// <summary>
        /// 追加集合后导出
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public async Task<ExportFileInfo> ExportAppendData(string fileName)
        {
           
            using (this._excelPackage)
            {
                fileName.CheckExcelFileName();
                var bytes = await _excelPackage.GetAsByteArrayAsync();
                ExportFileInfo exportFileInfo = bytes.ToExcelExportFileInfo(fileName);
                return exportFileInfo;
            }
        }


    }
}
