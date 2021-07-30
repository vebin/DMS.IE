using DMS.Excel.Extension;
using DMS.Excel.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.Excel
{
    /// <summary>
    /// 
    /// </summary>
    public class ExcelExporter : IExcelExporter
    {
        public ExcelExporter()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

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
            var helper = new ExportHelperV2<T>();
            if (helper.ExcelExporterSettings.MaxRowNumberOnASheet > 0 && dataItems.Count > helper.ExcelExporterSettings.MaxRowNumberOnASheet)
            {
                //using (helper.CurrentExcelPackage)
                //{
                //    var sheetCount = (int)(dataItems.Count / helper.ExporterSettings.MaxRowNumberOnASheet) +
                //                     ((dataItems.Count % helper.ExporterSettings.MaxRowNumberOnASheet) > 0
                //                         ? 1
                //                         : 0);
                //    for (int i = 0; i < sheetCount; i++)
                //    {
                //        var sheetDataItems = dataItems.Skip(i * helper.ExporterSettings.MaxRowNumberOnASheet)
                //            .Take(helper.ExporterSettings.MaxRowNumberOnASheet).ToList();
                //        helper.AddExcelWorksheet();
                //        helper.Export(sheetDataItems);
                //    }

                //    return Task.FromResult(helper.CurrentExcelPackage.GetAsByteArray());
                //}
                return null;
            }
            else
            {
                using (var ep = helper.Export(dataItems))
                {
                    return Task.FromResult(ep.GetAsByteArray());
                }
            }
        }


    }
}
