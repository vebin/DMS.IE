using DMS.Excel.Extension;
using DMS.Excel.Models;
using System;
using System.Collections.Generic;
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
        /// <summary>
        /// 
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
