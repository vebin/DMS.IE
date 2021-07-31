using DMS.Excel.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Threading.Tasks;

namespace DMS.Excel
{
    /// <summary>
    /// 导出
    /// </summary>
    public interface IExcelExporter
    {
        /// <summary>
        /// 最原始导入
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <param name="dataItems"></param>
        /// <returns></returns>
        Task ExportLoadFromCollection<T>(string fileName, ICollection<T> dataItems) where T : class, new();
        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="fileName">文件名称</param>
        /// <param name="dataItems">数据</param>
        /// <returns>文件</returns>
        Task<ExportFileInfo> Export<T>(string fileName, ICollection<T> dataItems) where T : class, new();

        /// <summary>
        /// 追加集合
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataItems"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        ExcelExporter Append<T>(ICollection<T> dataItems, string sheetName = null) where T : class, new();
        /// <summary>
        /// 追加集合后导出
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        Task<ExportFileInfo> ExportAppendData(string fileName);
    }
}
