using DMS.Excel.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace DMS.Excel.Extension
{
    public static class Extension
    {
        /// <summary>
        /// 将Bytes导出为Excel文件
        /// </summary>
        /// <param name="bytes">字节数组</param>
        /// <param name="fileName">文件路径</param>
        /// <returns></returns>
        public static ExportFileInfo ToExcelExportFileInfo(this byte[] bytes, string fileName)
        {
            fileName.CheckExcelFileName();
            File.WriteAllBytes(fileName, bytes);

            var file = new ExportFileInfo(fileName,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            return file;
        }

        /// <summary>
        /// 检查文件名
        /// </summary>
        /// <param name="fileName"></param>
        public static void CheckExcelFileName(this string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentNullException(Resource.FileNameShouldNotBeEmpty, nameof(fileName));
            if (!Path.GetExtension(fileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException(Resource.ExportingIsOnlySupportedXLSX, nameof(fileName));
            }
        }
    }
}
