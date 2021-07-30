using DMS.Excel;
using DMS.Excel.Models;
using System;
using System.IO;

namespace DMS.Excel.Extensions
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


        /// <summary>
        /// 验证路径是否为空
        /// 验证后缀是否是Excel文件
        /// 验证文件路径是否存在
        /// </summary>
        /// <param name="fileName"></param>
        public static void CheckExcelFilePath(this string filePath)
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
    }
}
