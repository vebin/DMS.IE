using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.ExcelV2.Extensions
{
    public static class Extension
    {
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
