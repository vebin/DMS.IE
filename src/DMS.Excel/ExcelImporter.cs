using DMS.Excel.Extensions;
using DMS.Excel.Result;
using OfficeOpenXml;
using System.IO;
using System.Threading.Tasks;

namespace DMS.Excel
{
    /// <summary>
    /// 
    /// </summary>
    public class ExcelImporter : IExcelImporter
    {
        public ExcelImporter()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }
        public Task<ImportResult<T>> Import<T>(string filePath) where T : class, new()
        {
            filePath.CheckExcelFilePath();
            var stream = new FileStream(filePath, FileMode.Open);
            return Import<T>(stream);
        }

        public Task<ImportResult<T>> Import<T>(Stream stream) where T : class, new()
        {
            using (var importer = new ImportHelper<T>())
            {
                return importer.Import(stream);
            }
        }
    }
}
