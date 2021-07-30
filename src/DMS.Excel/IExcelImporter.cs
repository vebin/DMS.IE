using DMS.Excel.Result;
using System.IO;
using System.Threading.Tasks;

namespace DMS.Excel
{
    public interface IExcelImporter
    {
        /// <summary>
        /// 导入模型验证数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <returns></returns>
        Task<ImportResult<T>> Import<T>(string filePath) where T : class, new();

        /// <summary>
        /// 导入模型验证数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream">文件流</param>
        /// <returns></returns>
        Task<ImportResult<T>> Import<T>(Stream stream) where T : class, new();
    }
}
