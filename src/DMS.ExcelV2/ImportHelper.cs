using DMS.ExcelV2.Result;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.ExcelV2
{
    public class ImportHelper<T> : IDisposable where T : class, new()
    {







        public Task<ImportResult<T>> Import(Stream stream)
        {
            ImportResult<T> importResult = new ImportResult<T>();

            using (stream)
            {
                using (var excelPackage = new ExcelPackage(stream))
                {
                    ParseHeader();
                }
            }


            return Task.FromResult(importResult);
        }

        /// <summary>
        /// 解析实体属性
        /// </summary>
        /// <returns></returns>
        protected bool ParseHeader()
        {
            var objProperties = typeof(T).GetProperties();
            foreach (var propertyInfo in objProperties)
            {

            }
            return true;
        }


        public void Dispose()
        {
            throw new NotImplementedException();
        }
    }
}
