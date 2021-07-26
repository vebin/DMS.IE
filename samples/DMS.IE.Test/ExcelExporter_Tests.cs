using DMS.Excel;
using DMS.IE.Test.Models.Export;
using System.Net;
using System.Threading.Tasks;
using Xunit;

namespace DMS.IE.Test
{
    public class ExcelExporter_Tests: TestBase
    {
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "DTO特性导出（测试格式化以及列头索引）")]
        public async Task AttrsExport_Test()
        {
            IExcelExporter exporter = new ExcelExporter();

            var filePath = GetTestFilePath($"{nameof(AttrsExport_Test)}.xlsx");

            DeleteFile(filePath);

            var data = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(100);
            foreach (var item in data)
            {
                item.LongNo = 458752665;
                item.Text = "测试长度超出单元格的字符串";
            }

            var result = await exporter.Export(filePath, data);
        }

        public static HttpWebResponse aaaa()
        {
            return null;
        }
    }
}
