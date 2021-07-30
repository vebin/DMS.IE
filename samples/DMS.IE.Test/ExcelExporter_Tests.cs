using DMS.Excel;
using DMS.IE.Test.Models.Export;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Threading.Tasks;
using Xunit;

namespace DMS.IE.Test
{
    public class ExcelExporter_Tests : TestBase
    {
        IExcelExporter exporter = new ExcelExporter();
        /// <summary>
        /// 最原始导入
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "最原始导入")]
        public async Task ExportLoadFromCollection_Test()
        {

            var filePath = GetTestFilePath($"{nameof(ExportLoadFromCollection_Test)}.xlsx");

            DeleteFile(filePath);

            var data = GenFu.GenFu.ListOf<ExportLoadFromCollection>(100);

            await exporter.ExportLoadFromCollection(filePath, data);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "DTO特性导出（测试格式化以及列头索引）")]
        public async Task ExportTestDataWithAttrs_Test()
        {
            var filePath = GetTestFilePath($"{nameof(ExportTestDataWithAttrs_Test)}.xlsx");
            DeleteFile(filePath);
            var data = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(100);
            foreach (var item in data)
            {
                item.LongNo = 458752665;
                item.Text = "测试长度超出单元格的字符串";
            }
            var result = await exporter.Export(filePath, data);
        }

        /// <summary>
        /// 不同的集合生成多个sheet
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "DTO特性导出（在同一个excel生成多个sheet）")]
        public async Task ExportTestDataWithAttrsGroup_Test()
        {
            var filePath = GetTestFilePath($"{nameof(ExportTestDataWithAttrsGroup_Test)}.xlsx");
            DeleteFile(filePath);
            var data = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(100);
            var data1 = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(100);
            foreach (var item in data)
            {
                item.LongNo = 458752665;
                item.Text = "测试长度超出单元格的字符串";
            }

            List<List<ExportTestDataWithAttrs>> datas = new List<List<ExportTestDataWithAttrs>>();
            datas.Add(data);
            datas.Add(data1);
            List<string> sheetNames = new List<string>() {
                "导出结果11",
                "导出结果22"
            };
            var result = await exporter.Export(filePath, datas, sheetNames);
        }



        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "DTO特性导出（测试格式化以及列头索引）")]
        public void MultilineHeaderExport_Test()
        {
            //List<ExportTestMultilineHeader> results = new List<ExportTestMultilineHeader>();

            //ExportTestMultilineHeader multilineHeader = new ExportTestMultilineHeader()
            //{

            //    Text = "1",
            //    Text2 = "2",
            //    Text3 = "3",
            //    companInfo = new CompanInfo()
            //    {
            //        Compan = "A",
            //        parents = new List<CompanParent>()
            //        {
            //            new CompanParent()
            //            {
            //                 Name="子1"
            //            },
            //            new CompanParent()
            //            {
            //                 Name="子2"
            //            }
            //        }
            //    },
            //};
            //results.Add(multilineHeader);

            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //using var package = new ExcelPackage();
            //var worksheet = package.Workbook.Worksheets.Add("结果导出");
            //worksheet.Cells.LoadFromCollection(results, true);
            //package.SaveAs(new FileStream("output.xlsx", FileMode.Create));

        }


    }
}
