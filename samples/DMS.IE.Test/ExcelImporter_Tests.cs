using DMS.Excel;
using DMS.IE.Test.Models.Import;
using Shouldly;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Xunit;

namespace DMS.IE.Test
{
    public class ExcelImporter_Tests
    {
        public IExcelImporter Importer = new ExcelImporter();

        /// <summary>
        /// 产品信息导入
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "产品信息导入")]
        public async Task ImportProductDto_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "产品导入模板.xlsx");
            var result = await Importer.Import<ImportProductDto>(filePath);
            result.ShouldNotBeNull();
        }

        /// <summary>
        /// 合并行数据导入
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "合并行数据导入")]
        public async Task ImportMergeRowsDto_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "合并行.xlsx");
            var import = await Importer.Import<ImportMergeRowsDto>(filePath);
            import.ShouldNotBeNull();
        }

        #region 图片测试

        /// <summary>
        /// 未测试完
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "导入图片测试")]
        public async Task ImportPicture_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "图片导入模板.xlsx");
            var import = await Importer.Import<ImportPictureDto>(filePath);
            import.ShouldNotBeNull();
           
        }
        #endregion
    }
}
