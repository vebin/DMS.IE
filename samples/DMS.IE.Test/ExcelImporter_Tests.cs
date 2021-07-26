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
        /// 测试：
        /// 表头行位置设置
        /// 导入逻辑测试
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "产品信息导入")]
        public async Task Importer_Test()
        {
            //第一列乱序

            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "产品导入模板.xlsx");
            var result = await Importer.Import<ImportProductDto>(filePath);
            //result.ShouldNotBeNull();

            //result.HasError.ShouldBeTrue();
            //result.RowErrors.Count.ShouldBe(1);
            //result.Data.ShouldNotBeNull();
            //result.Data.Count.ShouldBeGreaterThanOrEqualTo(2);
            //foreach (var item in result.Data)
            //{
            //    if (item.Name != null && item.Name.Contains("空格测试")) item.Name.ShouldBe(item.Name.Trim());

            //    if (item.Code.Contains("不去除空格测试")) item.Code.ShouldContain(" ");
            //    //去除中间空格测试
            //    item.BarCode.ShouldBe("123123");
            //}

            ////可为空类型测试
            //result.Data.ElementAt(4).Weight.HasValue.ShouldBe(true);
            //result.Data.ElementAt(5).Weight.HasValue.ShouldBe(false);
            ////提取性别公式测试
            //result.Data.ElementAt(0).Sex.ShouldBe("女");
            ////获取当前日期以及日期类型测试  如果时间不对，请打开对应的Excel即可更新为当前时间，然后再运行此单元测试
            ////import.Data[0].FormulaTest.Date.ShouldBe(DateTime.Now.Date);
            ////数值测试
            //result.Data.ElementAt(0).DeclareValue.ShouldBe(123123);
            //result.Data.ElementAt(0).Name.ShouldBe("1212");
            //result.Data.ElementAt(0).BarCode.ShouldBe("123123");
            //result.Data.ElementAt(0).ProductIdTest1.ShouldBe(Guid.Parse("C2EE3694-959A-4A87-BC8C-4003F6576352"));
            //result.Data.ElementAt(0).ProductIdTest2.ShouldBe(Guid.Parse("C2EE3694-959A-4A87-BC8C-4003F6576357"));
            //result.Data.ElementAt(1).Name.ShouldBe(null);
            //result.Data.ElementAt(2).Name.ShouldBe("左侧空格测试");

            //result.ImporterHeaderInfos.ShouldNotBeNull();
            //result.ImporterHeaderInfos.Count.ShouldBe(17);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "合并行数据导入")]
        public async Task MergeRowsImportTest()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "合并行.xlsx");
            var import = await Importer.Import<MergeRowsImportDto>(filePath);
            //import.ShouldNotBeNull();
            //if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

            //if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));
            //import.HasError.ShouldBeFalse();
            //import.Data.ShouldNotBeNull();
            //import.Data.Select(p => p.Sex).Take(8).All(p => p == "男").ShouldBeTrue();
            //import.Data.Select(p => p.Sex).Skip(8).All(p => p == "女").ShouldBeTrue();

            //import.Data.Select(p => p.Name).Take(3).All(p => p == "张三").ShouldBeTrue();
            //import.Data.Select(p => p.Name).Skip(3).Take(4).All(p => p == "李四").ShouldBeTrue();
            //import.Data.Select(p => p.Name).Skip(7).Take(6).All(p => p == "王五").ShouldBeTrue();
            //import.Data.Count.ShouldBe(13);
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

            //import.ShouldNotBeNull();
            //import.HasError.ShouldBeFalse();
            //if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

            //if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));
            //foreach (var item in import.Data)
            //{
            //    File.Exists(item.Img).ShouldBeTrue();
            //    File.Exists(item.Img1).ShouldBeTrue();
            //}

            ////添加严格校验，防止图片位置错误等问题
            //var image1 = new FileInfo(Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Images", "1.Jpeg"));
            //var image2 = new FileInfo(Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Images", "3.Jpeg"));
            //var image3 = new FileInfo(Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Images", "4.Jpeg"));
            //var image4 = new FileInfo(Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Images", "2.Jpeg"));
            //new FileInfo(import.Data.ElementAt(0).Img1).Length.ShouldBe(image1.Length);
            //new FileInfo(import.Data.ElementAt(0).Img).Length.ShouldBe(image2.Length);
            //new FileInfo(import.Data.ElementAt(1).Img1).Length.ShouldBe(image3.Length);
            //new FileInfo(import.Data.ElementAt(1).Img).Length.ShouldBe(image4.Length);
            //new FileInfo(import.Data.ElementAt(2).Img).Length.ShouldBe(image1.Length);
            //new FileInfo(import.Data.ElementAt(2).Img1).Length.ShouldBe(image1.Length);
        }
        #endregion
    }
}
