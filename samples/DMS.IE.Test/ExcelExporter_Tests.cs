using DMS.Excel;
using DMS.Excel.Attributes.Export;
using DMS.IE.Test.Models.Export;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Threading.Tasks;
using Xunit;

namespace DMS.IE.Test
{
    public class ExcelExporter_Tests : TestBase
    {

        /// <summary>
        /// 最原始导入
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "最原始导入")]
        public async Task ExportLoadFromCollection_Test()
        {
            IExcelExporter exporter = new ExcelExporter();
            var filePath = GetTestFilePath($"{nameof(ExportLoadFromCollection_Test)}.xlsx");

            DeleteFile(filePath);

            var data = GenFu.GenFu.ListOf<ExportLoadFromCollection>(100);

            await exporter.ExportLoadFromCollection(filePath, data);
        }

        /// <summary>
        /// 数据导出
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "数据导出")]
        public async Task ExportTestDataWithAttrs_Test()
        {
            IExcelExporter exporter = new ExcelExporter();
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
            var exporter = new ExcelExporter();
            var filePath = GetTestFilePath($"{nameof(ExportTestDataWithAttrsGroup_Test)}.xlsx");
            DeleteFile(filePath);
            var data = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(100);
            var data1 = GenFu.GenFu.ListOf<ExportLoadFromCollection>(100);
            foreach (var item in data)
            {
                item.LongNo = 458752665;
                item.Text = "测试长度超出单元格的字符串";
            }

            var result = await exporter
                .Append(data, "导出结果11")
                .Append(data1, "导出结果11")
                .ExportAppendData(filePath);
        }

        [Fact(DisplayName = "")]
        public async void DynamicHeaderExport_Test1()
        {

        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "")]
        public async void DynamicHeaderExport_Test()
        {
            //List<MainData> data = new List<MainData>() {

            // new MainData(){  Num=1, BasicTitle="硬质景观"},
            // new MainData(){  Num=2, BasicTitle="绿化工程"},
            // new MainData(){  Num=3, BasicTitle="小品布置"},
            // new MainData(){  Num=4, BasicTitle="电气工程"},
            // new MainData(){  Num=5, BasicTitle="给排水工程"},
            // new MainData(){  Num=6, BasicTitle="雨污水工程"},
            // new MainData(){  Num=7, BasicTitle="水景"},
            // new MainData(){  Num=8, BasicTitle="海绵城市"},
            // new MainData(){  Num=9, BasicTitle="围墙"},
            // new MainData(){  Num=10, BasicTitle="大门"},
            // new MainData(){  Num=10, BasicTitle="措施项目"},
            // new MainData(){  Num=10, BasicTitle="大门"}
            //};

            //var exporter = new ExcelExporter();
            //var filePath = GetTestFilePath($"{nameof(DynamicHeaderExport_Test)}.xlsx");
            //DeleteFile(filePath);
            //var result = await exporter.Export(filePath, data);


            IList<Gogo> list = new List<Gogo>
              {
                  new Gogo
                  {
                      Name = "张三",
                      Age = 18,
                      Card = "41234567890",
                      CreateTime = DateTime.Now,
                  },
                   new Gogo
                  {
                      Name = "李四",
                      Age = 20,
                      Card = "4254645461",
                      CreateTime = DateTime.Now,
                  },
              };
            //导出表头和字段集合
            ExportColumnCollective ecc = new ExportColumnCollective();
            //导出字段集合
            ecc.ExportColumnList = new List<ExportColumn>
              {
                  new ExportColumn{Field = "Name"},
                  new ExportColumn{Field = "Card"},
                  new ExportColumn{Field = "Age"},
                  new ExportColumn{Field = "CreateTime"},
              };
            //导出表头集合
            ecc.HeaderExportColumnList = new List<List<ExportColumn>>
              {
      	         //使用list是为了后续可能有多表头合并列的需求，这里只需要单个表头所以一个list就ok了
                  new List<ExportColumn>
                  {
                      new ExportColumn{Title = "姓名"},
                      new ExportColumn{Title = "身份号"},
                      new ExportColumn{Title = "年龄"},
                      new ExportColumn{Title = "添加时间"}
                  },
                  new List<ExportColumn>
                  {
                      new ExportColumn{Title = "子标题A",ColSpan = 1},
                      new ExportColumn{Title = "子标题B",ColSpan = 1}
                  },
                  new List<ExportColumn>
                  {
                      new ExportColumn{Title = "子标题2",ColSpan = 2},
                      new ExportColumn{Title = "子标题2",ColSpan = 3}
                  },
              };
            byte[] result = Export<Gogo>(list, ecc, "测试导出", true);


            File.WriteAllBytes("d:\\2.xlsx", result);


        }



        /// <summary>
        /// 生成excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dtSource">数据源</param>
        /// <param name="columns">导出字段表头合集</param>
        /// <param name="title">标题(Sheet名)</param>
        /// <param name="showTitle">是否显示标题</param>
        /// <returns></returns>
        public static byte[] Export<T>(IList<T> dtSource, ExportColumnCollective columns, string title, bool showTitle = true)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet workSheet = package.Workbook.Worksheets.Add(title);

                int maxColumnCount = columns.ExportColumnList.Count;
                int curRowIndex = 0;

                //Excel标题
                if (showTitle == true)
                {
                    curRowIndex++;
                    workSheet.Cells[curRowIndex, 1, 1, maxColumnCount].Merge = true;
                    workSheet.Cells[curRowIndex, 1].Value = title;
                    var headerStyle = workSheet.Workbook.Styles.CreateNamedStyle("headerStyle");
                    headerStyle.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    headerStyle.Style.Font.Bold = true;
                    headerStyle.Style.Font.Size = 20;
                    workSheet.Cells[curRowIndex, 1].StyleName = "headerStyle";

                    curRowIndex++;
                    //导出时间
                    workSheet.Cells[curRowIndex, 1, 2, maxColumnCount].Merge = true;
                    workSheet.Cells[curRowIndex, 1].Value = "导出时间：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                    workSheet.Cells[curRowIndex, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }

                //数据表格标题(列名)
                for (int i = 0, rowCount = columns.HeaderExportColumnList.Count; i < rowCount; i++)
                {
                    curRowIndex++;
                    workSheet.Cells[curRowIndex, 1, curRowIndex, maxColumnCount].Style.Font.Bold = true;
                    var curColSpan = 1;
                    for (int j = 0, colCount = columns.HeaderExportColumnList[i].Count; j < colCount; j++)
                    {
                        var colColumn = columns.HeaderExportColumnList[i][j];
                        var colSpan = FindSpaceCol(workSheet, curRowIndex, curColSpan);
                        if (j == 0) curColSpan = colSpan;
                        var toColSpan = colSpan + colColumn.ColSpan;
                        var cell = workSheet.Cells[curRowIndex, colSpan, colColumn.RowSpan + curRowIndex, toColSpan];
                        cell.Merge = true;
                        cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[curRowIndex, colSpan].Value = colColumn.Title;
                        curColSpan += colColumn.ColSpan;
                    }
                }
                workSheet.View.FreezePanes(curRowIndex + 1, 1);//冻结标题行

                Type type = typeof(T);
                PropertyInfo[] propertyInfos = type.GetProperties();
                if (propertyInfos.Count() == 0 && dtSource.Count > 0) propertyInfos = dtSource[0].GetType().GetProperties();

                //数据行
                for (int i = 0, sourceCount = dtSource.Count(); i < sourceCount; i++)
                {
                    curRowIndex++;
                    for (var j = 0; j < maxColumnCount; j++)
                    {
                        var column = columns.ExportColumnList[j];
                        var cell = workSheet.Cells[curRowIndex, j + 1];
                        foreach (var propertyInfo in propertyInfos)
                        {
                            if (column.Field == propertyInfo.Name)
                            {
                                object value = propertyInfo.GetValue(dtSource[i]);
                                var pType = propertyInfo.PropertyType;
                                pType = pType.Name == "Nullable`1" ? Nullable.GetUnderlyingType(pType) : pType;
                                if (pType == typeof(DateTime))
                                {
                                    cell.Style.Numberformat.Format = "yyyy-MM-dd hh:mm";
                                    cell.Value = Convert.ToDateTime(value);
                                }
                                else if (pType == typeof(int))
                                {
                                    cell.Style.Numberformat.Format = "#0";
                                    cell.Value = Convert.ToInt32(value);
                                }
                                else if (pType == typeof(double) || pType == typeof(decimal))
                                {
                                    if (column.Precision != null) cell.Style.Numberformat.Format = "#,##0.00";//保留两位小数

                                    cell.Value = Convert.ToDouble(value);
                                }
                                else
                                {
                                    cell.Value = value == null ? "" : value.ToString();
                                }
                            }
                        }
                    }
                }
                workSheet.Cells[workSheet.Dimension.Address].Style.Font.Name = "宋体";
                workSheet.Cells[workSheet.Dimension.Address].AutoFitColumns();//自动填充
                for (var i = 1; i <= workSheet.Dimension.End.Column; i++) { workSheet.Column(i).Width = workSheet.Column(i).Width + 2; }//在填充的基础上再加2

                return package.GetAsByteArray();
            }
        }

        private static int FindSpaceCol(ExcelWorksheet workSheet, int row, int col)
        {
            if (workSheet.Cells[row, col].Merge)
            {
                return FindSpaceCol(workSheet, row, col + 1);
            }
            return col;
        }


    }

    public class Gogo
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string Card { get; set; }
        public DateTime CreateTime { get; set; }
    }

    //导出所需要映射的字段和表头集合
    public class ExportColumnCollective
    {
        /// <summary>
        /// 字段列集合
        /// </summary>
        public List<ExportColumn> ExportColumnList { get; set; }
        /// <summary>
        /// 表头或多表头集合
        /// </summary>
        public List<List<ExportColumn>> HeaderExportColumnList { get; set; }
    }
    //映射excel实体
    public class ExportColumn
    {

        /// <summary>
        /// 标题
        /// </summary>
        [JsonProperty("title")]
        public string Title { get; set; }
        /// <summary>
        /// 字段
        /// </summary>
        [JsonProperty("field")]
        public string Field { get; set; }
        /// <summary>
        /// 精度(只对double、decimal有效)
        /// </summary>
        [JsonProperty("precision")]
        public int? Precision { get; set; }
        /// <summary>
        /// 跨列
        /// </summary>
        [JsonProperty("colSpan")]
        public int ColSpan { get; set; }
        /// <summary>
        /// 跨行
        /// </summary>
        [JsonProperty("rowSpan")]
        public int RowSpan { get; set; }
    }


    [ExcelExporter(Name = "动态列", AutoCenter = true, AutoFitAllColumn = true, IsBold = true, TableStyle = TableStyles.None)]
    public class MainData
    {
        /// <summary>
        /// 序号
        /// </summary>
        [ExporterHeader(DisplayName = "序号")]
        public int Num { get; set; }
        /// <summary>
        /// 分部分项名称
        /// </summary>
        [ExporterHeader(DisplayName = "分部分项名称")]
        public string BasicTitle { get; set; }
    }

    public class MainDataParent
    {
        /// <summary>
        /// 
        /// </summary>
        [ExporterHeader(DisplayName = "")]
        public string Title { get; set; }
        public List<DataItem> dataItems { get; set; }
    }

    public class DataItem
    {
        /// <summary>
        /// 含税金额
        /// </summary>
        [ExporterHeader(DisplayName = "含税金额（元）")]
        public double TotalPrice { get; set; }
        /// <summary>
        /// 单价
        /// </summary>
        [ExporterHeader(DisplayName = "单方（元 / m2）")]
        public double UnitPrice { get; set; }
    }
}
