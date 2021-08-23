using DMS.Excel;
using DMS.Excel.Attributes.Export;
using DMS.IE.Test;
using Newtonsoft.Json;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Reflection;
using Xunit;

namespace DMS.IE.Test_1
{
    public class ExcelExporter_Tests_1 : TestBase
    {

        [Fact(DisplayName = "")]
        public async void DynamicHeaderExport_Test1()
        {
            ////导出表头和字段集合
            //ExportColumnCollective ecc = new ExportColumnCollective();
            ////导出表头集合
            //ecc.HeaderExportColumnList = new List<List<ExportColumn>>
            //  {
            //    //使用list是为了后续可能有多表头合并列的需求，这里只需要单个表头所以一个list就ok了
            //      new List<ExportColumn>
            //      {
            //          new ExportColumn{Title = "姓名"},
            //          new ExportColumn{Title = "身份号"},
            //          new ExportColumn{Title = "年龄"},
            //          new ExportColumn{Title = "添加时间"}
            //      },
            //      new List<ExportColumn>
            //      {
            //          new ExportColumn{Title = "子标题A",ColSpan = 1},
            //          new ExportColumn{Title = "子标题B",ColSpan = 1}
            //      },

            //  };

            //    IDictionary<string, object> result = new ExpandoObject();
            //    result.Add("", "");
            //    var expandoObjectList = new List<ExpandoObject>()
            //    {
            //        new ExpandoObject(){ new { a=1} },
            //         ((IDictionary<string, object>)shapedObj).Add("序号", "一");
            //};

            var filePath = GetTestFilePath($"{nameof(DynamicHeaderExport_Test1)}.xlsx");

            DeleteFile(filePath);

            var exporter = new ExcelExporter();
            var expandoObjectList = new List<ExpandoObject>();

            List<TitleItem> titleItems = new List<TitleItem>() {
                new TitleItem(){ Num="一", Title="小四合院落架大修工程" },
                new TitleItem(){ Num="1", Title="拆除工程" },
                new TitleItem(){ Num="2", Title="仿古建筑工程" },
                new TitleItem(){ Num="3", Title="庭院工程" },
                new TitleItem(){ Num="4", Title="绿化工程" },
                new TitleItem(){ Num="5", Title="强电工程" },
                new TitleItem(){ Num="6", Title="弱电工程" },
                new TitleItem(){ Num="7", Title="给排水工程" },
            };

         

            List<List<SubTotalItem>> dyTotalItems = new List<List<SubTotalItem>>() {
                new List<SubTotalItem>()
                {
                    new SubTotalItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, Remark=""},
                    new SubTotalItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, Remark=""},
                    new SubTotalItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, Remark=""},
                    new SubTotalItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, Remark=""},
                    new SubTotalItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, Remark=""},
                    new SubTotalItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, Remark=""},
                    new SubTotalItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, Remark=""},
                    new SubTotalItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, Remark=""},
                 },
                new  List<SubTotalItem> ()
                {
                    new SubTotalItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, },
                    new SubTotalItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0,},
                    new SubTotalItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0,},
                    new SubTotalItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0,},
                    new SubTotalItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, },
                    new SubTotalItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0,},
                    new SubTotalItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0,},
                    new SubTotalItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, },
                },
    
            };
            //foreach (var item in titleItems)
            //{
            //    IDictionary<string, object> valuePairs = new ExpandoObject();
            //    valuePairs.Add("序号", item.Num);
            //    valuePairs.Add("项目名称", item.Title);
            //    expandoObjectList.Add((ExpandoObject)valuePairs);
            //}


            //var result = await exporter.ExportAsByteArray<ExpandoObject>(expandoObjectList);
            //File.WriteAllBytes(filePath, result);









            ExportColumnResult dataResult = new ExportColumnResult()
            {
                columns = new List<ExportColumn>()
                {
                    new ExportColumn() { Title = "招标清单部分不含税金额（元）" },
                    new ExportColumn() { Title = "不含税单方指标（元/景观总面积）" },
                    new ExportColumn() { Title = "税率" },
                    new ExportColumn() { Title = "税金（元）"  },
                    new ExportColumn() { Title = "含税总金额（元）"  },
                    new ExportColumn() { Title = "含税单方指标（元/景观总面积）" },
                    new ExportColumn() { Title = "备注"  },
                },
                items = new List<SubTotalItem>() {
                    new SubTotalItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, Remark=""},
                    new SubTotalItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, Remark=""},
                    new SubTotalItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, Remark=""},
                    new SubTotalItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, Remark=""},
                    new SubTotalItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, Remark=""},
                    new SubTotalItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, Remark=""},
                    new SubTotalItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, Remark=""},
                    new SubTotalItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, Remark=""},
                 },

            };

            var propertyInfoList = new List<PropertyInfo>();
            foreach (var field in dataResult.columns)
            {
                //var propertyName = field.BindName.Trim();
                //var propertyInfo = typeof(SubTotalItem).GetProperty(propertyName, BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Instance);

                //if (propertyInfo == null)
                //{
                //    throw new Exception($"Property: {propertyName} 没有找到：{typeof(SubTotalItem)}");
                //}
                //propertyInfoList.Add(propertyInfo);
            }



            foreach (var item in dataResult.items)
            {
                var shapedObj = new ExpandoObject();

                foreach (var propertyInfo in propertyInfoList)
                {
                    var propertyValue = propertyInfo.GetValue(item);
                    ((IDictionary<string, object>)shapedObj).Add(propertyInfo.Name, propertyValue);
                }

                expandoObjectList.Add(shapedObj);
            }








            Type type = typeof(SubTotalItem);
            var pros = type.GetProperties();
            foreach (var item in pros)
            {
                var itemType = item.PropertyType;
                var ca = TypeDescriptor.GetAttributes(itemType).OfType<DisplayNameAttribute>().FirstOrDefault();

                TypeDescriptor.AddAttributes(itemType, new DisplayNameAttribute("naughty"));

                ca = TypeDescriptor.GetAttributes(itemType).OfType<DisplayNameAttribute>().FirstOrDefault();
            }

            Type type1 = typeof(SubTotalItem);
            var pros1 = type.GetProperties();
            foreach (var item1 in pros1)
            {
                var itemType = item1.PropertyType;
                var ca = TypeDescriptor.GetAttributes(itemType).OfType<DisplayNameAttribute>().FirstOrDefault();
            }
            //var ca = TypeDescriptor.GetAttributes(typeof(SubTotalItem))
            //   .OfType<CategoryAttribute>().FirstOrDefault();


            //TypeDescriptor.AddAttributes(typeof(SubTotalItem), new CategoryAttribute("naughty"));

            //ca = TypeDescriptor.GetAttributes(typeof(SubTotalItem))
            //      .OfType<CategoryAttribute>().FirstOrDefault();

        }
    }

    public class ExportColumnResult
    {
        public List<ExportColumn> columns { get; set; }
        public List<SubTotalItem> items { get; set; }
    }

    /// <summary>
    /// 投标报价汇总表子项
    /// </summary>
    public class SubTotalItem
    {
        /// <summary>
        /// 不含税总价
        /// </summary>
        [DisplayName("")]
        public double TotalPriceN { get; set; }
        /// <summary>
        /// 不含税单方
        /// </summary>
        [DisplayName("")]
        public double UPriceN { get; set; }
        /// <summary>
        /// 税率
        /// </summary>
        [DisplayName("")]
        public double TaxRate { get; set; }
        /// <summary>
        /// 税金
        /// </summary>
        [DisplayName("")]
        public double Taxation { get; set; }
        /// <summary>
        /// 含税总价
        /// </summary>
        [DisplayName("")]
        public double TotalPrice { get; set; }
        /// <summary>
        /// 含税单方
        /// </summary>
        [DisplayName("")]
        public double UPrice { get; set; }
        /// <summary>
        /// 备注
        /// </summary>
        [DisplayName("")]
        public string Remark { get; set; }
    }


    public class TitleItem
    {
        /// <summary>
        /// 序号
        /// </summary>
        public string Num { get; set; }
        /// <summary>
        /// 单项工程 单位工程名称
        /// </summary>
        public string Title { get; set; }
    }



    //导出所需要映射的字段和表头集合
    public class ExportColumnCollective
    {
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
        public string Title { get; set; }
        /// <summary>
        /// 跨列
        /// </summary>
        public int ColSpan { get; set; }
        /// <summary>
        /// 跨行
        /// </summary>
        public int RowSpan { get; set; }
    }

}
