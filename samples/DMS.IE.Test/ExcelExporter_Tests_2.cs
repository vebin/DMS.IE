using DMS.Excel;
using DMS.Excel.Attributes.Export;
using DMS.Excel.Models;
using DMS.IE.Test;
using DMSN.Common.Extensions;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Reflection;
using Xunit;

namespace DMS.IE.Test_2
{
    public class ExcelExporter_Tests_2 : TestBase
    {
        public List<TitleItem> GetTitleItems()
        {
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

            return titleItems;
        }

        public List<BudgetItem> GetBudgetItems()
        {
            List<BudgetItem> budgetItems = new List<BudgetItem>() {
                 new BudgetItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, Remark="1"},
                    new BudgetItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, Remark="2"},
                    new BudgetItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, Remark="3"},
                    new BudgetItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, Remark="4"},
                    new BudgetItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, Remark="5"},
                    new BudgetItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, Remark="6"},
                    new BudgetItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, Remark="7"},
                    new BudgetItem(){ TotalPriceN=0,  UPriceN=1849296.57,TaxRate=0,Taxation=0,TotalPrice=1849296.57,UPrice=0, Remark="8"},
            };

            return budgetItems;
        }

        public List<List<SubTotalItem>> GetSubTotalItems()
        {
            List<List<SubTotalItem>> subTotalItems = new List<List<SubTotalItem>>() {
                new List<SubTotalItem>()
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
                }
            };

            return subTotalItems;
        }

        private List<ExporterHeaderInfo> _exporterHeaderInfoList;
        private int sumColumn = 0;
        /// <summary>
        /// 表头列表
        /// </summary>
        protected List<ExporterHeaderInfo> ExporterHeaderInfoList(Type type)
        {
            if (_exporterHeaderInfoList == null)
            {
                _exporterHeaderInfoList = new List<ExporterHeaderInfo>();
                var objProperties = type.GetProperties().OrderBy(p => p.GetAttribute<ExporterHeaderAttribute>()?.ColumnIndex ?? 10000).ToList();

                if (objProperties.Count == 0)
                    return _exporterHeaderInfoList;

                for (var i = 0; i < objProperties.Count; i++)
                {
                    var item = new ExporterHeaderInfo
                    {
                        Index = sumColumn = sumColumn + 1,
                        PropertyName = objProperties[i].Name,
                        ExporterHeaderAttribute = (objProperties[i].GetCustomAttributes(typeof(ExporterHeaderAttribute), true) as ExporterHeaderAttribute[])?.FirstOrDefault(),
                        CsTypeName = objProperties[i].PropertyType.GetTypeName(),
                        ExportImageFieldAttribute = objProperties[i].GetAttribute<ExportImageFieldAttribute>(true),

                    };

                    //设置列显示名
                    item.DisplayName = item.ExporterHeaderAttribute == null ||
                                                       item.ExporterHeaderAttribute.DisplayName.IsNullOrEmpty()
                                        ? item.PropertyName
                                        : item.ExporterHeaderAttribute.DisplayName;

                    ////设置Format
                    //if (item.ExporterHeaderAttribute != null && !item.ExporterHeaderAttribute.Format.IsNullOrEmpty())
                    //{
                    //    item.ExporterHeaderAttribute.Format = item.ExporterHeaderAttribute.Format;
                    //}

                    _exporterHeaderInfoList.Add(item);
                }
            }
            else
            {
                var objProperties = type.GetProperties().OrderBy(p => p.GetAttribute<ExporterHeaderAttribute>()?.ColumnIndex ?? 10000).ToList();
                if (objProperties.Count == 0)
                    return _exporterHeaderInfoList;

                for (var i = 0; i < objProperties.Count; i++)
                {
                    var item = new ExporterHeaderInfo
                    {
                        Index = sumColumn = sumColumn + 1,
                        PropertyName = objProperties[i].Name,
                        ExporterHeaderAttribute = (objProperties[i].GetCustomAttributes(typeof(ExporterHeaderAttribute), true) as ExporterHeaderAttribute[])?.FirstOrDefault(),
                        CsTypeName = objProperties[i].PropertyType.GetTypeName(),
                        ExportImageFieldAttribute = objProperties[i].GetAttribute<ExportImageFieldAttribute>(true),

                    };

                    //设置列显示名
                    item.DisplayName = item.ExporterHeaderAttribute == null ||
                                                       item.ExporterHeaderAttribute.DisplayName.IsNullOrEmpty()
                                        ? item.PropertyName
                                        : item.ExporterHeaderAttribute.DisplayName;

                    ////设置Format
                    //if (item.ExporterHeaderAttribute != null && !item.ExporterHeaderAttribute.Format.IsNullOrEmpty())
                    //{
                    //    item.ExporterHeaderAttribute.Format = item.ExporterHeaderAttribute.Format;
                    //}

                    _exporterHeaderInfoList.Add(item);
                }
            }

            return _exporterHeaderInfoList;

        }

        private ExcelWorksheet _excelWorksheet;
        private ExcelPackage _excelPackage;
        public ExcelPackage CurrentExcelPackage
        {
            get
            {
                if (_excelPackage == null)
                {
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    _excelPackage = new ExcelPackage();
                }

                return _excelPackage;
            }
            set => _excelPackage = value;
        }
        protected ExcelWorksheet CurrentExcelWorksheet
        {
            get
            {
                if (_excelWorksheet == null)
                {
                    AddExcelWorksheet();
                }

                return _excelWorksheet;
            }
            set => _excelWorksheet = value;
        }
        public ExcelWorksheet AddExcelWorksheet(string name = null)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                name = "导出结果";
            }

            _excelWorksheet = CurrentExcelPackage.Workbook.Worksheets.Add(name);
            _excelWorksheet.OutLineApplyStyle = true;
            return _excelWorksheet;
        }



        [Fact(DisplayName = "")]
        public  void DynamicHeaderExport_Test2()
        {
            var filePath = GetTestFilePath($"{nameof(DynamicHeaderExport_Test2)}.xlsx");

            DeleteFile(filePath);

            List<TitleItem> titleItems = GetTitleItems();
            List<BudgetItem> budgetItems = GetBudgetItems();
            List<List<SubTotalItem>> subTotalItems = GetSubTotalItems();
            int curRowIndex = 0;

            ExporterHeaderInfoList(typeof(TitleItem));
            ExporterHeaderInfoList(typeof(BudgetItem));
            foreach (var itemSubTotalHeader in subTotalItems)
            {
                ExporterHeaderInfoList(typeof(SubTotalItem));
            }
           


            SetHeader();

            foreach (var titleItem in titleItems)
            {
                curRowIndex++;
                CurrentExcelWorksheet.Cells[curRowIndex, 1].Value = titleItem.Num;
                CurrentExcelWorksheet.Cells[curRowIndex, 2].Value = titleItem.Title;
            }
            //curRowIndex = 0;
            //foreach (var budgetItem in budgetItems)
            //{
            //    curRowIndex++;
            //    workSheet.Cells[curRowIndex, titlePropertyCount + 1].Value = budgetItem.TotalPriceN;
            //    workSheet.Cells[curRowIndex, titlePropertyCount + 2].Value = budgetItem.UPriceN;
            //    workSheet.Cells[curRowIndex, titlePropertyCount + 3].Value = budgetItem.TaxRate;
            //    workSheet.Cells[curRowIndex, titlePropertyCount + 4].Value = budgetItem.Taxation;
            //    workSheet.Cells[curRowIndex, titlePropertyCount + 5].Value = budgetItem.TotalPrice;

            //    workSheet.Cells[curRowIndex, titlePropertyCount + 6].Value = budgetItem.UPrice;
            //    workSheet.Cells[curRowIndex, titlePropertyCount + 7].Value = budgetItem.Remark;
            //}

            File.WriteAllBytes("d:\\2.xlsx", CurrentExcelPackage.GetAsByteArray());

        }

        /// <summary>
        /// 设置头部样式
        /// </summary>
        protected void SetHeader()
        {
            foreach (var exporterHeader in _exporterHeaderInfoList)
            {
                var colCell = CurrentExcelWorksheet.Cells[1, exporterHeader.Index];
                colCell.Value = exporterHeader.DisplayName;

                var exporterHeaderAttribute = exporterHeader.ExporterHeaderAttribute;
                if (exporterHeaderAttribute != null)
                {
                    //colCell.Style.Font.Bold = exporterHeaderAttribute.IsBold;//当前字段加粗
                }
            }

        }
    }

    public class TitleItem
    {
        /// <summary>
        /// 序号
        /// </summary>
        [ExporterHeader(DisplayName = "序号")]
        public string Num { get; set; }
        /// <summary>
        /// 单项工程 单位工程名称
        /// </summary>
        [ExporterHeader(DisplayName = "项目名称")]
        public string Title { get; set; }
    }

    /// <summary>
    /// 预算价
    /// </summary>
    public class BudgetItem : SubTotalItem
    {
        /// <summary>
        /// 备注
        /// </summary>
        [ExporterHeader(DisplayName = "备注", ColumnIndex = 10)]
        public string Remark { get; set; }
    }


    /// <summary>
    /// 投标报价汇总表子项
    /// </summary>
    public class SubTotalItem
    {
        /// <summary>
        /// 招标清单部分不含税金额（元）
        /// </summary>
        [ExporterHeader(DisplayName = "招标清单部分不含税金额（元）")]
        public double TotalPriceN { get; set; }
        /// <summary>
        /// 不含税单方指标（元 / 景观总面积）
        /// </summary>
        [ExporterHeader(DisplayName = "不含税单方指标（元 / 景观总面积）")]
        public double UPriceN { get; set; }
        /// <summary>
        /// 税率
        /// </summary>
        [ExporterHeader(DisplayName = "税率")]
        public double TaxRate { get; set; }
        /// <summary>
        /// 税金（元）
        /// </summary>
        [ExporterHeader(DisplayName = "税金（元）")]
        public double Taxation { get; set; }
        /// <summary>
        /// 含税总金额（元）
        /// </summary>
        [ExporterHeader(DisplayName = "含税总金额（元）")]
        public double TotalPrice { get; set; }
        /// <summary>
        /// 含税单方指标（元 / 景观总面积）
        /// </summary>
        [ExporterHeader(DisplayName = "含税单方指标（元 / 景观总面积）")]
        public double UPrice { get; set; }
    }




}
