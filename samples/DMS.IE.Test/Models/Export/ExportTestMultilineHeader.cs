using DMS.Excel.Attributes;
using DMS.Excel.Attributes.Export;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;

namespace DMS.IE.Test.Models.Export
{
    [ExcelExporter(Name = "测试数据")]
    public class ExportTestMultilineHeader
    {
        /// <summary>
        /// Text：索引10
        /// </summary>
        [ExporterHeader(DisplayName = "加粗文本", IsBold = true, ColumnIndex = 10, WrapText = true)]
        public string Text { get; set; }
        /// <summary>
        /// Text2：索引1
        /// </summary>
        [ExporterHeader(DisplayName = "普通文本", ColumnIndex = 1, Hidden = true)]
        public string Text2 { get; set; }
        /// <summary>
        /// Text3:索引2
        /// </summary>
        [ExporterHeader(DisplayName = "文本3", ColumnIndex = 2)]
        public string Text3 { get; set; }
        [ExporterHeader(DisplayName = "文本4", ColumnIndex = 2)]
        public CompanInfo companInfo { get; set; }

    }

    public class CompanInfo
    {
        public string Compan { get; set; }
        public List<CompanParent> parents { get; set; }
    }

    public class CompanParent
    {
        public string Name { get; set; }
    }
}
