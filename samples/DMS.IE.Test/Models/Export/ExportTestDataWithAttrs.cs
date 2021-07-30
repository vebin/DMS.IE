using DMS.Excel.Attributes;
using DMS.Excel.Attributes.Export;
using OfficeOpenXml.Table;
using System;

namespace DMS.IE.Test.Models.Export
{
    [ExcelExporter(Name = "测试数据", AutoCenter = true, AutoFitAllColumn = true, IsBold = true, MaxRowNumberOnASheet =20, TableStyle = TableStyles.None)]
    public class ExportTestDataWithAttrs
    {
        /// <summary>
        /// Text：索引10
        /// </summary>
        [ExporterHeader(DisplayName = "加粗文本")]
        public string Text { get; set; }
        /// <summary>
        /// Text2：索引1
        /// </summary>
        [ExporterHeader(DisplayName = "普通文本")]
        public string Text2 { get; set; }
        /// <summary>
        /// Text3:索引2
        /// </summary>
        [ExporterHeader(DisplayName = "文本3")]
        public string Text3 { get; set; }
        /// <summary>
        /// Number:索引3
        /// </summary>
        [ExporterHeader(DisplayName = "数值", Format = "#,##0")]
        public int Number { get; set; }

        [ExporterHeader(DisplayName = "名称")]
        public string Name { get; set; }

        /// <summary>
        /// 时间测试
        /// </summary>
        [ExporterHeader(DisplayName = "日期1", Format = "yyyy-MM-dd")]
        public DateTime Time1 { get; set; }

        /// <summary>
        /// 时间测试
        /// </summary>
        [ExporterHeader(DisplayName = "日期2", Format = "yyyy-MM-dd HH:mm:ss")]
        public DateTime? Time2 { get; set; }

        [ExporterHeader(Width = 100)]
        public DateTime Time3 { get; set; }
        public DateTime Time4 { get; set; }

        /// <summary>
        /// 长数值测试
        /// </summary>
        [ExporterHeader(DisplayName = "长数值", Format = "#,##0")]
        public long LongNo { get; set; }
    }
}
