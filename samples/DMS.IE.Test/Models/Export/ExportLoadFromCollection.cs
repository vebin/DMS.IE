using DMS.Excel.Attributes.Export;
using System;
using System.Collections.Generic;
using System.Text;

namespace DMS.IE.Test.Models.Export
{
    /// <summary>
    /// 最原始导入
    /// 没有任何样式
    /// </summary>
    public class ExportLoadFromCollection
    {
        /// <summary>
        /// 
        /// </summary>
        public int? ID { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public string Name1 { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public string Name2 { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [ExporterHeader(DisplayName = "日期1", Format = "yyyy-MM-dd")]
        public DateTime Time1 { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public long LongNo { get; set; }
    }
}
