using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Text;

namespace DMS.Excel.Attributes.Export
{
    /// <summary>
    /// Excel导出特性
    /// </summary>
    public class ExcelExporterAttribute : ExporterAttribute
    {
        /// <summary>
        /// 自动居中(设置后为全局居中显示)
        /// </summary>
        public bool AutoCenter { get; set; }

        /// <summary>
        /// 表格样式风格
        /// </summary>
        public TableStyles TableStyle { get; set; } = TableStyles.None;





    }

   
}
