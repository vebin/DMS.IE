using System;

namespace DMS.Excel.Attributes.Import
{
    /// <summary>
    /// 导入
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class ImporterAttribute : Attribute
    {
        /// <summary>
        /// 数据起始行编号
        /// </summary>
        public int DataRowStartIndex { get; set; } = 1;
        /// <summary>
        /// 数据结束行编号
        /// </summary>
        public int? DataRowEndIndex { get; set; }

    }
}
