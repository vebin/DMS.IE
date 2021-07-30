using System;

namespace DMS.Excel.Attributes.Import
{
    /// <summary>
    /// 导入头部特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ImporterHeaderAttribute : Attribute
    {
        /// <summary>
        /// 显示名称
        /// </summary>
        public string Name { set; get; }

        /// <summary>
        /// 作者
        /// </summary>
        public string Author { set; get; } = "dylan;hailang";

        /// <summary>
        /// 列索引
        /// </summary>
        public int ColumnIndex { get; set; }
    }
}
