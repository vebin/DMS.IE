using System;

namespace DMS.Excel.Attributes.Export
{
    /// <summary>
    /// 导出属性特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ExporterHeaderAttribute : Attribute
    {
      
        /// <summary>
        /// 显示名称
        /// </summary>
        public string DisplayName { set; get; }
        /// <summary>
        /// 字体大小
        /// </summary>
        public float? FontSize { set; get; }
        /// <summary>
        /// 是否加粗
        /// </summary>
        public bool IsBold { set; get; }
        /// <summary>
        /// 格式化
        /// </summary>
        public string Format { get; set; }
        /// <summary>
        /// 宽度
        /// </summary>
        public int Width { get; set; }
        /// <summary>
        /// Hidden
        /// </summary>
        public bool Hidden { get; set; }

    }
}
