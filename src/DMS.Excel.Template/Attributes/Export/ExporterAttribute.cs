using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Text;

namespace DMS.Excel.Attributes.Export
{
    [AttributeUsage(AttributeTargets.Class)]
    public class ExporterAttribute : Attribute
    {
        /// <summary>
        /// 名称(比如当前Sheet 名称)
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 头部字体大小
        /// </summary>
        public float? HeaderFontSize { set; get; }

        /// <summary>
        /// 正文字体大小
        /// </summary>
        public float? FontSize { set; get; }

        /// <summary>
        /// 一个Sheet最大允许的行数，设置了之后将输出多个Sheet
        /// </summary>
        public int MaxRowNumberOnASheet { get; set; } = 0;

        /// <summary>
        /// 自适应所有列
        /// </summary>
        public bool AutoFitAllColumn { get; set; }

        /// <summary>
        /// 数据超过此行之后不启用自适应，默认关闭
        /// </summary>
        public int AutoFitMaxRows { get; set; }

        /// <summary>
        /// 作者
        /// </summary>
        public string Author { get; set; }

    }

   
}
