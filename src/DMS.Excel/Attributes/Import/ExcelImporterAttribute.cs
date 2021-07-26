using System;
using System.Collections.Generic;
using System.Text;

namespace DMS.Excel.Attributes.Import
{
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Property)]
    public class ExcelImporterAttribute : ImporterAttribute
    {
        /// <summary>
        /// 指定Sheet名称(获取指定Sheet名称)
        /// 为空则自动获取第一个
        /// </summary>
        //public string SheetName { get; set; }

        /// <summary>
        /// 指定Sheet下标（获取指定Sheet下标）
        /// </summary>
        /// <remarks>
        /// 在.NET Core+包括.NET5框架中下标从0开始，否则从 1 
        /// </remarks>
        public int SheetIndex { get; set; } =
#if NET461
            1
#else
            0
#endif
            ;


        /// <summary>
        /// Sheet顶部导入描述
        /// </summary>
        public string ImportDescription { get; set; }
    }
}
