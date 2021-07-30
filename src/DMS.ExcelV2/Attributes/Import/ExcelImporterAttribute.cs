using System;
using System.Collections.Generic;
using System.Text;

namespace DMS.ExcelV2.Attributes.Import
{
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Property)]
    public class ExcelImporterAttribute : ImporterAttribute
    {

        /// <summary>
        /// 指定Sheet下标（获取指定Sheet下标）
        /// </summary>
        public int SheetIndex { get; set; } = 0;
    }
}
