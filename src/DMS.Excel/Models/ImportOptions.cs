using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace DMS.Excel.Models
{
    public class ImportOptions
    {
        /// <summary>
        /// 工作表编号（默认1）
        /// <para>从 1 开始</para>
        /// <para>0：全部sheet，1：第一个sheet，2：第二个sheet，……</para>
        /// </summary>
        [Display(Name = "工作表编号")]
        [Range(0, int.MaxValue, ErrorMessage = "{0}最小值为{1}")]
        public int SheetIndex { get; set; }
        /// <summary>
        /// 表头行编号（默认1）
        /// <para>从 1 开始</para>
        /// </summary>
        [Display(Name = "表头行编号")]
        [Range(1, int.MaxValue, ErrorMessage = "{0}最小值为{1}")]
        public int HeaderRowIndex { get; set; }

        /// <summary>
        /// 数据起始行编号（默认2）
        /// <para>从 1 开始</para>
        /// </summary>
        [Display(Name = "数据起始行编号")]
        [Range(2, int.MaxValue, ErrorMessage = "{0}最小值为{1}")]
        public int DataRowStartIndex { get; set; }

        /// <summary>
        /// 数据结束行编号（默认最后一行）
        /// <para>从 1 开始</para>
        /// </summary>
        [Display(Name = "数据结束行编号")]
        [Range(2, int.MaxValue, ErrorMessage = "{0}最小值为{1}")]
        public int? DataRowEndIndex { get; set; }


        /// <summary>
        /// 构造
        /// </summary>
        public ImportOptions()
        {
            SheetIndex = 1;
            HeaderRowIndex = 1;
            DataRowStartIndex = 2;
            DataRowEndIndex = null;
        }

    }
}
