using System.Collections.Generic;

namespace DMS.Excel.Models
{
    /// <summary>
    /// 数据行错误信息
    /// </summary>
    public class DataRowErrorInfo
    {
        /// <summary>
        ///  
        /// </summary>
        public DataRowErrorInfo()
        {
            FieldErrors = new Dictionary<string, string>();
        }

        /// <summary>
        /// 序号
        /// </summary>
        public int RowIndex { get; set; }

        /// <summary>
        /// 字段错误信息
        /// </summary>
        public IDictionary<string, string> FieldErrors { get; set; }
    }
}
