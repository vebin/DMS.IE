using System.Collections.Generic;

namespace DMS.Excel.Result
{
    /// <summary>
    /// 导入结果
    /// </summary>
    public class ImportResult<T> where T : class
    {
        /// <summary>
        /// 导入数据
        /// </summary>
        public virtual ICollection<T> Data { get; set; }
    }
}
