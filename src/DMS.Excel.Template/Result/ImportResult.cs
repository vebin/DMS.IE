using DMS.Excel.Models;
using System.Collections.Generic;
using System.Linq;

namespace DMS.Excel.Result
{
    /// <summary>
    /// 导入结果
    /// </summary>
    public class ImportResult<T> where T : class
    {
        /// <summary>
        /// </summary>
        public ImportResult()
        {
            RowErrors = new List<DataRowErrorInfo>();
        }

        /// <summary>
        /// 导入数据
        /// </summary>
        public virtual ICollection<T> Data { get; set; }

        /// <summary>
        /// 验证错误
        /// </summary>
        public virtual IList<DataRowErrorInfo> RowErrors { get; set; }

        /// <summary>
        /// 模板错误
        /// </summary>
        public virtual IList<TemplateErrorInfo> TemplateErrors { get; set; }

        /// <summary>
        /// 是否存在导入错误
        /// </summary>
        public virtual bool HasError => (TemplateErrors?.Count() ?? 0) > 0 || (RowErrors?.Count ?? 0) > 0;

        /// <summary>
        ///     
        /// 
        /// </summary>
        public virtual IList<ImporterHeaderInfo> ImporterHeaderInfos { get; set; }
    }
}
