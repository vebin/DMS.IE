using DMS.Excel.Attributes.Import;
using System.Reflection;

namespace DMS.Excel.Models
{
    /// <summary>
    /// 导入列头设置
    /// </summary>
    public class ImporterHeaderInfo
    {
        /// <summary>
        /// 列名称
        /// </summary>
        public string PropertyName { get; set; }

        /// <summary>
        /// 列属性
        /// </summary>
        public ImporterHeaderAttribute HeaderAttribute { get; set; }

        /// <summary>
        /// 属性信息
        /// </summary>
        public PropertyInfo PropertyInfo { get; set; }
    }
}
