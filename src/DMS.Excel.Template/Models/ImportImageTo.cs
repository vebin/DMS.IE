using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.Excel.Models
{
    /// <summary>
    /// 图片导入类型
    /// </summary>
    public enum ImportImageTo
    {
        /// <summary>
        /// 导入到临时目录
        /// </summary>
        TempFolder,

        /// <summary>
        /// 导入为base64格式
        /// </summary>
        Base64
    }
}
