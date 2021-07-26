using DMS.Excel.Attributes;
using DMS.Excel.Attributes.Import;

namespace DMS.IE.Test.Models.Import
{
    /// <summary>
    /// 
    /// </summary>
    [Importer(HeaderRowIndex = 1)]
    public class MergeRowsImportDto
    {
        [ImporterHeader(Name = "学号")]
        public long No { get; set; }

        [ImporterHeader(Name = "姓名")]
        public string Name { get; set; }

        [ImporterHeader(Name = "性别")]
        public string Sex { get; set; }
    }
}
