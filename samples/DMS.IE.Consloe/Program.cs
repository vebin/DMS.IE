using Aspose.Pdf;
using System;

namespace DMS.IE.Consloe
{
    class Program
    {
        static void Main(string[] args)
        {
            pdf2word();
            Console.WriteLine("Hello World!");
        }

        /// <summary>
        /// 10.1.0.0版本生成
        /// 导出模板可用
        /// </summary>
        public static void pdf2word()
        {
            // 文档目录的路径。
            string dataFile = @"D:\导入Excel\融创集团华北区域公司土建类企业定额（2019）【天津】V6.0价税分离20.04(3).pdf";

            //打开源PDF文档
            Document pdfDocument = new Document(dataFile);

            // 实例化DocSaveOptions对象
            DocSaveOptions saveOptions = new DocSaveOptions();
            //将输出格式指定为DOCX
            //saveOptions.Format = DocSaveOptions.DocFormat.Doc;
            // 以docx格式保存文档
            pdfDocument.Save("ConvertToDOCX_out.doc", Aspose.Pdf.SaveFormat.Doc);



            //var pfile = new Aspose.Pdf.Document(dir + "template.pdf");
            //// save in different formats
            //pfile.Save(dir + "output.docx", Aspose.Pdf.SaveFormat.DocX);

        }
    }
}
