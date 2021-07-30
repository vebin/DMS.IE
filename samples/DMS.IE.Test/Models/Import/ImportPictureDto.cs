﻿using DMS.Excel.Attributes;
using DMS.Excel.Attributes.Import;
using DMS.Excel.Models;
using System;

namespace DMS.IE.Test.Models.Import
{
    /// <summary>
    /// 
    /// </summary>
    [Importer(DataRowStartIndex = 1)]
    public class ImportPictureDto
    {
        [ImporterHeader(Name = "加粗文本")]
        public string Text { get; set; }
        [ImporterHeader(Name = "普通文本")]
        public string Text2 { get; set; }

        /// <summary>
        /// 将图片写入到临时目录
        /// </summary>
        [ImportImageField(ImportImageTo = ImportImageTo.TempFolder)]
        [ImporterHeader(Name = "图1")]
        public string Img1 { get; set; }
        [ImporterHeader(Name = "数值")]
        public string Number { get; set; }
        [ImporterHeader(Name = "名称")]
        public string Name { get; set; }
        [ImporterHeader(Name = "日期")]
        public DateTime Time { get; set; }

        /// <summary>
        /// 将图片写入到临时目录
        /// </summary>
        [ImportImageField(ImportImageTo = ImportImageTo.TempFolder)]
        [ImporterHeader(Name = "图")]
        public string Img { get; set; }
    }
}
