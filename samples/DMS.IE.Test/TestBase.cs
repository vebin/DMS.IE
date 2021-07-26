﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace DMS.IE.Test
{
    public class TestBase
    {
        public TestBase()
        {
#if !NET461
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            //Encoding.UTF8
            Encoding.GetEncoding(65001);
#endif
        }
        /// <summary>
        ///     获取根目录
        /// </summary>
        /// <returns></returns>
        public string GetTestRootPath()
        {
            return Directory.GetCurrentDirectory();
        }

        /// <summary>
        ///     获取测试文件路径
        /// </summary>
        /// <param name="paths"></param>
        /// <returns></returns>
        public string GetTestFilePath(params string[] paths)
        {
            var rootPath = GetTestRootPath();
            var list = new List<string>
            {
                rootPath
            };
            list.AddRange(paths);
            return Path.Combine(list.ToArray());
        }

        /// <summary>
        ///     删除文件
        /// </summary>
        public void DeleteFile(string path)
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
