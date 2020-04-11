using System;
using System.IO;

namespace EmcReportWebApi.Common
{
    public static class FileUtil
    {
        //根据报告id创建报告文件夹
        public static string CreateReportDirectory(string fileFullName)
        {
            string datetimeStr = DateTime.Now.ToString("yyyyMMddhhmmss");
            string outputPath = string.Format("{0}\\{1}", fileFullName, datetimeStr);
            if (Directory.Exists(outputPath))
            {
                throw new Exception("文件夹已经存在");
            }
            else
            {
                Directory.CreateDirectory(outputPath);
            }
            return outputPath;
        }

        /// <summary>
        /// 判断文件是否存在
        /// </summary>
        public static bool FileExists(string fileFullName) {
            return File.Exists(fileFullName);
        }

        /// <summary>
        /// 获取拓展名
        /// </summary>
        public static string FilterExtendName(string fileFullName)
        {
            int index = fileFullName.LastIndexOf('.');
            string extendName = fileFullName.Substring(index, fileFullName.Length - index).ToLower();

            return extendName;
        }
    }
}