using System;
using System.IO;

namespace EmcReportWebApi.Common
{
    public static class FileUtil
    {
        //根据报告id创建报告文件夹
        public static string CreateReportDirectory(string filePath)
        {
            string datetimeStr = DateTime.Now.ToString("yyyyMMddhhmmss");
            string outputPath = string.Format("{0}\\{1}", filePath, datetimeStr);
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
    }
}