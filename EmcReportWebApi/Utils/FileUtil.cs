using System;
using System.IO;
using System.IO.Compression;
using EmcReportWebApi.Config;

namespace EmcReportWebApi.Utils
{
    /// <summary>
    /// 文件操作类
    /// </summary>
    public static class FileUtil
    {
        /// <summary>
        /// 创建目录
        /// </summary>
        public static string CreateDirectory(string fileFullName)
        {
            string datetimeStr = Guid.NewGuid().ToString();
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
        /// 创建报告目录文件
        /// </summary>
        /// <returns></returns>
        public static string CreateReportFilesDirectory()
        {
            string reportFiles = Guid.NewGuid().ToString();
            string outputPath = $@"{EmcConfig.ReportFilesPathRoot}{reportFiles}";
            if (Directory.Exists(outputPath))
            {
                throw new Exception($@"报告文件所需文件夹已经存在\{reportFiles}");
            }
            Directory.CreateDirectory(outputPath);
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
            return Path.GetExtension(fileFullName);
        }


        /// <summary>
        /// 解压zip文件
        /// </summary>
        /// <param name="zipPath"></param>
        /// <param name="outputDirectory"></param>
        public static void DecompressionZip(string zipPath, string outputDirectory)
        {
            DirectoryInfo di = new DirectoryInfo(outputDirectory);
            if (!di.Exists) { di.Create(); }
            ZipFile.ExtractToDirectory(zipPath, outputDirectory);
        }

        /// <summary>
        /// 打包zip文件
        /// </summary>
        /// <param name="inputDirectory"></param>
        /// <param name="zipPath"></param>
        public static void CreateFromDirectoryZip(string inputDirectory, string zipPath)
        {
            if (!Directory.Exists(inputDirectory))
            {
                throw new Exception("文件夹路径不存在");
            }
            ZipFile.CreateFromDirectory(inputDirectory, zipPath);
        }

    }
}