using System;
using System.IO;
using System.IO.Compression;

namespace EmcReportWebApi.Common
{
    public static class ZipFileHelper
    {
        /// <summary>
        /// 解压zip文件
        /// </summary>
        /// <param name="zipPath"></param>
        /// <param name="outputDirectory"></param>
        public static void DecompressionZip(string zipPath, string outputDirectory) {
            DirectoryInfo di = new DirectoryInfo(outputDirectory);
            if (!di.Exists) { di.Create(); }
            ZipFile.ExtractToDirectory(zipPath, outputDirectory);
        }

        /// <summary>
        /// 打包zip文件
        /// </summary>
        /// <param name="inputDirectory"></param>
        /// <param name="zipPath"></param>
        public static void CreateFromDirectoryZip(string inputDirectory, string zipPath) {
            if (!Directory.Exists(inputDirectory)) {
                throw new Exception("文件夹路径不存在");
            }
            ZipFile.CreateFromDirectory(inputDirectory, zipPath);
        }
    }
}