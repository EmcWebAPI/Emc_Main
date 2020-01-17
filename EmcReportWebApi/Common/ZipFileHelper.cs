using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.IO.Compression;

namespace EmcReportWebApi.Common
{
    public static class ZipFileHelper
    {
        public static void DecompressionZip(string zipPath, string outputDirectory) {
            DirectoryInfo di = new DirectoryInfo(outputDirectory);
            if (!di.Exists) { di.Create(); }
            ZipFile.ExtractToDirectory(zipPath, outputDirectory);
        }

        public static void CreateFromDirectoryZip(string inputDirectory, string zipPath) {
            if (!Directory.Exists(inputDirectory)) {
                throw new Exception("文件夹路径不存在");
            }
            ZipFile.CreateFromDirectory(inputDirectory, zipPath);
        }
    }
}