using System;
using System.Collections.Generic;
using System.IO;
using EmcReportWebApi.Business.ImplWordUtil;
using EmcReportWebApi.Config;
using EmcReportWebApi.Models;
using EmcReportWebApi.ReportComponent.Image;
using EmcReportWebApi.Utils;
using Newtonsoft.Json.Linq;

namespace EmcReportWebApi.StandardReportComponent
{
    /// <summary>
    /// 报告信息
    /// </summary>
    public class StandardReportInfo
    {
        /// <summary>
        /// 构造报告文件信息
        /// </summary>
        public StandardReportInfo(StandardReportParams para)
        {
            try
            {
                ZipFilesUrl = para.ZipFilesUrl;
                ReportId = para.OriginalRecord;
                ReportFilesPath = FileUtil.CreateReportFilesDirectory();
                TemplateFileFullName = CreateTemplateMiddle();
                ReportJsonObjectForWord = para.JsonObject;
                ReportZipFileFullPath = $@"{ReportFilesPath}\zip{Guid.NewGuid()}.zip";
                FileName = $"StandardReport{Guid.NewGuid()}.docx";
                OutFileFullName = $"{EmcConfig.ReportOutputPath}{FileName}";
                DecompressionReportFiles();
                //首页信息
                ReportFirstPage = new StandardReportFirstPage(this,this.ReportJsonObjectForWord);

                StandardReportResultInfo = new StandardReportResult
                {
                    FilePath = OutFileFullName,
                    ReportCode = ReportFirstPage.ReportCode,
                    FileName = FileName
                };
            }
            catch (Exception ex)
            {
                EmcConfig.ErrorLog.Error(ex.Message, ex);//设置错误信息
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public void HandleReportHeader(ReportStandardHandleWord wordUtil)
        {
            int pageCount = wordUtil.GetDocumnetPageCount() - 2;//获取文件页数(首页不算)
            Dictionary<int, Dictionary<string, string>> replaceDic = new Dictionary<int, Dictionary<string, string>>();
            Dictionary<string, string> valuePairs = new Dictionary<string, string>
            {
                {"bgbh", ReportFirstPage.ReportCode},//报告编号
                {"ypbh", ReportFirstPage.SampleCode},//样品编号
                {"page", pageCount.ToString()}
            };
            replaceDic.Add(3, valuePairs);//替换页眉

            wordUtil.ReplaceWritten(replaceDic);
        }

        /// <summary>
        /// 删除模板中间件文件夹
        /// </summary>
        public void DeleteTemplateMiddleDirectory()
        {
            DeleteDir(TemplateMiddleFilesPath);
            DeleteDir(ReportFilesPath);
        }

        /// <summary>
        /// 报告首页内容
        /// </summary>
        public StandardReportFirstPage ReportFirstPage { get; set; }

        /// <summary>
        /// 图片文件
        /// </summary>
        public ImageInfo ImageInfo { get; set; }

        /// <summary>
        /// 报告转word的Json
        /// </summary>
        public string ReportJsonStrForWord { get; set; }

        /// <summary>
        /// 报告编号
        /// </summary>
        public string ReportId { get; set; }

        /// <summary>
        /// 报告所需文件夹路径
        /// </summary>
        public string ReportFilesPath { get; set; }

        /// <summary>
        /// 报告zip文件
        /// </summary>
        public string ReportZipFileFullPath { get; set; }

        /// <summary>
        /// 报告转word的Json
        /// </summary>
        public JObject ReportJsonObjectForWord { get; set; }

        /// <summary>
        /// 报告文件名称
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// 输出文件全路径(路径+文件名)
        /// </summary>
        public string OutFileFullName { get; set; }

        /// <summary>
        /// 模板中间件文件夹
        /// </summary>
        public string TemplateMiddleFilesPath { get; set; }

        /// <summary>
        /// 解压文件的请求路径
        /// </summary>
        public string ZipFilesUrl { get; set; }

        /// <summary>
        /// 报告模板路径(路径+文件名)
        /// </summary>
        public string TemplateFileFullName { get; set; }

        /// <summary>
        /// 标准报告返回的结果
        /// </summary>
        public StandardReportResult StandardReportResultInfo { get; set; }

        private void DecompressionReportFiles()
        {
            if (ZipFilesUrl != null && !ZipFilesUrl.Equals(""))
            {
                string zipUrl = ZipFilesUrl;

                byte[] fileBytes = SyncHttpHelper.GetHttpRespponseForFile(zipUrl, ReportZipFileFullPath,
                    int.Parse(DateTime.Now.ToString("hhmmss")));
                if (fileBytes.Length <= 0)
                {
                    EmcConfig.ErrorLog.Error($"请求报告失败,报告id:{ReportId}");
                    throw new Exception($"请求报告文件失败,报告id{ReportId}");
                }
                //解压zip文件
                FileUtil.DecompressionZip(ReportZipFileFullPath, ReportFilesPath);
            }
            else
            {
                string reportZipFilesPath = $@"{EmcConfig.ReportFilesPathRoot}Test\{"QT2019-3015.zip"}";
                //解压zip文件
                FileUtil.DecompressionZip(reportZipFilesPath, ReportFilesPath);
            }
        }

        private string CreateTemplateMiddle()
        {
            string dateStr = Guid.NewGuid().ToString();
            TemplateMiddleFilesPath = $@"{EmcConfig.ReportTemplateMiddlewareFilePath}\{Guid.NewGuid()}\";
            string fileName = dateStr + ".docx";
            DirectoryInfo di = new DirectoryInfo(TemplateMiddleFilesPath);
            if (!di.Exists) { di.Create(); }

            string fileFullName = $"{TemplateMiddleFilesPath}{fileName}";
            FileInfo file = new FileInfo(EmcConfig.StandardReportTemplateFileFullName);
            if (File.Exists(EmcConfig.StandardReportTemplateFileFullName))
            {
                file.CopyTo(fileFullName);
                return fileFullName;
            }

            throw new Exception("模板不存在");
        }

        /// <summary>
        /// 删除模板中间件
        /// </summary>
        private void DeleteDir(string srcPath)
        {
            DirectoryInfo dir = new DirectoryInfo(srcPath);
            FileSystemInfo[] fileInfo = dir.GetFileSystemInfos();  //返回目录中所有文件和子目录
            foreach (FileSystemInfo i in fileInfo)
            {
                if (i is DirectoryInfo)            //判断是否文件夹
                {
                    DirectoryInfo directoryInfo = new DirectoryInfo(i.FullName);
                    directoryInfo.Delete(true);          //删除子目录和文件
                }
                else
                {
                    //如果 使用了 streamReader 在删除前 必须先关闭流 ，否则无法删除 sr.close();
                    File.Delete(i.FullName);      //删除指定文件
                }
            }
            Directory.Delete(srcPath);
        }
    }
}