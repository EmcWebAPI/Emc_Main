using System;
using System.Collections.Generic;
using System.IO;
using EmcReportWebApi.Business.ImplWordUtil;
using EmcReportWebApi.Config;
using EmcReportWebApi.Models;
using EmcReportWebApi.ReportComponent.Experiment;
using EmcReportWebApi.ReportComponent.FirstPage;
using EmcReportWebApi.ReportComponent.Image;
using EmcReportWebApi.ReportComponent.ReviewTable;
using EmcReportWebApi.Utils;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace EmcReportWebApi.ReportComponent
{
    /// <summary>
    /// 报告信息
    /// </summary>
    public class ReportInfo
    {
        /// <summary>
        /// 构造报告文件信息
        /// </summary>
        public ReportInfo(ReportParams para)
        {
            try
            {
                ReportJsonStrForWord = para.JsonStr;
                ZipFilesUrl = para.ZipFilesUrl;
                ReportId = para.ReportId;
                ReportFilesPath = FileUtil.CreateReportFilesDirectory();
                TemplateFileFullName = CreateTemplateMiddle();
                ReportJsonObjectForWord = JsonConvert.DeserializeObject<JObject>(this.ReportJsonStrForWord);
                DecompressionReportFiles();
                ReportId = string.IsNullOrEmpty(para.ReportId) ? "QW2018-698" : para.ReportId;
                ReportZipFileFullPath = $@"{ReportFilesPath}\zip{Guid.NewGuid()}.zip";
                FileName = $"Report{Guid.NewGuid()}.docx";
                OutFileFullName = $"{EmcConfig.ReportOutputPath}{FileName}";

                //首页信息
                ReportFirstPage = new ReportFirstPage(this.ReportJsonObjectForWord,this.ReportId);
                //审查表信息
                ReviewTableInfo = new ReviewTableInfo(this.ReportJsonObjectForWord,this.ReportFilesPath);
                //实验数据信息
                ExperimentInfo = new ExperimentInfo(this,ReportJsonObjectForWord);
                //标识文件
                IdentityTableInfo = new IdentityTableInfo(this.ReportJsonObjectForWord, this.ReportFilesPath);
                //样品图片
                ImageInfo = new ImageInfo(this,ReportJsonObjectForWord);
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
        public void HandleReportHeader(ReportHandleWord wordUtil)
        {
            int pageCount = wordUtil.GetDocumnetPageCount() - 1;//获取文件页数(首页不算)

            Dictionary<int, Dictionary<string, string>> replaceDic = new Dictionary<int, Dictionary<string, string>>();
            Dictionary<string, string> valuePairs = new Dictionary<string, string>
            {
                {"reportId", ReportFirstPage.ReportYmCode}, {"page", pageCount.ToString()}
            };
            replaceDic.Add(3, valuePairs);//替换页眉

            wordUtil.ReplaceWritten(replaceDic);
        }

        /// <summary>
        /// 删除模板中间件文件夹
        /// </summary>
        public void DeleteTemplateMiddleDirctory()
        {
            DeleteDir(TemplateMiddleFilesPath);
            DeleteDir(ReportFilesPath);
        }

        /// <summary>
        /// 报告首页内容
        /// </summary>
        public ReportFirstPageAbstract ReportFirstPage { get; set; }

        /// <summary>
        /// 审查表信息
        /// </summary>
        public ReviewTableInfoAbstract ReviewTableInfo { get; set; }

        /// <summary>
        /// 实验数据信息
        /// </summary>
        public ExperimentInfo ExperimentInfo { get; set; }

        /// <summary>
        /// 标识文件
        /// </summary>
        public ReviewTableInfoAbstract IdentityTableInfo { get; set; }

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
            FileInfo file = new FileInfo(EmcConfig.ReportTemplateFileFullName);
            if (File.Exists(EmcConfig.ReportTemplateFileFullName))
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
                    DirectoryInfo subdir = new DirectoryInfo(i.FullName);
                    subdir.Delete(true);          //删除子目录和文件
                }
                else
                {
                    //如果 使用了 streamreader 在删除前 必须先关闭流 ，否则无法删除 sr.close();
                    File.Delete(i.FullName);      //删除指定文件
                }
            }
            Directory.Delete(srcPath);
        }
    }
}