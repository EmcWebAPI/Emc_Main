using System;
using System.IO;
using EmcReportWebApi.Config;
using EmcReportWebApi.Models;
using EmcReportWebApi.ReportComponent.FirstPage;
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
        private readonly ReportParams _para;

        /// <summary>
        /// 构造报告文件信息
        /// </summary>
        public ReportInfo(ReportParams para)
        {
            try
            {
                _para = para;
                ReportFilesPath = FileUtil.CreateReportFilesDirectory();
                TemplateFileFullName = CreateTemplateMiddle(EmcConfig.ReportTemplateMiddlewareFilePath, "template",
                    EmcConfig.ReportTemplateFileFullName);
                this.ReportJsonObjectForWord = JsonConvert.DeserializeObject<JObject>(this.ReportJsonStrForWord);
                this.DecompressionReportFiles();
                this.ReportId = string.IsNullOrEmpty(_para.ReportId) ? "QW2018-698" : _para.ReportId;
                this.ReportZipFilesPath = $@"{ReportFilesPath}\zip{Guid.NewGuid()}.zip";
                this.FileName = $"Report{Guid.NewGuid()}.docx";
                this.OutFileFullName = $"{EmcConfig.ReportOutputPath}{FileName}";
                //首页信息
                ReportFirstPage = new ReportFirstPage(this.ReportJsonObjectForWord,this.ReportId);
                //审查表信息
                ReviewTableInfo = new ReviewTableInfo(this.ReportJsonObjectForWord,this.ReportFilesPath);
            }
            catch (Exception ex)
            {
                EmcConfig.ErrorLog.Error(ex.Message, ex);//设置错误信息
                throw new Exception(ex.Message);
            }
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
        /// 报告转word的Json
        /// </summary>
        public string ReportJsonStrForWord
        {
            get => _para.JsonStr;
            set => throw new NotImplementedException();
        }

        /// <summary>
        /// 报告编号
        /// </summary>
        public string ReportId { get; set; }

        /// <summary>
        /// 报告所需文件路径
        /// </summary>
        public string ReportFilesPath { get; set; }

        /// <summary>
        /// 报告zip文件路径
        /// </summary>
        public string ReportZipFilesPath { get; set; }

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
        /// 报告模板路径(路径+文件名)
        /// </summary>
        public string TemplateFileFullName { get; set; }

        private void DecompressionReportFiles()
        {
            if (_para.ZipFilesUrl != null && !_para.ZipFilesUrl.Equals(""))
            {
                string zipUrl = _para.ZipFilesUrl;

                byte[] fileBytes = SyncHttpHelper.GetHttpRespponseForFile(zipUrl, ReportZipFilesPath,
                    int.Parse(DateTime.Now.ToString("hhmmss")));
                if (fileBytes.Length <= 0)
                {
                    EmcConfig.ErrorLog.Error($"请求报告失败,报告id:{_para.ReportId}");
                    throw new Exception($"请求报告文件失败,报告id{_para.ReportId}");
                }
                //解压zip文件
                FileUtil.DecompressionZip(ReportZipFilesPath, ReportFilesPath);
            }
            else
            {
                string reportZipFilesPath = $@"{EmcConfig.ReportFilesPathRoot}Test\{"QT2019-3015.zip"}";
                //解压zip文件
                FileUtil.DecompressionZip(reportZipFilesPath, ReportFilesPath);
            }
        }

        private string CreateTemplateMiddle(string dir, string template, string filePath)
        {
            string dateStr = Guid.NewGuid().ToString();
            string fileName = template + dateStr + ".docx";
            DirectoryInfo di = new DirectoryInfo(dir);
            if (!di.Exists) { di.Create(); }

            string fileFullName = $"{dir}{fileName}";
            FileInfo file = new FileInfo(filePath);
            if (File.Exists(filePath))
            {
                file.CopyTo(fileFullName);
                return fileFullName;
            }

            throw new Exception("模板不存在");
        }
    }
}