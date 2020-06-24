using System;
using EmcReportWebApi.Config;
using EmcReportWebApi.Models;
using EmcReportWebApi.Utils;

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
            this.ReportFilesPath = FileUtil.CreateReportFilesDirectory();
            this.ReportId = string.IsNullOrEmpty(para.ReportId)?"QW2018-698":para.ReportId;
            this.ReportJsonStrForWord = para.JsonStr;
            if (para.ZipFilesUrl != null && !para.ZipFilesUrl.Equals(""))
            {
                string zipUrl = para.ZipFilesUrl;
                this.ReportZipFilesPath = $@"{ReportFilesPath}\zip{Guid.NewGuid()}.zip";

                byte[] fileBytes = SyncHttpHelper.GetHttpRespponseForFile(zipUrl, ReportZipFilesPath,
                    int.Parse(DateTime.Now.ToString("hhmmss")));
                if (fileBytes.Length <= 0)
                {
                    EmcConfig.ErrorLog.Error($"请求报告失败,报告id:{para.ReportId}");
                    throw new Exception($"请求报告文件失败,报告id{para.ReportId}");
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

        /// <summary>
        /// 报告编号
        /// </summary>
        public string ReportId { get; set; }

        /// <summary>
        /// 报告转word的Json
        /// </summary>
        public string ReportJsonStrForWord { get; set; }

        /// <summary>
        /// 报告所需文件路径
        /// </summary>
        public string ReportFilesPath { get; set; }

        /// <summary>
        /// 报告zip文件路径
        /// </summary>
        public string ReportZipFilesPath { get; set; }
    }
}