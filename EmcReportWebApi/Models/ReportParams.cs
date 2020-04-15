namespace EmcReportWebApi.Models
{
    /// <summary>
    /// 报告参数实体
    /// </summary>
    public class ReportParams
    {
        /// <summary>
        /// 需解析的json字符串
        /// </summary>
        public string JsonStr { get; set; }

        /// <summary>
        /// 报告编号
        /// </summary>
        public string ReportId { get; set; }

        /// <summary>
        /// 打包文件下载路径
        /// </summary>
        public string ZipFilesUrl { get; set; }
    }
}