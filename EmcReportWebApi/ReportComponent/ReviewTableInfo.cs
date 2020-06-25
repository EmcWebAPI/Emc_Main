using Newtonsoft.Json.Linq;

namespace EmcReportWebApi.ReportComponent
{
    /// <summary>
    /// 审查表信息
    /// </summary>
    public class ReviewTableInfo
    {
        /// <summary>
        /// 构造
        /// </summary>
        public ReviewTableInfo(JObject reportJsonObjectForWord,string reportFilesPath)
        {
            this.ReviewTableFileFullName = reportFilesPath + "\\" + (string)reportJsonObjectForWord["scbWord"];
        }
        /// <summary>
        /// 审查表路径
        /// </summary>
        public string ReviewTableFileFullName { get; set; }
    }
}