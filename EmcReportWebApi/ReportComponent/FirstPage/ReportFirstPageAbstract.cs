using EmcReportWebApi.Business.ImplWordUtil;
using EmcReportWebApi.Utils;
using Newtonsoft.Json.Linq;

namespace EmcReportWebApi.ReportComponent.FirstPage
{
    /// <summary>
    /// 首页信息抽象
    /// </summary>
    public abstract class ReportFirstPageAbstract
    {
        /// <summary>
        /// 写入word首页信息
        /// </summary>
        /// <param name="wordUtil"></param>
        public abstract void WriteFirstPage(ReportHandleWord wordUtil);

        /// <summary>
        /// 首页json
        /// </summary>
        public JObject FirstPageObject { get; set; }

        /// <summary>
        /// 首页上的报告编号
        /// </summary>
        public string ReportCode { get; set; }

        /// <summary>
        /// 首页上报告编号书签
        /// </summary>
        public string ReportCodeBookmark { get; set; } = "reportId";

        /// <summary>
        /// 页眉报告编号
        /// </summary>
        public string ReportYmCode { get; set; }
    }
}