using System;
using Newtonsoft.Json.Linq;

namespace EmcReportWebApi.ReportComponent
{
    /// <summary>
    /// 报告首页信息
    /// </summary>
    public class ReportFirstPage
    {
        private readonly string _reportId;

        /// <summary>
        /// 获取首页信息
        /// </summary>
        /// <param name="reportJsonObjectForWord"></param>
        /// <param name="reportId"></param>
        public ReportFirstPage(JObject reportJsonObjectForWord,string reportId)
        {
            _reportId = reportId;
            this.FirstPageObject = (JObject)(reportJsonObjectForWord["firstPage"] ?? throw new Exception("合同信息不能为null"));
            this.SetReportCode();
        }

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

        private void SetReportCode()
        {
            string[] reportArray = _reportId.Split('-');

            ReportCode = reportArray.Length >= 2 ? $"国医检(磁)字{reportArray[0]}第{reportArray[1]}号" : "国医检(磁)字QW2018第698号";
            ReportYmCode = reportArray.Length >= 2? $"国医检（磁）字{reportArray[0]}第{reportArray[1]}号": "国医检（磁）字QW2018第698号";
        }
    }
}