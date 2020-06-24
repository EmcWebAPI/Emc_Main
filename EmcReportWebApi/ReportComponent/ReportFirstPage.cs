using System;
using Newtonsoft.Json.Linq;

namespace EmcReportWebApi.ReportComponent
{
    /// <summary>
    /// 报告首页信息
    /// </summary>
    public class ReportFirstPage
    {
        private readonly JObject _reportJsonObjectForWord;

        /// <summary>
        /// 获取首页信息
        /// </summary>
        /// <param name="reportJsonObjectForWord"></param>
        public ReportFirstPage(JObject reportJsonObjectForWord)
        {
            _reportJsonObjectForWord = reportJsonObjectForWord;
        }
        /// <summary>
        /// 首页json
        /// </summary>
        public JObject FirstPageObject
        {
            get => (JObject)(_reportJsonObjectForWord["firstPage"] ?? throw new Exception("合同信息不能为null"));
            set => throw new NotImplementedException();
        }
    }
}