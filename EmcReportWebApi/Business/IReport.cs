using EmcReportWebApi.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using EmcReportWebApi.ReportComponent;

namespace EmcReportWebApi.Business
{
    /// <summary>
    /// 报告接口
    /// </summary>
    public interface IReport
    {
        /// <summary>
        /// 创建报告
        /// </summary>
        /// <param name="para">参数</param>
        /// <returns></returns>
        ReportResult<string> CreateReport(ReportParams para);

        /// <summary>
        /// json转成word
        /// </summary>
        /// <param name="reportInfo">报告信息</param>
        /// <returns></returns>
        string ReportJsonToWord(ReportInfo reportInfo);
    }
}