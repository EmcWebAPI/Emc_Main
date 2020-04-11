using EmcReportWebApi.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EmcReportWebApi.Business
{
    public interface IReport
    {
        /// <summary>
        /// 创建报告公共方法
        /// </summary>
        /// <param name="para">参数</param>
        /// <param name="reportType">报告类型</param>
        /// <returns></returns>
        ReportResult<string> CreateReportCommon(ReportParams para, int reportType);

        /// <summary>
        /// 报告json转成word
        /// </summary>
        /// <param name="reportId"></param>
        /// <param name="jsonStr"></param>
        /// <param name="reportFilesPath"></param>
        /// <returns></returns>
        string JsonToWord(string reportId, string jsonStr, string reportFilesPath);

        /// <summary>
        /// 生成标准报告方法
        /// </summary>
        /// <param name="reportId"></param>
        /// <param name="jsonStr"></param>
        /// <param name="reportFilesPath"></param>
        /// <returns></returns>
        string JsonToWordStandard(string reportId, string jsonStr, string reportFilesPath);
    }
}