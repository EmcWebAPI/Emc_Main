using EmcReportWebApi.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EmcReportWebApi.Business
{
    public interface IReportStandard
    {
        /// <summary>
        /// 创建报告
        /// </summary>
        /// <param name="para">参数</param>
        /// <returns></returns>
        ReportResult<string> CreateReportStandard(ReportParams para);

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