﻿using EmcReportWebApi.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EmcReportWebApi.Business
{
    public interface IReport
    {
        /// <summary>
        /// 创建报告
        /// </summary>
        /// <param name="para">参数</param>
        /// <returns></returns>
        ReportResult<string> CreateReport(ReportParams para);

        /// <summary>
        /// 报告json转成word
        /// </summary>
        /// <param name="reportId"></param>
        /// <param name="jsonStr"></param>
        /// <param name="reportFilesPath"></param>
        /// <returns></returns>
        string JsonToWord(string reportId, string jsonStr, string reportFilesPath);
        
    }
}