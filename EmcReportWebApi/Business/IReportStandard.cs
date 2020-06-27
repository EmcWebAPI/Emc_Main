using EmcReportWebApi.Models;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using EmcReportWebApi.StandardReportComponent;

namespace EmcReportWebApi.Business
{
    /// <summary>
    /// 标准报告接口
    /// </summary>
    public interface IReportStandard
    {
        /// <summary>
        /// 创建报告
        /// </summary>
        /// <param name="para">参数</param>
        /// <returns></returns>
        ReportResult<string> CreateReportStandard(StandardReportParams para);

        /// <summary>
        /// 生成标准报告方法
        /// </summary>
        /// <param name="standardReportInfo"></param>
        /// <returns></returns>
        StandardReportResult JsonToWordStandard(StandardReportInfo standardReportInfo);
        
    }
}