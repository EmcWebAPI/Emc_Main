using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EmcReportWebApi.Models
{
    /// <summary>
    /// 报告返回结果
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ReportResult<T>
    {
        /// <summary>
        /// 返回内容
        /// </summary>
        public object Content { get; set; }

        /// <summary>
        /// 返回信息
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// 提交结果是否成功
        /// </summary>
        public bool SumbitResult { get; set; }
    }
}