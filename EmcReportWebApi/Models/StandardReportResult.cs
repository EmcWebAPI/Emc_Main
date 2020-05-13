using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EmcReportWebApi.Models
{
    public class StandardReportResult
    {
        /// <summary>
        /// 返回文件路径
        /// </summary>
        public string FilePath { get; set; }

        /// <summary>
        /// 生成文件名称
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// 报告编号
        /// </summary>
        public string ReportCode { get; set; }

        /// <summary>
        /// 返回状态码
        /// </summary>
        public bool Status { get; set; }

        /// <summary>
        /// 返回信息
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// 回调函数请求路径
        /// </summary>
        public string CallBackUrl { get; set; }

        /// <summary>
        /// 报告id
        /// </summary>
        public string ReportId { get; set; }

        /// <summary>
        /// 合同id
        /// </summary>
        public string ContractId { get; set; }
    }
}