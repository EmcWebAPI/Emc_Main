using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EmcReportWebApi.Models
{
    public class StandardReportParams
    {
        /// <summary>
        /// 报告编号
        /// </summary>
        public string ReportId { get; set; }

        /// <summary>
        /// 合同编号
        /// </summary>
        public string ContractId { get; set; }

        /// <summary>
        /// 文件解压路径
        /// </summary>
        public string ZipFilesUrl { get; set; }

    }
}