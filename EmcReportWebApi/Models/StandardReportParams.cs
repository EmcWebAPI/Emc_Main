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
        public string ReportCode { get; set; }

        /// <summary>
        /// 合同编号
        /// </summary>
        public string ContractCode { get; set; }


    }
}