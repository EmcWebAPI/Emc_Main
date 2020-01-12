using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EmcReportWebApi.Models
{
    public class ReportResult<T>
    {
        public object Content { get; set; }

        public string Message { get; set; }

        public bool SumbitResult { get; set; }
    }
}