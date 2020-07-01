using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http.Filters;

namespace EmcReportWebApi.Config
{
    /// <summary>
    /// 内容压缩特性
    /// </summary>
    public class CompressContentAttribute : ActionFilterAttribute
    {
        /// <summary>
        /// 方法结束时内容压缩
        /// </summary>
        /// <param name="context"></param>
        public override void OnActionExecuted(HttpActionExecutedContext context)
        {
            var acceptedEncoding = context.Response.RequestMessage.Headers.AcceptEncoding.First().Value;
            if (!acceptedEncoding.Equals("gzip", StringComparison.InvariantCultureIgnoreCase)
            && !acceptedEncoding.Equals("deflate", StringComparison.InvariantCultureIgnoreCase))
            {
                return;
            }
            context.Response.Content = new CompressContent(context.Response.Content, acceptedEncoding);
        }

    }
}