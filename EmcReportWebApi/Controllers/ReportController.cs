using EmcReportWebApi.Business;
using EmcReportWebApi.Business.Implement;
using EmcReportWebApi.Common;
using EmcReportWebApi.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;

namespace EmcReportWebApi.Controllers
{
    public class ReportController : ApiController
    {

        private IReport _report;
        private IReportStandard _reportStandard;

        public ReportController(IReport report, IReportStandard reportStandard)
        {
            _report = report;
            _reportStandard = reportStandard;
        }


        /// <summary>
        /// 默认输出
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public IEnumerable<string> Get()
        {
            return new string[] { "Emc", "生成报告" };
        }

        /// <summary>
        /// 生成报告
        /// </summary>
        /// <param name="para">参数</param>
        /// <returns></returns>
        [HttpPost]
        public IHttpActionResult CreateReport(ReportParams para)
        {
            ReportResult<string> result = _report.CreateReport(para);
            return Json<ReportResult<string>>(result);
        }

        /// <summary>
        /// 生成标准报告
        /// </summary>
        /// <param name="para">参数</param>
        /// <returns></returns>
        [HttpPost]
        public IHttpActionResult CreateStandardReport(StandardReportParams para)
        {
            ReportResult<string> result = _reportStandard.CreateReportStandard(para);
            return Json<ReportResult<string>>(result);
        }

        /// <summary>
        /// 下载文件
        /// </summary>
        [HttpPost]
        public IHttpActionResult DownloadFiles(FileParams para)
        {
            try
            {
                if (para == null)
                {
                    throw new Exception("参数为null");
                }
                string fileName = para.FileName;
                var browser = String.Empty;
                if (HttpContext.Current.Request.UserAgent != null)
                {
                    browser = HttpContext.Current.Request.UserAgent.ToUpper();
                }
                string fileFullName = string.Format(@"{0}Files\OutPut\{1}", EmcConfig.CurrRoot, fileName);
                if (!FileUtil.FileExists(fileFullName))
                {
                    throw new Exception(string.Format("文件{0},不存在", fileName));
                }

                HttpResponseMessage httpResponseMessage = new HttpResponseMessage(HttpStatusCode.OK);
                FileStream fileStream = File.OpenRead(fileFullName);
                httpResponseMessage.Content = new StreamContent(fileStream);
                httpResponseMessage.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                httpResponseMessage.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                {
                    FileName =
                        browser.Contains("FIREFOX")
                            ? Path.GetFileName(fileFullName)
                            : HttpUtility.UrlEncode(Path.GetFileName(fileFullName))
                    //FileName = HttpUtility.UrlEncode(Path.GetFileName(filePath))
                };
                EmcConfig.InfoLog.Info("下载成功:" + fileName);
                return ResponseMessage(httpResponseMessage);
            }
            catch (Exception ex)
            {
                ReportResult<string> result = new ReportResult<string>();
                EmcConfig.ErrorLog.Error(ex.Message, ex);
                result.Message = string.Format("下载失败,错误信息:{0}", ex.Message);
                result.SumbitResult = false;
                return Json<ReportResult<string>>(result);
            }
        }

    }
}
