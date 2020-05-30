using EmcReportWebApi.Business;
using EmcReportWebApi.Config;
using EmcReportWebApi.Models;
using EmcReportWebApi.Utils;
using System;
using System.Collections.Generic;
using System.IO;
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
        [HiddenApi]
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

        /// <summary>
        /// word转pdf 只传文件
        /// 参数:signAndIssue:1为写入签发日期|
        ///      qrCodeStr:二维码字符串 不传值不生成|
        ///      auditor:审核人
        /// </summary>
        /// <returns></returns>
        public IHttpActionResult WordConvertPdf()
        {
            //下载文件到本地
            ReportResult<string> result = new ReportResult<string>();
            HttpRequest request = HttpContext.Current.Request;
            HttpFileCollection filelist = HttpContext.Current.Request.Files;
            string convertExtendName = ".pdf";
            if (request["extendName"] != null)
                convertExtendName = request["extendName"];
            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            if (filelist != null && filelist.Count > 0)
            {
                for (int i = 0; i < filelist.Count; i++)
                {
                    try
                    {
                        HttpPostedFile file = filelist[i];
                        string filename = request["fileName"]!=null? request["fileName"].ToString():file.FileName;
                        if (filename.Equals(""))
                        {
                            EmcConfig.ErrorLog.Error("上传失败:上传的文件信息不存在！");
                            result = SetReportResult<string>("下载失败:上传的文件信息不存在！", false, "");
                        }
                        string filePath = currRoot + "Files\\WordConvert\\";
                        string forceName = "upload";
                        string extendName = FilterExtendName(filename);
                        string newName = Guid.NewGuid().ToString();
                        string templateFileName = forceName + newName + extendName;
                        string outFileName = filePath + templateFileName;
                        DirectoryInfo di = new DirectoryInfo(filePath);
                        if (!di.Exists) { di.Create(); }

                        file.SaveAs(outFileName);

                        // result = SetReportResult<string>(string.Format("上传成功:{0}", filename), true, templateFileName);
                        //EmcConfig.InfoLog.Info(result);

                        string outPictureFullName = "";
                        //如果有二维码字符串先生成二维码
                        if (request["qrCodeStr"] != null && !request["qrCodeStr"].ToString().Equals(""))
                        {
                            string qrCodeStr = request["qrCodeStr"].ToString();
                            outPictureFullName = filePath + newName + ".jpg";
                            QRCodeUtil.QRCode(outPictureFullName, qrCodeStr);
                        }

                        string convertFileName = newName + convertExtendName;
                        string convertFileFullName = filePath + convertFileName;
                        //转pdf
                        using (WordUtil wu = new WordUtil(convertFileFullName, outFileName))
                        {
                            //签发日期
                            if (request["signAndIssue"]!=null&&request["signAndIssue"].ToString().Equals("1")) {
                                string signStr = wu.InsertContentToWordByBookmark(DateTime.Now.ToString("yyyy年MM月dd日"), "qfrq");
                                if (signStr.Contains("未找到书签"))
                                    EmcConfig.ErrorLog.Error(filename+"错误消息:" + signStr);
                            }
                            if (!outPictureFullName.Equals("")) {
                                wu.AddPictureToWord(outPictureFullName, "main_qrcode",630f,60f, 80f, 80f);
                            }
                            //审核人
                            if (request["auditor"] != null && !request["auditor"].ToString().Equals("")) {
                                string signStr = wu.InsertContentToWordByBookmark(request["auditor"].ToString(), "shry");
                                if (signStr.Contains("未找到书签"))
                                    EmcConfig.ErrorLog.Error(filename + "错误消息:" + signStr);
                            }
                        }

                        //result = SetReportResult<string>(string.Format("转化成功:{0}", templateFileName), true, convertFileName);
                        //EmcConfig.InfoLog.Info(result);

                        //写入文件流
                        if (!FileUtil.FileExists(convertFileFullName))
                        {
                            throw new Exception(string.Format("文件{0},不存在", convertFileName));
                        }
                        var browser = String.Empty;
                        HttpResponseMessage httpResponseMessage = new HttpResponseMessage(HttpStatusCode.OK);
                        FileStream fileStream = File.OpenRead(convertFileFullName);
                        httpResponseMessage.Content = new StreamContent(fileStream);
                        httpResponseMessage.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                        httpResponseMessage.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                        {
                            FileName =
                                browser.Contains("FIREFOX")
                                    ? Path.GetFileName(convertFileFullName)
                                    : HttpUtility.UrlEncode(Path.GetFileName(convertFileFullName))
                            //FileName = HttpUtility.UrlEncode(Path.GetFileName(filePath))
                        };
                        result = SetReportResult<string>(string.Format("转化成功:{0}", filename), true, convertFileName);
                        EmcConfig.InfoLog.Info(string.Format("转化成功:{0}", filename));
                        return ResponseMessage(httpResponseMessage);
                    }
                    catch (Exception ex)
                    {
                        EmcConfig.ErrorLog.Error(ex.Message, ex);
                        result = SetReportResult<string>(string.Format("转化失败：{0}", ex.Message), false, "");
                    }
                }
            }
            else
            {
                EmcConfig.ErrorLog.Error("转化失败:上传的文件信息不存在！");
                result = SetReportResult<string>("转化失败:上传的文件信息不存在！", false, "");
            }
            return Json<ReportResult<string>>(result);
        }


        private static string FilterExtendName(string fileFullName)
        {
            int index = fileFullName.LastIndexOf('.');
            string extendName = fileFullName.Substring(index, fileFullName.Length - index);

            return extendName;
        }


        /// <summary>
        /// 返回结果参数
        /// </summary>
        private ReportResult<T> SetReportResult<T>(string message, bool submitResult, T content)
        {
            Type type = content.GetType();
            ReportResult<T> reportResult = new ReportResult<T>();
            reportResult.Message = message;
            reportResult.SumbitResult = submitResult;
            reportResult.Content = content;
            return reportResult;
        }

    }
}
