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
    /// <summary>
    /// 报告操作接口
    /// </summary>
    public class ReportController : ApiController
    {

        private readonly IReport _report;
        private readonly IReportStandard _reportStandard;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="report">报告接口</param>
        /// <param name="reportStandard">标准报告接口</param>
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
            return new [] { "Emc", "生成报告" };
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
            return Json(result);
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
            return Json(result);
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
                string fileFullName = $@"{EmcConfig.CurrentRoot}Files\OutPut\{fileName}";
                if (!FileUtil.FileExists(fileFullName))
                {
                    throw new Exception($"文件{fileName},不存在");
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
                result.Message = $"下载失败,错误信息:{ex.Message}";
                result.SumbitResult = false;
                return Json(result);
            }
        }

        /// <summary>
        /// 参数:signAndIssue:1为写入签发日期
        /// 
        /// qrCodeStr:二维码字符串 不传值不生成
        /// 
        /// auditor:审核人
        /// 
        /// 结果:response.headers上获取批准人高度比例 关键字approver.vertical.proportion
        /// </summary>
        /// <returns></returns>
        public IHttpActionResult WordConvertPdf()
        {
            //下载文件到本地
            ReportResult<string> result = new ReportResult<string>();
            HttpRequest request = HttpContext.Current.Request;
            HttpFileCollection fileCollection = HttpContext.Current.Request.Files;
            string convertExtendName = ".pdf";
            if (request["extendName"] != null)
                convertExtendName = request["extendName"];
            if (fileCollection.Count > 0)
            {
                for (int i = 0; i < fileCollection.Count; i++)
                {
                    try
                    {
                        HttpPostedFile file = fileCollection[i];
                        string filename = request["fileName"] ?? file.FileName;
                        if (filename.Equals(""))
                        {
                            EmcConfig.ErrorLog.Error("上传失败:上传的文件信息不存在！");
                            result = SetReportResult("下载失败:上传的文件信息不存在！", false, "");
                            return Json(result);
                        }
                        string filePath = EmcConfig.CurrentRoot + "Files\\WordConvert\\";
                        string forceName = "upload";
                        string extendName = FilterExtendName(filename);
                        string newName = Guid.NewGuid().ToString();
                        string templateFileName = forceName + newName + extendName;
                        string outFileName = filePath + templateFileName;
                        DirectoryInfo di = new DirectoryInfo(filePath);
                        if (!di.Exists) { di.Create(); }

                        file.SaveAs(outFileName);
                        string outPictureFullName = "";
                        //如果有二维码字符串先生成二维码
                        if (request["qrCodeStr"] != null && !request["qrCodeStr"].Equals(""))
                        {
                            string qrCodeStr = request["qrCodeStr"];
                            outPictureFullName = filePath + newName + ".jpg";
                            QRCodeUtil.QRCode(outPictureFullName, qrCodeStr);
                        }

                        string convertFileName = newName + convertExtendName;
                        string convertFileFullName = filePath + convertFileName;

                        double approverHeightProportion = 0;
                        //转pdf
                        using (WordUtil wu = new WordUtil(convertFileFullName, outFileName))
                        {
                            //签发日期
                            if (request["signAndIssue"]!=null&&request["signAndIssue"].Equals("1")) {
                                string signStr = wu.InsertContentToWordByBookmark(DateTime.Now.ToString("yyyy年M月d日"), "qfrq");
                                if (signStr.Contains("未找到书签"))
                                    EmcConfig.ErrorLog.Error(filename+"错误消息:" + signStr);
                            }
                            if (!outPictureFullName.Equals("")) {
                                wu.AddPictureToWord(outPictureFullName, "main_qrcode",630f,60f, 80f, 80f);
                            }
                            //审核人
                            if (request["auditor"] != null && !request["auditor"].Equals("")) {
                                string signStr = wu.InsertContentToWordByBookmark(request["auditor"], "shry");
                                if (signStr.Contains("未找到书签"))
                                    EmcConfig.ErrorLog.Error(filename + "错误消息:" + signStr);
                            }

                            approverHeightProportion= wu.GetBookmarkHeightProportion("shry");
                        }

                        //result = SetReportResult<string>(string.Format("转化成功:{0}", templateFileName), true, convertFileName);
                        //EmcConfig.InfoLog.Info(result);

                        //写入文件流
                        if (!FileUtil.FileExists(convertFileFullName))
                        {
                            throw new Exception($"文件{convertFileName},不存在");
                        }
                        var browser = String.Empty;
                        HttpResponseMessage httpResponseMessage = new HttpResponseMessage(HttpStatusCode.OK);
                        FileStream fileStream = File.OpenRead(convertFileFullName);

                        if (approverHeightProportion != 0)
                        {
                            httpResponseMessage.Headers.Add("approver.vertical.proportion", approverHeightProportion.ToString());
                        }

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
                        result = SetReportResult($"转化成功:{filename}", true, convertFileName);
                        EmcConfig.InfoLog.Info($"{result.Message},目标文件:{result.Content}");
                        return ResponseMessage(httpResponseMessage);
                    }
                    catch (Exception ex)
                    {
                        EmcConfig.ErrorLog.Error(ex.Message, ex);
                        result = SetReportResult($"转化失败：{ex.Message}", false, "");
                    }
                }
            }
            else
            {
                EmcConfig.ErrorLog.Error("转化失败:上传的文件信息不存在！");
                result = SetReportResult("转化失败:上传的文件信息不存在！", false, "");
            }
            return Json(result);
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
            ReportResult<T> reportResult = new ReportResult<T>();
            reportResult.Message = message;
            reportResult.SumbitResult = submitResult;
            reportResult.Content = content;
            return reportResult;
        }

    }
}
