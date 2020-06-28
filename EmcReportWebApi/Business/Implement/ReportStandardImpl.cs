using EmcReportWebApi.Config;
using EmcReportWebApi.Utils;
using EmcReportWebApi.Models;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using EmcReportWebApi.Business.ImplWordUtil;
using EmcReportWebApi.StandardReportComponent;

namespace EmcReportWebApi.Business.Implement
{
    /// <summary>
    /// 标准报告实现类
    /// </summary>
    public class ReportStandardImpl : ReportBase, IReportStandard
    {
        /// <summary>
        /// 生成标准报告
        /// </summary>
        /// <param name="para"></param>
        /// <returns></returns>
        public ReportResult<string> CreateReportStandard(StandardReportParams para)
        {
            Task<ReportResult<string>> task = new Task<ReportResult<string>>(() => CreateReportStandardAsync(para));
            task.Start();
            //ReportResult<string> result = task.Result;
            ReportResult<string> result = SetReportResult("报告生成中......", true, para.OriginalRecord);
            return result;
        }

        private ReportResult<string> CreateReportStandardAsync(StandardReportParams para)
        {
            ReportResult<string> result;
            try
            {
                EmcConfig.SemLim.Wait();
                //计时
                TimerUtil tu = new TimerUtil(new Stopwatch());
                StandardReportInfo standardReportInfo = new StandardReportInfo(para);

                //生成报告
                StandardReportResult srr = JsonToWordStandard(standardReportInfo);
             
                result = SetReportResult(string.Format("报告生成成功,用时:" + tu.StopTimer()), true, srr.FileName);
                EmcConfig.InfoLog.Info("报告:" + result.Content + ",信息:" + result.Message);
                CallbackReqSuccess(srr.FilePath, srr.ReportCode, result.Message, para.CallbackUrl, para.OriginalRecord);

            }
            catch (Exception ex)
            {
                string message = $"报告生成失败,reportId:{para.OriginalRecord},错误信息:{ex.Message}";
                EmcConfig.ErrorLog.Error(message, ex);//设置错误信息
                result = SetReportResult(message, false, "");
                CallbackReqFail(message, para.CallbackUrl, para.OriginalRecord);
                return result;
            }
            finally
            {
                //保存参数用作排查bug
                SaveParams(para);
                EmcConfig.SemLim.Release();
            }
            return result;
        }

        /// <summary>
        /// 解析json字符串
        /// </summary>
        /// <param name="standardReportInfo">报告信息</param>
        /// <returns></returns>
        //public StandardReportResult JsonToWordStandard(string reportId, JObject mainObj, string reportFilesPath)
        public StandardReportResult JsonToWordStandard(StandardReportInfo standardReportInfo)
        {
            try
            {
                var mainObj = standardReportInfo.ReportJsonObjectForWord;

                //生成报告
                using (ReportStandardHandleWord wordUtil = new ReportStandardHandleWord(standardReportInfo.OutFileFullName, standardReportInfo.TemplateFileFullName))
                {
                    //写首页内容
                    var firstPageInfo = standardReportInfo.ReportFirstPage;
                    firstPageInfo.WriteFirstPage(wordUtil);
                    //先画附表再画标准内容
                    //附表测试数据
                    if (mainObj["attach"] != null && !mainObj["attach"].ToString().Equals(""))
                    {
                        JArray attachArray = (JArray)mainObj["attach"];
                        AddAttachTable(wordUtil, attachArray, "standard");
                    }

                    if (mainObj["standard"] != null && !mainObj["standard"].ToString().Equals(""))
                    {
                        //标准内容
                        JArray standardArray = (JArray)mainObj["standard"];
                        wordUtil.TableSplit(standardArray, "standard", firstPageInfo.ContractDataInfo.ColSpan ?? 0);
                        //添加续
                        //wordUtil.TableSplit("standard");
                    }

                    //样品图片
                    if (mainObj["yptp"] != null && !mainObj["yptp"].ToString().Equals(""))
                    {
                        JArray yptp = (JArray)mainObj["yptp"];
                        if (yptp.Count > 0)
                            InsertImageToWordYptp(wordUtil, yptp, standardReportInfo.ReportFilesPath);
                        else
                        {
                            wordUtil.RemovePhotoTable("photo");
                        }
                    }
                    else
                    {
                        wordUtil.RemovePhotoTable("photo");
                    }

                    if (mainObj["standard"] != null && !mainObj["standard"].ToString().Equals(""))
                        wordUtil.TableSplit("standard", mainObj["yptp"] != null && !mainObj["yptp"].ToString().Equals("") && ((JArray)mainObj["yptp"]).Count > 0);

                    //替换页眉内容
                    standardReportInfo.HandleReportHeader(wordUtil);

                }
                //删除中间件文件夹
                standardReportInfo.DeleteTemplateMiddleDirectory();

                return standardReportInfo.StandardReportResultInfo;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw e;
            }
           
        }

        private void AddAttachTable(ReportStandardHandleWord wordUtil, JArray array, string bookmark)
        {
            for (int i = array.Count - 1; i >= 0; i--)
            {
                string title = array[i]["header"].ToString();
                JArray attachList = (JArray)array[i]["list"];
                wordUtil.AddAttachTable(title, attachList, bookmark);
            }
        }

        /// <summary>
        /// 照片和说明
        /// </summary>
        private void InsertImageToWordYptp(ReportStandardHandleWord wordUtil, JArray array, string reportFilesPath)
        {
            List<string> list = new List<string>();
            foreach (var jToken in array)
            {
                var item = (JObject) jToken;
                list.Add(reportFilesPath + "\\" + item["fileName"] + "," + item["content"]);
            }
            wordUtil.InsertPhotoToWord(list, "photo");
        }

        /// <summary>
        /// 模拟带参数的表单上传
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="reportCode">报告编号</param>
        /// <param name="message">生成报告消息</param>
        /// <param name="url">请求路径</param>
        /// <param name="original">原始数据</param>
        /// <returns></returns>
        private string CallbackReqSuccess(string filePath, string reportCode, string message, string url, string original)
        {
            string fileName = reportCode + ".docx";

            try
            {
                byte[] fileContentByte = new byte[1024]; // 文件内容二进制

                #region 将文件转成二进制

                FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                fileContentByte = new byte[fs.Length]; // 二进制文件
                fs.Read(fileContentByte, 0, Convert.ToInt32(fs.Length));
                fs.Close();

                #endregion


                #region 定义请求体中的内容 并转成二进制

                string boundary = "cb";
                string Enter = "\r\n";

                string fileContentStr = "--" + boundary + Enter
                        + "Content-Type:application/octet-stream" + Enter
                        + "Content-Disposition: form-data; name=\"file\"; filename=\"" + fileName + "\"" + Enter + Enter;

                string reportIdStr = Enter + "--" + boundary + Enter
                      + "Content-Disposition: form-data; name=\"original\"" + Enter + Enter
                      + original;

                string statusStr = Enter + "--" + boundary + Enter
                        + "Content-Disposition: form-data; name=\"status\"" + Enter + Enter
                        + true.ToString();

                string messageStr = Enter + "--" + boundary + Enter
                       + "Content-Disposition: form-data; name=\"message\"" + Enter + Enter
                       + message + Enter + "--" + boundary + "--";


                var fileContentStrByte = Encoding.UTF8.GetBytes(fileContentStr);//fileContent一些名称等信息的二进制（不包含文件本身）

                var reportIdStrByte = Encoding.UTF8.GetBytes(reportIdStr);//reportId所有字符串二进制

                var statusStrByte = Encoding.UTF8.GetBytes(statusStr);//contractId所有字符串二进制
                var messageStrByte = Encoding.UTF8.GetBytes(messageStr);//contractId所有字符串二进制


                #endregion

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = "POST";
                request.ContentType = "multipart/form-data;boundary=" + boundary;

                Stream myRequestStream = request.GetRequestStream();//定义请求流

                #region 将各个二进制 安顺序写入请求流

                myRequestStream.Write(fileContentStrByte, 0, fileContentStrByte.Length);
                myRequestStream.Write(fileContentByte, 0, fileContentByte.Length);

                myRequestStream.Write(reportIdStrByte, 0, reportIdStrByte.Length);

                myRequestStream.Write(statusStrByte, 0, statusStrByte.Length);
                myRequestStream.Write(messageStrByte, 0, messageStrByte.Length);

                #endregion

                HttpWebResponse response = (HttpWebResponse)request.GetResponse();//发送

                Stream myResponseStream = response.GetResponseStream();//获取返回值
                StreamReader myStreamReader = new StreamReader(myResponseStream, Encoding.GetEncoding("utf-8"));

                string retString = myStreamReader.ReadToEnd();

                myStreamReader.Close();
                myResponseStream.Close();

                return retString;
            }
            catch (Exception ex)
            {
                throw new Exception("报告生成成功但是回调函数失败,错误信息:" + ex.Message);
            }


        }

        private string CallbackReqFail(string message, string url, string original)
        {

            try
            {
                #region 定义请求体中的内容 并转成二进制

                string boundary = "cb";
                string Enter = "\r\n";

                string reportIdStr = "--" + boundary + Enter
                        + "Content-Disposition: form-data; name=\"original\"" + Enter + Enter
                        + original;

                string statusStr = Enter + "--" + boundary + Enter
                        + "Content-Disposition: form-data; name=\"status\"" + Enter + Enter
                        + false.ToString();

                string messageStr = Enter + "--" + boundary + Enter
                       + "Content-Disposition: form-data; name=\"message\"" + Enter + Enter
                       + message + Enter + "--" + boundary + "--";

                var reportIdStrByte = Encoding.UTF8.GetBytes(reportIdStr);//reportId所有字符串二进制

                var statusStrByte = Encoding.UTF8.GetBytes(statusStr);//contractId所有字符串二进制
                var messageStrByte = Encoding.UTF8.GetBytes(messageStr);//contractId所有字符串二进制

                #endregion

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = "POST";
                request.ContentType = "multipart/form-data;boundary=" + boundary;

                Stream myRequestStream = request.GetRequestStream();//定义请求流

                #region 将各个二进制 安顺序写入请求流 modelIdStr -> (fileContentStr + fileContent) -> uodateTimeStr -> encryptStr

                myRequestStream.Write(reportIdStrByte, 0, reportIdStrByte.Length);
                myRequestStream.Write(statusStrByte, 0, statusStrByte.Length);
                myRequestStream.Write(messageStrByte, 0, messageStrByte.Length);


                #endregion

                HttpWebResponse response = (HttpWebResponse)request.GetResponse();//发送

                Stream myResponseStream = response.GetResponseStream();//获取返回值
                StreamReader myStreamReader = new StreamReader(myResponseStream, Encoding.GetEncoding("utf-8"));

                string retString = myStreamReader.ReadToEnd();

                myStreamReader.Close();
                myResponseStream.Close();

                return retString;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
    }
}