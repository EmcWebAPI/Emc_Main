using EmcReportWebApi.Common;
using EmcReportWebApi.Models;
using EmcReportWebApi.Models.Repository;
using EmcReportWebApi.Repository;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace EmcReportWebApi.Business.Implement
{
    public class ReportStandardImpl : ReportBase, IReportStandard
    {
        private IReportStandardInfos _reportStandardInfos;
        public ReportStandardImpl(IReportStandardInfos reportStandardInfos)
        {
            _reportStandardInfos = reportStandardInfos;
        }

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
            ReportResult<string> result = SetReportResult<string>(string.Format("报告生成中......"), true, para.OriginalRecord);
            return result;
        }

        private ReportResult<string> CreateReportStandardAsync(StandardReportParams para)
        {
            ReportResult<string> result = new ReportResult<string>();
            try
            {
                EmcConfig.SemLim.Wait();
                //计时
                Stopwatch sw = new Stopwatch();
                sw.Start();
                string reportId = para.OriginalRecord;
                //获取zip文件 
                string reportFilesPath = FileUtil.CreateReportDirectory(string.Format("{0}\\Files\\ReportFiles", EmcConfig.CurrRoot));
                string reportZipFilesPath = string.Format("{0}\\zip{1}.zip", reportFilesPath, Guid.NewGuid().ToString());
                if (para.ZipFilesUrl != null && !para.ZipFilesUrl.Equals(""))
                {
                    string zipUrl = para.ZipFilesUrl;
                    byte[] fileBytes = SyncHttpHelper.GetHttpRespponseForFile(zipUrl, reportZipFilesPath,
                 int.Parse(DateTime.Now.ToString("hhmmss")));
                    if (fileBytes.Length <= 0)
                    {
                        result = SetReportResult<string>("请求报告文件失败", false, para.OriginalRecord.ToString());
                        EmcConfig.ErrorLog.Error(string.Format("请求报告失败,报告id:{0}", para.OriginalRecord));
                        return result;
                    }
                    //解压zip文件
                    ZipFileHelper.DecompressionZip(reportZipFilesPath, reportFilesPath);
                }

                //生成报告
                StandardReportResult srr = JsonToWordStandard(reportId.Equals("") ? "QW2018-698" : reportId, para.JsonObject, reportFilesPath);
                //string content = JsonToWordStandardNew(reportId.Equals("") ? "QW2018-698" : reportId, para.ContractId, reportFilesPath);
                sw.Stop();
                //报告生成时间
                double time1 = (double)sw.ElapsedMilliseconds / 1000;

                result = SetReportResult<string>(string.Format("报告生成成功,用时:" + time1.ToString()), true, srr.FileName);
                EmcConfig.InfoLog.Info("报告:" + result.Content + ",信息:" + result.Message);

                CallbackReqSuccess(srr.FilePath, srr.ReportCode, result.Message, para.CallbackUrl, para.OriginalRecord);

            }
            catch (Exception ex)
            {
                EmcConfig.ErrorLog.Error(ex.Message, ex);//设置错误信息
                string message = string.Format("报告生成失败,reportId:{0},错误信息:{1}", para.OriginalRecord, ex.Message);
                result = SetReportResult<string>(message, false, "");
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
        /// <param name="reportId">报告编号</param>
        /// <param name="mainObj">需解析的json字符串</param>
        /// <param name="reportFilesPath">解压出的报告文件路径</param>
        /// <returns></returns>
        public StandardReportResult JsonToWordStandard(string reportId, JObject mainObj, string reportFilesPath)
        {

            StandardReportResult srr = new StandardReportResult();

            //解析json字符串
            // JObject mainObj = (JObject)JsonConvert.DeserializeObject(jsonStr);
            string outfileName = string.Format("StandardReport{0}.docx", Guid.NewGuid().ToString());//输出文件名称
            string outfilePth = string.Format(@"{0}\Files\OutPut\{1}", EmcConfig.CurrRoot, outfileName);//输出文件路径
            string filePath = string.Format(@"{0}\Files\{1}", EmcConfig.CurrRoot, ConfigurationManager.AppSettings["StandardTemplateName"].ToString());//模板文件

            string middleDir = EmcConfig.CurrRoot + "\\Files\\TemplateMiddleware\\" + Guid.NewGuid().ToString();
            filePath = CreateTemplateMiddle(middleDir, "template", filePath);
            string result = "保存成功1";
            string reportStr = "";
            //生成报告
            using (WordUtil wordUtil = new WordUtil(outfilePth, filePath))
            {
                //首页内容 object
                ContractData contractInfo = mainObj["firstPage"].ToObject<ContractData>();
                JObject firstPage = ContractDataToJObject(contractInfo);
                //样品编号
                string ypbh = firstPage["ypbh"] != null ? firstPage["ypbh"].ToString() : "";
                //报告编号
                reportStr = firstPage["bgbh"] != null ? firstPage["bgbh"].ToString() : "";
                result = InsertContentToWord(wordUtil, firstPage);

                if (!result.Equals("保存成功"))
                {
                    srr.Status = false;
                    srr.Message = "报告生成失败";
                    return srr;
                }

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
                    wordUtil.TableSplit(standardArray, "standard");
                    //添加续
                    //wordUtil.TableSplit("standard");
                }



                //样品图片
                if (mainObj["yptp"] != null && !mainObj["yptp"].ToString().Equals(""))
                {
                    JArray yptp = (JArray)mainObj["yptp"];
                    if (yptp.Count > 0)
                        InsertImageToWordYptp(wordUtil, yptp, reportFilesPath);
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
                    wordUtil.TableSplit("standard");

                //替换页眉内容
                int pageCount = wordUtil.GetDocumnetPageCount() - 1;//获取文件页数(首页不算)
                Dictionary<int, Dictionary<string, string>> replaceDic = new Dictionary<int, Dictionary<string, string>>();
                Dictionary<string, string> valuePairs = new Dictionary<string, string>();
                valuePairs.Add("bgbh", reportStr);//报告编号
                valuePairs.Add("ypbh", ypbh);//样品编号
                valuePairs.Add("page", pageCount.ToString());
                replaceDic.Add(3, valuePairs);//替换页眉

                wordUtil.ReplaceWritten(replaceDic);

            }
            //using (WordUtil wordUtil = new WordUtil(outfilePth))
            //{

            //}
            //删除中间件文件夹
            DelectDir(middleDir);
            DelectDir(reportFilesPath);

            srr.FilePath = outfilePth;
            srr.ReportCode = reportStr;
            srr.FileName = outfileName;

            return srr;
        }

        //设置首页内容
        public override string InsertContentToWord(WordUtil wordUtil, JObject jo1)
        {
            foreach (var item in jo1)
            {
                string key = item.Key.ToString();
                string value = item.Value.ToString();
                if (key.Equals("main_wtf") || key.Equals("main_ypmc") || key.Equals("main_xhgg") || key.Equals("main_jylb"))
                {
                    value = CheckFirstPage(value);
                    wordUtil.InsertContentToWordByBookmark(value, key, true);
                }
                else
                    wordUtil.InsertContentToWordByBookmark(value, key);
            }
            return "保存成功";
        }
        //首页内容特殊处理
        public override string CheckFirstPage(string itemValue)
        {
            int fontCount = 38;
            int valueCount = System.Text.Encoding.Default.GetBytes(itemValue).Length;
            if (fontCount > valueCount)
            {
                int spaceCount = (fontCount - valueCount) / 2;
                for (int i = 0; i < spaceCount; i++)
                {
                    itemValue = " " + itemValue + " ";
                }
            }

            return itemValue;
        }

        private string AddAttachTable(WordUtil wordUtil, JArray array, string bookmark)
        {
            string result = "";

            for (int i = array.Count - 1; i >= 0; i--)
            {
                string title = array[i]["header"].ToString();
                JArray attachList = (JArray)array[i]["list"];
                result = wordUtil.AddAttachTable(title, attachList, bookmark);
            }
            return result;
        }

        /// <summary>
        /// 照片和说明
        /// </summary>
        private string InsertImageToWordYptp(WordUtil wordUtil, JArray array, string reportFilesPath)
        {
            List<string> list = new List<string>();
            foreach (JObject item in array)
            {
                list.Add(reportFilesPath + "\\" + item["fileName"].ToString() + "," + item["content"].ToString());
            }
            return wordUtil.InsertPhotoToWord(list, "photo");
        }

        /// <summary>
        /// 合同信息转成jobject供报告使用
        /// </summary>
        private JObject ContractInfoToJObject(ContractInfo contractInfo)
        {
            JObject jObject = new JObject();
            foreach (var item in EmcConfig.ContractToJObject)
            {
                string key = item.Key;
                string value = item.Value;

                var property = contractInfo.Data.GetType().GetProperty(value);
                string obj = (property == null || property.GetValue(contractInfo.Data, null) == null) ? "" : contractInfo.Data.GetType().GetProperty(value).GetValue(contractInfo.Data, null).ToString();
                jObject.Add(key, obj);
            }
            return jObject;
        }

        /// <summary>
        /// 合同信息转成jobject供报告使用
        /// </summary>
        private JObject ContractDataToJObject(ContractData contractData)
        {
            JObject jObject = new JObject();
            foreach (var item in EmcConfig.ContractToJObject)
            {
                string key = item.Key;
                string value = item.Value;

                var property = contractData.GetType().GetProperty(value);
                string obj = (property == null || property.GetValue(contractData, null) == null) ? "" : contractData.GetType().GetProperty(value).GetValue(contractData, null).ToString();
                jObject.Add(key, obj);
            }
            return jObject;
        }

        /// <summary>
        /// 报告标准内容
        /// </summary>
        /// <returns></returns>
        private JArray GetStandardToJArray()
        {
            JArray ja = new JArray();

            return new JArray();
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
            // string url = @"http://192.168.30.10:9081/hydra/std/readXls";
            //string url = para.CallbackUrl;
            //string reportId = para.ReportId;
            //string contractId = para.ContractId;

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