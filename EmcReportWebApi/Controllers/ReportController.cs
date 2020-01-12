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
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;

namespace EmcReportWebApi.Controllers
{
    public class ReportController : ApiController
    {

        [HttpGet]
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        /// <summary>
        /// 上传文件
        /// </summary>
        [HttpPost]
        public ReportResult<string> UploadFiles()
        {
            ReportResult<string> result = new ReportResult<string>();
            HttpFileCollection filelist = HttpContext.Current.Request.Files;
            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            if (filelist != null && filelist.Count > 0)
            {
                for (int i = 0; i < filelist.Count; i++)
                {
                    try
                    {
                        HttpPostedFile file = filelist[i];
                        string filename = file.FileName;
                        if (filename.Equals(""))
                        {
                            MyTools.ErrorLog.Error("上传失败:上传的文件信息不存在！");
                            result = SetReportResult<string>("下载失败:上传的文件信息不存在！", false, "");
                        }
                        string extendName = MyTools.FilterExtendName(filename);
                        string filePath = currRoot + "\\Files\\Upload\\";
                        string forceName = "";
                        //判断上传的文件
                        switch (extendName)
                        {
                            case ".jpg":
                            case ".png":
                                filePath = currRoot + "\\Files\\Upload\\Image\\";
                                forceName = "image";
                                break;
                            case ".rtf":
                                filePath = currRoot + "\\Files\\Upload\\Rtf\\";
                                forceName = "rtf";
                                break;
                            default:
                                filePath = currRoot + "\\Files\\Upload\\";
                                forceName = "upload";
                                break;
                        }
                        string templateFileName = forceName + DateTime.Now.ToString("yyyyMMddHHmmssfff") + extendName;

                        DirectoryInfo di = new DirectoryInfo(filePath);
                        if (!di.Exists) { di.Create(); }

                        file.SaveAs(filePath + templateFileName);
                        MyTools.InfoLog.Info(result);
                        result = SetReportResult<string>(string.Format("上传成功:{0}", filename), true, templateFileName);
                    }
                    catch (Exception ex)
                    {
                        MyTools.ErrorLog.Error(ex.Message,ex);
                        result = SetReportResult<string>(string.Format("上传文件写入失败：{0}", ex.Message), false, "");
                    }
                }
            }
            else
            {
                MyTools.ErrorLog.Error("上传失败:上传的文件信息不存在！");
                result = SetReportResult<string>("下载失败:上传的文件信息不存在！", false, "");
            }

            return result;
        }

        /// <summary>
        /// 下载文件
        /// </summary>
        [HttpPost]
        public async Task<HttpResponseMessage> DownloadFiles(ReportParams para)
        {
            string fileName = para.FileName;
            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            try
            {
                string extendName = MyTools.FilterExtendName(fileName);
                string fileFullName = "";
                //判断上传的文件
                switch (extendName)
                {
                    case ".jpg":
                    case ".png":
                        fileFullName = GetImagePath(fileName);
                        break;
                    case ".rtf":
                        fileFullName = GetRtfPath(fileName);
                        break;
                    default:
                        fileFullName = GetWordPath(fileName);
                        break;
                }
                if (!string.IsNullOrWhiteSpace(fileFullName) && File.Exists(fileFullName))
                {
                    var stream = new FileStream(fileFullName, FileMode.Open);
                    HttpResponseMessage resp = new HttpResponseMessage(HttpStatusCode.OK)
                    {
                        Content = new StreamContent(stream)
                    };
                    resp.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                    {
                        FileName = fileName
                    };
                    resp.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                    resp.Content.Headers.ContentLength = stream.Length;

                    MyTools.InfoLog.Info("下载成功");//下载记录
                    return await Task.FromResult(resp);
                }
            }
            catch (Exception ex)
            {
                MyTools.ErrorLog.Error("下载失败:" + ex.Message, ex);
                throw ex;
            }
            return new HttpResponseMessage(HttpStatusCode.NoContent);
        }
        
        [HttpGet]
        public string Get2()
        {
            Stopwatch sw = new Stopwatch();

            sw.Start();

            //string result = JsonToWord(jsonStr);
            string result = "";
            sw.Stop();
            double time1 = (double)sw.ElapsedMilliseconds / 1000;
            return result + ":" + time1.ToString();
        }


        [HttpPost]
        public string CreateReport(ReportParams para)
        {
            string miwen = MD5Helper.MD5Encrypt(ConfigurationManager.AppSettings["ReportToken"].ToString());
            if (!para.Token.Equals(miwen))
            {
                throw new Exception("无访问此方法的权限");
            }

            string jsonStr = para.JsonStr;
            string result = "创建成功";
            try
            {
                result = JsonToWord(jsonStr);
            }
            catch (Exception ex)
            {

                throw ex;
            }


            return result;
        }



        #region 私有方法

        private ReportResult<T> SetReportResult<T>(string message, bool submitResult, object content)
        {
            Type type = content.GetType();
            ReportResult<T> reportResult = new ReportResult<T>();
            reportResult.Message = message;
            reportResult.SumbitResult = submitResult;
            reportResult.Content = content;
            return reportResult;
        }

        private string GetImagePath(string fileName)
        {
            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            string imageFullFileName = string.Format(@"{0}\Files\Upload\Image\{1}", currRoot, fileName);
            return imageFullFileName;
        }

        private string GetRtfPath(string fileName)
        {
            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            string imageFullFileName = string.Format(@"{0}\Files\Upload\Rtf\{1}", currRoot, fileName);
            return imageFullFileName;
        }

        private string GetWordPath(string fileName)
        {
            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            string imageFullFileName = string.Format(@"{0}\Files\OutPut\{1}", currRoot, fileName);
            return imageFullFileName;
        }

        #region 生成报表方法
        private string JsonToWord(string jsonStr)
        {
            //解析json字符串
            JObject mainObj = (JObject)JsonConvert.DeserializeObject(jsonStr);
            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            string outfileName = string.Format("out{0}.docx", MyTools.GetTimestamp(DateTime.Now));//输出文件名称
            string outfilePth = string.Format(@"{0}\Files\OutPut\{1}", currRoot, outfileName);//输出文件路径
            string filePath = string.Format(@"{0}\Files\{1}", currRoot, "国医检(磁)字QW2018第698号模板改造.docx");//模板文件

            //生成报告
            using (WordUtil wordUtil = new WordUtil(outfilePth, filePath))
            {
                //首页内容 object

                JObject firstPage = (JObject)mainObj["firstPage"];
                InsertContentToWord(wordUtil, firstPage);

                //受检样品描述 object
                JArray ypgcList = (JArray)mainObj["ypgcList"];
                InsertListIntoTable(wordUtil, ypgcList, 2, "ypgclist");

                //样品构成 list

                //样品连接图 图片
                JArray graphList = (JArray)mainObj["connectionGraph"];
                InsertImageToWord(wordUtil, graphList, "connectionGraph");

                //样品运行模式 list

                //样品电缆 list

                //测试设备list

                //辅助设备 list

                //实验数据
            }

            return "创建成功";
        }

        //设置首页内容
        private string InsertContentToWord(WordUtil wordUtil, JObject jo1)
        {
            foreach (var item in jo1)
            {
                wordUtil.InsertContentToWord(item.Value.ToString(), item.Key);
            }
            return "保存成功";
        }

        private string InsertListIntoTable(WordUtil wordUtil, JArray array, int mergeColumn, string bookmark)
        {
            List<string> list = JarrayToList(array);

            wordUtil.InsertListToTable(list, bookmark, mergeColumn);

            return "保存成功";
        }

        private void InsertImageToWord(WordUtil wordUtil, JArray array, string bookmark)
        {
            List<string> list = new List<string>();
            foreach (JObject item in array)
            {
                string jTemp = "";
                int iTemp = 0;
                foreach (var item2 in item)
                {
                    iTemp++;
                    string tempValue = item2.Value.ToString();
                    if (iTemp == 2)
                    {
                        tempValue = GetImagePath(tempValue);
                    }
                    if (iTemp != item.Count)
                        jTemp += (tempValue + ",");
                    else
                        jTemp += tempValue;
                }
                list.Add(jTemp);
            }

            wordUtil.InsertImageToWord(list, bookmark);

        }

        private List<string> JarrayToList(JArray array)
        {
            List<string> list = new List<string>();

            foreach (JObject item in array)
            {
                string jTemp = "";
                int iTemp = 0;
                foreach (var item2 in item)
                {
                    iTemp++;
                    if (iTemp != item.Count)
                        jTemp += (item2.Value + ",");
                    else
                        jTemp += item2.Value;
                }
                list.Add(jTemp);
            }

            return list;
        }
        #endregion


        #region 测试
        private string JsonStrToJObject()
        {
            string jsonStr = "{\"FirstPage\":[{\"wtf\":\"飞利浦(中国)投资有限公司1\",\"ypmc\":\"病人监护仪1\",\"xhgg\":\"M8102A1\",\"jylb\":\"委托检验1\"},{\"wtf\":\"飞利浦(中国)投资有限公司2\",\"ypmc\":\"病人监护仪2111\",\"xhgg\":\"M8102A211\",\"jylb\":\"委托检验2111\"},{\"wtf\":\"飞利浦(中国)投资有限公司2\",\"ypmc\":\"病人监护仪2111\",\"xhgg\":\"M8102A211\",\"jylb\":\"委托检验2111\"},{\"wtf\":\"飞利浦(中国)投资有限公司1\",\"ypmc\":\"病人监护仪1\",\"xhgg\":\"M8102A1\",\"jylb\":\"委托检验1\"}]}";
            JObject jo1 = (JObject)JsonConvert.DeserializeObject(jsonStr);
            JArray firstPage = JArray.Parse(jo1["FirstPage"].ToString());
            //首页内容
            foreach (JObject item in firstPage)
            {
                //wordUtil.InsertContentInBookmark(item.Value.ToString(), item.Name);
            }

            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            string currDateStr = MyTools.GetTimestamp(DateTime.Now);
            string outfilePth = string.Format(@"{0}\Files\OutPut\output{1}.docx", currRoot, currDateStr);
            string filePath = string.Format(@"{0}\Files\{1}", currRoot, "国医检(磁)字QW2018第698号模板改造.docx");
            using (WordUtil wordUtil = new WordUtil(outfilePth, filePath))
            {
                //首页内容
                foreach (JArray item in firstPage)
                {
                    //wordUtil.InsertContentInBookmark(item.Value.ToString(), item.Name);
                }
            }
            return "转化成功";
        }

        private string InsertListIntoTable()
        {
            string jsonStr = "{\"FirstPage\":[{\"wtf\":\"飞利浦(中国)投资有限公司1\",\"ypmc\":\"病人监护仪1\",\"xhgg\":\"M8102A1\",\"jylb\":\"委托检验1\",\"t1\":\"t1\",\"t2\":\"t2\"},{\"wtf\":\"飞利浦(中国)投资有限公司2\",\"ypmc\":\"病人监护仪2111\",\"xhgg\":\"M8102A211\",\"jylb\":\"委托检验2111\",\"t1\":\"t1\",\"t2\":\"t2\"},{\"wtf\":\"飞利浦(中国)投资有限公司2\",\"ypmc\":\"病人监护仪2111\",\"xhgg\":\"M8102A211\",\"jylb\":\"委托检验2111\",\"t1\":\"t1\",\"t2\":\"t2\"},{\"wtf\":\"飞利浦(中国)投资有限公司1\",\"ypmc\":\"病人监护仪2111\",\"xhgg\":\"M8102A1\",\"jylb\":\"委托检验1\",\"t1\":\"t1\",\"t2\":\"t2\"},{\"wtf\":\"飞利浦(中国)投资有限公司1\",\"ypmc\":\"病人监护仪1\",\"xhgg\":\"M8102A1\",\"jylb\":\"委托检验1\",\"t1\":\"t1\",\"t2\":\"t2\"},{\"wtf\":\"飞利浦(中国)投资有限公司2\",\"ypmc\":\"病人监护仪2111\",\"xhgg\":\"M8102A211\",\"jylb\":\"委托检验2111\",\"t1\":\"t1\",\"t2\":\"t2\"},{\"wtf\":\"飞利浦(中国)投资有限公司2\",\"ypmc\":\"病人监护仪2111\",\"xhgg\":\"M8102A211\",\"jylb\":\"委托检验2111\",\"t1\":\"t1\",\"t2\":\"t2\"},{\"wtf\":\"飞利浦(中国)投资有限公司1\",\"ypmc\":\"病人监护仪1\",\"xhgg\":\"M8102A1\",\"jylb\":\"委托检验1\",\"t1\":\"t1\",\"t2\":\"t2\"},{\"wtf\":\"飞利浦(中国)投资有限公司1\",\"ypmc\":\"病人监护仪1\",\"xhgg\":\"M8102A1\",\"jylb\":\"委托检验1\",\"t1\":\"t1\",\"t2\":\"t2\"},{\"wtf\":\"飞利浦(中国)投资有限公司2\",\"ypmc\":\"病人监护仪2111\",\"xhgg\":\"M8102A211\",\"jylb\":\"委托检验2111\",\"t1\":\"t1\",\"t2\":\"t2\"},{\"wtf\":\"飞利浦(中国)投资有限公司2\",\"ypmc\":\"病人监护仪2111\",\"xhgg\":\"M8102A211\",\"jylb\":\"委托检验2111\",\"t1\":\"t1\",\"t2\":\"t2\"},{\"wtf\":\"飞利浦(中国)投资有限公司1\",\"ypmc\":\"病人监护仪1\",\"xhgg\":\"M8102A1\",\"jylb\":\"委托检验1\",\"t1\":\"t1\",\"t2\":\"t2\"}]}";
            JObject jo1 = (JObject)JsonConvert.DeserializeObject(jsonStr);
            JArray firstPage = (JArray)jo1["FirstPage"];

            List<string> list = new List<string>();

            foreach (JObject item in firstPage)
            {
                string jTemp = "";
                int iTemp = 0;
                foreach (var item2 in item)
                {
                    iTemp++;
                    if (iTemp != item.Count)
                        jTemp += (item2.Value + ",");
                    else
                        jTemp += item2.Value;
                }
                list.Add(jTemp);
            }

            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            string currDateStr = MyTools.GetTimestamp(DateTime.Now);
            string outfilePth = string.Format(@"{0}\Files\OutPut\output{1}.docx", currRoot, currDateStr);
            string filePath = string.Format(@"{0}\Files\{1}", currRoot, "TestListToTable.docx");
            using (WordUtil wordUtil = new WordUtil(outfilePth, filePath))
            {
                wordUtil.InsertListToTable(list, "bookmark41", 2);
            }

            return "保存成功";
        }

        private string InsertRtfIntoReport(string fileName, string htmlstr)
        {
            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            string currDateStr = MyTools.GetTimestamp(DateTime.Now);
            string outfilePth = string.Format(@"{0}\Files\OutPut\output{1}.docx", currRoot, currDateStr);
            string filePath = string.Format(@"{0}\Files\{1}", currRoot, fileName);

            MyTools.KillWordProcess();

            //string htmlfilePath = string.Format(@"{0}\Files\Html\{1}", currRoot, "testhtml.html");
            string htmlfilePath = CreateHtmlFile(htmlstr);

            string result = "创建成功";

            using (WordUtil wordUtil = new WordUtil(outfilePth, filePath))
            {

                wordUtil.CopyContentToWord(htmlfilePath, "bookmark1");

                //获取文件中的table插入到当前文件
                string rtfFileName = "ZC2018-128  生物安全柜 模式1 CE L.Rtf";
                string rtfFullName = string.Format(@"{0}\Files\检测设备产出文档\{1}", currRoot, rtfFileName);

                RtfTableInfo rtfTableInfo = MyTools.RtfTableInfos.Where(p => rtfFullName.Contains(p.RtfType)).FirstOrDefault();

                if (rtfTableInfo == null)
                {
                    throw new Exception("rtf配置文件未找到(" + rtfFullName + ")相关文件信息");
                }

                int startIndex = rtfTableInfo.StartIndex;
                Dictionary<int, string> dic = rtfTableInfo.ColumnInfoDic;
                string bookmark = rtfTableInfo.Bookmark;

                wordUtil.CopyOtherFileTableForColByTableIndex(rtfFullName, startIndex, dic, bookmark, false);

                RtfPictureInfo rtfPictureInfo = MyTools.RtfPictureInfos.Where(p => rtfFullName.Contains(p.RtfType)).FirstOrDefault();
                startIndex = rtfPictureInfo.StartIndex;
                bookmark = rtfPictureInfo.Bookmark;

                wordUtil.CopyOtherFilePictureToWord(rtfFullName, startIndex, bookmark);
            }

            MyTools.InfoLog.Info(result);
            MyTools.ErrorLog.Error("创建失败");
            return result;
        }

        private string CreateHtmlFile(string htmlStr)
        {
            string dateStr = DateTime.Now.ToString("yyyyMMddhhmmss");
            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            string htmlpath = currRoot + "Files\\Html\\reportHtml" + dateStr + ".html";
            FileStream fs = new FileStream(htmlpath, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);
            sw.Write(htmlStr);
            sw.Close();
            sw.Dispose();
            fs.Close();
            fs.Dispose();
            return htmlpath;
        }
        #endregion
        #endregion

    }
}
