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
            //保存参数用作排查bug
            SaveParams(para);
            string reportId = para.ReportId;
            //计时
            Stopwatch sw = new Stopwatch();
            sw.Start();

            ReportResult<string> result = new ReportResult<string>();
            try
            {
                //获取zip文件 
                string reportFilesPath = FileUtil.CreateReportDirectory(string.Format("{0}\\Files\\ReportFiles", MyTools.CurrRoot));
                string reportZipFilesPath = string.Format("{0}\\zip{1}.zip", reportFilesPath, DateTime.Now.ToString("yyyyMMddhhmmss"));
                string zipUrl = ConfigurationManager.AppSettings["ReportFilesUrl"].ToString() + reportId + "?timestamp=" + MyTools.GetTimestamp(DateTime.Now);
                if (para.ZipFilesUrl != null && !para.ZipFilesUrl.Equals(""))
                {
                    zipUrl = para.ZipFilesUrl;
                }
                byte[] fileBytes = SyncHttpHelper.GetHttpRespponseForFile(zipUrl, reportZipFilesPath,
                int.Parse(DateTime.Now.ToString("hhmmss")));
                if (fileBytes.Length <= 0)
                {
                    result = SetReportResult<string>("请求报告文件失败", false, para.ReportId.ToString());
                    MyTools.ErrorLog.Error(string.Format("请求报告失败,报告id:{0}", para.ReportId));
                    return Json<ReportResult<string>>(result);
                }
                //解压zip文件
                ZipFileHelper.DecompressionZip(reportZipFilesPath, reportFilesPath);
                //生成报告
                string content = JsonToWord(reportId.Equals("") ? "QW2018-698" : reportId, para.JsonStr, reportFilesPath);
                sw.Stop();
                double time1 = (double)sw.ElapsedMilliseconds / 1000;
                result = SetReportResult<string>(string.Format("报告生成成功,用时:" + time1.ToString()), true, content);
                MyTools.InfoLog.Info("报告:" + result.Content + ",信息:" + result.Message);

            }
            catch (Exception ex)
            {
                MyTools.ErrorLog.Error(ex.Message, ex);//设置错误信息
                result = SetReportResult<string>(string.Format("报告生成失败,reportId:{0},错误信息:{1}", reportId, ex.Message), false, "");
                return Json<ReportResult<string>>(result);
            }

            return Json<ReportResult<string>>(result);
        }

        /// <summary>
        /// 生成标准报告
        /// </summary>
        /// <param name="para">参数</param>
        /// <returns></returns>
        [HttpPost]
        public IHttpActionResult CreateStandardReport(ReportParams para)
        {
            //保存参数用作排查bug
            SaveParams(para);
            string reportId = para.ReportId;
            Stopwatch sw = new Stopwatch();
            sw.Start();

            ReportResult<string> result = new ReportResult<string>();
            try
            {
                //获取zip文件 
                string reportFilesPath = FileUtil.CreateReportDirectory(string.Format("{0}\\Files\\ReportFiles", MyTools.CurrRoot));
                string reportZipFilesPath = string.Format("{0}\\zip{1}.zip", reportFilesPath, DateTime.Now.ToString("yyyyMMddhhmmss"));
                string zipUrl = ConfigurationManager.AppSettings["ReportFilesUrl"].ToString() + reportId + "?timestamp=" + MyTools.GetTimestamp(DateTime.Now);
                if (para.ZipFilesUrl != null && !para.ZipFilesUrl.Equals(""))
                {
                    zipUrl = para.ZipFilesUrl;
                }
                byte[] fileBytes = SyncHttpHelper.GetHttpRespponseForFile(zipUrl, reportZipFilesPath,
                int.Parse(DateTime.Now.ToString("hhmmss")));
                if (fileBytes.Length <= 0)
                {
                    result = SetReportResult<string>("请求报告文件失败", false, para.ReportId.ToString());
                    MyTools.ErrorLog.Error(string.Format("请求报告失败,报告id:{0}", para.ReportId));
                    return Json<ReportResult<string>>(result);
                }
                //解压zip文件
                ZipFileHelper.DecompressionZip(reportZipFilesPath, reportFilesPath);
                //生成报告
                string content = JsonToWordStandard(reportId.Equals("") ? "QW2018-698" : reportId, para.JsonStr, reportFilesPath);
                sw.Stop();
                double time1 = (double)sw.ElapsedMilliseconds / 1000;
                result = SetReportResult<string>(string.Format("报告生成成功,用时:" + time1.ToString()), true, content);
                MyTools.InfoLog.Info("报告:" + result.Content + ",信息:" + result.Message);

            }
            catch (Exception ex)
            {
                MyTools.ErrorLog.Error(ex.Message, ex);//设置错误信息
                result = SetReportResult<string>(string.Format("报告生成失败,reportId:{0},错误信息:{1}", reportId, ex.Message), false, "");
                return Json<ReportResult<string>>(result);
            }

            return Json<ReportResult<string>>(result);
        }

        /// <summary>
        /// 下载文件
        /// </summary>
        [HttpPost]
        public IHttpActionResult DownloadFiles(ReportParams para)
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
                string fileFullName = GetWordPath(fileName);
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
                MyTools.InfoLog.Info("下载成功:" + fileName);
                return ResponseMessage(httpResponseMessage);
            }
            catch (Exception ex)
            {
                ReportResult<string> result = new ReportResult<string>();
                MyTools.ErrorLog.Error(ex.Message, ex);
                result = SetReportResult<string>(string.Format("下载失败,错误信息:{0}", ex.Message), false, "");
                return Json<ReportResult<string>>(result);
            }
        }

        #region 私有方法

        #region 生成报表方法
        protected string JsonToWord(string reportId, string jsonStr, string reportFilesPath)
        {
            //解析json字符串
            JObject mainObj = (JObject)JsonConvert.DeserializeObject(jsonStr);
            string outfileName = string.Format("report{0}.docx", MyTools.GetTimestamp(DateTime.Now));//输出文件名称
            string outfilePth = string.Format(@"{0}\Files\OutPut\{1}", MyTools.CurrRoot, outfileName);//输出文件路径
            string filePath = string.Format(@"{0}\Files\{1}", MyTools.CurrRoot, ConfigurationManager.AppSettings["TemplateName"].ToString());//模板文件
            string middleDir = MyTools.CurrRoot + "\\Files\\TemplateMiddleware\\" + DateTime.Now.ToString("yyyyMMddhhmmss");
            filePath = CreateTemplateMiddle(middleDir, "template", filePath);
            string result = "保存成功1";
            //生成报告
            using (WordUtil wordUtil = new WordUtil(outfilePth, filePath))
            {
                //审查表 //测试数据
                string scbWord = reportFilesPath + "\\" + (string)mainObj["scbWord"];

                //首页内容 object
                JObject firstPage = (JObject)mainObj["firstPage"];
                result = InsertContentToWord(wordUtil, firstPage);
                //报告编号
                string[] reportArray = reportId.Split('-');
                string reportStr = "国医检(磁)字QW2018第698号";
                string reportYMStr = "国医检（磁）字QW2018第698号";
                if (reportArray.Length >= 2)
                {
                    reportStr = string.Format("国医检(磁)字{0}第{1}号", reportArray[0], reportArray[1]);
                    reportYMStr = string.Format("国医检（磁）字{0}第{1}号", reportArray[0], reportArray[1]);
                }
                wordUtil.InsertContentToWordByBookmark(reportStr, "reportId");

                //设置页眉

                if (!result.Equals("保存成功"))
                {
                    return result;
                }

                //受检样品描述 object  sjypms (审查表)
                GetTableFromReview(wordUtil, "sjypms", scbWord, 3, false);

                //样品构成 list ypgcList (审查表)
                GetTableFromReview(wordUtil, "ypgcList", scbWord, 4, false);

                //样品连接图 图片 connectionGraph (审查表)
                GetImageFomReview(wordUtil, "connectionGraph", scbWord, false);

                //样品运行模式 list ypyxList (审查表)
                GetTableFromReview(wordUtil, "ypyxList", scbWord, 6, false);

                //样品电缆 list ypdlList (审查表)
                GetTableFromReview(wordUtil, "ypdlList", scbWord, 7, false);

                //测试设备list cssbList 不动
                JArray cssbList = (JArray)mainObj["cssbList"];
                result = InsertListIntoTable(wordUtil, cssbList, 1, "cssblist");
                if (!result.Equals("保存成功"))
                {
                    return result;
                }

                //辅助设备 list fzsbList (审查表)
                GetTableFromReview(wordUtil, "fzsbList", scbWord, 5, true);

                //实验数据
                JArray experiment = (JArray)mainObj["experiment"];

                int experimentCount = experiment.Count;
                int k = 1;
                string newBookmark = "experiment";
                foreach (JObject item in experiment)
                {
                    //判断模板是否存在
                    if (!File.Exists(GetTemplatePath(item["name"] + ".docx")) && !item["name"].ToString().Equals("电压暂降/短时中断"))
                    {
                        MyTools.ErrorLog.Error(string.Format("{0}模板不存在", item["name"]));
                        continue;
                    }

                    if (item["name"].ToString().Equals("传导发射实验") || item["name"].ToString().Equals("传导发射"))
                        newBookmark = SetEmissionCommon(wordUtil, item, newBookmark, "CE", middleDir, reportFilesPath, 1, k != experimentCount);
                    else if (item["name"].ToString().Equals("辐射发射试验") || item["name"].ToString().Equals("辐射发射"))
                        newBookmark = SetEmissionCommon(wordUtil, item, newBookmark, "RE", middleDir, reportFilesPath, 1, k != experimentCount);
                    else if (item["name"].ToString().Equals("谐波失真"))
                        newBookmark = SetEmissionCommon(wordUtil, item, newBookmark, "谐波", middleDir, reportFilesPath, 2, k != experimentCount);
                    else if (item["name"].ToString().Equals("电压波动和闪烁"))
                        newBookmark = SetEmissionCommon(wordUtil, item, newBookmark, "波动", middleDir, reportFilesPath, 2, k != experimentCount);
                    else if (item["name"].ToString().Equals("电快速瞬变脉冲群") || item["name"].ToString().Equals("电压暂降和短时中断") || item["name"].ToString().Contains("电压暂降"))
                        newBookmark = SetPulseEmission(wordUtil, item, newBookmark, "", middleDir, reportFilesPath, k != experimentCount);
                    else
                        newBookmark = SetEmissionCommon(wordUtil, item, newBookmark, "", middleDir, reportFilesPath, 3, k != experimentCount);
                    k++;
                }
                wordUtil.FormatCurrentWord(k);
                //样品图片
                if (mainObj["yptp"] != null && !mainObj["yptp"].ToString().Equals(""))
                {
                    JArray yptp = (JArray)mainObj["yptp"];
                    InsertImageToWordYptp(wordUtil, yptp, reportFilesPath);
                }

                //替换页眉内容
                int pageCount = wordUtil.GetDocumnetPageCount() - 1;//获取文件页数(首页不算)

                Dictionary<int, Dictionary<string, string>> replaceDic = new Dictionary<int, Dictionary<string, string>>();
                Dictionary<string, string> valuePairs = new Dictionary<string, string>();
                valuePairs.Add("reportId", reportYMStr);
                valuePairs.Add("page", pageCount.ToString());
                replaceDic.Add(3, valuePairs);//替换页眉

                wordUtil.ReplaceWritten(replaceDic);



            }
            //删除中间件文件夹
            DelectDir(middleDir);
            DelectDir(reportFilesPath);

            return outfileName;
        }

        //设置首页内容
        protected string InsertContentToWord(WordUtil wordUtil, JObject jo1)
        {
            foreach (var item in jo1)
            {
                string key = item.Key.ToString();
                string value = item.Value.ToString();
                if (key.Equals("main_wtf") || key.Equals("main_ypmc") || key.Equals("main_xhgg") || key.Equals("main_jylb"))
                {
                    value = CheckFirstPage(value);
                }
                wordUtil.InsertContentToWordByBookmark(value, key);
            }
            return "保存成功";
        }
        //首页内容特殊处理
        public string CheckFirstPage(string itemValue)
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

        //测试工具
        protected string InsertListIntoTable(WordUtil wordUtil, JArray array, int mergeColumn, string bookmark, bool isNeedNumber = true)
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

            string result = wordUtil.InsertListToTable(list, bookmark, mergeColumn, isNeedNumber);

            return result;
        }

        //从审查表中取table数据
        protected void GetTableFromReview(WordUtil wordUtil, string bookmark, string scbWordPath, int tableIndex, bool isCloseTheFile)
        {
            wordUtil.CopyTableToWord(scbWordPath, bookmark, tableIndex, isCloseTheFile);
        }

        //从审查表中取连接图
        protected void GetImageFomReview(WordUtil wordUtil, string bookmark, string scbWordPath, bool isCloseTheFile)
        {
            wordUtil.CopyImageToWord(scbWordPath, bookmark, isCloseTheFile);
        }

        /// <summary>
        /// 实验数据
        /// </summary>
        /// <param name="funType">1.传导发射实验,辐射发射实验 2.谐波失真 3.其他html表单实验</param>
        /// <returns>新建的书签供下个实验使用</returns>
        protected string SetEmissionCommon(WordUtil wordUtil, JObject jObject, string bookmark, string rtfType, string middleDir, string reportFilesPath, int funType, bool isNewBookmark)
        {
            string templateName = jObject["name"].ToString();
            string templateFullPath = CreateTemplateMiddle(middleDir, "experiment", GetTemplatePath(templateName + ".docx"));
            string sysjTemplateFilePath = CreateTemplateMiddle(middleDir, "sysj", GetTemplatePath("RTFTemplate.docx"));

            foreach (var item in jObject)
            {
                if (!item.Key.Equals("sysj") && !item.Key.Equals("name") && !item.Key.Equals("syljt") && !item.Key.Equals("sybzt"))
                    wordUtil.InsertContentInBookmark(templateFullPath, item.Value.ToString(), item.Key, false);
            }

            JArray sysj = (JArray)jObject["sysj"];

            RtfTableInfo rtfTableInfo = MyTools.RtfTableInfos.Where(p => p.RtfType == rtfType).FirstOrDefault();
            RtfPictureInfo rtfPictureInfo = MyTools.RtfPictureInfos.Where(p => p.RtfType == rtfType).FirstOrDefault();

            int startIndex = 0;
            int endIndex = 0;
            int titleRow = 0;
            string mainTitle = "";
            Dictionary<int, string> dic = new Dictionary<int, string>();
            string rtfbookmark = "";
            int imageStartIndex = 0;
            string imageBookmark = "";

            switch (funType)
            {
                case 1:
                    startIndex = rtfTableInfo.StartIndex;
                    endIndex = rtfTableInfo.EndIndex;
                    titleRow = rtfTableInfo.TitleRow;
                    mainTitle = rtfTableInfo.MainTitle;
                    dic = rtfTableInfo.ColumnInfoDic;
                    rtfbookmark = rtfTableInfo.Bookmark;

                    imageStartIndex = rtfPictureInfo.StartIndex;
                    imageBookmark = rtfPictureInfo.Bookmark;

                    break;
                case 2:
                    startIndex = rtfTableInfo.StartIndex;
                    endIndex = rtfTableInfo.EndIndex;
                    titleRow = rtfTableInfo.TitleRow;
                    mainTitle = rtfTableInfo.MainTitle;
                    dic = rtfTableInfo.ColumnInfoDic;
                    rtfbookmark = rtfTableInfo.Bookmark;
                    break;
                default:
                    break;

            }

            int i = 0;
            foreach (JObject item in sysj)
            {
                //插入实验数据信息 (画表格)

                List<string> contentList = new List<string>();
                if (item["sygdy"] != null && !item["sygdy"].ToString().Equals(""))
                    contentList.Add("试验供电电源：" + item["sygdy"].ToString());
                if (item["syplfw"] != null && !item["syplfw"].ToString().Equals(""))
                    contentList.Add("试验频率范围：" + item["syplfw"].ToString());
                if (item["ypyxms"] != null && !item["ypyxms"].ToString().Equals(""))
                    contentList.Add("样品运行模式：" + item["ypyxms"].ToString());
                if (item["mccfpl"] != null && !item["mccfpl"].ToString().Equals(""))
                    contentList.Add("脉冲重复频率（kHz）：" + item["mccfpl"].ToString());
                if (item["sycxsj"] != null && !item["sycxsj"].ToString().Equals(""))
                    contentList.Add("试验持续时间（s）：" + item["sycxsj"].ToString());
                if (item["cfpl"] != null && !item["cfpl"].ToString().Equals(""))
                    contentList.Add("重复频率（s）：" + item["cfpl"].ToString());
                if (item["cs"] != null && !item["cs"].ToString().Equals(""))
                    contentList.Add("次数（次）：" + item["cs"].ToString());
                if (item["sycfcs"] != null && !item["sycfcs"].ToString().Equals(""))
                    contentList.Add("试验重复次数（次）：" + item["sycfcs"].ToString());
                if (item["sysjjg"] != null && !item["sysjjg"].ToString().Equals(""))
                    contentList.Add("试验时间间隔（s）：" + item["sysjjg"].ToString());
                if (item["sypl"] != null && !item["sypl"].ToString().Equals(""))
                    contentList.Add("试验频率（Hz）：" + item["sypl"].ToString());

                wordUtil.CreateTableToWord(sysjTemplateFilePath, contentList, "sysj", false, i != 0);

                switch (funType)
                {
                    case 1:
                        if (item["rtf"] != null && !item["rtf"].Equals(""))
                        {
                            JArray rtf = (JArray)item["rtf"];
                            int rtfCount = rtf.Count;
                            int j = 0;
                            try
                            {
                                foreach (JObject rtfObj in (JArray)item["rtf"])
                                {
                                    //需要画表格和插入rtf内容
                                    wordUtil.CopyOtherFileTableForColByTableIndex(sysjTemplateFilePath, reportFilesPath + "\\" + rtfObj["name"].ToString(), startIndex, endIndex, dic, rtfbookmark, titleRow, mainTitle, false, true, false);

                                    wordUtil.CopyOtherFilePictureToWord(sysjTemplateFilePath, reportFilesPath + "\\" + rtfObj["name"].ToString(), imageStartIndex, imageBookmark, false, true, j == rtfCount - 1);
                                    j++;
                                }
                            }
                            catch (Exception)
                            {
                                throw new Exception(string.Format("实验:{0}rtf文件内容不正确", templateName));
                            }
                        }
                        break;
                    case 2:
                        if (item["rtf"] != null && !item["rtf"].Equals(""))
                        {
                            JArray rtf1 = (JArray)item["rtf"];
                            int rtfCount1 = rtf1.Count;
                            int k = 0;
                            try
                            {
                                foreach (JObject rtfObj in (JArray)item["rtf"])
                                {
                                    //需要画表格和插入rtf内容
                                    wordUtil.CopyOtherFileTableForColByTableIndex(sysjTemplateFilePath, reportFilesPath + "\\" + rtfObj["name"].ToString(), startIndex, endIndex, dic, rtfbookmark, titleRow, mainTitle, false, true, k == rtfCount1 - 1);
                                    k++;
                                }
                            }
                            catch (Exception)
                            {
                                throw new Exception(string.Format("实验:{0}rtf文件内容不正确", templateName));
                            }
                        }


                        break;
                    default:
                        if (item["html"] != null && !item["html"].Equals(""))
                        {
                            JArray html = (JArray)item["html"];
                            int htmlCount = html.Count;
                            int m = 0;

                            foreach (JObject rtfObj in html)
                            {
                                //生成html并将内容插入到模板中
                                string htmlstr = (string)rtfObj["table"];
                                string htmlfullname = CreateHtmlFile(htmlstr, middleDir);
                                wordUtil.CopyHtmlContentToTemplate(htmlfullname, sysjTemplateFilePath, "sysj", true, true, false);
                                m++;
                            }
                        }
                        break;
                }

                i++;
            }

            wordUtil.CopyOtherFileContentToWord(sysjTemplateFilePath, templateFullPath, "sysj", true);

            List<string> list = new List<string>();

            //插入图片
            if (jObject["syljt"] != null && !jObject["syljt"].ToString().Equals(""))
            {
                JArray syljt = (JArray)jObject["syljt"];

                foreach (JObject item in syljt)
                {
                    list.Add(reportFilesPath + "\\" + item["name"].ToString() + "," + item["content"].ToString());
                }

                wordUtil.InsertImageToTemplate(templateFullPath, list, "syljt", false);
            }

            if (jObject["sybzt"] != null && !jObject["sybzt"].ToString().Equals(""))
            {
                JArray sybzt = (JArray)jObject["sybzt"];
                list = new List<string>();
                foreach (JObject item in sybzt)
                {
                    list.Add(reportFilesPath + "\\" + item["name"].ToString() + "," + item["content"].ToString());
                }
                wordUtil.InsertImageToTemplate(templateFullPath, list, "sybzt", false);
            }

            string result = wordUtil.CopyOtherFileContentToWordReturnBookmark(templateFullPath, bookmark, isNewBookmark);

            return result;

        }

        //电快速瞬变脉冲群 电压暂降和短时中断
        protected string SetPulseEmission(WordUtil wordUtil, JObject jObject, string bookmark, string rtfType, string middleDir, string reportFilesPath, bool isNewBookmark)
        {

            string templateName = jObject["name"].ToString();
            if (templateName.Contains("电压暂降") || templateName.Contains("短时中断"))
                templateName = "电压暂降和短时中断";
            string templateFullPath = CreateTemplateMiddle(middleDir, "experiment", GetTemplatePath(templateName + ".docx"));
            string sysjTemplateFilePath = CreateTemplateMiddle(middleDir, "sysj", GetTemplatePath("RTFTemplate.docx"));

            foreach (var item in jObject)
            {
                if (!item.Key.Equals("sysj") && !item.Key.Equals("name") && !item.Key.Equals("syljt") && !item.Key.Equals("sybzt"))
                    wordUtil.InsertContentInBookmark(templateFullPath, item.Value.ToString(), item.Key, false);
            }

            JArray sysj = (JArray)jObject["sysj"];

            //交、直流电源线
            int i = 0;
            foreach (JObject item in sysj)
            {
                if ((item["sysjTitle"] != null && (item["sysjTitle"].ToString().Equals("交、直流电源线")) || item["sysjTitle"].ToString().Contains("电源线")) ||
                    (item["sysjTitle"] != null && item["sysjTitle"].ToString().Equals("电压暂降"))
                    )
                {
                    //插入实验数据信息 (画表格)
                    List<string> contentList = new List<string>();
                    if (item["sygdy"] != null && !item["sygdy"].ToString().Equals(""))
                        contentList.Add("试验供电电源：" + item["sygdy"].ToString());
                    if (item["syplfw"] != null && !item["syplfw"].ToString().Equals(""))
                        contentList.Add("试验频率范围：" + item["syplfw"].ToString());
                    if (item["ypyxms"] != null && !item["ypyxms"].ToString().Equals(""))
                        contentList.Add("样品运行模式：" + item["ypyxms"].ToString());
                    if (item["mccfpl"] != null && !item["mccfpl"].ToString().Equals(""))
                        contentList.Add("脉冲重复频率（kHz）：" + item["mccfpl"].ToString());
                    if (item["sycxsj"] != null && !item["sycxsj"].ToString().Equals(""))
                        contentList.Add("试验持续时间（s）：" + item["sycxsj"].ToString());
                    if (item["cfpl"] != null && !item["cfpl"].ToString().Equals(""))
                        contentList.Add("重复频率（s）：" + item["cfpl"].ToString());
                    if (item["cs"] != null && !item["cs"].ToString().Equals(""))
                        contentList.Add("次数（次）：" + item["cs"].ToString());
                    if (item["sycfcs"] != null && !item["sycfcs"].ToString().Equals(""))
                        contentList.Add("试验重复次数（次）：" + item["sycfcs"].ToString());
                    if (item["sysjjg"] != null && !item["sysjjg"].ToString().Equals(""))
                        contentList.Add("试验时间间隔（s）：" + item["sysjjg"].ToString());
                    if (item["sypl"] != null && !item["sypl"].ToString().Equals(""))
                        contentList.Add("试验频率（Hz）：" + item["sypl"].ToString());


                    wordUtil.CreateTableToWord(sysjTemplateFilePath, contentList, "sysj", false, i != 0);

                    JArray html = (JArray)item["html"];
                    int htmlCount = html.Count;
                    int m = 0;

                    foreach (JObject rtfObj in html)
                    {
                        //生成html并将内容插入到模板中
                        string htmlstr = (string)rtfObj["table"];
                        string htmlfullname = CreateHtmlFile(htmlstr, middleDir);
                        wordUtil.CopyHtmlContentToTemplate(htmlfullname, sysjTemplateFilePath, "sysj", true, true, false);
                        m++;
                    }
                }
            }
            wordUtil.CopyOtherFileContentToWord(sysjTemplateFilePath, templateFullPath, "sysj1", true);

            //信号电缆和互连电缆
            int j = 0;
            foreach (JObject item in sysj)
            {
                if ((item["sysjTitle"] != null && (item["sysjTitle"].ToString().Equals("信号电缆和互连电缆") || item["sysjTitle"].ToString().Contains("电缆"))) ||
                    (item["sysjTitle"] != null && item["sysjTitle"].ToString().Equals("短时中断"))
                    )
                {
                    //插入实验数据信息 (画表格)
                    List<string> contentList = new List<string>();
                    if (item["sygdy"] != null && !item["sygdy"].ToString().Equals(""))
                        contentList.Add("试验供电电源：" + item["sygdy"].ToString());
                    if (item["syplfw"] != null && !item["syplfw"].ToString().Equals(""))
                        contentList.Add("试验频率范围：" + item["syplfw"].ToString());
                    if (item["ypyxms"] != null && !item["ypyxms"].ToString().Equals(""))
                        contentList.Add("样品运行模式：" + item["ypyxms"].ToString());
                    if (item["mccfpl"] != null && !item["mccfpl"].ToString().Equals(""))
                        contentList.Add("脉冲重复频率（kHz）：" + item["mccfpl"].ToString());
                    if (item["sycxsj"] != null && !item["sycxsj"].ToString().Equals(""))
                        contentList.Add("试验持续时间（s）：" + item["sycxsj"].ToString());
                    if (item["cfpl"] != null && !item["cfpl"].ToString().Equals(""))
                        contentList.Add("重复频率（s）：" + item["cfpl"].ToString());
                    if (item["cs"] != null && !item["cs"].ToString().Equals(""))
                        contentList.Add("次数（次）：" + item["cs"].ToString());
                    if (item["sycfcs"] != null && !item["sycfcs"].ToString().Equals(""))
                        contentList.Add("试验重复次数（次）：" + item["sycfcs"].ToString());
                    if (item["sysjjg"] != null && !item["sysjjg"].ToString().Equals(""))
                        contentList.Add("试验时间间隔（s）：" + item["sysjjg"].ToString());
                    if (item["sypl"] != null && !item["sypl"].ToString().Equals(""))
                        contentList.Add("试验频率（Hz）：" + item["sypl"].ToString());


                    wordUtil.CreateTableToWord(sysjTemplateFilePath, contentList, "sysj", false, j != 0);

                    JArray html = (JArray)item["html"];
                    int htmlCount = html.Count;
                    int m = 0;

                    foreach (JObject rtfObj in html)
                    {
                        //生成html并将内容插入到模板中
                        string htmlstr = (string)rtfObj["table"];
                        string htmlfullname = CreateHtmlFile(htmlstr, middleDir);
                        wordUtil.CopyHtmlContentToTemplate(htmlfullname, sysjTemplateFilePath, "sysj", true, true, false);
                        m++;
                    }
                }
            }
            wordUtil.CopyOtherFileContentToWord(sysjTemplateFilePath, templateFullPath, "sysj2", true);

            List<string> list = new List<string>();

            //插入图片
            if (jObject["syljt"] != null && !jObject["syljt"].ToString().Equals(""))
            {
                JArray syljt = (JArray)jObject["syljt"];

                foreach (JObject item in syljt)
                {
                    list.Add(reportFilesPath + "\\" + item["name"].ToString() + "," + item["content"].ToString());
                }

                wordUtil.InsertImageToTemplate(templateFullPath, list, "syljt", false);
            }

            if (jObject["sybzt"] != null && !jObject["sybzt"].ToString().Equals(""))
            {
                JArray sybzt = (JArray)jObject["sybzt"];
                list = new List<string>();
                foreach (JObject item in sybzt)
                {
                    list.Add(reportFilesPath + "\\" + item["name"].ToString() + "," + item["content"].ToString());
                }
                wordUtil.InsertImageToTemplate(templateFullPath, list, "sybzt", false);
            }

            string result = wordUtil.CopyOtherFileContentToWordReturnBookmark(templateFullPath, bookmark, isNewBookmark);

            return result;
        }

        //样品图片
        protected string InsertImageToWordYptp(WordUtil wordUtil, JArray array, string reportFilesPath)
        {
            List<string> list = new List<string>();
            foreach (JObject item in array)
            {
                list.Add(reportFilesPath + "\\" + item["fileName"].ToString() + "," + item["content"].ToString());
            }
            return wordUtil.InsertImageToWord2(list, "yptp");
        }
        #endregion

        #region 生成标准报表方法
        protected string JsonToWordStandard(string reportId, string jsonStr, string reportFilesPath)
        {
            //解析json字符串
            JObject mainObj = (JObject)JsonConvert.DeserializeObject(jsonStr);
            string outfileName = string.Format("report2{0}.docx", MyTools.GetTimestamp(DateTime.Now));//输出文件名称
            string outfilePth = string.Format(@"{0}\Files\OutPut\{1}", MyTools.CurrRoot, outfileName);//输出文件路径
            string filePath = string.Format(@"{0}\Files\{1}", MyTools.CurrRoot, ConfigurationManager.AppSettings["StandardTemplateName"].ToString());//模板文件

            string middleDir = MyTools.CurrRoot + "\\Files\\TemplateMiddleware\\" + DateTime.Now.ToString("yyyyMMddhhmmss");
            filePath = CreateTemplateMiddle(middleDir, "template", filePath);
            string result = "保存成功1";
            //生成报告
            using (WordUtil wordUtil = new WordUtil(outfilePth, filePath))
            {
                //首页内容 object
                JObject firstPage = (JObject)mainObj["firstPage"];
                result = InsertContentToWord(wordUtil, firstPage);

                if (!result.Equals("保存成功"))
                {
                    return result;
                }
                //报告编号
                string[] reportArray = reportId.Split('-');
                string reportStr = "国医检(磁)字QW2018第698号";
                if (reportArray.Length >= 2)
                {
                    reportStr = string.Format("国医检(磁)字{0}第{1}号", reportArray[0], reportArray[1]);
                }
                wordUtil.InsertContentToWordByBookmark(reportStr, "reportId");

                //标准内容

                JArray standardArray = (JArray)mainObj["standard"];

                wordUtil.TableSplit(standardArray, "standard");


                //替换页眉内容
                int pageCount = wordUtil.GetDocumnetPageCount() - 1;//获取文件页数(首页不算)

                Dictionary<int, Dictionary<string, string>> replaceDic = new Dictionary<int, Dictionary<string, string>>();
                Dictionary<string, string> valuePairs = new Dictionary<string, string>();
                valuePairs.Add("reportId", reportStr);
                valuePairs.Add("page", pageCount.ToString());
                replaceDic.Add(3, valuePairs);//替换页眉

                wordUtil.ReplaceWritten(replaceDic);

            }
            //删除中间件文件夹
            DelectDir(middleDir);
            DelectDir(reportFilesPath);

            return outfileName;
        }
        #endregion

        /// <summary>
        /// 创建模板中间件
        /// </summary>
        protected string CreateTemplateMiddle(string dir, string template, string filePath)
        {

            string dateStr = DateTime.Now.ToString("yyyyMMddhhmmss");
            string fileName = template + dateStr + ".docx";
            DirectoryInfo di = new DirectoryInfo(dir);
            if (!di.Exists) { di.Create(); }

            string htmlpath = dir + "\\" + fileName;
            FileInfo file = new FileInfo(filePath);
            if (File.Exists(filePath))
            {
                file.CopyTo(htmlpath);
                return htmlpath;
            }
            else
            {
                return "模板不存在";
            }

        }

        /// <summary>
        /// 删除模板中间件
        /// </summary>
        public void DelectDir(string srcPath)
        {
            try
            {
                DirectoryInfo dir = new DirectoryInfo(srcPath);
                FileSystemInfo[] fileinfo = dir.GetFileSystemInfos();  //返回目录中所有文件和子目录
                foreach (FileSystemInfo i in fileinfo)
                {
                    if (i is DirectoryInfo)            //判断是否文件夹
                    {
                        DirectoryInfo subdir = new DirectoryInfo(i.FullName);
                        subdir.Delete(true);          //删除子目录和文件
                    }
                    else
                    {
                        //如果 使用了 streamreader 在删除前 必须先关闭流 ，否则无法删除 sr.close();
                        File.Delete(i.FullName);      //删除指定文件
                    }
                }
                Directory.Delete(srcPath);
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// html字符串转word
        /// </summary>

        protected string CreateHtmlFile(string htmlStr, string dirPath)
        {
            string dateStr = DateTime.Now.ToString("yyyyMMddHHmmss");
            string htmlpath = dirPath + "\\reportHtml" + dateStr + ".html";
            FileStream fs = new FileStream(htmlpath, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);
            sw.Write(htmlStr);
            sw.Close();
            sw.Dispose();
            fs.Close();
            fs.Dispose();
            return htmlpath;
        }
        /// <summary>
        /// 保存参数文件
        /// </summary>
        protected void SaveParams(ReportParams para)
        {
            string dateStr = DateTime.Now.ToString("yyyyMMddHHmmss");
            string txtPath = string.Format("{0}Log\\Params\\{1}.txt", MyTools.CurrRoot, dateStr);
            if (!System.IO.File.Exists(txtPath))
            {
                //没有则创建这个文件
                FileStream fs1 = new FileStream(txtPath, FileMode.Create, FileAccess.Write);//创建写入文件      
                StreamWriter sw = new StreamWriter(fs1);
                sw.WriteLine("ReportId:" + para.ReportId);
                sw.WriteLine("ZipFilesUrl:" + para.ZipFilesUrl);
                sw.WriteLine("JsonStr:" + para.JsonStr);
                sw.Close();
                fs1.Close();
            }
        }

        /// <summary>
        /// 返回结果参数
        /// </summary>
        protected ReportResult<T> SetReportResult<T>(string message, bool submitResult, T content)
        {
            Type type = content.GetType();
            ReportResult<T> reportResult = new ReportResult<T>();
            reportResult.Message = message;
            reportResult.SumbitResult = submitResult;
            reportResult.Content = content;
            return reportResult;
        }

        /// <summary>
        /// 获取报告文件路径
        /// </summary>
        protected string GetWordPath(string fileName)
        {
            return string.Format(@"{0}\Files\OutPut\{1}", MyTools.CurrRoot, fileName); ;
        }

        /// <summary>
        /// 获取模板路径
        /// </summary>
        protected string GetTemplatePath(string fileName)
        {
            return string.Format(@"{0}\Files\ExperimentTemplate\{1}", MyTools.CurrRoot, fileName);
        }
        #endregion
    }
}
