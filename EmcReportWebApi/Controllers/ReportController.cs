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

        
        /// <summary>
        /// 上传文件
        /// </summary>
        [HttpPost]
        public string UploadFiles()
        {
            string result = "上传成功";
            HttpFileCollection filelist = HttpContext.Current.Request.Files;
            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            if (filelist != null && filelist.Count > 0)
            {
                for (int i = 0; i < filelist.Count; i++)
                {
                    HttpPostedFile file = filelist[i];
                    string filename = file.FileName;
                    string extendName = MyTools.FilterExtendName(filename);
                    string templateFileName = "upload" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + extendName;
                    string FilePath = currRoot + "\\Files\\Upload\\";
                    DirectoryInfo di = new DirectoryInfo(FilePath);
                    if (!di.Exists) { di.Create(); }
                    try
                    {
                        file.SaveAs(FilePath + templateFileName);
                        result = string.Format("上传成功{0}", filename);
                        MyTools.InfoLog.Info(result);
                        result = templateFileName;
                    }
                    catch (Exception ex)
                    {
                        MyTools.ErrorLog.Error(result);
                        result = "上传文件写入失败：" + ex.Message;
                    }
                }
            }
            else
            {
                result = "上传的文件信息不存在！";
            }

            return result;
        }

        /// <summary>
        /// 下载文件
        /// </summary>
        [HttpPost]
        public async Task<HttpResponseMessage> DownloadFiles(ReportParams para)
        {
            try
            {
                string fileName = para.FileName;
                string currRoot = AppDomain.CurrentDomain.BaseDirectory;
                string filePath = string.Format(@"{0}\Files\OutPut\{1}", currRoot, fileName);
                if (!string.IsNullOrWhiteSpace(filePath) && File.Exists(filePath))
                {
                    var stream = new FileStream(filePath, FileMode.Open);
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
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        [HttpGet]
        public string Get2()
        {
            Stopwatch sw = new Stopwatch();

            sw.Start();
            string jsonStr = "{\"firstPage\":{\"main_wtf\":\"飞利浦(中国)投资有限公司1\",\"main_ypmc\":\"病人监护仪1\",\"main_xhgg\":\"M8102A1\",\"main_jylb\":\"委托检验1\",\"ypmc\":\"病人监护仪\",\"sb\":\"\",\"wtf\":\"飞利浦（中国）投资有限公司\",\"wtfdz\":\"上海市静安区灵石路718号A幢\",\"scdw\":\"PhilipsMedizinSystemeBoeblingenGmbH\",\"sjdw\":\"PhilipsMedizinSystemeBoeblingenGmbH\",\"cydw\":\"\",\"cydd\":\"\",\"cyrq\":\"\",\"dyrq\":\"2018年5月8日\",\"jyxm\":\"YY0505全项目、YY0601中36、YY0667中36、YY0668中36、YY0783中36、YY0784中36\",\"jyyj\":\"YY0505-2012《医用电气设备第1-2部分：安全通用要求并列标准电磁兼容要求和试验》、YY0601-2009《医用电气设备呼吸气体监护仪的基本安全和主要性能专用要求》、YY0667-2008《医用电气设备第2-30部分：自动循环无创血压监护设备的安全和基本性能专用要求》、YY0668-2008《医用电气设备第2-49部分：多参数患者监护设备安全专用要求》、YY0783-2010《医用电气设备第2-34部分：有创血压监测设备的安全和基本性能专用要求》、YY0784-2010《医用电气设备医用脉搏血氧仪设备基本安全和主要性能专用要求》\",\"jyjl\":\"被检样品符合YY0505-2012标准要求、符合YY0601-2009标准第36章要求、符合YY0667-2008标准第36章要求、符合YY0668-2008标准第36章要求、符合YY0783-2008标准第36章要求、符合YY0784-2010标准第36章要求\",\"bz\":\"报告中“/”表示此项空白，“—”表示不适用。\",\"ypbh\":\"QW2018-0698\",\"xhgg\":\"M8102A\",\"jylb\":\"委托检验\",\"cpbhph\":\"DE65528125\",\"cydbh\":\"\",\"scrq\":\"2018-02-16\",\"ypsl\":\"1台\",\"cyjs\":\"\",\"jydd\":\"本所实验室\",\"jyrq\":\"2018年5月22日~2018年7月13日\",\"jydd\":\"本所实验室\",\"ypms\":\"见本报告第3页“1受检样品信息”。\",\"xhgghqtsm\":\"1.检测结果不包括不确定度的估算值。2.ECG附件有63个型号：M1631A、M1671A、M1984A、M1611A、M1968A、M1625A、M1639A、M1675A、M1602A、M1974A、M1601A、M1635A、M1678A、M1976A、M1672A、M1673A、M1533A、M1971A、M1973A、M1684A、M1613A、M1681A、M1558A、M1609A、M1683A、M1621A、M1674A、M1604A、M1685A、M1603A、M1619A、M1669A、M1645A、M1510A、M1500A、M1520A、M1979A、M1530A、M1557A、M1644A、M1605A、M1680A、M1537A、M1647A、M1532A、M1978A、M1615A、M1633A、M1668A、M1629A、M1663A、M1667A、M1623A、M1538A、M1665A、M1682A、M1540C、M1550C、M1560C、M1570C、989803170171、989803170181、989803143201。其电气原理和材料组成完全一致,   仅导联数与长度有所区别。本次检测了M1663A，M1978A，M1971A。SpO2附件有5个型号：M1192A、M1193A、M1194A、M1195A、M1196A，其电气原理和材料组成完全一致，仅长度和适用人群有所区别。本次检测了M1196A。CO2附件有17个型号：M2516A、M2761A、M2772A、M2751A、M2750A、M2745A、M2756A、M2757A、M2501A、M2768A、M2773A、M2741A、M2536A、M2746A、M2776A、M2777A、M1920A。其产品结构及原理均相同。本次检测了M2741A。温度探头有11个型号：21075A、21076A、21078A、M1837A、21091A、21093A、21094A、21095A、21090A、21082A、21082B。其电气原理和材料组成完全一致，仅长度和适用范围有所区别，本次检测了M21075A。袖带（含连接管）共有8个型号：M1571A、M1572A、M1573A、M1574A、M1575A、M1576A、M1598B、M1599B。其电气原理及材料组成完全一致，仅围度和连接管长度有所区别。本次检测了M1598B和M1574A。\",\"sjyp_ypmc\":\"病人监护仪\",\"sjyp_ypxh\":\"M8102A\",\"sjyp_ypbhph\":\"DE65528125\",\"sjyp_srdy\":\"AC100-240V\",\"sjyp_pl\":\"50/60Hz\",\"sjyp_edsrglhdl\":\"1.3-0.7A\",\"sjyp_dclx\":\"锂锰电池\",\"sjyp_gddy\":\"DC11.1V\",\"sjyp_ypcc\":\"199mm×146mm×89mm\"},\"ypgcList\":[{\"xh\":\"1\",\"bjmc\":\"主机\",\"bjfl\":\"\",\"xhbbh\":\"M8102A\",\"xlh\":\"DE65528125\",\"bz\":\"\"},{\"xh\":\"2\",\"bjmc\":\"模块\",\"bjfl\":\"\",\"xhbbh\":\"M3014A\",\"xlh\":\"DE45454454\",\"bz\":\"\"},{\"xh\":\"2\",\"bjmc\":\"模块\",\"bjfl\":\"\",\"xhbbh\":\"M3015B\",\"xlh\":\"DE45619953\",\"bz\":\"\"},{\"xh\":\"3\",\"bjmc\":\"外部电源配件\",\"bjfl\":\"\",\"xhbbh\":\"M8023A\",\"xlh\":\"DE21977324\",\"bz\":\"\"},{\"xh\":\"4\",\"bjmc\":\"锂电子电池\",\"bjfl\":\"\",\"xhbbh\":\"M4607A\",\"xlh\":\"\",\"bz\":\"\"},{\"xh\":\"5\",\"bjmc\":\"外接电池盒\",\"bjfl\":\"\",\"xhbbh\":\"865297/M4605A\",\"xlh\":\"865297：DE43610244\",\"bz\":\"\"}]}";

            string result = JsonToWord(jsonStr);
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
                this.InsertContentToWord(wordUtil, firstPage);

                //受检样品描述 object
                JArray ypgcList = (JArray)mainObj["ypgcList"];
                this.InsertListIntoTable(wordUtil, ypgcList,2);

                //样品构成 list

                //样品连接图 图片

                //样品运行模式 list

                //样品电缆 list

                //测试设备list

                //辅助设备 list

                //实验数据
            }

            return "创建成功";
        }

        //设置首页内容
        private string InsertContentToWord(WordUtil wordUtil,JObject jo1)
        {
            foreach (var item in jo1)
            {
                wordUtil.InsertContentToWord(item.Value.ToString(), item.Key);
            }
            return "保存成功";
        }

        private string InsertListIntoTable(WordUtil wordUtil,JArray array,int mergeColumn)
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
            
            wordUtil.InsertListToTable(list, "ypgclist", mergeColumn);

            return "保存成功";
        }
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
                wordUtil.InsertListToTable(list, "bookmark41",2);
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
