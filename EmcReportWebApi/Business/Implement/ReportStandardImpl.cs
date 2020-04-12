using EmcReportWebApi.Common;
using EmcReportWebApi.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Web;

namespace EmcReportWebApi.Business.Implement
{
    public class ReportStandardImpl:ReportBase,IReportStandard
    {
        /// <summary>
        /// 生成报告
        /// </summary>
        /// <param name="para"></param>
        /// <returns></returns>
        public ReportResult<string> CreateReportStandard(ReportParams para)
        {
            //计时
            Stopwatch sw = new Stopwatch();
            sw.Start();
            //保存参数用作排查bug
            SaveParams(para);
            ReportResult<string> result = new ReportResult<string>();
            try
            {
                string reportId = para.ReportId;
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
                    return result;
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
                result = SetReportResult<string>(string.Format("报告生成失败,reportId:{0},错误信息:{1}", para.ReportId, ex.Message), false, "");
                return result;
            }
            return result;
        }

        public string JsonToWordStandard(string reportId, string jsonStr, string reportFilesPath)
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
    }
}