using EmcReportWebApi.Common;
using EmcReportWebApi.Models;
using EmcReportWebApi.Models.Repository;
using EmcReportWebApi.Repository;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Threading.Tasks;

namespace EmcReportWebApi.Business.Implement
{
    public class ReportStandardImpl : ReportBase, IReportStandard
    {
        private IReportStandardInfos _reportStandardInfos;
        public ReportStandardImpl(IReportStandardInfos reportStandardInfos) {
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
            ReportResult<string> result = task.Result;
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
                string reportId = para.ReportId;
                //获取zip文件 
                string reportFilesPath = FileUtil.CreateReportDirectory(string.Format("{0}\\Files\\ReportFiles", EmcConfig.CurrRoot));
                string reportZipFilesPath = string.Format("{0}\\zip{1}.zip", reportFilesPath, Guid.NewGuid().ToString());
                string zipUrl = ConfigurationManager.AppSettings["ReportFilesUrl"].ToString() + reportId + "?timestamp=" + EmcConfig.GetTimestamp(DateTime.Now);
                if (para.ZipFilesUrl != null && !para.ZipFilesUrl.Equals(""))
                {
                    zipUrl = para.ZipFilesUrl;
                }
                byte[] fileBytes = SyncHttpHelper.GetHttpRespponseForFile(zipUrl, reportZipFilesPath,
                int.Parse(DateTime.Now.ToString("hhmmss")));
                if (fileBytes.Length <= 0)
                {
                    result = SetReportResult<string>("请求报告文件失败", false, para.ReportId.ToString());
                    EmcConfig.ErrorLog.Error(string.Format("请求报告失败,报告id:{0}", para.ReportId));
                    return result;
                }
                //解压zip文件
                ZipFileHelper.DecompressionZip(reportZipFilesPath, reportFilesPath);
                //生成报告
                //string content = JsonToWordStandard(reportId.Equals("") ? "QW2018-698" : para.JsonStr, reportFilesPath);
                string content = JsonToWordStandardNew(reportId.Equals("") ? "QW2018-698" : reportId, para.ContractId, reportFilesPath);
                sw.Stop();
                double time1 = (double)sw.ElapsedMilliseconds / 1000;
                result = SetReportResult<string>(string.Format("报告生成成功,用时:" + time1.ToString()), true, content);
                EmcConfig.InfoLog.Info("报告:" + result.Content + ",信息:" + result.Message);

            }
            catch (Exception ex)
            {
                EmcConfig.ErrorLog.Error(ex.Message, ex);//设置错误信息
                result = SetReportResult<string>(string.Format("报告生成失败,reportId:{0},错误信息:{1}", para.ReportId, ex.Message), false, "");
                return result;
            }
            finally
            {
                EmcConfig.SemLim.Release();
            }
            return result;
        }

        /// <summary>
        /// 解析json字符串(测试用,方法通了移除)
        /// </summary>
        /// <param name="reportId">报告编号</param>
        /// <param name="jsonStr">需解析的json字符串</param>
        /// <param name="reportFilesPath">解压出的报告文件路径</param>
        /// <returns></returns>
        public string JsonToWordStandard(string reportId, string jsonStr, string reportFilesPath)
        {
            //解析json字符串
            JObject mainObj = (JObject)JsonConvert.DeserializeObject(jsonStr);
            string outfileName = string.Format("StandardReport{0}.docx", Guid.NewGuid().ToString());//输出文件名称
            string outfilePth = string.Format(@"{0}\Files\OutPut\{1}", EmcConfig.CurrRoot, outfileName);//输出文件路径
            string filePath = string.Format(@"{0}\Files\{1}", EmcConfig.CurrRoot, ConfigurationManager.AppSettings["StandardTemplateName"].ToString());//模板文件

            string middleDir = EmcConfig.CurrRoot + "\\Files\\TemplateMiddleware\\" + Guid.NewGuid().ToString();
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


                //先画附表再画标准内容
                //附表测试数据
                JArray attachArray = (JArray)mainObj["attach"];
                this.AddAttachTable(wordUtil, "附表202", attachArray, "standard");

                //标准内容
                JArray standardArray = (JArray)mainObj["standard"];
                wordUtil.TableSplit(standardArray, "standard");

                
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

        /// <summary>
        /// 解析生成报告
        /// </summary>
        /// <param name="reportId"></param>
        /// <param name="contractId"></param>
        /// <param name="reportFilesPath"></param>
        /// <returns></returns>
        public string JsonToWordStandardNew(string reportId, string contractId,string reportFilesPath)
        {

            //获取合同信息
            ContractInfo contractInfo = _reportStandardInfos.GetContract(contractId);
            JObject firstPage = ContractInfoToJObject(contractInfo);

            //获取报告标准内容
            //JArray standardArray = (JArray)mainObj["standard"];

            //获取附表内容
            //JArray attachArray = (JArray)mainObj["attach"];

            //获取报告图片内容
            //JArray imageArray=

            //数据库报告文件相对内容

            string outfileName = string.Format("StandardReport{0}.docx", Guid.NewGuid().ToString());//输出文件名称
            string outfilePth = string.Format(@"{0}\Files\OutPut\{1}", EmcConfig.CurrRoot, outfileName);//输出文件路径
            string filePath = string.Format(@"{0}\Files\{1}", EmcConfig.CurrRoot, ConfigurationManager.AppSettings["StandardTemplateName"].ToString());//模板文件

            string middleDir = EmcConfig.CurrRoot + "\\Files\\TemplateMiddleware\\" + Guid.NewGuid().ToString();
            filePath = CreateTemplateMiddle(middleDir, "template", filePath);
            string result = "保存成功1";
            //生成报告
            using (WordUtil wordUtil = new WordUtil(outfilePth, filePath))
            {
                //首页内容 object
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

                //wordUtil.TableSplit(standardArray, "standard");

                //样品图片
                //if (mainObj["yptp"] != null && !mainObj["yptp"].ToString().Equals(""))
                //{
                //    JArray yptp = (JArray)mainObj["yptp"];
                //    InsertImageToWordYptp(wordUtil, yptp, reportFilesPath);
                //}

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


        private string AddAttachTable(WordUtil wordUtil,string title,JArray array,string bookmark) {
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

           return wordUtil.AddAttachTable(title, list, bookmark);
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
        /// 报告标准内容
        /// </summary>
        /// <returns></returns>
        private JArray GetStandardToJArray() {
            JArray ja = new JArray();

            return new JArray();
        }
        
    }
}