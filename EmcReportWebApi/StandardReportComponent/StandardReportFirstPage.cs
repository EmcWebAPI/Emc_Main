using System;
using EmcReportWebApi.Business.ImplWordUtil;
using EmcReportWebApi.Config;
using EmcReportWebApi.Models;
using EmcReportWebApi.ReportComponent.FirstPage;
using Newtonsoft.Json.Linq;

namespace EmcReportWebApi.StandardReportComponent
{
    /// <summary>
    /// 报告首页信息
    /// </summary>
    public class StandardReportFirstPage 
    {
        /// <summary>
        /// 获取首页信息
        /// </summary>
        /// <param name="standardReportInfo"></param>
        /// <param name="reportJsonObjectForWord"></param>
        public StandardReportFirstPage(StandardReportInfo standardReportInfo,JObject reportJsonObjectForWord)
        {
            SourceFirstPageObject = (JObject)(reportJsonObjectForWord["firstPage"] ?? throw new Exception("合同信息不能为null"));
            ContractDataInfo=SourceFirstPageObject.ToObject<ContractData>();
            FirstPageObject = ContractDataToJObject();
            ReportCode = FirstPageObject["bgbh"] != null ? FirstPageObject["bgbh"].ToString() : string.Empty;
            SampleCode = FirstPageObject["ypbh"] != null ? FirstPageObject["ypbh"].ToString() : string.Empty;
            
        }
        /// <summary>
        /// 写入首页内容
        /// </summary>
        /// <param name="wordUtil"></param>
        public void WriteFirstPage(ReportStandardHandleWord wordUtil)
        {
            try
            {
                foreach (var item in FirstPageObject)
                {
                    string key = item.Key.ToString();
                    string value = item.Value.ToString();
                    if (key.Equals("main_ypmc"))
                    {
                        string[] values = value.Split('\n');
                        if (values.Length > 1)
                        {
                            value = values[0];

                            for (int i = values.Length - 1; i <= 1; i++)
                            {
                                var tempValue = values[i];
                                wordUtil.TableAddRowForY("main_ypmc", tempValue);
                            }

                        }
                        //wordUtil.InsertContentToWordByBookmark(value, key);

                    }
                    wordUtil.InsertContentToWordByBookmark(value, key);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw e;
            }
            
        }

        /// <summary>
        /// 合同信息转成jobject供报告使用
        /// </summary>
        private JObject ContractDataToJObject()
        {
            JObject jObject = new JObject();
            foreach (var item in EmcConfig.ContractToJObject)
            {
                string key = item.Key;
                string value = item.Value;

                var property = ContractDataInfo.GetType().GetProperty(value);
                string obj = (property == null || property.GetValue(ContractDataInfo, null) == null) ? "" : ContractDataInfo.GetType().GetProperty(value)?.GetValue(ContractDataInfo, null).ToString();
                jObject.Add(key, obj);
            }
            return jObject;
        }


        /// <summary>
        /// 首页json
        /// </summary>
        public JObject FirstPageObject { get; set; }

        /// <summary>
        /// 首页源json
        /// </summary>
        public JObject SourceFirstPageObject { get; set; }

        /// <summary>
        /// 合同信息
        /// </summary>
        public ContractData ContractDataInfo { get; set; }

        /// <summary>
        /// 首页上的报告编号
        /// </summary>
        public string ReportCode { get; set; }

        /// <summary>
        /// 首页上报告编号书签
        /// </summary>
        public string ReportCodeBookmark { get; set; } = "reportId";

        /// <summary>
        /// 样品编号
        /// </summary>
        public string SampleCode { get; set; }

    }
}