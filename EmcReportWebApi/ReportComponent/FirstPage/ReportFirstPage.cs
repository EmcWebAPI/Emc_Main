using System;
using EmcReportWebApi.Business.ImplWordUtil;
using EmcReportWebApi.Utils;
using Newtonsoft.Json.Linq;

namespace EmcReportWebApi.ReportComponent.FirstPage
{
    /// <summary>
    /// 报告首页信息
    /// </summary>
    public class ReportFirstPage:ReportFirstPageAbstract
    {
        private readonly string _reportId;

        /// <summary>
        /// 获取首页信息
        /// </summary>
        /// <param name="reportJsonObjectForWord"></param>
        /// <param name="reportId"></param>
        public ReportFirstPage(JObject reportJsonObjectForWord,string reportId)
        {
            _reportId = reportId;
            this.FirstPageObject = (JObject)(reportJsonObjectForWord["firstPage"] ?? throw new Exception("合同信息不能为null"));
            this.SetReportCode();
        }
        /// <summary>
        /// 写入首页内容
        /// </summary>
        /// <param name="wordUtil"></param>
        public override void WriteFirstPage(ReportHandleWord wordUtil)
        {
            foreach (var item in FirstPageObject)
            {
                string key = item.Key;
                string value = item.Value.ToString();
                if (key.Equals("main_wtf") || key.Equals("main_ypmc") || key.Equals("main_xhgg") || key.Equals("main_jylb"))
                {
                    value = CheckFirstPage(value);
                }
                wordUtil.InsertContentToWordByBookmark(value, key);
            }
            //报告编号
            wordUtil.InsertContentToWordByBookmark(ReportCode, ReportCodeBookmark);
        }
        private string CheckFirstPage(string itemValue)
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

        private void SetReportCode()
        {
            string[] reportArray = _reportId.Split('-');

            ReportCode = reportArray.Length >= 2 ? $"国医检(磁)字{reportArray[0]}第{reportArray[1]}号" : "国医检(磁)字QW2018第698号";
            ReportYmCode = reportArray.Length >= 2? $"国医检（磁）字{reportArray[0]}第{reportArray[1]}号": "国医检（磁）字QW2018第698号";
        }
    }
}