using System;
using System.Collections.Generic;
using System.Linq;
using EmcReportWebApi.Business.ImplWordUtil;
using EmcReportWebApi.Config;
using Newtonsoft.Json.Linq;

namespace EmcReportWebApi.ReportComponent.ReviewTable
{
    /// <summary>
    /// 审查表信息
    /// 1.受检样品描述 2.样品构成 3.辅助设备 4.样品运行模式 5.样品电缆 6.样品连接图
    /// </summary>
    public class ReviewTableInfo: ReviewTableInfoAbstract
    {

        /// <summary>
        /// 构造函数
        /// </summary>
        public ReviewTableInfo(JObject reportJsonObjectForWord,string reportFilesPath)
        {
            this.ReviewTableFileFullName = reportJsonObjectForWord["scbWord"]==null?throw new Exception("scbWord不能为null"): $@"{reportFilesPath}\{(string)reportJsonObjectForWord["scbWord"]}";
            this.TestEquipmentArray = reportJsonObjectForWord["cssbList"] != null
                ? (JArray) reportJsonObjectForWord["cssbList"]
                : new JArray();
            this.ItemInfos = EmcConfig.ReviewTableItemInfos.ToList();
        }
        
        /// <summary>
        /// 写入审查表信息
        /// </summary>
        /// <param name="wordUtil"></param>
        public override void WriteReviewTableInfo(ReportHandleWord wordUtil)
        {
            for (int i = 0; i < ItemInfos.Count; i++)
            {
                var itemInfo =  ItemInfos[i];
                itemInfo.IsCloseTheFile = (ItemInfos.Count - 1) == i;
                itemInfo.SetItemFromReview(wordUtil,this);
            }
        }
    }
}