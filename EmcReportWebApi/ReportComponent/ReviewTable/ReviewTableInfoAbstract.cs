using System.Collections.Generic;
using EmcReportWebApi.Business.ImplWordUtil;
using Newtonsoft.Json.Linq;

namespace EmcReportWebApi.ReportComponent.ReviewTable
{
    /// <summary>
    /// 审查表信息
    /// </summary>
    public abstract class ReviewTableInfoAbstract
    {
        /// <summary>
        /// 写入word信息
        /// </summary>
        /// <param name="wordUtil"></param>
        public abstract void WriteReviewTableInfo(ReportHandleWord wordUtil);

        /// <summary>
        /// 审查表路径
        /// </summary>
        public string ReviewTableFileFullName { get; set; }

        /// <summary>
        /// 审查表中
        /// </summary>
        public List<ReviewTableItemInfo> ItemInfos { get; set; }

        /// <summary>
        /// 测试测试集合
        /// </summary>
        public JArray TestEquipmentArray { get; set; }
    }
}