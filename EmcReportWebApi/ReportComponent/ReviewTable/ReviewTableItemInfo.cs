using EmcReportWebApi.Business.ImplWordUtil;
using Newtonsoft.Json.Linq;

namespace EmcReportWebApi.ReportComponent.ReviewTable
{
    /// <summary>
    /// 审查表内信息
    /// </summary>
    public class ReviewTableItemInfo
    {
        /// <summary>
        /// 写入word
        /// </summary>
        /// <param name="wordUtil"></param>
        /// <param name="reviewTableInfo"></param>
        public void SetItemFromReview(ReportHandleWord wordUtil, ReviewTableInfo reviewTableInfo)
        {
            switch (ItemType)
            {
                case ReviewTableItemType.Table:
                    wordUtil.CopyTableToWord(reviewTableInfo.ReviewTableFileFullName, ReportBookmark, ReviewTableItemIndex, IsCloseTheFile);
                    break;
                case ReviewTableItemType.Image:
                    wordUtil.CopyImageToWord(reviewTableInfo.ReviewTableFileFullName, ReportBookmark, IsCloseTheFile);
                    break;
                case ReviewTableItemType.TestEquipment:
                    wordUtil.InsertListToTable(reviewTableInfo.TestEquipmentArray, ReportBookmark);
                    break;
            }
        }

        /// <summary>
        /// 报告中的书签
        /// </summary>
        public string ReportBookmark { get; set; }

        /// <summary>
        /// 审查表内容序号
        /// </summary>
        public int ReviewTableItemIndex { get; set; }

        /// <summary>
        /// 审查表内容类型
        /// </summary>
        public ReviewTableItemType ItemType { get; set; }

        /// <summary>
        /// 修改内容后是否关闭审查表
        /// </summary>
        public bool IsCloseTheFile { get; set; }
    }


    /// <summary>
    /// 审查表内容类型
    /// </summary>
    public enum ReviewTableItemType
    {
        /// <summary>
        /// 表格
        /// </summary>
        Table,
        /// <summary>
        /// 图片
        /// </summary>
        Image,
        /// <summary>
        /// 测试设备
        /// </summary>
        TestEquipment
      
    }
}