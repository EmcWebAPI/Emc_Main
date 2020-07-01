using System.Linq;
using EmcReportWebApi.Business.ImplWordUtil;
using EmcReportWebApi.Config;
using Newtonsoft.Json.Linq;

namespace EmcReportWebApi.ReportComponent.ReviewTable
{
    /// <summary>
    /// 标识文件
    /// </summary>
    public class IdentityTableInfo:ReviewTableInfoAbstract
    {
        /// <summary>
        /// new 
        /// </summary>
        /// <param name="reportJsonObjectForWord"></param>
        /// <param name="reportFilesPath"></param>
        public IdentityTableInfo(JObject reportJsonObjectForWord, string reportFilesPath)
        {
            this.ReviewTableFileFullName = reportJsonObjectForWord["bsWord"] == null ? string.Empty : $@"{reportFilesPath}\{(string)reportJsonObjectForWord["bsWord"]}";
            this.ItemInfos = EmcConfig.ReviewTableItemInfos.ToList();
        }

        /// <summary>
        /// 写入文件信息
        /// </summary>
        /// <param name="wordUtil"></param>
        public override void WriteReviewTableInfo(ReportHandleWord wordUtil)
        {
            if(this.ReviewTableFileFullName.Equals(string.Empty))
                return;
            wordUtil.CopyOtherFileContentToWord(ReviewTableFileFullName, "bsWord");
        }
    }
}