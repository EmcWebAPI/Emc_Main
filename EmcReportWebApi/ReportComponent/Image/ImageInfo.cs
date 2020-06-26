using System.Collections.Generic;
using EmcReportWebApi.Business.ImplWordUtil;
using EmcReportWebApi.ReportComponent.Experiment;
using Newtonsoft.Json.Linq;

namespace EmcReportWebApi.ReportComponent.Image
{
    /// <summary>
    /// 报告图片
    /// </summary>
    public class ImageInfo
    {
        /// <summary>
        /// new
        /// </summary>
        /// <param name="reportInfo"></param>
        /// <param name="reportJsonObjectForWord"></param>
        public ImageInfo(ReportInfo reportInfo, JObject reportJsonObjectForWord)
        {
            if (reportJsonObjectForWord["yptp"] != null)
            {
                if (this.ImageInfos == null)
                    this.ImageInfos = new List<ImageInfoAbstract>();
                foreach (var item in (JArray)reportJsonObjectForWord["yptp"])
                {
                    JObject image = (JObject)item;
                    this.ImageInfos.Add(new SampleImageInfo
                    {
                        Content = image["content"] != null ? image["content"].ToString() : string.Empty,
                        ImageName = item["fileName"].ToString(),
                        ImageFileFullName = $@"{reportInfo.ReportFilesPath}\{image["fileName"]}"
                    });
                }
            }
        }

        /// <summary>
        /// 描画图片
        /// </summary>
        /// <param name="wordUtil"></param>
        public void WriteImages(ReportHandleWord wordUtil)
        {
            wordUtil.InsertImageToWordSample(ImageInfos, "yptp");
        }

        /// <summary>
        /// 图片集合
        /// </summary>
        public IList<ImageInfoAbstract> ImageInfos { get; set; }


    }
}