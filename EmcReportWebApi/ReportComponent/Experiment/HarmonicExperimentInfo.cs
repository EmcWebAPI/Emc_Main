using System.Collections.Generic;
using System.Linq;
using EmcReportWebApi.Business.ImplWordUtil;
using EmcReportWebApi.Config;
using EmcReportWebApi.ReportComponent.ExperimentData;
using Newtonsoft.Json.Linq;

namespace EmcReportWebApi.ReportComponent.Experiment
{
    /// <summary>
    /// 谐波失真实验
    /// </summary>
    public sealed class HarmonicExperimentInfo:ExperimentInfoAbstract
    {
        /// <summary>
        /// new
        /// </summary>
        /// <param name="reportInfo"></param>
        /// <param name="experimentInfo"></param>
        /// <param name="experimentName"></param>
        /// <param name="experimentJObject"></param>
        public HarmonicExperimentInfo(ReportInfo reportInfo, ExperimentInfo experimentInfo, string experimentName, JObject experimentJObject)
        {
            ExperimentInfo = experimentInfo;
            this.ReportInfo = reportInfo;
            this.ExperimentName = experimentName;
            this.ExperimentJObject = experimentJObject;
            this.ExperimentTemplateFileFullName = CreateTemplateMiddle($@"{EmcConfig.ExperimentTemplateFilePath}{ExperimentName}.docx");
            this.ExperimentDataTemplateFileFullname = CreateTemplateMiddle($@"{EmcConfig.ExperimentTemplateFilePath}RTFTemplate.docx");
            RtfTableInfo = EmcConfig.RtfTableInfos.FirstOrDefault(p => p.RtfType.Equals("谐波"));
            RtfPictureInfo = EmcConfig.RtfPictureInfos.FirstOrDefault(p => p.RtfType.Equals("谐波"));

            if (experimentJObject["sysj"] != null)
            {
                if (ExperimentDataInfos == null)
                    ExperimentDataInfos = new List<ExperimentDataInfoAbstract>();
                foreach (var item in (JArray)experimentJObject["sysj"])
                {
                    JObject experimentDataJObject = (JObject)item;
                    this.ExperimentDataInfos.Add(new HarmonicExperimentDataInfo(reportInfo, this, experimentDataJObject));
                }
            }

            if (experimentJObject["syljt"] != null)
            {
                if (this.ConnectionImages == null)
                    this.ConnectionImages = new List<ExperimentImage>();
                foreach (var item in (JArray)experimentJObject["syljt"])
                {
                    JObject image = (JObject)item;
                    this.ConnectionImages.Add(new ExperimentImage
                    {
                        Content = image["content"] != null ? image["content"].ToString() : string.Empty,
                        ImageName = item["name"].ToString(),
                        ImageFileFullName = $@"{reportInfo.ReportFilesPath}\{image["name"]}"
                    });
                }
            }

            if (experimentJObject["sybzt"] != null)
            {
                if (this.ArrangementImages == null)
                    this.ArrangementImages = new List<ExperimentImage>();
                foreach (var item in (JArray)experimentJObject["sybzt"])
                {
                    JObject image = (JObject)item;
                    this.ArrangementImages.Add(new ExperimentImage
                    {
                        Content = image["content"] != null ? image["content"].ToString() : string.Empty,
                        ImageName = image["name"].ToString(),
                        ImageFileFullName = $@"{reportInfo.ReportFilesPath}\{image["name"]}"
                    });
                }
            }
        }
        /// <summary>
        /// 写入实验信息
        /// </summary>
        /// <param name="wordUtil"></param>
        public override void WriteExperimentInfo(ReportHandleWord wordUtil)
        {
            foreach (var item in ExperimentJObject)
            {
                if (EmcConfig.ExperimentBaseInfo.Contains(item.Key))
                    wordUtil.InsertContentInBookmark(this.ExperimentTemplateFileFullName, item.Value.ToString(), item.Key, false);
            }
            int index = 0;
            foreach (var experimentDataInfo in ExperimentDataInfos)
            {
                experimentDataInfo.WriteExperimentDataInfo(wordUtil, index != 0);
                index++;
            }
            wordUtil.CopyOtherFileContentToWord(ExperimentDataTemplateFileFullname, ExperimentTemplateFileFullName, "sysj");
            wordUtil.InsertConnectionImageToTemplate(ExperimentTemplateFileFullName, ConnectionImages, "syljt", false);
            wordUtil.InsertImageToTemplate(ExperimentTemplateFileFullName, ArrangementImages, "sybzt", false);
            ExperimentInfo.NewBookmark = wordUtil.CopyOtherFileContentToWordReturnBookmark(ExperimentTemplateFileFullName, ExperimentInfo.NewBookmark, IsNeedNewBookmark);//是否需要新的书签
        }
    }
}