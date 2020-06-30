using System;
using System.Collections.Generic;
using EmcReportWebApi.Business.ImplWordUtil;
using EmcReportWebApi.Config;
using EmcReportWebApi.ReportComponent.Experiment;
using Newtonsoft.Json.Linq;

namespace EmcReportWebApi.ReportComponent.ExperimentData
{
    /// <summary>
    /// 谐波失真 电压波动和闪烁
    /// </summary>
    public class FluctuationExperimentDataInfo : ExperimentDataInfoAbstract
    {
        private readonly ReportInfo _reportInfo;
        private readonly ExperimentInfoAbstract _experimentInfo;

        /// <summary>
        /// 实验数据信息
        /// </summary>
        /// <param name="reportInfo"></param>
        /// <param name="experimentInfo"></param>
        /// <param name="experimentDataJObject"></param>
        public FluctuationExperimentDataInfo(ReportInfo reportInfo, ExperimentInfoAbstract experimentInfo, JObject experimentDataJObject)
        {
            _reportInfo = reportInfo;
            _experimentInfo = experimentInfo;
            this.ExperimentDataJObject = experimentDataJObject;
            if (ExperimentDataTitleInfos == null)
                ExperimentDataTitleInfos = new List<string>();
            foreach (var title in EmcConfig.ExperimentDataTitleInfo)
            {
                if (ExperimentDataJObject[title.Key] != null)
                {
                    ExperimentDataTitleInfos.Add($"{title.Value}{ExperimentDataJObject[title.Key]}");
                }
            }

            this.ExperimentDataRtfJArray = experimentDataJObject["rtf"] != null
                ? (JArray)experimentDataJObject["rtf"]
                : new JArray();
        }

        /// <summary>
        /// 写入实验数据
        /// </summary>
        /// <param name="wordUtil"></param>
        /// <param name="isNeedBreak"></param>
        public override void WriteExperimentDataInfo(ReportHandleWord wordUtil, bool isNeedBreak)
        {
            wordUtil.CreateTableToWord(_experimentInfo.ExperimentDataTemplateFileFullname, ExperimentDataTitleInfos, "sysj", false, isNeedBreak);
            int j = 0;
            int rtfCount = ExperimentDataRtfJArray.Count;
            foreach (var rtf in ExperimentDataRtfJArray)
            {
                try
                {
                    var rtfObj = (JObject)rtf;
                    //需要画表格和插入rtf内容
                    wordUtil.CopyFluctuationFileTableForColByTableIndex(_experimentInfo.ExperimentDataTemplateFileFullname,
                        _reportInfo.ReportFilesPath + "\\" + rtfObj["name"].ToString(), _experimentInfo.RtfTableInfo.StartIndex,
                        _experimentInfo.RtfTableInfo.EndIndex,
                        _experimentInfo.RtfTableInfo.ColumnInfoDic,
                        _experimentInfo.RtfTableInfo.Bookmark,
                        _experimentInfo.RtfTableInfo.TitleRow,
                        _experimentInfo.RtfTableInfo.MainTitle, false, true, j == rtfCount - 1, j == rtfCount - 1);
                }
                catch (Exception)
                {
                    throw new Exception($"实验:{_experimentInfo.ExperimentName}rtf文件内容不正确");
                }

                j++;
            }
        }
    }
}