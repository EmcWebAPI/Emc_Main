using System;
using System.Collections.Generic;
using System.IO;
using EmcReportWebApi.Business.ImplWordUtil;
using EmcReportWebApi.Config;
using EmcReportWebApi.ReportComponent.Experiment;
using Newtonsoft.Json.Linq;

namespace EmcReportWebApi.ReportComponent.ExperimentData
{
    /// <summary>
    /// 电快速瞬变脉冲群 电压暂降和短时中断
    /// </summary>
    public class AcDcExperimentDataInfo : ExperimentDataInfoAbstract
    {
        private readonly ReportInfo _reportInfo;
        private readonly ExperimentInfoAbstract _experimentInfo;
        /// <summary>
        /// 默认实验数据
        /// </summary>
        /// <param name="reportInfo"></param>
        /// <param name="experimentInfo"></param>
        /// <param name="experimentDataJObject"></param>
        public AcDcExperimentDataInfo(ReportInfo reportInfo, ExperimentInfoAbstract experimentInfo, JObject experimentDataJObject)
        {
            _reportInfo = reportInfo;
            _experimentInfo = experimentInfo;
            ExperimentDataJObject = experimentDataJObject;
            ExperimentDataTitle = ExperimentDataJObject["sysjTitle"] != null? ExperimentDataJObject["sysjTitle"].ToString():string.Empty;
            if (ExperimentDataTitleInfos == null)
                ExperimentDataTitleInfos = new List<string>();
            foreach (var title in EmcConfig.ExperimentDataTitleInfo)
            {
                if (ExperimentDataJObject[title.Key] != null)
                {
                    ExperimentDataTitleInfos.Add($"{title.Value}{ExperimentDataJObject[title.Key]}");
                }
            }

            this.ExperimentDataHtmlJArray = experimentDataJObject["html"] != null
                ? (JArray)experimentDataJObject["html"]
                : new JArray();
        }

        /// <summary>
        /// 重写写入实验数据
        /// </summary>
        /// <param name="wordUtil"></param>
        /// <param name="isNeedBreak"></param>
        public override void WriteExperimentDataInfo(ReportHandleWord wordUtil, bool isNeedBreak)
        {
            wordUtil.CreateTableToWord(_experimentInfo.ExperimentDataTemplateFileFullname, ExperimentDataTitleInfos, "sysj", false, isNeedBreak);
            int j = 0;
            int rtfCount = ExperimentDataHtmlJArray.Count;
            foreach (var rtf in ExperimentDataHtmlJArray)
            {
                try
                {
                    var rtfObj = (JObject)rtf;
                    string htmlStr = (string)rtfObj["table"];
                    string htmlFileFullName = this.CreateHtmlFile(htmlStr, _reportInfo.ReportFilesPath);
                    wordUtil.CopyHtmlContentToTemplate(htmlFileFullName, _experimentInfo.ExperimentDataTemplateFileFullname, "sysj", true, true, false);
                }
                catch (Exception e)
                {
                    throw new Exception($"实验:{_experimentInfo.ExperimentName}html文件内容不正确");
                }

                j++;
            }
        }

        private string CreateHtmlFile(string htmlStr, string dirPath)
        {
            string dateStr = Guid.NewGuid().ToString();
            string htmlFullPath = dirPath + "\\reportHtml" + dateStr + ".html";
            FileStream fs = new FileStream(htmlFullPath, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);
            sw.Write(htmlStr);
            sw.Close();
            sw.Dispose();
            fs.Close();
            fs.Dispose();
            return htmlFullPath;
        }
    }
}