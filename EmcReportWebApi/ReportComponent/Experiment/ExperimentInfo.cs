using System.Collections.Generic;
using System.IO;
using EmcReportWebApi.Business.ImplWordUtil;
using EmcReportWebApi.Config;
using Newtonsoft.Json.Linq;

namespace EmcReportWebApi.ReportComponent.Experiment
{
    /// <summary>
    /// 实验数据整合
    /// </summary>
    public class ExperimentInfo
    {
        /// <summary>
        /// new
        /// </summary>
        public ExperimentInfo(ReportInfo reportInfo,JObject reportJsonObjectForWord)
        {
            this.ReportInfo = reportInfo;
            this.NewBookmark = "experiment";
            this.ExperimentInfosJArray = (JArray)reportJsonObjectForWord["experiment"];

            foreach (var item in ExperimentInfosJArray)
            {
                JObject experimentInfo = (JObject)item;

               string experimentName = experimentInfo["name"].ToString();
                //判断模板是否存在
                if (!File.Exists($@"{EmcConfig.ExperimentTemplateFilePath}\{experimentName}.docx")&&!experimentName.Equals("电压暂降/短时中断"))
                {
                    EmcConfig.ErrorLog.Error($"{experimentInfo["name"]}模板不存在");
                    continue;
                }

                if (ExperimentInfos == null)
                    ExperimentInfos = new List<ExperimentInfoAbstract>();
                switch (experimentName)
                {
                    case "传导发射":
                        ExperimentInfos.Add(new CeExperimentInfo(reportInfo, this, experimentName, experimentInfo));
                        break;
                    case "辐射发射":
                        ExperimentInfos.Add(new ReExperimentInfo(reportInfo, this, experimentName, experimentInfo));
                        break;
                    case "谐波失真":
                        ExperimentInfos.Add(new HarmonicExperimentInfo(reportInfo, this, experimentName, experimentInfo));
                        break;
                    case "电压波动和闪烁":
                        ExperimentInfos.Add(new FluctuationExperimentInfo(reportInfo, this, experimentName, experimentInfo));
                        break;
                    case "电快速瞬变脉冲群":
                        ExperimentInfos.Add(new AcDcExperimentInfo(reportInfo, this, experimentName, experimentInfo));
                        break;
                    case "电压暂降/短时中断":
                    case "电压暂降和短时中断":
                        ExperimentInfos.Add(new SagBreakExperimentInfo(reportInfo, this, "电压暂降和短时中断", experimentInfo));
                        break;
                    default:
                        ExperimentInfos.Add(new DefaultExperimentInfo(reportInfo, this, experimentName, experimentInfo));
                        break;
                }
            }

        }
        /// <summary>
        /// 写入所有实验
        /// </summary>
        /// <param name="wordUtil"></param>
        public void WriteExperimentInfoAll(ReportHandleWord wordUtil)
        {
            var k = 1;
            foreach (var experimentInfo in ExperimentInfos)
            {
                experimentInfo.IsNeedNewBookmark = (k != ExperimentInfos.Count);
                experimentInfo.WriteExperimentInfo(wordUtil);
                k++;
            }
            wordUtil.FormatCurrentWord(ExperimentInfos.Count);
        }

        /// <summary>
        /// 实验数据的集合
        /// </summary>
        public IList<ExperimentInfoAbstract> ExperimentInfos { get; set; }

        /// <summary>
        /// 实验数据
        /// </summary>
        public JArray ExperimentInfosJArray { get; set; }

        /// <summary>
        /// 实验新的bookmark
        /// </summary>
        public string NewBookmark { get; set; }

        /// <summary>
        /// 报告信息 导航属性
        /// </summary>
        public ReportInfo ReportInfo { get; set; }
    }
}