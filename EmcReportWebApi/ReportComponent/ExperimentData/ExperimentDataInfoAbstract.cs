using System.Collections.Generic;
using EmcReportWebApi.Business.ImplWordUtil;
using Newtonsoft.Json.Linq;

namespace EmcReportWebApi.ReportComponent.ExperimentData
{
    /// <summary>
    /// 实验数据信息
    /// </summary>
    public abstract class ExperimentDataInfoAbstract
    {
        /// <summary>
        /// 写入实验数据信息
        /// </summary>
        /// <param name="wordUtil"></param>
        /// <param name="isNeedBreak"></param>
        public abstract void WriteExperimentDataInfo(ReportHandleWord wordUtil, bool isNeedBreak);

        /// <summary>
        /// 实验数据头信息集合
        /// </summary>
        public IList<string> ExperimentDataTitleInfos { get; set; }

        /// <summary>
        /// 实验数据
        /// </summary>
        public JObject ExperimentDataJObject { get; set; }

        /// <summary>
        /// 实验数据rtf信息的集合
        /// </summary>
        public JArray ExperimentDataRtfJArray { get; set; }

        /// <summary>
        /// 实验数据html数据集合
        /// </summary>
        public JArray ExperimentDataHtmlJArray { get; set; }

        /// <summary>
        /// 交直流数据线 电压暂降和短时中断
        /// </summary>
        public string ExperimentDataTitle { get; set; }
    }
}