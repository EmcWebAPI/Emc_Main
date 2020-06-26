using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using EmcReportWebApi.Business.ImplWordUtil;
using EmcReportWebApi.Config;
using EmcReportWebApi.Models;
using EmcReportWebApi.ReportComponent.ExperimentData;
using Newtonsoft.Json.Linq;

namespace EmcReportWebApi.ReportComponent.Experiment
{
    /// <summary>
    /// 实验项目
    /// </summary>
    public abstract class ExperimentInfoAbstract
    {
        /// <summary>
        /// 实验数据的集合
        /// </summary>
        public JArray ExperimentJArray { get; set; }

        /// <summary>
        /// 实验数据
        /// </summary>
        public JObject ExperimentJObject { get; set; }

        /// <summary>
        /// 实验名称
        /// </summary>
        public string ExperimentName { get; set; }

        /// <summary>
        /// 报告的bookmark
        /// </summary>
        public string ReportBookmark { get; set; }

        /// <summary>
        /// 实验类型
        /// </summary>
        public ExperimentType ExperimentType { get; set; }

        /// <summary>
        /// 实验所需模板全路径(手动copy)
        /// </summary>
        public string ExperimentTemplateFileFullName { get; set; }

        /// <summary>
        /// 实验数据所需模板全路径(手动copy)
        /// </summary>
        public string ExperimentDataTemplateFileFullname { get; set; }


        /// <summary>
        /// rtf中的表格信息
        /// </summary>
        public RtfTableInfo RtfTableInfo { get; set; }

        /// <summary>
        /// rtf中的文件信息
        /// </summary>
        public RtfPictureInfo RtfPictureInfo { get; set; }

        /// <summary>
        /// 连接图集合
        /// </summary>
        public IList<ExperimentImage> ConnectionImages { get; set; }

        /// <summary>
        /// 布置图信息
        /// </summary>
        public IList<ExperimentImage> ArrangementImages { get; set; }

        /// <summary>
        /// 报告信息 导航属性
        /// </summary>
        public ReportInfo ReportInfo { get; set; }

        /// <summary>
        /// 实验总的信息 导航属性
        /// </summary>
        public ExperimentInfo ExperimentInfo { get; set; }

        /// <summary>
        /// 是否需要创建新的书签
        /// </summary>
        public bool IsNeedNewBookmark { get; set; }


        /// <summary>
        /// 实验数据的集合
        /// </summary>
        public IList<ExperimentDataInfoAbstract> ExperimentDataInfos { get; set; }

        /// <summary>
        /// 写入实验数据
        /// </summary>
        public abstract void WriteExperimentInfo(ReportHandleWord wordUtil);

        /// <summary>
        /// 创建模板中间件
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        protected virtual string CreateTemplateMiddle(string filePath)
        {
            string fileName = Guid.NewGuid() + ".docx";
            DirectoryInfo di = new DirectoryInfo(ReportInfo.TemplateMiddleFilesPath);
            if (!di.Exists) { di.Create(); }

            string fileFullName = ReportInfo.TemplateMiddleFilesPath + fileName;
            FileInfo file = new FileInfo(filePath);
            if (File.Exists(filePath))
            {
                file.CopyTo(fileFullName);
                return fileFullName;
            }
            else
            {
                return $"{ExperimentName}模板不存在";
            }
        }

    }
    /// <summary>
    /// 实验图片
    /// </summary>
    public class ExperimentImage
    {
        /// <summary>
        /// 图片内容
        /// </summary>
        public string Content { get; set; }

        /// <summary>
        /// 图片名称
        /// </summary>
        public string ImageName { get; set; }

        /// <summary>
        /// 图片路径(路径+名称)
        /// </summary>
        public string ImageFileFullName { get; set; }
    }

    /// <summary>
    /// 实验类型
    /// </summary>
    public enum ExperimentType
    {
        /// <summary>
        /// 
        /// </summary>
        传导发射,
        /// <summary>
        /// 
        /// </summary>
        辐射发射,
        /// <summary>
        /// 
        /// </summary>
        谐波失真,
        /// <summary>
        /// 
        /// </summary>
        电压波动和闪烁,
        /// <summary>
        /// 
        /// </summary>
        静电放电,
        /// <summary>
        /// 
        /// </summary>
        射频电磁场辐射,
        /// <summary>
        /// 
        /// </summary>
        电快速瞬变脉冲群,
        /// <summary>
        /// 
        /// </summary>
        浪涌,
        /// <summary>
        /// 
        /// </summary>
        射频感应的的传导骚扰,
        /// <summary>
        /// 
        /// </summary>
        电压暂降和短时中断,
        /// <summary>
        /// 
        /// </summary>
        工频磁场
    }
}