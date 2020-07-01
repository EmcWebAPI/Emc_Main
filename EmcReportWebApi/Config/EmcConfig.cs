using EmcReportWebApi.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using EmcReportWebApi.ReportComponent.ReviewTable;

namespace EmcReportWebApi.Config
{
    /// <summary>
    /// 项目配置信息
    /// </summary>
    public static class EmcConfig
    {
        /// <summary>
        /// 报表日志记录
        /// </summary>
        public static log4net.ILog InfoLog = log4net.LogManager.GetLogger("InfoLogger");
        /// <summary>
        /// 错误日志
        /// </summary>
        public static log4net.ILog ErrorLog = log4net.LogManager.GetLogger("ErrorLogger");
        /// <summary>
        /// 当前程序路径
        /// </summary>
        public static string CurrentRoot = AppDomain.CurrentDomain.BaseDirectory;
        /// <summary>
        /// 报告存放文件的根目录
        /// </summary>
        public static string ReportFilesPathRoot = $@"{CurrentRoot}Files\ReportFiles\";
        /// <summary>
        /// 报告生成文件目录
        /// </summary>
        public static string ReportOutputPath = $@"{CurrentRoot}Files\OutPut\";
        /// <summary>
        /// emc报告模板文件路径
        /// </summary>
        public static string ReportTemplateFileFullName =
            $@"{CurrentRoot}Files\{ConfigurationManager.AppSettings["TemplateName"]}";
        /// <summary>
        /// 标准报告模板文件路径
        /// </summary>
        public static string StandardReportTemplateFileFullName =
            $@"{CurrentRoot}Files\{ConfigurationManager.AppSettings["StandardTemplateName"]}";
        /// <summary>
        /// 模板中间件文件路径
        /// </summary>
        public static string ReportTemplateMiddlewareFilePath =
            $@"{CurrentRoot}Files\TemplateMiddleware\";

        /// <summary>
        /// 模板中间件文件路径
        /// </summary>
        public static string ExperimentTemplateFilePath =
            $@"{CurrentRoot}Files\ExperimentTemplate\";

        /// <summary>
        /// 线程池任务信号量
        /// </summary>
        public static SemaphoreSlim SemLim = new SemaphoreSlim(int.Parse(ConfigurationManager.AppSettings["TaskCount"]));

        /// <summary>
        /// 获取时间戳
        /// </summary>
        /// <param name="d"></param>
        /// <returns></returns>
        public static string GetTimestamp(DateTime d)
        {
            TimeSpan ts = d.ToUniversalTime() - new DateTime(1970, 1, 1);
            return ts.TotalMilliseconds.ToString(CultureInfo.CurrentCulture);     //精确到毫秒
        }

        /// <summary>
        /// 删除进程
        /// </summary>
        public static void KillWordProcess()
        {
            Process[] wordProcess = Process.GetProcessesByName("winword");
            foreach (Process pro in wordProcess) //这里是找到那些没有界面的Word进程
            {
                pro.Kill();
            }
        }


        #region Emc报告
        /// <summary>
        /// 审查表内容信息
        /// </summary>
        public static IEnumerable<ReviewTableItemInfo> ReviewTableItemInfos = ConfigReviewTableItemInfos();

        /// <summary>
        /// 获取审查表内容信息
        /// </summary>
        /// <returns></returns>
        public static IEnumerable<ReviewTableItemInfo> ConfigReviewTableItemInfos()
        {
            return new List<ReviewTableItemInfo>
            {
                new ReviewTableItemInfo
                {
                    ReportBookmark = "sjypms",//受检样品描述
                    ReviewTableItemIndex = 3,
                    ItemType = ReviewTableItemType.Table
                },
                new ReviewTableItemInfo
                {
                    ReportBookmark = "ypgcList",//样品构成
                    ReviewTableItemIndex = 4,
                    ItemType = ReviewTableItemType.Table
                },
                new ReviewTableItemInfo
                {
                    ReportBookmark = "fzsbList",//辅助设备
                    ReviewTableItemIndex = 5,
                    ItemType = ReviewTableItemType.Table
                },
                new ReviewTableItemInfo
                {
                    ReportBookmark = "ypyxList",
                    ReviewTableItemIndex = 6,//样品运行模式
                    ItemType = ReviewTableItemType.Table
                },
                new ReviewTableItemInfo
                {
                    ReportBookmark = "ypdlList",//样品电缆
                    ReviewTableItemIndex = 7,
                    ItemType = ReviewTableItemType.Table
                },new ReviewTableItemInfo
                {
                    ReportBookmark = "connectionGraph",//样品连接图
                    ReviewTableItemIndex = 1,
                    ItemType = ReviewTableItemType.Image
                },new ReviewTableItemInfo
                {
                    ReportBookmark = "cssbList",
                    ItemType = ReviewTableItemType.TestEquipment
                }
            };
        }

        /// <summary>
        /// rtf中获取表格的配置
        /// </summary>
        public static List<RtfTableInfo> RtfTableInfos = GetRtfTableInfo();
        /// <summary>
        /// rtf中获取图片的配置
        /// </summary>
        public static List<RtfPictureInfo> RtfPictureInfos = GetRtfPictureInfo();
        /// <summary>
        /// rtf内容信息
        /// </summary>
        /// <returns></returns>
        public static List<RtfTableInfo> GetRtfTableInfo()
        {
            //初始化xml信息
            XDocument docXml = XDocument.Load(CurrentRoot + "RtfConfig.xml");
            var data = new List<RtfTableInfo>();

            if (docXml.Root != null)
                foreach (var item in docXml.Root.Elements())
                {
                    string itemType = item.Attribute("Type")?.Value;
                    data = data.Concat((from d in item.Elements().Where(p => p.Name == "Table")
                                        select new RtfTableInfo
                                        {
                                            StartIndex = int.Parse(d.Attribute("StartIndex")?.Value.ToString() ?? string.Empty),
                                            EndIndex = int.Parse(d.Attribute("EndIndex")?.Value.ToString() ?? string.Empty),
                                            MainTitle = d.Attribute("MainTitle")?.Value.ToString(),
                                            TitleRow = int.Parse(d.Attribute("TitleRow")?.Value.ToString() ?? string.Empty),
                                            RtfType = itemType,
                                            Bookmark = d.Attribute("Bookmark")?.Value.ToString(),
                                            ColumnInfoDic = (from f in d.Elements()
                                                             select new
                                                             {
                                                                 ColumnIndex = int.Parse(f.Element("ColumnIndex")?.Value ?? string.Empty),
                                                                 Title = f.Element("Title")?.Value.ToString()
                                                             }
                                                ).ToDictionary(x => x.ColumnIndex, x => x.Title)
                                        }).ToList()).ToList();
                }

            return data;
        }
        /// <summary>
        /// rtf图片信息
        /// </summary>
        /// <returns></returns>
        public static List<RtfPictureInfo> GetRtfPictureInfo()
        {
            //初始化xml信息
            XDocument docXml = XDocument.Load(CurrentRoot + "RtfConfig.xml");
            var data = new List<RtfPictureInfo>();
            if (docXml.Root != null)
                foreach (var item in docXml.Root.Elements())
                {
                    string itemType = item.Attribute("Type")?.Value;
                    data = data.Concat((from d in item.Elements().Where(p => p.Name == "Picture")
                                        select new RtfPictureInfo
                                        {
                                            StartIndex = int.Parse(d.Attribute("StartIndex")?.Value.ToString() ?? string.Empty),
                                            RtfType = itemType,
                                            Bookmark = d.Attribute("Bookmark")?.Value.ToString()
                                        }).ToList()).ToList();
                }

            return data;
        }

        /// <summary>
        /// 实验基础信息
        /// </summary>
        public static IList<string> ExperimentBaseInfo = new List<string>
        {
            "syjg",
            "jyrq",
            "wd",
            "xdsd",
            "dqyl"
        };
        /// <summary>
        /// 实验数据头信息
        /// </summary>
        public static Dictionary<string,string> ExperimentDataTitleInfo = new Dictionary<string, string>
        {
            {"sygdy","试验供电电源："},
            {"syplfw","试验频率范围："},
            {"ypyxms","样品运行模式："},
            {"mccfpl","脉冲重复频率（kHz）："},
            {"sycxsj","试验持续时间（s）："},
            {"cfpl","重复频率（s）："},
            {"cs","次数（次）："},
            {"sycfcs","试验重复次数（次）："},
            {"sysjjg","试验时间间隔（s）："},
            {"sypl","试验频率（Hz）："},
            {"gczqsj","观察周期/时间："}
        };

        #endregion

        #region 标准报告
        /// <summary>
        /// 合同内容对应的书签
        /// </summary>
        public static Dictionary<string, string> ContractToJObject = new Dictionary<string, string>() {
            { "bgbh","ReportCode" },//报告编号
            { "main_wtf","ContractClient"},//委托方
            { "main_jylb","DetectType"},//检验类别
            { "main_ypmc","SampleNameRPT"},//样品名称
            { "main_xhgg","SampleModelSpecificationRPT"},//型号规格
            { "ypmc","SampleName"},//样品名称
            { "sb","SampleTrademark"},//商标
            { "wtf","ContractClient"},//委托方
            { "wtfdz","AddressOrIdCard"},//委托方地址
            { "scdw","ManufactureCompany"},//生产单位
            { "sjdw","DetectCompany"},//受检单位
            { "cydw","SamplingCompany"},//抽样单位
            { "cydd","SamplingAddress"},//抽样地点
            { "cyrq","SamplingDate"},//抽样日期
            { "dyrq","SampleReceiptDateRPT"},//到样日期
            { "jyxm","Content"},//检验项目
            { "jyyj","SampleTestBasis"},//检验依据
            { "jyjl","Conclusion"},//检验结论
            { "bz","TestRemark"},//备注
            { "ypbh","SampleNumber"},//样品编号
            { "xhgg","SampleModelSpecificationRPT"},//型号规格
            { "jylb","DetectType"},//检验类别
            { "cpbhph","BatchNumberRPT"},//产品编号/批号
            { "cydbh","SamplingNumber"},//抽样单编号
            { "scrq","SampleProductionDate"},//生产日期
            { "ypsl","SampleQuantity"},//样品数量
            { "cyjs","SamplingBase"},//抽样基数
            { "jydd","AfterTreatmentMethod"},//检验地点
            { "jyrq","InspectionDate"},//检验日期
            { "ypms","SampleDescription"},//样品描述
            { "xhgghqtsm","OtherDescription"},//型号规格及其他说明
            { "zjgcs","ChiefInspection"},//主检工程师(检验)
            { "syxz","SampleAcquisitionModeSy"},//送样选中
            { "cyxz","SampleAcquisitionModeCy"}//抽样选中
        };

        /// <summary>
        /// 公式的集合
        /// </summary>
        public static IList<string> FormulaType= new List<string>
        {
            "avg",
            "absFv",
            "absFc",
            "uva",
            "uvb",
            "lamv",
            "<上标>",
            "<下标>",
            "<上下标>"
        };

        #endregion
    }
}