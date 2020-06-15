/*
 * 
 * 
 * 
 */
using EmcReportWebApi.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

namespace EmcReportWebApi.Config
{
    public static class EmcConfig
    {
        /// <summary>
        /// 报表日志记录
        /// </summary>
        public static log4net.ILog ErrorLog = log4net.LogManager.GetLogger("ErrorLogger");
        public static log4net.ILog InfoLog = log4net.LogManager.GetLogger("InfoLogger");

        /// <summary>
        /// 当前程序路径
        /// </summary>
        public static string CurrRoot = AppDomain.CurrentDomain.BaseDirectory;

        /// <summary>
        /// rtf中获取表格的配置
        /// </summary>
        public static List<RtfTableInfo> RtfTableInfos = GetRtfTableInfo();
        /// <summary>
        /// rtf中获取图片的配置
        /// </summary>
        public static List<RtfPictureInfo> RtfPictureInfos = GetRtfPictueInfo();
        
        /// <summary>
        /// 线程池任务信号量
        /// </summary>
        public static SemaphoreSlim SemLim = new SemaphoreSlim(int.Parse(ConfigurationManager.AppSettings["TaskCount"].ToString()));

        /// <summary>
        /// 获取时间戳
        /// </summary>
        /// <param name="d"></param>
        /// <returns></returns>
        public static string GetTimestamp(DateTime d)
        {
            TimeSpan ts = d.ToUniversalTime() - new DateTime(1970, 1, 1);
            return ts.TotalMilliseconds.ToString();     //精确到毫秒
        }

        /// <summary>
        /// 删除进程
        /// </summary>
        public static void KillWordProcess()
        {
            Process myProcess = new Process();
            Process[] wordProcess = Process.GetProcessesByName("winword");
            foreach (Process pro in wordProcess) //这里是找到那些没有界面的Word进程
            {
                pro.Kill();
            }
        }

        public static List<RtfTableInfo> GetRtfTableInfo()
        {
            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            //初始化xml信息
            XDocument docXml = XDocument.Load(currRoot + "\\RtfConfig.xml");
            var data = new List<RtfTableInfo>();

            foreach (var item in docXml.Root.Elements())
            {
                string itemType = item.Attribute("Type").Value;
                data = data.Concat((from d in item.Elements().Where(p => p.Name == "Table")
                                    select new RtfTableInfo
                                    {
                                        StartIndex = int.Parse(d.Attribute("StartIndex").Value.ToString()),
                                        EndIndex = int.Parse(d.Attribute("EndIndex").Value.ToString()),
                                        MainTitle =d.Attribute("MainTitle").Value.ToString(),
                                        TitleRow = int.Parse(d.Attribute("TitleRow").Value.ToString()),
                                        RtfType = itemType,
                                        Bookmark = d.Attribute("Bookmark").Value.ToString(),
                                        ColumnInfoDic = (from f in d.Elements()
                                                         select new
                                                         {
                                                             ColumnIndex = int.Parse(f.Element("ColumnIndex").Value),
                                                             Title = f.Element("Title").Value.ToString()
                                                         }
                                                     ).ToDictionary(x => x.ColumnIndex, x => x.Title)
                                    }).ToList()).ToList();
            }

            return data;
        }
        public static List<RtfPictureInfo> GetRtfPictueInfo()
        {
            string currRoot = AppDomain.CurrentDomain.BaseDirectory;
            //初始化xml信息
            XDocument docXml = XDocument.Load(currRoot + "\\RtfConfig.xml");
            var data = new List<RtfPictureInfo>();
            foreach (var item in docXml.Root.Elements())
            {
                string itemType = item.Attribute("Type").Value;
                data = data.Concat((from d in item.Elements().Where(p => p.Name == "Picture")
                                    select new RtfPictureInfo
                                    {
                                        StartIndex = int.Parse(d.Attribute("StartIndex").Value.ToString()),
                                        RtfType = itemType,
                                        Bookmark = d.Attribute("Bookmark").Value.ToString()
                                    }).ToList()).ToList();
            }

            return data;
        }


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
            { "dyrq","SampleReceiptDate"},//到样日期
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
            { "cyjs","SamplingBase"},//抽样技术
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
            "lamv"
        };

        #endregion
    }
}