/*
 * 
 * 
 * 
 */
using EmcReportWebApi.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

namespace EmcReportWebApi.Common
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
        /// 当前运行的接口数
        /// </summary>
        public static List<Guid> ReportQueue = new List<Guid>();

        /// <summary>
        /// 待执行的任务数
        /// </summary>
        public static Queue<Guid> TaskQueue = new Queue<Guid>(); 
        
        /// <summary>
        /// 线程池任务信号量
        /// </summary>
        public static SemaphoreSlim SemLim = new SemaphoreSlim(1);

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
            { "main_wtf","ContractClient"},
            { "main_jylb","DetectType"},
            { "main_ypmc","SampleName"},
            { "main_xhgg","SampleModelSpecification"},
            { "ypmc","SampleName"},
            { "sb",""},
            { "wtf","ContractClient"},
            { "wtfdz","AddressOrIdCard"},
            { "scdw","ManufactureCompany"},
            { "sjdw","DetectCompany"},
            { "cydw","SamplingCompany"},
            { "cydd","SamplingAddress"},
            { "cyrq","SamplingDate"},
            { "dyrq","SampleReceiptDate"},
            { "jyxm","Content"},
            { "jyyj","SampleTestBasis"},
            { "jyjl",""},
            { "bz","TestRemark"},
            { "ypbh","SampleNumber"},
            { "xhgg","SampleModelSpecification"},
            { "jylb","DetectType"},
            { "cpbhph","SampleModelSpecification"},
            { "cydbh","SamplingNumber"},
            { "scrq","SampleProductionDate"},
            { "ypsl","SampleQuantity"},
            { "cyjs","SamplingBase"},
            { "jydd","AfterTreatmentMethod"},
            { "jyrq",""},
            { "ypms","SampleTrademark"},
            { "xhgghqtsm","TestRemark"}
        };
        #endregion
    }
}