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
using System.Xml.Linq;

namespace EmcReportWebApi.Common
{
    public static class MyTools
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

        public static List<RtfTableInfo> RtfTableInfos = GetRtfTableInfo();
        public static List<RtfPictureInfo> RtfPictureInfos = GetRtfPictueInfo();

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
    }
}