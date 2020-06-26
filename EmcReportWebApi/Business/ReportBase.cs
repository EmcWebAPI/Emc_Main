using EmcReportWebApi.Config;
using EmcReportWebApi.Models;
using System;
using System.IO;
using System.Reflection;

namespace EmcReportWebApi.Business
{
    /// <summary>
    /// 参数设置
    /// </summary>
    public class ReportBase
    {
        /// <summary>
        /// 保存参数文件
        /// </summary>
        protected void SaveParams<T>(T para)
        {
            string dateStr = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            string txtPath = $"{EmcConfig.CurrentRoot}Log\\Params\\{dateStr}.txt";
            if (!File.Exists(txtPath))
            {
                //没有则创建这个文件
                FileStream fs1 = new FileStream(txtPath, FileMode.Create, FileAccess.Write);//创建写入文件      
                StreamWriter sw = new StreamWriter(fs1);

                PropertyInfo[] propertyInfos= para.GetType().GetProperties();
                foreach (PropertyInfo item in propertyInfos)
                {
                    if (item.GetValue(para) != null)
                        sw.WriteLine(item.Name + ":" + item.GetValue(para, null));
                }
                sw.Close();
                fs1.Close();
            }
        }

        /// <summary>
        /// 返回结果参数
        /// </summary>
        protected ReportResult<T> SetReportResult<T>(string message, bool submitResult, T content)
        {
            ReportResult<T> reportResult = new ReportResult<T>();
            reportResult.Message = message;
            reportResult.SumbitResult = submitResult;
            reportResult.Content = content;
            return reportResult;
        }
    }
}