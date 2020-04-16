using EmcReportWebApi.Common;
using EmcReportWebApi.Models;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace EmcReportWebApi.Business.Implement
{
    public class ReportBase
    {
        //设置首页内容
        protected string InsertContentToWord(WordUtil wordUtil, JObject jo1)
        {
            foreach (var item in jo1)
            {
                string key = item.Key.ToString();
                string value = item.Value.ToString();
                if (key.Equals("main_wtf") || key.Equals("main_ypmc") || key.Equals("main_xhgg") || key.Equals("main_jylb"))
                {
                    value = CheckFirstPage(value);
                }
                wordUtil.InsertContentToWordByBookmark(value, key);
            }
            return "保存成功";
        }
        //首页内容特殊处理
        protected string CheckFirstPage(string itemValue)
        {
            int fontCount = 38;
            int valueCount = System.Text.Encoding.Default.GetBytes(itemValue).Length;
            if (fontCount > valueCount)
            {
                int spaceCount = (fontCount - valueCount) / 2;
                for (int i = 0; i < spaceCount; i++)
                {
                    itemValue = " " + itemValue + " ";
                }
            }

            return itemValue;
        }

        /// <summary>
        /// 创建模板中间件
        /// </summary>
        protected string CreateTemplateMiddle(string dir, string template, string filePath)
        {

            string dateStr = DateTime.Now.ToString("yyyyMMddhhmmss");
            string fileName = template + dateStr + ".docx";
            DirectoryInfo di = new DirectoryInfo(dir);
            if (!di.Exists) { di.Create(); }

            string htmlpath = dir + "\\" + fileName;
            FileInfo file = new FileInfo(filePath);
            if (File.Exists(filePath))
            {
                file.CopyTo(htmlpath);
                return htmlpath;
            }
            else
            {
                return "模板不存在";
            }

        }

        /// <summary>
        /// 删除模板中间件
        /// </summary>
        public void DelectDir(string srcPath)
        {
            try
            {
                DirectoryInfo dir = new DirectoryInfo(srcPath);
                FileSystemInfo[] fileinfo = dir.GetFileSystemInfos();  //返回目录中所有文件和子目录
                foreach (FileSystemInfo i in fileinfo)
                {
                    if (i is DirectoryInfo)            //判断是否文件夹
                    {
                        DirectoryInfo subdir = new DirectoryInfo(i.FullName);
                        subdir.Delete(true);          //删除子目录和文件
                    }
                    else
                    {
                        //如果 使用了 streamreader 在删除前 必须先关闭流 ，否则无法删除 sr.close();
                        File.Delete(i.FullName);      //删除指定文件
                    }
                }
                Directory.Delete(srcPath);
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// html字符串转word
        /// </summary>
        protected string CreateHtmlFile(string htmlStr, string dirPath)
        {
            string dateStr = DateTime.Now.ToString("yyyyMMddHHmmss");
            string htmlpath = dirPath + "\\reportHtml" + dateStr + ".html";
            FileStream fs = new FileStream(htmlpath, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);
            sw.Write(htmlStr);
            sw.Close();
            sw.Dispose();
            fs.Close();
            fs.Dispose();
            return htmlpath;
        }

        /// <summary>
        /// 保存参数文件
        /// </summary>
        protected void SaveParams(ReportParams para)
        {
            string dateStr = DateTime.Now.ToString("yyyyMMddHHmmss");
            string txtPath = string.Format("{0}Log\\Params\\{1}.txt", EmcConfig.CurrRoot, dateStr);
            if (!System.IO.File.Exists(txtPath))
            {
                //没有则创建这个文件
                FileStream fs1 = new FileStream(txtPath, FileMode.Create, FileAccess.Write);//创建写入文件      
                StreamWriter sw = new StreamWriter(fs1);
                sw.WriteLine("ReportId:" + para.ReportId);
                sw.WriteLine("ZipFilesUrl:" + para.ZipFilesUrl);
                sw.WriteLine("JsonStr:" + para.JsonStr);
                sw.Close();
                fs1.Close();
            }
        }

        /// <summary>
        /// 返回结果参数
        /// </summary>
        protected ReportResult<T> SetReportResult<T>(string message, bool submitResult, T content)
        {
            Type type = content.GetType();
            ReportResult<T> reportResult = new ReportResult<T>();
            reportResult.Message = message;
            reportResult.SumbitResult = submitResult;
            reportResult.Content = content;
            return reportResult;
        }

        /// <summary>
        /// 获取模板路径
        /// </summary>
        protected string GetTemplatePath(string fileName)
        {
            return string.Format(@"{0}\Files\ExperimentTemplate\{1}", EmcConfig.CurrRoot, fileName);
        }
    }
}