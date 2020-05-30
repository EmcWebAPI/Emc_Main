using EmcReportWebApi.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EmcReportWebApi.Business.ImplWordUtil
{
    /// <summary>
    /// 报告实现word工具类
    /// </summary>
    public class ReportHandleWord:WordUtil
    {

        /// <summary>
        /// 打开现有文件操作
        /// </summary>
        /// <param name="fileFullName">需保存文件的路径</param>
        public ReportHandleWord(string fileFullName) : base(fileFullName) {

        }

        /// <summary>
        /// 打开现有文件操作
        /// </summary>
        /// <param name="outFileFullName">生成文件路径</param>
        /// <param name="fileFullName">引用文件路径</param>
        /// <param name="isSaveAs">是否另存文件</param>
        public ReportHandleWord(string outFileFullName, string fileFullName = "", bool isSaveAs = true) : base(outFileFullName, fileFullName, isSaveAs)
        {

        }
    }
}