using System.Collections.Generic;

namespace EmcReportWebApi.Models
{
    /// <summary>
    /// Rtf中table
    /// </summary>
    public class RtfTableInfo
    {
        /// <summary>
        /// 开始的table下标 从1开始
        /// </summary>
        public int StartIndex { get; set; }
        /// <summary>
        /// 结束下标 
        /// </summary>
        public int EndIndex { get; set; }
        /// <summary>
        /// 实验类型(传导发射,辐射发射,谐波,电压波动)
        /// </summary>
        public string RtfType { get; set; }
        /// <summary>
        /// 模板中书签
        /// </summary>
        public string Bookmark { get; set; }
        /// <summary>
        /// 标题行
        /// </summary>
        public int TitleRow { get; set; }
        /// <summary>
        /// table上的主标题(现在只有谐波失真有)
        /// </summary>
        public string MainTitle { get; set; }
        /// <summary>
        /// 列信息的集合
        /// </summary>
        public Dictionary<int, string> ColumnInfoDic { get; set; }
    }
    /// <summary>
    /// rtf图片信息
    /// </summary>
    public class RtfPictureInfo
    {
        /// <summary>
        /// 图片开始的下标(没有结束下标 现在是所有图片都取)
        /// </summary>
        public int StartIndex { get; set; }
        /// <summary>
        /// word模板中的书签
        /// </summary>
        public string Bookmark { get; set; }
        /// <summary>
        /// 实验类型(传导发射,辐射发射,谐波,电压波动)
        /// </summary>
        public string RtfType { get; set; }
    }
}