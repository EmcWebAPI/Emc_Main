using System.Collections.Generic;

namespace EmcReportWebApi.Models
{
    public class RtfTableInfo
    {
        public int StartIndex { get; set; }

        public string RtfType { get; set; }

        public string Bookmark { get; set; }

        public int TitleRow { get; set; }

        public string MainTitle { get; set; }
        public Dictionary<int, string> ColumnInfoDic { get; set; }
    }

    public class RtfPictureInfo
    {
        public int StartIndex { get; set; }

        public string Bookmark { get; set; }

        public string RtfType { get; set; }
    }
}