using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Office.Interop.Word;

namespace EmcReportWebApi.Models
{
    /// <summary>
    /// 单元格信息
    /// </summary>
    public class CellInfo
    {
        /// <summary>
        /// new
        /// </summary>
        public CellInfo() { }

        /// <summary>
        /// new
        /// </summary>
        public CellInfo(string cellText, int row, int column,int pageIndex,Cell realCell)
        {
            this.CellText = cellText;
            this.RowNumber = row;
            this.ColumnNumber = column;
            this.PageIndex = pageIndex;
            this.RealCell = realCell;
        }
        /// <summary>
        /// 单元格内容
        /// </summary>
        public string CellText { get; set; }

        /// <summary>
        /// 行
        /// </summary>
        public int RowNumber { get; set; }

        /// <summary>
        /// 列
        /// </summary>
        public int ColumnNumber { get; set; }

        /// <summary>
        /// 所在文档页
        /// </summary>
        public int PageIndex { get; set; }

        /// <summary>
        /// 源单元格
        /// </summary>
        public Cell RealCell { get; set; }
    }
}