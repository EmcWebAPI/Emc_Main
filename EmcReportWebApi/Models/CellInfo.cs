using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Office.Interop.Word;

namespace EmcReportWebApi.Models
{
    public class CellInfo
    {

        public CellInfo() { }

        public CellInfo(string cellText, int row, int column,int pageIndex,Cell realCell)
        {
            this.CellText = cellText;
            this.RowNumber = row;
            this.ColumnNumber = column;
            this.PageIndex = pageIndex;
            this.RealCell = realCell;
        }
        
        public string CellText { get; set; }

        public int RowNumber { get; set; }

        public int ColumnNumber { get; set; }

        public int PageIndex { get; set; }

        public Cell RealCell { get; set; }
    }
}