using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EmcReportWebApi.Models
{
    public class CellInfo
    {

        public CellInfo(string cellText,int row,int column) {
            this.CellText = cellText;
            this.RowNumber = row;
            this.ColumnNumber = column;
        }

        public string CellText { get; set; }

        public int RowNumber { get; set; }

        public int ColumnNumber { get; set; }
    }
}