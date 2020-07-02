using EmcReportWebApi.Models;
using EmcReportWebApi.Utils;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EmcReportWebApi.Business.ImplWordUtil
{
    /// <summary>
    /// standard report concrete word utils class
    /// </summary>
    public class ReportStandardHandleWord : WordUtil
    {
        /// <summary>
        /// 打开现有文件操作
        /// </summary>
        /// <param name="fileFullName">需保存文件的路径</param>
        public ReportStandardHandleWord(string fileFullName) : base(fileFullName)
        {

        }

        /// <summary>
        /// 打开现有文件操作
        /// </summary>
        /// <param name="outFileFullName">生成文件路径</param>
        /// <param name="fileFullName">引用文件路径</param>
        /// <param name="isSaveAs">是否另存文件</param>
        public ReportStandardHandleWord(string outFileFullName, string fileFullName = "", bool isSaveAs = true) : base(outFileFullName, fileFullName, isSaveAs)
        {

        }
        #region 标准报告业务相关

        /// <summary>
        /// 单元格右下添加内容的集合
        /// </summary>
        private readonly Dictionary<string, string> _lowerRightCornerCells = new Dictionary<string, string>();
        private int _colSpan;

        /// <summary>
        /// 表格拆分合并 添加"续"
        /// </summary>
        public virtual int TableSplit(string bookmark, bool hasPhoto)
        {
            try
            {
                List<CellInfo> cellList = new List<CellInfo>();

                int pageIndex = 0;
                Range tableRange = GetBookmarkRank(_currentWord, bookmark);
                Table table = tableRange.Tables[1];

                Cells cells = table.Range.Cells;

                foreach (Cell cell in cells)
                {

                    Range r = cell.Range;
                    int rowNumber = (int)r.Information[WdInformation.wdStartOfRangeRowNumber];
                    int columnNumber = (int)r.Information[WdInformation.wdStartOfRangeColumnNumber];
                    int pageNumber = (int)r.Information[WdInformation.wdActiveEndPageNumber];
                    cellList.Add(new CellInfo(r.Text, rowNumber, columnNumber, pageNumber, cell));
                    if (r.Text.Equals("\r\a") && !r.Text.Contains("^^"))
                    {
                        continue;
                    }
                    r.Select();
                    if (_wordApp.Selection.Bookmarks.Exists("photo"))
                        break;

                    if (pageNumber >= 6 && pageNumber != pageIndex)
                    {
                        //判断第一列最后一个单元格 的高度
                        CellInfo cellInfo = cellList.LastOrDefault(p => p.ColumnNumber == 1);

                        if (cellInfo != null && (int)cellInfo.RealCell.Range.Information[WdInformation.wdActiveEndPageNumber] ==
                            pageNumber)
                        {
                            pageIndex = pageNumber;
                            continue;
                        }

                        if (cellInfo != null)
                        {
                            Cell firstLastCell = cellInfo.RealCell;
                            if (HandleFirstColumnCellAddRow(firstLastCell, pageNumber))
                            {
                                pageIndex = pageNumber;
                                continue;
                            }
                        }
                        r.Select();
                        _wordApp.Selection.SplitTable();

                        List<Cell> cellNextList = new List<Cell>();
                        for (int j = 1; j <= 11; j++)
                        {
                            Cell cellNext = TableContinueContent(table, j, cellList);
                            if (cellNext != null)
                                cellNextList.Add(cellNext);
                        }

                        //处理单项结论
                        HandleConclusion(cellList, cellNextList);

                        _wordApp.Selection.Delete(WdUnits.wdCharacter, 1);
                        pageIndex = pageNumber;
                    }
                }
                //替换字符
                Dictionary<int, Dictionary<string, string>> replaceDic = new Dictionary<int, Dictionary<string, string>>();
                Dictionary<string, string> valuePairs = new Dictionary<string, string> { { "^^", "" } };
                //报告编号
                replaceDic.Add(1, valuePairs);//替换全部内容
                ReplaceWritten(replaceDic);

                #region 此处空白
                table.Select();
                Cell lastCellOrDefault = table.Range.Cells.Cast<Cell>().LastOrDefault();
                if (lastCellOrDefault != null)
                {
                    lastCellOrDefault.Select();
                    int currentPageNumber =
                        (int)lastCellOrDefault.Range.Information[WdInformation.wdActiveEndPageNumber];
                    float cellPositionTop =
                        (float)lastCellOrDefault.Range.Information[WdInformation.wdVerticalPositionRelativeToPage];
                    float pageHeight = lastCellOrDefault.Range.PageSetup.PageHeight;
                    //页眉高度大约62.37
                    float cellToPageBottom = pageHeight - cellPositionTop - lastCellOrDefault.Height;
                    bool result = cellToPageBottom > 200;
                    if (result)
                    {
                        lastCellOrDefault.Select();
                        _wordApp.Selection.InsertRowsBelow(1);
                        _wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        _wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop;

                        _wordApp.Selection.Font.NameFarEast = "宋体";
                        _wordApp.Selection.Font.NameAscii = "宋体";
                        _wordApp.Selection.Font.NameOther = "宋体";

                        _wordApp.Selection.Cells.Merge();
                        Cell blankCell = _wordApp.Selection.Cells[1];
                        blankCell.SetHeight(cellToPageBottom - 200, WdRowHeightRule.wdRowHeightAtLeast);
                        blankCell.Range.Text = hasPhoto ? "此处空白" : "以下空白";
                        while (blankCell.Range.Information[WdInformation.wdActiveEndPageNumber] !=
                               currentPageNumber)
                        {
                            blankCell.SetHeight(blankCell.Height - 1, WdRowHeightRule.wdRowHeightAtLeast);
                        }
                    }
                }

                #endregion

                return 1;
            }
            catch (Exception ex)
            {
                _needWrite = false;
                Dispose();
                throw new Exception($"错误信息:{ex.StackTrace}.{ex.Message}");
            }

        }

        /// <summary>
        /// 处理上一页第一列最后一个单元格
        /// </summary>
        private bool HandleFirstColumnCellAddRow(Cell firstLastCell, int pageNumber)
        {
            firstLastCell.Select();
            float cellPositionTop =
                (float)firstLastCell.Range.Information[WdInformation.wdVerticalPositionRelativeToPage];
            float pageHeight = firstLastCell.Range.PageSetup.PageHeight;
            bool result = pageHeight - cellPositionTop < 190;
            if (result)
            {
                _wordApp.Selection.InsertRowsAbove(1);
                _wordApp.Selection.Cells.Merge();
                _wordApp.Selection.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleNone;
                _wordApp.Selection.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleNone;
                _wordApp.Selection.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
                Cell secondLastCell = _wordApp.Selection.Cells[1];

                while ((int)firstLastCell.Range.Information[WdInformation.wdActiveEndPageNumber] != pageNumber)
                {
                    secondLastCell.SetHeight(secondLastCell.Height + 1, WdRowHeightRule.wdRowHeightAtLeast);
                }
            }

            return result;
        }

        private void HandleConclusion(List<CellInfo> cellList, List<Cell> nextCellList)
        {

            //新的单项结论表格(第二个表格倒数第二行)
            if (nextCellList.Count < 2)
            {
                return;
            }
            Cell nextCellInfo = nextCellList[nextCellList.Count - 2];

            //找到最后一个带^^的单元格
            CellInfo cellInfo = cellList.LastOrDefault(p => p.CellText.Contains("^^"));
            if (nextCellInfo.Range.Text.Equals("") || nextCellInfo.Range.Text.Equals("\r\a"))
                if (cellInfo != null)
                    nextCellInfo.Range.InsertAfter(cellInfo.CellText.Replace("\r\a", ""));
        }

        private Cell TableContinueContent(Table table, int column, List<CellInfo> list)
        {

            Cell cellInfo;
            Range range = table.Range.GoToNext(WdGoToItem.wdGoToTable);
            Table tableNext = range.Tables[1];
            try
            {
                cellInfo = tableNext.Cell(1, column);
            }
            catch (Exception)
            {
                return null;
            }
            if (!tableNext.Cell(1, column).Range.Text.Equals("\r\a") || column > 6)
                return cellInfo;

            var cellText = list.Where(p => p.ColumnNumber == column).OrderByDescending(p => p.RowNumber).First().CellText;
            if (!cellText.Contains("续") && column == 1)
                cellText = "续\r\a" + cellText.Replace("\r\a", "");
            tableNext.Cell(1, column).Range.InsertAfter(cellText.Replace("\r\a", ""));
            return cellInfo;
        }

        /// <summary>
        /// 表格拆分
        /// </summary>
        /// <param name="array">拆分数组</param>
        /// <param name="bookmark">书签</param>
        /// <param name="colSpan">检验结果是否有加列</param>
        /// <returns></returns>
        public virtual int TableSplit(JArray array, string bookmark, int colSpan)
        {
            try
            {
                Range tableRange = GetBookmarkRank(_currentWord, bookmark);

                Table table = tableRange.Tables[1];

                if (colSpan > 1)
                {
                    _colSpan = colSpan;
                    table.Cell(1, 4).SetWidth(table.Cell(1, 4).Width - 28.35f, WdRulerStyle.wdAdjustNone);
                    table.Cell(1, 5).SetWidth(table.Cell(1, 5).Width + 28.35f, WdRulerStyle.wdAdjustNone);
                }

                for (int i = array.Count - 1; i >= 0; i--)
                {
                    TableSplit((JObject)array[i], i + 1, table);
                }

                table.Cell(1, 1).Select();
                _wordApp.Selection.Rows.HeadingFormat = -1;
                ClearTableFormat(table);

                if (_lowerRightCornerCells.Count > 0)
                {
                    foreach (var item in _lowerRightCornerCells)
                    {
                        string newBookmark = item.Key;
                        string content = item.Value;
                        AddCellLowerRightCornerContent(newBookmark, content);
                    }
                }

                return 1;
            }
            catch (Exception ex)
            {
                _needWrite = false;
                Dispose();
                throw new Exception($"错误信息:{ex.StackTrace}.{ex.Message}");
            }
        }

        private void ClearTableFormat(Table table)
        {
            table.Select();
            _wordApp.Selection.ParagraphFormat.SpaceBeforeAuto = 0;
            _wordApp.Selection.ParagraphFormat.SpaceAfterAuto = 0;
            _wordApp.Selection.ParagraphFormat.AutoAdjustRightIndent = 0;
            _wordApp.Selection.ParagraphFormat.DisableLineHeightGrid = -1;
            _wordApp.Selection.ParagraphFormat.WordWrap = -1;
        }



        /// <summary>
        /// 表格拆分
        /// </summary>
        /// <param name="jObject">拆分对象</param>
        /// <param name="serialNumber">序号</param>
        /// <param name="table">需拆分表格</param>
        /// <returns></returns>
        public virtual int TableSplit(JObject jObject, int serialNumber, Table table)
        {

            table.Cell(1, 1).Select();
            _wordApp.Selection.InsertRowsBelow(1);
            _wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            _wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop;

            _wordApp.Selection.Font.NameFarEast = "宋体";
            _wordApp.Selection.Font.NameAscii = "宋体";
            _wordApp.Selection.Font.NameOther = "宋体";
            // Row newRow = table.Rows[2];

            for (int i = 5; i <= 7; i++)
            {
                CellAlignCenter(table.Cell(2, i));
            }

            //序号
            table.Cell(2, 1).Range.Text = serialNumber.ToString();

            //检验项目
            table.Cell(2, 2).Range.Text = jObject["itemContent"].ToString();
            //单项结论
            if (jObject["comment"] != null)
                table.Cell(2, 6).Range.Text = "^^" + jObject["comment"];
            //备注
            if (jObject["reMark"] != null)
                table.Cell(2, 7).Range.Text = jObject["reMark"].ToString().Equals(string.Empty) ? "/" : jObject["reMark"].ToString();

            JArray firstItems = (JArray)jObject["list"];

            if (firstItems.Count != 0)
            {
                int firstItemsCount = firstItems.Count;
                //标准条款
                Cell cell3 = table.Cell(2, 3);
                cell3.Split(firstItemsCount, 1);
                for (int i = 0; i < firstItemsCount; i++)
                {
                    table.Cell(2 + i, 3).Range.Text = firstItems[i]["stdItmNo"].ToString();
                }

                //标准要求
                Cell cell4 = table.Cell(2, 4);
                Cell cell5 = table.Cell(2, 5);
                cell4.Split(firstItemsCount, 1);
                cell5.Split(firstItemsCount, 1);

                //拆分备注列
                Cell cell7 = table.Cell(2, 7);
                cell7.Split(firstItemsCount, 1);

                Dictionary<string, JObject> noStdNameDictionary = new Dictionary<string, JObject>();
                Dictionary<string, JObject> stdNameDictionary = new Dictionary<string, JObject>();
                for (int i = 0; i < firstItemsCount; i++)
                {
                    Cell tempCell = table.Cell(2 + i, 4);
                    JObject firstItem = (JObject)firstItems[i];
                    if (firstItem["stdName"] == null)
                    {
                        noStdNameDictionary.Add((2 + i) + "," + 4, firstItem);
                    }
                    else
                    {
                        tempCell.Range.Text = firstItem["stdName"].ToString();
                        stdNameDictionary.Add((2 + i) + "," + 4, firstItem);
                    }
                }

                //遍历节点拆分单元格
                if (stdNameDictionary.Count != 0)
                {
                    bool whileBool = true;
                    while (whileBool)
                    {
                        //cellCol4Dic = AddCellAndSplitNoStdName(table, cellCol4Dic, fist);
                        stdNameDictionary = AddCellAndSplit(table, stdNameDictionary);
                        whileBool = (stdNameDictionary.Count != 0);
                    }
                }

                if (noStdNameDictionary.Count != 0)
                {
                    bool whileBool = true;
                    bool fist = true;
                    while (whileBool)
                    {
                        noStdNameDictionary = AddCellAndSplitNoStdName(table, noStdNameDictionary, fist);
                        whileBool = (noStdNameDictionary.Count != 0);
                        fist = false;
                    }
                }
            }

            return 1;
        }


        /// <summary>
        /// 遍历节点拆分单元格
        /// </summary>
        /// <returns></returns>
        private Dictionary<string, JObject> AddCellAndSplit(Table table, Dictionary<string, JObject> cellCol6Dic)
        {
            Dictionary<string, JObject> cellCol7Dic = new Dictionary<string, JObject>();
            int incr = 0;
            foreach (var item in cellCol6Dic)
            {
                string c = item.Key;
                JObject j = item.Value;
                int cRow = int.Parse(c.Split(',')[0]) + incr;
                int cCol = int.Parse(c.Split(',')[1]);
                if (j["list"] == null)
                    continue;
                JArray secondItems = (JArray)j["list"];
                int secondItemsCount = secondItems.Count;
                if (secondItemsCount > 0)
                {

                    if (secondItemsCount != 1)
                    {
                        //检验结果列拆分
                        // table.Cell(cRow, cCol + 1).Split(secondItemsCount, 1);

                        for (int i = 0; i < secondItemsCount - 1; i++)
                        {
                            table.Cell(cRow, cCol + 1).Select();
                            table.Cell(cRow, cCol + 1).Split(2, 1);
                        }


                        //备注列拆分
                        table.Cell(cRow, cCol + 3).Split(secondItemsCount, 1);
                    }


                    table.Cell(cRow, cCol).Select();
                    var splitCellText = table.Cell(cRow, cCol).Range.Text;
                    table.Cell(cRow, cCol).Split(secondItemsCount, 2);

                    table.Cell(cRow, cCol).Range.Select();
                    if (_wordApp.Selection.OMaths.Count <= 0)
                    {
                        //拆分之后重新赋值
                        int splitCellLength = splitCellText.Length;
                        if (!splitCellText.Equals("\r\a") && !splitCellText.Equals(""))
                        {
                            try
                            {
                                table.Cell(cRow, cCol).Range.Text = splitCellText.Substring(splitCellLength - 2, 2).Equals("\r\a") ?
                                    splitCellText.Substring(0, splitCellLength - 2) : splitCellText;
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine(e);
                            }
                        }
                    }

                    for (int i = 0; i < secondItemsCount; i++)
                    {
                        table.Cell(cRow + i, cCol).SetWidth(45f, WdRulerStyle.wdAdjustFirstColumn);

                        //table.Cell(cRow+i, cCol).PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints;
                        //table.Cell(cRow+i, cCol).PreferredWidth = 40f;
                    }

                    if (secondItemsCount != 1)
                        table.Cell(cRow, cCol).Merge(table.Cell(cRow + secondItemsCount - 1, cCol));
                    // table.Cell(cRow, cCol).SetWidth(40f, WdRulerStyle.wdAdjustFirstColumn);//拆分单元格后设置列宽

                    //结果有拆分的
                    int resultIndex = 0;


                    for (int i = 0; i < secondItemsCount; i++)
                    {
                        Cell tempCell = table.Cell(cRow + i + resultIndex, cCol + 1);
                        JObject secondItem = (JObject)secondItems[i];
                        string itemContent = secondItem["stdItmNo"] != null
                            ? secondItem["stdItmNo"] + secondItem["itemContent"].ToString()
                            : secondItem["itemContent"].ToString();


                        Cell previousTempCell = null;
                        try
                        {
                            previousTempCell = table.Cell(cRow + i + resultIndex - 1, cCol + 1);
                        }
                        catch (Exception)
                        {
                            //resultCell.Range.Text = resultList.First["result"].ToString();
                        }

                        if (previousTempCell != null)
                        {
                            string previousText = previousTempCell.Range.Text;
                            if (previousText.Replace("\r\a", "").Equals("$", StringComparison.OrdinalIgnoreCase))
                            {
                                previousTempCell.Range.Text = "";

                                previousTempCell.Select();
                                previousTempCell.Merge(tempCell);

                                tempCell = previousTempCell;
                            }
                        }

                        tempCell.Range.Text = itemContent;
                        this.FindHtmlLabel(tempCell.Range);

                        if (secondItem["rightContent"] != null && !secondItem["rightContent"].ToString().Equals(""))
                        {
                            //this.AddCellLowerRightCornerContent(tempCell, secondItem["rightContent"].ToString());

                            string newBookmark = "cellBookmark" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + new Random().Next(999);
                            _wordApp.Selection.Bookmarks.Add(newBookmark, tempCell.Range);
                            _lowerRightCornerCells.Add(newBookmark, secondItem["rightContent"].ToString());
                        }

                        if (secondItem["reMark"] != null)
                        {
                            try
                            {
                                table.Cell(cRow + i + resultIndex, cCol + 4).Range.Text = secondItem["reMark"].ToString().Equals(string.Empty) ? "/" : secondItem["reMark"].ToString();
                            }
                            catch (Exception)
                            {
                                table.Cell(cRow + i + resultIndex, cCol + 2).Range.Text = secondItem["reMark"].ToString().Equals(string.Empty) ? "/" : secondItem["reMark"].ToString();
                            }

                        }
                        //检验结果
                        if (secondItem["controls"] != null && !secondItem["controls"].ToString().Equals("") && (secondItem["list"] == null || ((JArray)secondItem["list"]).Count == 0))
                        {
                            Cell resultCell = table.Cell(cRow + i + resultIndex, cCol + 2);
                            JArray resultList = JArray.Parse(secondItem["controls"].ToString());
                            int resultCount = resultList.Count;
                            if (resultCount > 1)
                            {
                                resultCell.Select();
                                resultCell.Split(resultCount, 2);
                                for (int k = 0; k < resultCount; k++)
                                {
                                    //序号列的单元格
                                    Cell xuhaoCell = table.Cell(cRow + i + resultIndex + k, cCol + 2);
                                    xuhaoCell.SetWidth(26f, WdRulerStyle.wdAdjustFirstColumn);

                                }
                                for (int k = 0; k < resultCount; k++)
                                {
                                    Cell xuhaoCell = table.Cell(cRow + i + resultIndex + k, cCol + 2);
                                    xuhaoCell.Range.Text = "#" + (k + 1).ToString();

                                    SetResult(table.Cell(cRow + i + resultIndex + k, cCol + 2 + 1), resultList[k]["result"].ToString(), 2);
                                }
                                resultIndex = resultIndex + resultCount - 1;
                            }
                            else
                            {

                                if (resultList.First["result"].ToString().Equals("@", StringComparison.OrdinalIgnoreCase))
                                {
                                    var previousCell = table.Cell(cRow + i + resultIndex, cCol + 1);
                                    var rangText = previousCell.Range.Text;
                                    var textLength = previousCell.Range.Text.Length;
                                    resultCell.Select();
                                    if (textLength > 2 && rangText.Substring(0, textLength - 2).Length > 10)
                                    {
                                        _wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                                        _wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop;
                                    }

                                    resultCell.Merge(previousCell);
                                }
                                else
                                {
                                    Cell previous = null;
                                    try
                                    {
                                        previous = table.Cell(cRow + i + resultIndex - 1, cCol + 2);
                                    }
                                    catch (Exception)
                                    {
                                        //resultCell.Range.Text = resultList.First["result"].ToString();
                                    }
                                    SetResult(resultCell, resultList.First["result"].ToString(), 1);

                                    if (previous != null)
                                    {
                                        string previousText = previous.Range.Text;
                                        if (previousText.Replace("\r\a", "").Equals("$", StringComparison.OrdinalIgnoreCase))
                                        {
                                            previous.Range.Text = "";

                                            previous.Select();
                                            previous.Merge(resultCell);
                                        }
                                    }
                                }
                            }

                        }
                        cellCol7Dic.Add((cRow + i + resultIndex).ToString() + "," + (cCol + 1).ToString(), secondItem);
                    }
                    incr = incr + secondItemsCount + resultIndex - 1;
                }
            }

            return cellCol7Dic;
        }

        /// <summary>
        /// 遍历无标准内容节点拆分单元格
        /// </summary>
        /// <returns></returns>
        private Dictionary<string, JObject> AddCellAndSplitNoStdName(Table table, Dictionary<string, JObject> cellCol6Dic, bool first)
        {
            Dictionary<string, JObject> cellCol7Dic = new Dictionary<string, JObject>();
            int incr = 0;
            foreach (var item in cellCol6Dic)
            {
                string c = item.Key;
                JObject j = item.Value;
                int cRow = int.Parse(c.Split(',')[0]) + incr;
                int cCol = int.Parse(c.Split(',')[1]);
                if (j["list"] == null)
                    continue;
                JArray secondItems = (JArray)j["list"];
                int secondItemsCount = secondItems.Count;
                if (secondItemsCount > 0)
                {

                    if (secondItemsCount != 1)
                    {
                        //检验结果列拆分
                        for (int i = 0; i < secondItemsCount - 1; i++)
                        {
                            table.Cell(cRow, cCol + 1).Select();
                            table.Cell(cRow, cCol + 1).Split(2, 1);
                        }
                        //备注列拆分
                        table.Cell(cRow, cCol + 3).Split(secondItemsCount, 1);
                    }


                    table.Cell(cRow, cCol).Select();
                    var splitCellText = table.Cell(cRow, cCol).Range.Text;
                    table.Cell(cRow, cCol).Split(secondItemsCount, (cCol == 4 && first) ? 1 : 2);

                    table.Cell(cRow, cCol).Range.Select();
                    if (_wordApp.Selection.OMaths.Count <= 0)
                    {
                        //拆分之后重新赋值
                        int splitCellLength = splitCellText.Length;
                        if (!splitCellText.Equals("\r\a") && !splitCellText.Equals(""))
                        {
                            try
                            {
                                table.Cell(cRow, cCol).Range.Text = splitCellText.Substring(splitCellLength - 2, 2).Equals("\r\a") ?
                                    splitCellText.Substring(0, splitCellLength - 2) : splitCellText;
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine(e);
                            }
                        }
                    }

                    if (cCol != 4 || !first)
                    {
                        for (int i = 0; i < secondItemsCount; i++)
                        {
                            table.Cell(cRow + i, cCol).SetWidth(45f, WdRulerStyle.wdAdjustFirstColumn);
                        }
                    }
                    if (secondItemsCount != 1 && (cCol != 4 || !first))
                        table.Cell(cRow, cCol).Merge(table.Cell(cRow + secondItemsCount - 1, cCol));

                    //结果有拆分的
                    int resultIndex = 0;

                    if (cCol == 4 && first)
                    {
                        cCol = 3;
                    }

                    for (int i = 0; i < secondItemsCount; i++)
                    {
                        Cell tempCell = table.Cell(cRow + i + resultIndex, cCol + 1);
                        JObject secondItem = (JObject)secondItems[i];
                        string itemContent = secondItem["stdItmNo"] != null
                            ? secondItem["stdItmNo"] + secondItem["itemContent"].ToString()
                            : secondItem["itemContent"].ToString();

                       

                        Cell previousTempCell = null;
                        try
                        {
                            previousTempCell = table.Cell(cRow + i + resultIndex - 1, cCol + 1);
                        }
                        catch (Exception)
                        {
                            //resultCell.Range.Text = resultList.First["result"].ToString();
                        }

                        if (previousTempCell != null)
                        {
                            string previousText = previousTempCell.Range.Text;
                            if (previousText.Replace("\r\a", "").Equals("$", StringComparison.OrdinalIgnoreCase))
                            {
                                previousTempCell.Range.Text = "";

                                previousTempCell.Select();
                                previousTempCell.Merge(tempCell);

                                tempCell = previousTempCell;
                            }
                        }

                        tempCell.Range.Text = itemContent;
                        this.FindHtmlLabel(tempCell.Range);

                        if (secondItem["rightContent"] != null && !secondItem["rightContent"].ToString().Equals(""))
                        {
                            //this.AddCellLowerRightCornerContent(tempCell, secondItem["rightContent"].ToString());

                            string newBookmark = "cellBookmark" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + new Random().Next(999);
                            _wordApp.Selection.Bookmarks.Add(newBookmark, tempCell.Range);
                            _lowerRightCornerCells.Add(newBookmark, secondItem["rightContent"].ToString());
                        }

                        if (secondItem["reMark"] != null)
                        {
                            try
                            {
                                table.Cell(cRow + i + resultIndex, cCol + 4).Range.Text = secondItem["reMark"].ToString().Equals(string.Empty) ? "/" : secondItem["reMark"].ToString();
                            }
                            catch (Exception)
                            {
                                table.Cell(cRow + i + resultIndex, cCol + 2).Range.Text = secondItem["reMark"].ToString().Equals(string.Empty) ? "/" : secondItem["reMark"].ToString();
                            }

                        }
                        //检验结果
                        if (secondItem["controls"] != null && !secondItem["controls"].ToString().Equals("") && (secondItem["list"] == null || ((JArray)secondItem["list"]).Count == 0))
                        {
                            Cell resultCell = table.Cell(cRow + i + resultIndex, cCol + 2);
                            JArray resultList = JArray.Parse(secondItem["controls"].ToString());
                            int resultCount = resultList.Count;
                            if (resultCount > 1)
                            {
                                resultCell.Select();
                                resultCell.Split(resultCount, 2);
                                for (int k = 0; k < resultCount; k++)
                                {
                                    //序号列的单元格
                                    Cell xuhaoCell = table.Cell(cRow + i + resultIndex + k, cCol + 2);
                                    xuhaoCell.SetWidth(26f, WdRulerStyle.wdAdjustFirstColumn);

                                }
                                for (int k = 0; k < resultCount; k++)
                                {
                                    Cell xuhaoCell = table.Cell(cRow + i + resultIndex + k, cCol + 2);
                                    xuhaoCell.Range.Text = "#" + (k + 1).ToString();


                                    SetResult(table.Cell(cRow + i + resultIndex + k, cCol + 2 + 1), resultList[k]["result"].ToString(), 2);

                                }
                                resultIndex = resultIndex + resultCount - 1;
                            }
                            else
                            {
                                if (resultList.First["result"].ToString().Equals("@", StringComparison.OrdinalIgnoreCase))
                                {
                                    var previousCell = table.Cell(cRow + i + resultIndex, cCol + 1);
                                    var rangText = previousCell.Range.Text;
                                    var textLength = previousCell.Range.Text.Length;
                                    resultCell.Select();
                                    if (textLength > 2 && rangText.Substring(0, textLength - 2).Length > 10)
                                    {
                                        _wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                                        _wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop;
                                    }

                                    resultCell.Merge(previousCell);
                                }
                                else
                                {
                                    Cell previous = null;
                                    try
                                    {
                                        previous = table.Cell(cRow + i + resultIndex - 1, cCol + 2);
                                    }
                                    catch (Exception)
                                    {
                                        //resultCell.Range.Text = resultList.First["result"].ToString();
                                    }

                                    SetResult(resultCell, resultList.First["result"].ToString(), 1);

                                    if (previous != null)
                                    {
                                        string previousText = previous.Range.Text;
                                        if (previousText.Replace("\r\a", "").Equals("$", StringComparison.OrdinalIgnoreCase))
                                        {
                                            previous.Range.Text = "";

                                            previous.Select();
                                            previous.Merge(resultCell);
                                        }
                                    }
                                }
                            }

                        }
                        cellCol7Dic.Add((cRow + i + resultIndex) + "," + (cCol + 1), secondItem);
                    }
                    incr = incr + secondItemsCount + resultIndex - 1;
                }
            }

            return cellCol7Dic;
        }
        //设置检验结果  1.merge 2.split
        private void SetResult(Cell tCell, string tString, int resultType)
        {
            if (tString.Contains("～") || tString.Contains("~"))
            {
                if (_colSpan > 1 && resultType == 1 && System.Text.Encoding.Default.GetBytes(tString).Length > 12)
                {
                    tString= SetResult(tCell, tString);
                }
                else if (System.Text.Encoding.Default.GetBytes(tString).Length > 8)
                {
                    tString= SetResult(tCell, tString);
                }
            }

            tCell.Range.Text = tString;

            this.FindHtmlLabel(tCell.Range);
        }

        private string SetResult(Cell tCell, string tString)
        {
            var indexStr = string.Empty;
            indexStr = tString.Contains("～") ? "～" : "~";
            if (indexStr.Equals(string.Empty))
            {
                return tString;
            }
            string[] strArraySplit = tString.Split(Convert.ToChar(indexStr));
            tString = "";
            for (int k = 0; k < strArraySplit.Length; k++)
            {
                if (k == strArraySplit.Length - 1)
                {
                    tString += strArraySplit[k];
                }
                else
                {
                    tString += strArraySplit[k] + indexStr + "\n";
                }
            }
            tCell.Select();
            _wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            return tString;
        }


        /// <summary>
        /// 单元格加段 加内容
        /// </summary>
        /// <param name="bookmark"></param>
        /// <param name="content"></param>
        public virtual void AddCellLowerRightCornerContent(string bookmark, string content)
        {
            try
            {
                Range range = GetBookmarkRank(_currentWord, bookmark);
                range.Select();
                object unite = WdUnits.wdLine;
                _wordApp.Selection.EndKey(ref unite, ref _missing);
                _wordApp.Selection.TypeParagraph();
                _wordApp.Selection.TypeText(content);
                _wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            }
            catch (Exception ex)
            {
                _needWrite = false;
                Dispose();
                throw new Exception($"错误信息:{ex.StackTrace}.{ex.Message}");
            }

        }

        /// <summary>
        /// 添加附表
        /// </summary>
        public virtual string AddAttachTable(string title, JArray array, string bookmark)
        {
            try
            {
                Range tableRange = GetBookmarkRank(_currentWord, bookmark);
                Table table = tableRange.Tables[1];
                table.Cell(1, 1).Select();
                _wordApp.Selection.InsertRowsBelow(1);
                _wordApp.Selection.Font.NameFarEast = "宋体";
                _wordApp.Selection.Font.NameAscii = "宋体";
                _wordApp.Selection.Font.NameOther = "宋体";
                _wordApp.Selection.Cells.Merge();
                _wordApp.Selection.Range.Text = title;

                int rowIndex = 2;

                for (int i = array.Count - 1; i >= 0; i--)
                {
                    JObject item = (JObject)array[i];
                    table.Cell(rowIndex, 1).Select();
                    _wordApp.Selection.InsertRowsBelow(1);
                    _wordApp.Selection.Font.NameFarEast = "宋体";
                    _wordApp.Selection.Font.NameAscii = "宋体";
                    _wordApp.Selection.Font.NameOther = "宋体";
                    Cell cell = table.Cell(rowIndex + 1, 1);
                    cell.Select();
                    int itemCount = item["isTitle"] != null ? item.Count - 1 : item.Count;
                    cell.Split(1, itemCount);
                    int cellColumnIndex = 1;
                    foreach (var item2 in item)
                    {
                        if (!item2.Key.Equals("isTitle"))
                        {
                            Cell tempCell = table.Cell(rowIndex + 1, cellColumnIndex);
                            tempCell.Range.Text = item2.Value.ToString();
                            if (cellColumnIndex == 3)
                            {
                                tempCell.SetWidth(64f, WdRulerStyle.wdAdjustSameWidth);
                            }
                            else if (cellColumnIndex == 5)
                            {
                                tempCell.SetWidth(80f, WdRulerStyle.wdAdjustSameWidth);
                            }

                            cellColumnIndex++;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _needWrite = false;
                Dispose();
                throw new Exception($"错误信息:{ex.StackTrace},{ex.Message}");
            }
            return "创建成功";
        }

        /// <summary>
        /// 照片和说明
        /// </summary>
        /// <param name="list"></param>
        /// <param name="bookmark"></param>
        /// <returns></returns>
        public virtual string InsertPhotoToWord(List<string> list, string bookmark)
        {
            try
            {
                int listCount = list.Count;
                if (listCount == 0)
                {
                    return "没有照片";
                }
                Range range = GetBookmarkRank(_currentWord, bookmark);
                Table table = range.Tables[1];
                float tableWidth = 0f;
                foreach (Column item in table.Columns)
                {
                    tableWidth += item.Width;
                }
                table.Cell(1, 1).Select();

                int rowIndex = 1;
                string frontStr = "№";

                for (int i = 0; i < listCount; i++)
                {
                    _wordApp.Selection.InsertRowsBelow(1);
                    rowIndex = rowIndex + i + 1;
                    Cell currentCell = table.Cell(rowIndex, 1);
                    currentCell.SetHeight(270.603f, WdRowHeightRule.wdRowHeightAtLeast);
                    string[] arrStr = list[i].Split(',');
                    string fileName = arrStr[0];
                    string content = arrStr[1];
                    if (!fileName.Equals(""))
                    {
                        AddPictureForStandard(fileName, _currentWord, currentCell.Range, tableWidth - 40, tableWidth - 240);
                    }
                    string templateStr = frontStr + (i + 1).ToString();
                    CreateAndGoToNextParagraph(currentCell.Range, true, false);
                    currentCell.Range.InsertAfter(templateStr + content);
                }
                table.Select();
                _wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                _wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                table.Cell(1, 1).Select();
                _wordApp.Selection.Rows.HeadingFormat = -1;
            }
            catch (Exception ex)
            {
                _needWrite = false;
                Dispose();
                throw new Exception($"错误信息:{ex.StackTrace}.{ex.Message}");
            }

            return "创建成功";
        }

        /// <summary>
        /// 判断图片大小
        /// </summary>
        private void AddPictureForStandard(string picFileName, Document doc, Range range, float width = 0,
            float height = 0)
        {
            float imageWidth;
            float imageHeight;
            using (FileStream fs = new FileStream(picFileName, FileMode.Open, FileAccess.Read))
            {
                System.Drawing.Image sourceImage = System.Drawing.Image.FromStream(fs);
                imageWidth = float.Parse(sourceImage.Width.ToString());
                imageHeight = float.Parse(sourceImage.Height.ToString());
            }

            InlineShape image = doc.InlineShapes.AddPicture(picFileName, ref _missing, ref _missing, range);

            if (imageWidth > width && imageWidth > imageHeight)
            {

                if (imageHeight * (width / imageWidth) > height)
                {
                    image.Height = height;
                    image.Width = imageWidth * (height / imageHeight);
                }
                else
                {
                    image.Width = width;
                    image.Height = imageHeight * (width / imageWidth);
                }

            }

            else if (imageHeight > height && imageHeight >= imageWidth)
            {
                if (imageWidth * (height / imageHeight) > width)
                {
                    image.Width = width;
                    image.Height = imageHeight * (width / imageWidth);
                }
                else
                {
                    image.Height = height;
                    image.Width = imageWidth * (height / imageHeight);
                }

            }
        }

        /// <summary>
        /// 移除图片和说明
        /// </summary>
        /// <param name="bookmark"></param>
        /// <returns></returns>
        public virtual string RemovePhotoTable(string bookmark)
        {
            Range range = GetBookmarkRank(_currentWord, bookmark);

            Table table = range.Tables[1];
            table.Select();
            _wordApp.Selection.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = true;
            table.Delete();
            _wordApp.Selection.TypeBackspace();
            _wordApp.Selection.Delete(WdUnits.wdCharacter, 1);
            return "成功";
        }

        /// <summary>
        /// 样品名称添加一行
        /// </summary>
        public void TableAddRowForY(string bookmark, string value)
        {
            Range range = GetBookmarkRank(_currentWord, bookmark);
            Cell cell = range.Cells[1];
            cell.Select();
            _wordApp.Selection.InsertRowsBelow(1);
            _wordApp.Selection.Cells[2].Range.Text = value;
        }

        #endregion

    }
}