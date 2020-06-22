using EmcReportWebApi.Utils;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;

namespace EmcReportWebApi.Business.ImplWordUtil
{
    /// <summary>
    /// 报告实现word工具类
    /// </summary>
    public class ReportHandleWord : WordUtil
    {

        /// <summary>
        /// 打开现有文件操作
        /// </summary>
        /// <param name="fileFullName">需保存文件的路径</param>
        public ReportHandleWord(string fileFullName) : base(fileFullName)
        {

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

        #region emc报告业务相关 

        /// <summary>
        /// 将html内容导入模板
        /// </summary>
        /// <returns></returns>
        public virtual string CopyHtmlContentToTemplate(string htmlFilePath, string TemplateFilePath, string bookmark, bool isNeedBreak, bool isCloseTheFile, bool isCloseTemplateFile)
        {
            try
            {
                Document htmlDoc = OpenWord(htmlFilePath);
                htmlDoc.Select();
                htmlDoc.Content.Copy();

                Document templateDoc = OpenWord(TemplateFilePath);
                templateDoc.Select();
                Range range = GetBookmarkRank(templateDoc, bookmark);
                range.Select();
                int tableCount = range.Tables.Count;

                Range tableRange = range.Tables[tableCount].Range;

                CreateAndGoToNextParagraph(tableRange, true, true);
                CreateAndGoToNextParagraph(tableRange, true, true);
                tableRange.Paste();

                foreach (Table item in range.Tables)
                {
                    item.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
                }
                if (isCloseTemplateFile)
                {
                    CloseWord(templateDoc, TemplateFilePath);
                }
                if (isCloseTheFile)
                    CloseWord(htmlDoc, htmlFilePath);
            }
            catch (Exception ex)
            {

                _needWrite = false;
                Dispose();
                throw new Exception(string.Format("错误信息:{0}.{1}", ex.StackTrace.ToString(), ex.Message));
            }

            return "创建成功";
        }

        /// <summary>
        /// 在书签位置插入内容
        /// </summary>
        /// <returns></returns>
        public virtual string InsertContentInBookmark(string fileFullPath, string content, string bookmark, bool isCloseTheFile = true)
        {
            try
            {
                Document openWord = OpenWord(fileFullPath);
                Range range = GetBookmarkRank(openWord, bookmark);
                range.InsertAfter(content);
                if (isCloseTheFile)
                    CloseWord(openWord, fileFullPath);
            }
            catch (Exception ex)
            {
                _needWrite = false;
                Dispose();
                throw new Exception(string.Format("错误信息:{0}.{1}", ex.StackTrace.ToString(), ex.Message));
            }

            return "插入成功";
        }

        /// <summary>
        /// 向table中插入list(不需要合并单元格)
        /// </summary>
        /// <param name="list"></param>
        /// <param name="bookmark"></param>
        /// <returns></returns>
        public virtual string InsertListToTable(List<string> list, string bookmark)
        {
            Range range = GetBookmarkRank(_currentWord, bookmark);
            Table table = range.Tables[1];
            int tableRowIndex = table.Rows.Count;
            int listCount = list.Count;
            for (int i = 0; i < listCount; i++)
            {

                if (i != 0)
                {
                    table.Cell(tableRowIndex, 1).Select();
                    _wordApp.Selection.InsertRowsBelow(1);
                    tableRowIndex++;
                }

                string[] arrStr = list[i].Split(',');
                for (int j = 0; j < arrStr.Length; j++)
                {
                    table.Cell(tableRowIndex, j + 1).Range.Text = arrStr[j];
                }
            }

            return "保存成功";
        }

        /// <summary>
        /// 向table中插入list
        /// </summary>
        /// <param name="list">内容集合</param>
        /// <param name="bookmark">插入内容的位置 用bookmark获取</param>
        /// <param name="mergeColumn">需要合并的列</param>
        /// <param name="isNeedNumber">是否需要添加序号</param>
        /// <returns></returns>
        public virtual string InsertListToTable(List<string> list, string bookmark, int mergeColumn, bool isNeedNumber = true)
        {
            if (mergeColumn < 1)
            {
                return "合并列不能小于1";
            }
            try
            {
                //获取bookmark位置的table
                Range range = GetBookmarkRank(_currentWord, bookmark);
                range.Select();
                Table table = range.Tables[1];
                int rowCount = table.Rows.Count;
                int columnCount = table.Columns.Count;
                //设置合并第二列相邻的相同内容

                int startRow = 0;
                int endRow = 0;
                string mergeContent = "";
                string nextColumnStr = "tempStr";
                bool isAddRow = false;

                foreach (var item in list)
                {
                    string[] arrStr = item.Split(',');
                    if (table.Columns.Count != arrStr.Length)
                    {
                        return "列和list集合不匹配";
                    }
                    if (isAddRow)
                    {
                        table.Rows.Add(ref _missing);
                        rowCount++;
                    }
                    isAddRow = true;

                    for (int i = 0; i < arrStr.Length; i++)
                    {
                        if (i == mergeColumn && arrStr[i].Equals(""))
                        {
                            nextColumnStr = arrStr[i];
                            endRow = rowCount;
                        }

                        if (i == mergeColumn - 1)
                        {

                            if (mergeContent == arrStr[i])
                            {
                                endRow = rowCount;
                            }
                            else
                            {
                                if (endRow != 0)
                                {
                                    //备注
                                    if (startRow != endRow)
                                    {
                                        string tempText = table.Cell(startRow, columnCount).Range.Text;
                                        MergeCell(table, startRow, columnCount, endRow, columnCount);
                                        table.Select();
                                        table.Cell(startRow, columnCount).Range.Text = tempText;
                                    }

                                    MergeCell(table, startRow, i + 1, endRow, i + (nextColumnStr.Equals("") ? 2 : 1));
                                    //合并序号列
                                    if (startRow != endRow)
                                    {
                                        MergeCell(table, startRow, 1, endRow, 1);

                                    }

                                    endRow = 0;
                                    nextColumnStr = "tempStr";
                                }
                                mergeContent = arrStr[i];
                                startRow = rowCount;
                            }
                            table.Cell(startRow, i + 1).Range.Text = arrStr[i];
                        }

                        else
                        {
                            table.Cell(rowCount, i + 1).Range.Text = arrStr[i];
                        }
                    }
                }

                //判断最后一行是否需要合并
                if (endRow != 0)
                {
                    MergeCell(table, startRow, mergeColumn, endRow, mergeColumn - 1 + (nextColumnStr.Equals("") ? 2 : 1));
                    //合并序号列
                    if (startRow != endRow)
                        MergeCell(table, startRow, 1, endRow, 1);
                }

                //写序号
                if (isNeedNumber)
                    AddTableNumber(table, 1);

                SetAutoFitContentForTable(table);
            }
            catch (Exception ex)
            {
                _needWrite = false;
                Dispose();
                throw new Exception($"错误信息:{ex.StackTrace}.{ex.Message}");
            }


            return "保存成功";
        }

        /// <summary>
        /// 样品连接图
        /// </summary>
        /// <returns></returns>
        public string InsertImageToWord(List<string> list, string bookmark)
        {
            try
            {
                //获取bookmark位置的table
                Range range = GetBookmarkRank(_currentWord, bookmark);
                range.Select();
                foreach (var item in list)
                {
                    string[] arrStr = item.Split(',');
                    string content = arrStr[0];
                    string fileName = arrStr[1];
                    CreateAndGoToNextParagraph(range, true, true);
                    range.InsertAfter(content);
                    CreateAndGoToNextParagraph(range, true, true);
                    AddPicture(fileName, range.Application.ActiveDocument, range);
                }
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
        /// 样品图片用到的
        /// </summary>
        /// <param name="list"></param>
        /// <param name="bookmark"></param>
        /// <returns></returns>
        public virtual string InsertImageToWordSample(List<string> list, string bookmark)
        {
            try
            {
                Range range = GetBookmarkRank(_currentWord, bookmark);

                int listCount = list.Count;
                //创建表格
                range.Select();
                Table table = _wordApp.Selection.Tables.Add(range, listCount, 1, ref _missing, ref _missing);
                float tableWidth = 0f;
                foreach (Column item in table.Columns)
                {
                    tableWidth += item.Width;
                }

                string frontStr = "№";

                for (int i = 0; i < listCount; i++)
                {
                    string[] arrStr = list[i].Split(',');
                    string fileName = arrStr[0];
                    string content = arrStr[1];
                    table.Select();
                    Range cellRange = _wordApp.Selection.Cells[i + 1].Range;
                    cellRange.Select();

                    if (!fileName.Equals(""))
                    {
                        AddPicture(fileName, _currentWord, cellRange, tableWidth - 56, tableWidth - 280);
                    }
                    string templateStr = frontStr + (i + 1).ToString();
                    CreateAndGoToNextParagraph(cellRange, true, false);
                    cellRange.InsertAfter(templateStr + content);
                }
                table.Select();
                //设置table格式
                table.Borders.Enable = (int)WdLineStyle.wdLineStyleSingle;
                _wordApp.Selection.SelectCell();
                _wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                _wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
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
        /// 将图片插入模板文件
        /// </summary>
        /// <returns></returns>
        public virtual string InsertConnectionImageToTemplate(string fileFullPath, List<string> list, string bookmark, bool isCloseTheFile = true)
        {
            try
            {
                Document doc = OpenWord(fileFullPath);
                Range range = GetBookmarkRank(doc, bookmark);

                int listCount = list.Count;
                //创建表格
                range.Select();
                Table table = _wordApp.Selection.Tables.Add(range, listCount, 1, ref _missing, ref _missing);
                float tableWidth = 0f;
                foreach (Column item in table.Columns)
                {
                    tableWidth += item.Width;
                }

                for (int i = 0; i < listCount; i++)
                {
                    string[] arrStr = list[i].Split(',');
                    string fileName = arrStr[0];
                    string content = arrStr[1];
                    table.Select();
                    Range cellRange = _wordApp.Selection.Cells[i + 1].Range;
                    cellRange.Select();

                    if (!fileName.Equals(""))
                    {
                            InlineShape image = AddPicture(fileName, doc, cellRange, tableWidth - 56, tableWidth - 280);
                    }
                    //CreateAndGoToNextParagraph(cellRange, true, false);
                    //cellRange.InsertAfter(content);
                }
                table.Select();
                //设置table格式
                _wordApp.Selection.SelectCell();
                _wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                _wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                if (isCloseTheFile)
                    CloseWord(doc, fileFullPath);
            }
            catch (Exception ex)
            {
                _needWrite = false;
                Dispose();
                throw new Exception($"错误信息:{ex.StackTrace}.{ex.Message}");
            }

            return "插入图片成功";
        }

        /// <summary>
        /// 将图片插入模板文件
        /// </summary>
        /// <returns></returns>
        public virtual string InsertImageToTemplate(string fileFullPath, List<string> list, string bookmark, bool isCloseTheFile = true)
        {
            try
            {
                Document doc = OpenWord(fileFullPath);
                Range range = GetBookmarkRank(doc, bookmark);

                int listCount = list.Count;
                int rowCount = listCount / 2;
                int columnCount = 2;
                if (listCount % 2 != 0)
                {
                    rowCount++;
                }
                if (listCount == 1)
                    columnCount = 1;
                //创建表格
                range.Select();
                Table table = _wordApp.Selection.Tables.Add(range, rowCount, columnCount, ref _missing, ref _missing);
                float tableWidth = 0f;
                foreach (Column item in table.Columns)
                {
                    tableWidth += item.Width;
                }

                for (int i = 0; i < listCount; i++)
                {
                    string[] arrStr = list[i].Split(',');
                    string fileName = arrStr[0];
                    string content = arrStr[1];
                    table.Select();
                    Range cellRange = _wordApp.Selection.Cells[i + 1].Range;
                    cellRange.Select();

                    if (!fileName.Equals(""))
                    {
                        if (columnCount == 1)
                        {
                            InlineShape image = AddPicture(fileName, doc, cellRange, tableWidth - 56, tableWidth - 280);
                        }
                        else
                        {
                            InlineShape image = AddPicture(fileName, doc, cellRange, tableWidth / 2 - 33, tableWidth / 2 - 66);
                        }
                    }
                    CreateAndGoToNextParagraph(cellRange, true, false);
                    cellRange.InsertAfter(content);
                }
                table.Select();
                //设置table格式
                _wordApp.Selection.SelectCell();
                _wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                _wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                if (isCloseTheFile)
                    CloseWord(doc, fileFullPath);
            }
            catch (Exception ex)
            {
                _needWrite = false;
                Dispose();
                throw new Exception($"错误信息:{ex.StackTrace}.{ex.Message}");
            }

            return "插入图片成功";
        }
        
        /// <summary>
        /// 复制其他文件内容到当前word并创建一个新的书签
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="bookmark"></param>
        /// <param name="isNewBookmark"></param>
        /// <param name="isCloseTheFile"></param>
        /// <returns></returns>
        public virtual string CopyOtherFileContentToWordReturnBookmark(string filePath, string bookmark, bool isNewBookmark, bool isCloseTheFile = true)
        {
            string newBookmark = "bookmark" + DateTime.Now.ToString("yyyyMMddHHmmssfff");
            try
            {
                Document htmlDoc = OpenWord(filePath) ?? throw new ArgumentNullException(nameof(filePath));
                if (isNewBookmark)
                {
                    Range rangeContent = htmlDoc.Content;
                    rangeContent.Select();
                    object unite = WdUnits.wdStory;
                    _wordApp.Selection.EndKey(ref unite, ref _missing);
                    object breakPage = WdBreakType.wdPageBreak;//分页符
                    _wordApp.ActiveWindow.Selection.InsertBreak(breakPage);
                    rangeContent = rangeContent.Sections.Last.Range;
                    CreateAndGoToNextParagraph(rangeContent, false, true);
                    rangeContent.Select();
                    _wordApp.Selection.Bookmarks.Add(newBookmark, rangeContent);
                }
                Range range = GetBookmarkRank(_currentWord, bookmark);
                htmlDoc.Content.Copy();
                range.Paste();
                range.Select();
                int tableCount = _wordApp.Selection.Tables.Count;
                if (tableCount > 0)
                {
                    foreach (Table table in _wordApp.Selection.Tables)
                    {
                        table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
                    }
                }

                if (isCloseTheFile)
                    CloseWord(htmlDoc, filePath);
            }
            catch (Exception ex)
            {
                _needWrite = false;
                Dispose();
                throw new Exception($"错误信息:{ex.StackTrace}.{ex.Message}");
            }

            return newBookmark;
        }

        /// <summary>
        /// 插入实验数据后设置word的格式
        /// </summary>
        /// <param name="intbackspace"></param>
        /// <returns></returns>
        public virtual string FormatCurrentWord(int intbackspace)
        {
            intbackspace = intbackspace - 1;
            _currentWord.Content.Select();
            object unite = WdUnits.wdSection;
            _wordApp.Selection.Expand(unite);
            Range range = GetBookmarkRank(_currentWord, "experimentEnd");
            range.Select();
            unite = WdUnits.wdParagraph;
            _wordApp.Selection.MoveUp(ref unite, 1);
            for (int i = 1; i <= intbackspace; i++)
            {
                _wordApp.Selection.TypeBackspace();
            }
            return "修改成功";
        }

        #region rtf操作

        /// <summary>
        /// 插入其他文件的单一表格 保留表格列
        /// </summary>
        /// <param name="copyFileFullPath">其他文件的路径</param>
        /// <param name="copyFileTableIndex">文件中表格的序列</param>
        /// <param name="copyTableColDic">保留表格的列并替换表头</param>
        /// <param name="wordBookmark">需要插入内容的书签</param>
        /// <param name="isCloseTheFile">是否关闭新打开的文件</param>
        /// <returns></returns>
        public virtual string CopyOtherFileTableForCol(string copyFileFullPath, int copyFileTableIndex, Dictionary<int, string> copyTableColDic, string wordBookmark, bool isCloseTheFile)
        {
            try
            {
                Document rtfDoc = OpenWord(copyFileFullPath, true);

                Table copyTable = rtfDoc.Tables[copyFileTableIndex];

                int copyTableColCount = copyTable.Columns.Count;

                object wdDeleteCellsRow = WdDeleteCells.wdDeleteCellsEntireRow;
                object wdDeleteCellsCol = WdDeleteCells.wdDeleteCellsEntireColumn;

                List<int> removeCols = new List<int>();
                for (int i = copyTableColCount; i >= 1; i--)
                {
                    if (!copyTableColDic.ContainsKey(i))
                    {
                        copyTable.Cell(1, i).Delete(ref wdDeleteCellsCol);
                    }
                    else
                    {
                        copyTable.Cell(1, i).Range.Text = copyTableColDic[i];
                    }
                }

                copyTable.Select();
                copyTable.Range.Copy();

                Range wordTable = GetBookmarkRank(_currentWord, wordBookmark);
                wordTable.Paste();
                ClearFormatTable(wordTable.Tables[1]);
                if (isCloseTheFile)
                    CloseWord(rtfDoc, copyFileFullPath);
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
        /// 创建table
        /// </summary>
        /// <param name="contentList"></param>
        /// <param name="bookmark"></param>
        /// <param name="isNeedBreak"></param>
        /// <returns></returns>
        public virtual string CreateTableToWord(List<string> contentList, string bookmark, bool isNeedBreak)
        {
            return CreateTableToWord(_currentWord, contentList, bookmark, isNeedBreak);
        }

        /// <summary>
        /// 创建table
        /// </summary>
        /// <param name="otherFileFullName"></param>
        /// <param name="contentList"></param>
        /// <param name="bookmark"></param>
        /// <param name="isCloseTemplateFile"></param>
        /// <param name="isNeedBreak"></param>
        /// <param name="isCloseTheFile"></param>
        /// <returns></returns>
        public string CreateTableToWord(string otherFileFullName, List<string> contentList, string bookmark, bool isCloseTemplateFile, bool isNeedBreak, bool isCloseTheFile = true)
        {
            Document otherFileDoc = OpenWord(otherFileFullName);
            if (isCloseTemplateFile)
                CloseWord(otherFileDoc);
            return CreateTableToWord(otherFileDoc, contentList, bookmark, isNeedBreak, isCloseTheFile);
        }

        /// <summary>
        /// 
        /// </summary>
        private string CreateTableToWord(Document doc, List<string> contentList, string bookmark, bool isNeedBreak, bool isCloseTheFile = true)
        {
            Range table = GetBookmarkRank(doc, bookmark);
            table.Select();
            if (isNeedBreak)
            {
                //InsertBreakPage(true);
                object unite = WdUnits.wdStory;
                _wordApp.Selection.EndKey(ref unite, ref _missing);
                object breakPage = WdBreakType.wdSectionBreakNextPage;//分页符
                _wordApp.Selection.InsertBreak(breakPage);

                table = _wordApp.Selection.Range.Sections.Last.Range;
                //CreateAndGoToNextParagraph(table, true, true);
                //CreateAndGoToNextParagraph(table, true, true);
            }
            int numRows = 1;
            int numColumns = 2;
            switch (contentList.Count)
            {
                case 3:
                case 4:
                    numRows = 2;
                    numColumns = 2;
                    break;
                case 5:
                case 6:
                    numRows = 3;
                    numColumns = 2;
                    break;

            }
            table.Select();
            _wordApp.Selection.Tables.Add(table, numRows, numColumns, ref _missing, ref _missing);
            //设置表格格式
            Table table1 = table.Tables[1];
            SetTabelFormat(table1);
            for (int i = 0; i < contentList.Count; i++)
            {
                table1.Select();
                _wordApp.Selection.Cells[i + 1].Range.Text = contentList[i];
            }

            //合并最后一行
            if (contentList.Count % 2 != 0)
            {
                MergeCell(table1, numRows, 1, numRows, numColumns);
            }

            return "创建成功";
        }

        /// <summary>
        /// 插入其他文件的表格 保留表格列
        /// </summary>
        public string CopyOtherFileTableForColByTableIndex(string copyFileFullPath, int copyFileTableStartIndex, int copyFileTableEndIndex, Dictionary<int, string> copyTableColDic, string wordBookmark, int titleRow, string mainTitle, bool isNeedBreak, bool isCloseTheFile = true)
        {
            string result;
            try
            {
                Document rtfDoc = OpenWord(copyFileFullPath, true);
                result = CopyOtherFileTableForColByTableIndex(_currentWord, rtfDoc, copyFileTableStartIndex, copyFileTableEndIndex, copyTableColDic, wordBookmark, titleRow, mainTitle, isNeedBreak);
                if (isCloseTheFile)
                    CloseWord(rtfDoc, copyFileFullPath);

            }
            catch (Exception ex)
            {
                _needWrite = false;
                Dispose();
                throw new Exception($"错误信息:{ex.StackTrace}.{ex.Message}");
            }

            return result;
        }


        /// <summary>
        /// 从其他文件取内容到word
        /// </summary>
        public string CopyOtherFileTableForColByTableIndex(string templateFullPath, string copyFileFullPath, int copyFileTableStartIndex, int copyFileTableEndIndex, Dictionary<int, string> copyTableColDic, string wordBookmark, int titleRow, string mainTitle, bool isCloseTemplateFile, bool isNeedBreak, bool isCloseTheFile = true)
        {
            try
            {
                Document templateDoc = OpenWord(templateFullPath);
                Document rtfDoc = OpenWord(copyFileFullPath, true);
                CopyOtherFileTableForColByTableIndex(templateDoc, rtfDoc, copyFileTableStartIndex, copyFileTableEndIndex, copyTableColDic, wordBookmark, titleRow, mainTitle, isNeedBreak);
                if (isCloseTemplateFile)
                    CloseWord(templateDoc, templateFullPath);
                if (isCloseTheFile)
                    CloseWord(rtfDoc, copyFileFullPath);
            }
            catch (Exception ex)
            {
                _needWrite = false;
                Dispose();
                //throw new Exception("rtf文件内容不正确");
                throw new Exception($"错误信息:{ex.StackTrace}.{ex.Message}");
            }

            return "创建成功";
        }


        /// <summary>
        /// 
        /// </summary>
        private string CopyOtherFileTableForColByTableIndex(Document templateDoc, Document rtfDoc, int copyFileTableStartIndex, int copyFileTableEndIndex, Dictionary<int, string> copyTableColDic, string wordBookmark, int titleRow, string mainTitle, bool isNeedBreak)
        {
            try
            {
                //Document rtfDoc = OpenWord(copyFileFullPath, true);

                int rtfTableCount = rtfDoc.Tables.Count;

                Range wordTable = GetBookmarkRank(templateDoc, wordBookmark);
                wordTable.Select();
                if (isNeedBreak)
                {
                    object unite = WdUnits.wdStory;
                    _wordApp.Selection.EndKey(ref unite, ref _missing);

                    object breakType = WdBreakType.wdLineBreak;//换行符
                    _wordApp.ActiveWindow.Selection.InsertBreak(breakType);

                    wordTable = _wordApp.Selection.Range.Sections.Last.Range;
                }

                //判断主表头是否为null
                string[] mainTitleArray = null;
                if (!mainTitle.Equals(""))
                {
                    mainTitleArray = mainTitle.Split(',');
                }
                int m = 0;

                int forCount = copyFileTableEndIndex == 0 ? rtfTableCount : copyFileTableEndIndex;

                List<int> removeCols = new List<int>();
                for (int i = copyFileTableStartIndex; i <= forCount; i++)
                {
                    Table copyTable = rtfDoc.Tables[i];

                    int copyTableColCount = copyTable.Columns.Count;

                    object wdDeleteCellsCol = WdDeleteCells.wdDeleteCellsEntireColumn;

                    for (int j = copyTableColCount; j >= 1; j--)
                    {
                        if (!copyTableColDic.ContainsKey(j))
                        {
                            copyTable.Cell(titleRow, j).Delete(ref wdDeleteCellsCol);
                        }
                        else
                        {
                            copyTable.Cell(titleRow, j).Range.Text = copyTableColDic[j];
                        }
                    }
                    if (mainTitleArray != null && mainTitleArray[m] != null)
                    {
                        copyTable.Cell(1, 1).Range.Text = mainTitleArray[m];
                        m++;
                    }

                    if (i != copyFileTableStartIndex)
                        CreateAndGoToNextParagraph(wordTable, (i != copyFileTableStartIndex) || isNeedBreak, (i != copyFileTableStartIndex) || isNeedBreak);//获取下一个range
                    CreateAndGoToNextParagraph(wordTable, (i != copyFileTableStartIndex) || isNeedBreak, (i != copyFileTableStartIndex) || isNeedBreak);//InsertBR(wordTable, i <= rtfTableCount);//添加回车
                    copyTable.Range.Copy();
                    wordTable.Paste();

                    //检索table的最后一列 耗时间 找更好的办法
                    //Table table1 = wordTable.Tables[1];
                    //table1.Select();
                    //Cells cells = _wordApp.Selection.Cells;
                    //foreach (Cell item in cells)
                    //{
                    //    if (item.Range.Text.Contains("PASS")) {
                    //        item.Range.Text = "符合";
                    //    }
                    //}

                    ClearFormatTable(wordTable.Tables[1]);
                }

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
        /// 从其他文件取图片到word
        /// </summary>
        public virtual string CopyOtherFilePictureToWord(string templalteFileFullName, string copyFileFullPath, int copyFilePictureStartIndex, string workBookmark, bool isCloseTemplateFile, bool isNeedBreak, bool isPage, bool isCloseTheFile = true)
        {
            string result;
            try
            {
                Document templateDoc = OpenWord(templalteFileFullName);
                Document copyFileDoc = OpenWord(copyFileFullPath, true);
                result = CopyOtherFilePictureToWord(templateDoc,
                    copyFileDoc,
                    copyFilePictureStartIndex,
                    workBookmark,
                    isNeedBreak,
                    isPage);
                if (isCloseTemplateFile)
                {
                    CloseWord(templateDoc, templalteFileFullName);
                }
                if (isCloseTheFile)
                    CloseWord(copyFileDoc, copyFileFullPath);
            }
            catch (Exception ex)
            {
                _needWrite = false;
                Dispose();
                throw new Exception($"错误信息:{ex.StackTrace}.{ex.Message}");
            }
            return result;
        }


        /// <summary>
        /// 插入其他文件的图片 从第几个图片开始
        /// </summary>
        /// <param name="copyFileFullPath">其他文件的路径</param>
        /// <param name="copyFilePictureStartIndex">开始的图片</param>
        /// <param name="workBookmark">插入文件的书签位置</param>
        /// <param name="isPage"></param>
        /// <param name="isCloseTheFile">是否关闭其他文件</param>
        /// <param name="isNeedBreak"></param>
        /// <returns></returns>
        public virtual string CopyOtherFilePictureToWord(string copyFileFullPath, int copyFilePictureStartIndex, string workBookmark, bool isNeedBreak, bool isPage, bool isCloseTheFile = true)
        {
            string result;
            try
            {
                Document copyFileDoc = OpenWord(copyFileFullPath, true);
                result = CopyOtherFilePictureToWord(_currentWord,
                    copyFileDoc,
                    copyFilePictureStartIndex,
                    workBookmark,
                    isNeedBreak,
                    isPage);
                if (isCloseTheFile)
                    CloseWord(copyFileDoc, copyFileFullPath);
            }
            catch (Exception ex)
            {
                _needWrite = false;
                Dispose();
                throw new Exception($"错误信息:{ex.StackTrace}.{ex.Message}");
            }
            return result;
        }

        /// <summary>
        /// 从其他文件取图片到word
        /// </summary>
        public virtual string CopyOtherFilePictureToWord(Document fileDoc, Document copyFileDoc, int copyFilePictureStartIndex, string workBookmark, bool isNeedBreak, bool isPage)
        {
            try
            {
                Range bookmarkPic = GetBookmarkRank(fileDoc, workBookmark);
                bookmarkPic.Select();
                if (isNeedBreak)
                {
                    InsertBreakPage(false);
                    bookmarkPic = _wordApp.Selection.Range.Sections.Last.Range;
                }

                copyFileDoc.Select();//选中当前文档进行操作
                int i = 1;
                foreach (InlineShape shape in copyFileDoc.InlineShapes)
                {
                    //判断类型
                    if (shape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        if (i >= copyFilePictureStartIndex)
                        {
                            //利用剪贴板保存数据
                            shape.Select(); //选定当前图片
                            shape.Range.Copy();
                            // WordApp.Selection.Copy();//copy当前图片


                            bookmarkPic.Paste();
                            CreateAndGoToNextParagraph(bookmarkPic, true, true);
                        }
                        i++;
                    }
                }
                if (isPage)
                {
                    InsertBreakPage(false);
                }

            }
            catch (Exception ex)
            {
                _needWrite = false;
                Dispose();
                throw new Exception($"错误信息:{ex.StackTrace}.{ex.Message}");
            }
            return "创建成功";
        }

        #endregion

        #region history 以后可能会用到

        /// <summary>
        /// 设置title并插入list 实验结果概述要用到
        /// </summary>
        /// <returns></returns>
        public string InsertListIntoTableByTitle(Dictionary<string, List<string>> dic, string bookmark, bool isNeedNumber = true)
        {

            Range range = GetBookmarkRank(_currentWord, bookmark);
            range.Select();
            Table table = range.Tables[1];
            table.Select();
            table.Rows.Add(ref _missing);
            int columnCount = 4;
            int rowCount = 2;
            //创建模板

            string mergeContent = "tempStr";
            string mergeContent2 = "tempStr";
            int startRow1 = 0;
            int endRow1 = 0;
            int startRow2 = 0;
            int endRow2 = 0;
            int dicFor = 0;
            foreach (var item in dic)
            {
                string title = item.Key;
                List<string> list = item.Value;
                if (dicFor != 0)
                {
                    table.Rows.Add(ref _missing);
                    table.Rows.Add(ref _missing);
                    rowCount = rowCount + 2;
                }
                //设置title
                table.Select();
                table.Cell(rowCount - 1, 1).Range.Text = title;
                //添加列头
                table.Cell(rowCount, 1).Range.Text = "条款";
                table.Cell(rowCount, 2).Range.Text = "项目";
                table.Cell(rowCount, 3).Range.Text = "实验结果";
                table.Cell(rowCount, 4).Range.Text = "备注";
                MergeCell(table, rowCount - 1, 1, rowCount - 1, columnCount);
                table.Cell(rowCount - 1, 1).Select();
                FontBoldLeft();

                foreach (var listItem in list)
                {
                    table.Select();
                    table.Rows.Add(ref _missing);
                    rowCount++;
                    string[] arrStr = listItem.Split(',');
                    for (int i = 0; i < arrStr.Length; i++)
                    {
                        if (i == 0)
                        {
                            if (mergeContent == arrStr[i])
                            {
                                endRow1 = rowCount;
                            }
                            else
                            {
                                if (endRow1 != 0)
                                {
                                    MergeCell(table, startRow1, i + 1, endRow1, i + 1);
                                    endRow1 = 0;
                                }
                                mergeContent = arrStr[i];
                                startRow1 = rowCount;
                            }
                        }
                        else if (i == 3)
                        {
                            if (mergeContent2 == arrStr[i])
                            {
                                endRow2 = rowCount;
                            }
                            else
                            {
                                if (endRow2 != 0)
                                {
                                    MergeCell(table, startRow2, i + 1, endRow2, i + 1);
                                    endRow2 = 0;
                                }
                                mergeContent2 = arrStr[i];
                                startRow2 = rowCount;
                            }
                        }

                        table.Cell(rowCount, i + 1).Range.Text = arrStr[i];
                    }
                }
                dicFor++;
            }
            return "保存成功";
        }

        #endregion
        #endregion
    }
}