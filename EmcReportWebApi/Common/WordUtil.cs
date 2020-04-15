using Microsoft.Office.Interop.Word;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace EmcReportWebApi.Common
{
    public class WordUtil : IDisposable
    {
        private Application _wordApp;//word用程序
        private Document _currentWord;//当前操作的word
        private bool _disposed;
        private bool _needWrite = false;
        private bool _isSaveAs = false;
        private string _outFilePath = "";

        private object _missing = System.Reflection.Missing.Value;
        private object _objFalse = false;
        private object _objTrue = true;
        private object wdReplaceAll = WdReplace.wdReplaceAll;//替换所有文字
        private object wdReplaceOne = WdReplace.wdReplaceOne;//替换第一个文字

        private Dictionary<string, Document> _fileDic;

        public WordUtil()
        {

        }
        public WordUtil(string fileFullName, bool isSaveAs)
        {
            NewApp();
        }

        /// <summary>
        /// 打开现有文件操作
        /// </summary>
        /// <param name="outFileFullName">生成文件路径</param>
        /// <param name="fileFullName">引用文件路径</param>
        /// <param name="isSaveAs">是否另存文件</param>
        public WordUtil(string outFileFullName, string fileFullName = "", bool isSaveAs = true)
        {
            if (outFileFullName.Equals(""))
                throw new Exception("输出文件路径不能为空");
            if (fileFullName.Equals(""))
                _currentWord = CreatWord();
            else
            {
                _currentWord = OpenWord(fileFullName);
            }

            _outFilePath = outFileFullName;
            _isSaveAs = isSaveAs;

            _disposed = false;
            _needWrite = true;
        }

        #region 标准报告业务相关
        
        /// <summary>
        /// 表格拆分
        /// </summary>
        /// <param name="array">拆分数组</param>
        /// <param name="bookmark">书签</param>
        /// <returns></returns>
        public int TableSplit(JArray array, string bookmark)
        {
            try
            {
                Range tableRange = GetBookmarkRank(_currentWord, bookmark);

                Table table = tableRange.Tables[1];

                for (int i = array.Count - 1; i >= 0; i--)
                {
                    TableSplit((JObject)array[i], i + 1, table);
                }

                table.Cell(1, 1).Select();
                _wordApp.Selection.Rows.HeadingFormat = -1;
                return 1;
            }
            catch (Exception ex)
            {
                _needWrite = false;
                Dispose();
                throw new Exception(string.Format("错误信息:{0}.{1}", ex.StackTrace.ToString(), ex.Message));
            }

        }

        /// <summary>
        /// 表格拆分
        /// </summary>
        /// <param name="jObject">拆分对象</param>
        /// <param name="xuhao">序号</param>
        /// <param name="table">需拆分表格</param>
        /// <returns></returns>
        public int TableSplit(JObject jObject, int xuhao, Table table)
        {

            table.Cell(1, 1).Select();
            _wordApp.Selection.InsertRowsBelow(1);

            // Row newRow = table.Rows[2];

            //序号
            table.Cell(2, 1).Range.Text = jObject["idxNo"].ToString();

            //检验项目
            table.Cell(2, 2).Range.Text = jObject["itemContent"].ToString();
            //单项结论
            if(jObject["comment"]!=null&& !jObject["comment"].Equals(""))
                table.Cell(2, 6).Range.Text = jObject["comment"].ToString();
            //备注
            if (jObject["reMark"] != null && !jObject["reMark"].Equals(""))
                table.Cell(2, 7).Range.Text = jObject["reMark"].ToString();

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

                Dictionary<JObject, string> cellCol4Dic = new Dictionary<JObject, string>();
                for (int i = 0; i < firstItemsCount; i++)
                {
                    Cell tempCell = table.Cell(2 + i, 4);
                    JObject firstItem = (JObject)firstItems[i];
                    tempCell.Range.Text = firstItem["stdName"].ToString();
                    cellCol4Dic.Add(firstItem, (2 + i).ToString() + "," + 4.ToString());

                }
                //遍历节点拆分单元格
                bool whilebool = true;
                while (whilebool) {
                    cellCol4Dic = AddCellAndSplit(table, cellCol4Dic);
                    whilebool = (cellCol4Dic.Count != 0);
                }
            }

            return 1;
        }
        /// <summary>
        /// 遍历节点拆分单元格
        /// </summary>
        /// <returns></returns>
        private Dictionary<JObject, string> AddCellAndSplit(Table table,Dictionary<JObject, string> cellCol6Dic) {
            Dictionary<JObject, string> cellCol7Dic = new Dictionary<JObject, string>();
            int incr = 0;
            foreach (var item in cellCol6Dic)
            {
                JObject j = item.Key;
                string c = item.Value;
                int cRow = int.Parse(c.Split(',')[0]) + incr;
                int cCol = int.Parse(c.Split(',')[1]);

                JArray secondItems = (JArray)j["list"];
                int secondItemsCount = secondItems.Count;
                if (secondItemsCount > 0)
                {
                    if (secondItemsCount != 1)
                        table.Cell(cRow, cCol + 1).Split(secondItemsCount, 1);
                    table.Cell(cRow, cCol).Split(secondItemsCount, 2);

                    if (secondItemsCount != 1)
                        table.Cell(cRow, cCol).Merge(table.Cell(cRow + secondItemsCount - 1, cCol));
                    // table.Cell(cRow, cCol).SetWidth(40f, WdRulerStyle.wdAdjustFirstColumn);//拆分单元格后设置列宽

                    table.Cell(cRow, cCol).PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints;
                    table.Cell(cRow, cCol).PreferredWidth = 40f;


                    for (int i = 0; i < secondItemsCount; i++)
                    {
                        Cell tempCell = table.Cell(cRow + i, cCol + 1);
                        JObject secondItem = (JObject)secondItems[i];
                        tempCell.Range.Text = secondItems[i]["itemContent"].ToString();
                        //检验结果
                        if (secondItems[i]["result"] != null && !secondItems[i]["result"].Equals("")) {
                            Cell resultCell = table.Cell(cRow + i, cCol + 2);
                            resultCell.Range.Text = secondItems[i]["result"].ToString();
                        }
                        cellCol7Dic.Add(secondItem, (cRow + i).ToString() + "," + (cCol + 1).ToString());
                    }
                    incr = incr + secondItemsCount - 1;
                }
            }

            return cellCol7Dic;
        }

        #endregion

        #region emc报告业务相关 

        /// <summary>
        /// 将html内容导入模板
        /// </summary>
        /// <returns></returns>
        public string CopyHtmlContentToTemplate(string htmlFilePath, string TemplateFilePath, string bookmark, bool isNeedBreak, bool isCloseTheFile, bool isCloseTemplateFile)
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
        public string InsertContentInBookmark(string fileFullPath, string content, string bookmark, bool isCloseTheFile = true)
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
        /// 向table中插入list
        /// </summary>
        /// <param name="list">内容集合</param>
        /// <param name="bookmark">插入内容的位置 用bookmark获取</param>
        /// <param name="mergeColumn">需要合并的列</param>
        /// <param name="isNeedNumber">是否需要添加序号</param>
        /// <returns></returns>
        public string InsertListToTable(List<string> list, string bookmark, int mergeColumn, bool isNeedNumber = true)
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
                throw new Exception(string.Format("错误信息:{0}.{1}", ex.StackTrace.ToString(), ex.Message));
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
                throw new Exception(string.Format("错误信息:{0}.{1}", ex.StackTrace.ToString(), ex.Message));
            }

            return "创建成功";
        }

        /// <summary>
        /// 样品图片用到的
        /// </summary>
        /// <param name="list"></param>
        /// <param name="bookmark"></param>
        /// <returns></returns>
        public string InsertImageToWordSample(List<string> list, string bookmark)
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
                        InlineShape image = AddPicture(fileName, _currentWord, cellRange, tableWidth - 56, tableWidth - 280);
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
                throw new Exception(string.Format("错误信息:{0}.{1}", ex.StackTrace.ToString(), ex.Message));
            }

            return "创建成功";
        }

        /// <summary>
        /// 将图片插入模板文件
        /// </summary>
        /// <returns></returns>
        public string InsertImageToTemplate(string fileFullPath, List<string> list, string bookmark, bool isCloseTheFile = true)
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
                throw new Exception(string.Format("错误信息:{0}.{1}", ex.StackTrace.ToString(), ex.Message));
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
        public string CopyOtherFileContentToWordReturnBookmark(string filePath, string bookmark, bool isNewBookmark, bool isCloseTheFile = true)
        {
            string newBookmark = "bookmark" + DateTime.Now.ToString("yyyyMMddhhmmss");
            try
            {
                Document htmldoc = OpenWord(filePath);
                if (isNewBookmark)
                {
                    Range rangeContent = htmldoc.Content;
                    rangeContent.Select();
                    InsertBreakPage(true);
                    rangeContent = rangeContent.Sections.Last.Range;
                    CreateAndGoToNextParagraph(rangeContent, false, true);
                    rangeContent.Select();
                    _wordApp.Selection.Bookmarks.Add(newBookmark, rangeContent);
                }
                Range range = GetBookmarkRank(_currentWord, bookmark);
                htmldoc.Content.Copy();
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
                    CloseWord(htmldoc, filePath);
            }
            catch (Exception ex)
            {
                _needWrite = false;
                Dispose();
                throw new Exception(string.Format("错误信息:{0}.{1}", ex.StackTrace.ToString(), ex.Message));
            }

            return newBookmark;
        }
        
        /// <summary>
        /// 插入实验数据后设置word的格式
        /// </summary>
        /// <param name="intbackspace"></param>
        /// <returns></returns>
        public string FormatCurrentWord(int intbackspace)
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
        public string CopyOtherFileTableForCol(string copyFileFullPath, int copyFileTableIndex, Dictionary<int, string> copyTableColDic, string wordBookmark, bool isCloseTheFile)
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
                throw new Exception(string.Format("错误信息:{0}.{1}", ex.StackTrace.ToString(), ex.Message));
            }

            return "创建成功";
        }
        public string CreateTableToWord(List<string> contentList, string bookmark, bool isNeedBreak)
        {
            return CreateTableToWord(_currentWord, contentList, bookmark, isNeedBreak);
        }

        public string CreateTableToWord(string otherFileFullName, List<string> contentList, string bookmark, bool isCloseTemplateFile, bool isNeedBreak, bool isCloseTheFile = true)
        {
            Document otherFileDoc = OpenWord(otherFileFullName);
            if (isCloseTemplateFile)
                CloseWord(otherFileDoc);
            return CreateTableToWord(otherFileDoc, contentList, bookmark, isNeedBreak, isCloseTheFile);
        }

        public string CreateTableToWord(Document doc, List<string> contentList, string bookmark, bool isNeedBreak, bool isCloseTheFile = true)
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
        /// <param name="copyFileFullPath">其他文件的路径</param>
        /// <param name="copyFileTableStartIndex">从第几个开始获取表格</param>
        /// <param name="copyTableColDic">保留表格的列并替换表头</param>
        /// <param name="wordBookmark">需要插入内容的书签</param>
        /// <param name="isCloseTheFile">是否关闭新打开的文件</param>
        /// <returns></returns>
        public string CopyOtherFileTableForColByTableIndex(string copyFileFullPath, int copyFileTableStartIndex, int copyFileTableEndIndex, Dictionary<int, string> copyTableColDic, string wordBookmark, int titleRow, string mainTitle, bool isNeedBreak, bool isCloseTheFile = true)
        {
            string result = "创建成功";
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
                throw new Exception(string.Format("错误信息:{0}.{1}", ex.StackTrace.ToString(), ex.Message));
            }

            return result;
        }


        public string CopyOtherFileTableForColByTableIndex(string templateFullPath, string copyFileFullPath, int copyFileTableStartIndex, int copyFileTableEndIndex, Dictionary<int, string> copyTableColDic, string wordBookmark, int titleRow, string mainTitle, bool isCloseTemplateFile, bool isNeedBreak, bool isCloseTheFile = true)
        {
            string result = "创建成功";
            try
            {
                Document templateDoc = OpenWord(templateFullPath);
                Document rtfDoc = OpenWord(copyFileFullPath, true);
                result = CopyOtherFileTableForColByTableIndex(templateDoc, rtfDoc, copyFileTableStartIndex, copyFileTableEndIndex, copyTableColDic, wordBookmark, titleRow, mainTitle, isNeedBreak);
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
                throw new Exception(string.Format("错误信息:{0}.{1}", ex.StackTrace.ToString(), ex.Message));
            }

            return "创建成功";
        }


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

                for (int i = copyFileTableStartIndex; i <= forCount; i++)
                {
                    Table copyTable = rtfDoc.Tables[i];

                    int copyTableColCount = copyTable.Columns.Count;

                    object wdDeleteCellsRow = WdDeleteCells.wdDeleteCellsEntireRow;
                    object wdDeleteCellsCol = WdDeleteCells.wdDeleteCellsEntireColumn;

                    List<int> removeCols = new List<int>();
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
                throw new Exception(string.Format("错误信息:{0}.{1}", ex.StackTrace.ToString(), ex.Message));
            }

            return "创建成功";
        }


        public string CopyOtherFilePictureToWord(string templalteFileFullName, string copyFileFullPath, int copyFilePictureStartIndex, string workBookmark, bool isCloseTemplateFile, bool isNeedBreak, bool isPage, bool isCloseTheFile = true)
        {
            string result = "创建失败";
            try
            {
                Document templateDoc = OpenWord(templalteFileFullName);
                Document copyFileDoc = OpenWord(copyFileFullPath, true);
                result = CopyOtherFilePictureToWord(templateDoc, copyFileDoc, copyFilePictureStartIndex, workBookmark, isNeedBreak, isPage);
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
                throw new Exception(string.Format("错误信息:{0}.{1}", ex.StackTrace.ToString(), ex.Message));
            }
            return "创建成功";
        }


        /// <summary>
        /// 插入其他文件的图片 从第几个图片开始
        /// </summary>
        /// <param name="copyFileFullPath">其他文件的路径</param>
        /// <param name="copyFilePictureStartIndex">开始的图片</param>
        /// <param name="workBookmark">插入文件的书签位置</param>
        /// <param name="isCloseTheFile">是否关闭其他文件</param>
        /// <returns></returns>
        public string CopyOtherFilePictureToWord(string copyFileFullPath, int copyFilePictureStartIndex, string workBookmark, bool isNeedBreak, bool isPage, bool isCloseTheFile = true)
        {
            string result = "创建失败";
            try
            {
                Document copyFileDoc = OpenWord(copyFileFullPath, true);
                result = CopyOtherFilePictureToWord(_currentWord, copyFileDoc, copyFilePictureStartIndex, workBookmark, isNeedBreak, isPage);
                if (isCloseTheFile)
                    CloseWord(copyFileDoc, copyFileFullPath);
            }
            catch (Exception ex)
            {
                _needWrite = false;
                Dispose();
                throw new Exception(string.Format("错误信息:{0}.{1}", ex.StackTrace.ToString(), ex.Message));
            }
            return "创建成功";
        }

        public string CopyOtherFilePictureToWord(Document fileDoc, Document copyFileDoc, int copyFilePictureStartIndex, string workBookmark, bool isNeedBreak, bool isPage)
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
                throw new Exception(string.Format("错误信息:{0}.{1}", ex.StackTrace.ToString(), ex.Message));
            }
            return "设置成功";
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

        #region 公共方法

        /// <summary>
        /// 获取操作word的页数
        /// </summary>
        /// <returns></returns>
        public int GetDocumnetPageCount()
        {
            return _currentWord.ComputeStatistics(WdStatistic.wdStatisticPages, ref _missing);
        }

        /// <summary>
        /// 复制其他文件的表格到当前word
        /// </summary>
        /// <returns></returns>
        public string CopyTableToWord(string otherFilePath, string bookmark, int tableIndex, bool isCloseTheFile)
        {
            try
            {
                Document otherFile = OpenWord(otherFilePath);
                Range range = GetBookmarkRank(_currentWord, bookmark);
                otherFile.Tables[tableIndex].Range.Copy();
                range.Paste();
                range.Tables[1].AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
                if (isCloseTheFile)
                    CloseWord(otherFile, otherFilePath);
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
        /// 复制其他文件的文本框到当前word
        /// </summary>
        /// <returns></returns>
        public string CopyImageToWord(string otherFilePath, string bookmark, bool isCloseTheFile)
        {
            try
            {

                Document otherFile = OpenWord(otherFilePath);
                otherFile.Select();

                Range bookmarkPic = GetBookmarkRank(_currentWord, bookmark);

                try
                {
                    ShapeRange shapeRange = otherFile.Shapes.Range(1);
                    InlineShape inlineShape = shapeRange.ConvertToInlineShape();
                    inlineShape.Select();
                    _wordApp.Selection.Copy();
                }
                catch (Exception)
                {
                    Shape shape = otherFile.Shapes[1];
                    Frame inlineShape = shape.ConvertToFrame();
                    inlineShape.Select();
                    _wordApp.Selection.Copy();
                }
                finally
                {
                    bookmarkPic.Paste();
                    if (isCloseTheFile)
                        CloseWord(otherFile, otherFilePath);
                }

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
        /// 根据书签向word中插入内容
        /// </summary>
        /// <returns></returns>
        public string InsertContentToWordByBookmark(string content, string bookmark)
        {
            try
            {
                Range range = GetBookmarkReturnNull(_currentWord, bookmark);
                if (range == null)
                    return "未找到书签:" + bookmark;
                range.Select();
                range.Text = content;
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
        /// 复制其他文件内容到当前文件
        /// </summary>
        /// <returns></returns>
        public string CopyOtherFileContentToWord(string filePth, string bookmark, bool isCloseTheFile = true)
        {
            try
            {
                Document htmldoc = OpenWord(filePth);
                htmldoc.Content.Copy();
                Range range = GetBookmarkRank(_currentWord, bookmark);
                range.Select();
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
                    CloseWord(htmldoc, filePth);
            }
            catch (Exception ex)
            {
                _needWrite = false;
                Dispose();
                throw new Exception(string.Format("错误信息:{0}.{1}", ex.StackTrace.ToString(), ex.Message));
            }

            return "保存成功";
        }

        /// <summary>
        /// 复制第二个文件内容到第一个文件
        /// </summary>
        /// <returns></returns>
        public string CopyOtherFileContentToWord(string firstFilePath, string secondFilePath, string bookmark, bool isCloseTheFile = true)
        {
            try
            {
                Document htmldoc = OpenWord(firstFilePath);
                Document secondFile = OpenWord(secondFilePath);
                Range range = GetBookmarkRank(secondFile, bookmark);
                htmldoc.Content.Copy();
                range.Select();
                range.PasteAndFormat(WdRecoveryType.wdPasteDefault);
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
                    CloseWord(htmldoc, firstFilePath);
            }
            catch (Exception ex)
            {
                _needWrite = false;
                Dispose();
                throw new Exception(string.Format("错误信息:{0}.{1}", ex.StackTrace.ToString(), ex.Message));
            }

            return "保存成功";
        }

        /// <summary>
        /// 替换文字
        /// </summary>
        /// <returns></returns>
        public string ReplaceWritten(Dictionary<int, Dictionary<string, string>> replaceDic, int replaceType = 2)
        {
            try
            {
                _currentWord.Content.Select();

                foreach (var item in replaceDic)
                {
                    int type = item.Key;
                    foreach (var itemV in item.Value)
                    {
                        Replace(type, itemV.Key, itemV.Value, replaceType);
                    }
                }
            }
            catch (Exception ex)
            {
                _needWrite = false;
                throw new Exception(string.Format("错误信息:{0}.{1}", ex.StackTrace.ToString(), ex.Message));
            }
            finally
            {
                Dispose();
            }

            return "替换成功";
        }

        /// <summary>
        /// word文档中插入excel
        /// </summary>
        /// <param name="excelFileFullName">excel文件</param>
        /// <param name="bookmark">书签名称 excel插入位置书签名称 默认为"" 插入文档末尾</param>
        /// <param name="excelWidht">excel宽</param>
        /// <param name="excelHeight">excel高</param>
        /// <returns>创建结果或错误信息</returns>
        public string InsertExcel(string excelFileFullName, string bookmark, int excelWidht = 450, int excelHeight = 200)
        {
            if (excelFileFullName.Equals(""))
            {
                throw new Exception("excel文件不能为空");
            }
            try
            {
                _currentWord.Select();

                object bk = bookmark;
                Range excelRange = null;

                if (_currentWord.Bookmarks.Exists(bookmark))
                {
                    excelRange = GetBookmarkRank(_currentWord, bookmark);
                }
                else
                {
                    InsertBreakPage(false);
                }

                object fileType = @"Excel.Sheet.12";//插入的excel 格式，HKEY_CLASSES_ROOT，所以是.12  Excel.Chart.8也可以
                object filename = excelFileFullName;//插入的excel的位置
                object linkToFile = true;
                object rangeOLE = excelRange;
                //添加一个OLEObject对象
                InlineShape sp = _wordApp.Selection.InlineShapes.AddOLEObject(ref fileType,
                    ref filename,
                    ref _missing,//真 若要将 OLE 对象链接到创建它的文件
                    ref _missing,//真 图标
                    ref _missing,//图标链接
                    ref _missing,
                    ref _missing,
                    ref rangeOLE);//位置
                sp.Height = excelHeight;//200
                sp.Width = excelWidht;
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
        /// 合并word
        /// </summary>
        /// <returns></returns>
        public int MerageWord(List<string> fileFullNamelist, bool isPage = false)
        {
            try
            {
                _currentWord.Activate();
                for (int i = 0; i < fileFullNamelist.Count; i++)
                {
                    InsertWord(fileFullNamelist[i], isPage ? i > 0 : isPage);
                }
            }
            catch (Exception ex)
            {
                _needWrite = false;
                Dispose();
                throw new Exception(string.Format("错误信息:{0}.{1}", ex.StackTrace.ToString(), ex.Message));
            }

            return fileFullNamelist.Count;
        }
        
        /// <summary>
        /// 设置复选框选中
        /// </summary>
        /// <returns></returns>
        public string SelecionCheckbox(string bookmark, int controlIndex, bool isCheck = true)
        {
            try
            {
                Range range = GetBookmarkRank(_currentWord, bookmark);
                range.Select();
                range = _wordApp.Selection.Cells[1].Range;
                range.Select();

                int i = 1;

                foreach (InlineShape shape in _wordApp.Selection.InlineShapes)
                {
                    if (shape.Type == WdInlineShapeType.wdInlineShapeOLEControlObject)
                    {
                        object oleControl = shape.OLEFormat.Object;
                        Type oleControlType = oleControl.GetType();
                        //设置复选框选中
                        if (i == controlIndex)
                            oleControlType.InvokeMember("Value", System.Reflection.BindingFlags.SetProperty, null, oleControl, new object[] { isCheck.ToString() });
                        i++;
                    }
                }

                return "设置成功";
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region 私有方法
        private void NewApp()
        {
            _wordApp = new Application();
        }

        //关闭application
        private void CloseApp()
        {
            object wdSaveOptions = WdSaveOptions.wdDoNotSaveChanges;
            foreach (Document item in _wordApp.Documents)
            {
                item.Close(ref wdSaveOptions, ref _missing, ref _missing);
            }
            _wordApp.Application.Quit(ref _objFalse, ref _missing, ref _missing);
            _wordApp = null;
        }

        //创建一个新的word
        private Document CreatWord()
        {
            if (_wordApp == null)
                NewApp();
            return _wordApp.Documents.Add(ref _missing, ref _missing, ref _missing, ref _objFalse);
        }

        //打开word
        private Document OpenWord(string fileFullPath, bool isOtherFormat = false)
        {
            if (_wordApp == null)
                NewApp();
            Document openWord = null;

            //判断文件是否打开
            if (_fileDic != null && _fileDic.ContainsKey(fileFullPath))
                openWord = _fileDic[fileFullPath];
            else
            {
                if (_fileDic == null)
                {
                    _fileDic = new Dictionary<string, Document>();
                }
                object otherFormat = isOtherFormat ? true : false;//其他格式文件用word打开
                object obj = fileFullPath;
                openWord = _wordApp.Documents.Open(ref obj,
                ref _missing,//文档名（可包含路径）
                ref _missing,//True 显示 转换文件对话框中，如果该文件不是 Microsoft Word 格式
                ref _missing,// 如果该属性值为 True , 则以只读方式打开文档
                ref _missing,//true 要将文件名添加到列表中最近使用的文件在文件菜单的底部
                ref _missing,//打开文档时所需的密码
                ref _missing,//打开模板时所需的密码
                ref _missing,//为 False，则激活打开的文档
                ref _missing,
                ref _missing,
                ref _missing,
                ref _objFalse, //在可见窗口中打开文档。 默认值为 True
                ref _missing,
                ref _missing,
                ref _missing,
                ref _missing);

                _fileDic.Add(fileFullPath, openWord);
            }


            return openWord;
        }

        //关闭word
        private void CloseWord(Document doc, string fileFulleName = "")
        {
            doc.Close(ref _objFalse, ref _missing, ref _missing);
            if (fileFulleName != null)
                _fileDic.Remove(fileFulleName);
        }

        //保存word
        private void SaveWord(Document doc)
        {
            doc.Save();
        }

        //另存word
        private void SaveAsWord(Document doc, string outFileFullName)
        {
            //筛选保存格式
            object defFormat = FilterExtendName(outFileFullName);
            object path = outFileFullName;
            doc.SaveAs(
                ref path,      //FileName
                ref defFormat,     //FileFormat
                ref _missing,     //LockComments
                ref _missing,     //PassWord     
                ref _missing,     //AddToRecentFiles
                ref _missing,     //WritePassword
                ref _missing,     //ReadOnlyRecommended
                ref _missing,     //EmbedTrueTypeFonts
                ref _missing,     //SaveNativePictureFormat
                ref _missing,     //SaveFormsData
                ref _missing,     //SaveAsAOCELetter,
                ref _missing,     //Encoding
                ref _objFalse,     //InsertLineBreaks
                ref _missing,     //AllowSubstitutions
                ref _missing,     //LineEnding
                ref _missing      //AddBiDiMarks
          );
        }

        /// <summary>
        /// 根据文件拓展名获取文件类型
        /// </summary>
        /// <param name="fileFullName"></param>
        /// <returns></returns>
        private object FilterExtendName(string fileFullName)
        {
            int index = fileFullName.LastIndexOf('.');
            string extendName = fileFullName.Substring(index, fileFullName.Length - index);
            object resultFormat = null;
            switch (extendName)
            {
                case ".pdf":
                    resultFormat = WdSaveFormat.wdFormatPDF;
                    break;
                case ".rtf":
                    resultFormat = WdSaveFormat.wdFormatRTF;
                    break;
                case ".xml":
                    resultFormat = WdSaveFormat.wdFormatXMLDocument;
                    break;
                case ".docx":
                default:
                    resultFormat = WdSaveFormat.wdFormatDocumentDefault;
                    break;
            }
            return resultFormat;
        }

        /// <summary>
       /// 获取文件名称
       /// </summary>
       /// <param name="fileFullName"></param>
       /// <returns></returns>
        private string FilterFileName(string fileFullName)
        {
            int index = fileFullName.LastIndexOf('\\');
            return fileFullName.Substring(index, fileFullName.Length - index);
        }


        //在application内插入文件(合并word)
        private void InsertWord(string fileName, bool ifBreakPage = false)
        {
            if (ifBreakPage)
                InsertBreakPage(true);
            _wordApp.Selection.InsertFile(
                        fileName,
                        ref _missing,
                        ref _missing,      //confirmConversion = false;
                        ref _missing,                   //link = false;
                        ref _missing              //attachment = false;
                        );
        }

        /// <summary>
        /// 为当前选中区域添加分页符
        /// </summary>
        /// <param name="isPage"></param>
        private void InsertBreakPage(bool isPage)
        {
            object unite = WdUnits.wdStory;
            _wordApp.Selection.EndKey(ref unite, ref _missing);

            object breakType = WdBreakType.wdSectionBreakContinuous;//分节符
            _wordApp.ActiveWindow.Selection.InsertBreak(breakType);

            if (isPage)
            {
                object breakPage = WdBreakType.wdPageBreak;//分页符
                _wordApp.ActiveWindow.Selection.InsertBreak(breakPage);
            }
        }

        /// <summary>
        /// 当前word插入图片
        /// </summary>
        /// <returns></returns>
        private InlineShape AddPicture(string picFileName, Document doc, Range range, float width = 0, float height = 0)
        {
            InlineShape image = doc.InlineShapes.AddPicture(picFileName, ref _missing, ref _missing, range);
            if (width != 0 && height != 0)
            {
                image.Width = width;
                image.Height = height;
            }
            return image;
        }

        /// <summary>
        /// 获取书签的位置
        /// </summary>
        /// <param name="word"></param>
        /// <param name="bookmark"></param>
        /// <returns></returns>
        private Range GetBookmarkRank(Document word, string bookmark)
        {
            object bk = bookmark;
            Range bookmarkRank = null;
            if (word.Bookmarks.Exists(bookmark))
                bookmarkRank = word.Bookmarks.get_Item(ref bk).Range;
            else
                throw new Exception(string.Format("未找到书签:{0}", bookmark));

            return bookmarkRank;
        }

        /// <summary>
        /// 获取书签位置,如果不存在返回null
        /// </summary>
        /// <param name="word"></param>
        /// <param name="bookmark"></param>
        /// <returns></returns>
        private Range GetBookmarkReturnNull(Document word, string bookmark)
        {
            object bk = bookmark;
            Range bookmarkRank = null;
            if (word.Bookmarks.Exists(bookmark))
                bookmarkRank = word.Bookmarks.get_Item(ref bk).Range;
            //else
            //    throw new Exception(string.Format("未找到书签:{0}", bookmark));

            return bookmarkRank;
        }

        /// <summary>
        /// word清除域代码格式
        /// </summary>
        /// <returns></returns>
        private int ClearWordCodeFormat()
        {
            _currentWord.Select();
            ClearCode();
            return 1;
        }

        /// <summary>
        /// 域代码转文本
        /// </summary>
        private void ClearCode()
        {
            ShowCodesAndUnlink(_currentWord.Content);
            for (int i = 1; i <= _wordApp.Selection.Sections.Count; i++)
            {
                Section wordSection = _wordApp.Selection.Sections[i];
                Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                ShowCodesAndUnlink(footerRange);

                Microsoft.Office.Interop.Word.Range footerRange1 = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                ShowCodesAndUnlink(footerRange1);

                Microsoft.Office.Interop.Word.Range headerRange = wordSection.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                ShowCodesAndUnlink(headerRange);

                Microsoft.Office.Interop.Word.Range headerRange1 = wordSection.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                ShowCodesAndUnlink(headerRange1);
            }
        }

        /// <summary>
        /// 将域代码替换为文本域值
        /// </summary>
        /// <param name="range"></param>
        private void ShowCodesAndUnlink(Range range)
        {
            range.Fields.ToggleShowCodes();
            range.Fields.Unlink();
        }

        //全文替换文本 1.文本 2. 页脚 3. 页眉
        private void Replace(int type, string oldWord, string newWord, int replaceType)
        {
            object repalceTypObj = replaceType == 1 ? wdReplaceOne : wdReplaceAll;
            switch (type)
            {
                //1:为文本
                default:
                    _wordApp.Selection.Find.Replacement.ClearFormatting();
                    _wordApp.Selection.Find.ClearFormatting();
                    _wordApp.Selection.Find.Text = oldWord;//需要被替换的文本
                    _wordApp.Selection.Find.Replacement.Text = newWord;//替换文本 
                    try
                    {
                        //执行替换操作
                        _wordApp.Selection.Find.Execute(
                        ref _missing, ref _missing, ref _missing,
                        ref _missing, ref _missing, ref _missing,
                        ref _missing, ref _missing, ref _missing,
                        ref _missing, ref repalceTypObj,// 指定要执行替换的个数：一个、全部或者不替换。 可以是任何WdReplace常量:wdReplaceAll wdReplaceNone wdReplaceOne
                        ref _missing, ref _missing, ref _missing,
                        ref _missing);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    break;
                case 2://页脚
                    try
                    {
                        for (int i = 1; i <= _wordApp.Selection.Sections.Count; i++)
                        {

                            Section wordSection = _wordApp.Selection.Sections[i];

                            Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

                            //footerRange.Fields.ToggleShowCodes();//显示域代码
                            footerRange.Find.ClearFormatting();
                            footerRange.Find.Replacement.ClearFormatting();
                            footerRange.Find.Text = oldWord;
                            footerRange.Find.Replacement.Text = newWord;
                            footerRange.Find.Execute(ref _missing, ref _missing, ref _missing,
                                                   ref _missing, ref _missing, ref _missing,
                                                   ref _missing, ref _missing, ref _missing,
                                                   ref _missing, ref repalceTypObj, ref _missing,
                                                   ref _missing, ref _missing, ref _missing);
                            // footerRange.Fields.Update();//更新域代码


                            Microsoft.Office.Interop.Word.Range footerRange1 = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                            // footerRange1.Fields.ToggleShowCodes();//显示域代码
                            footerRange1.Find.ClearFormatting();
                            footerRange1.Find.Replacement.ClearFormatting();
                            footerRange1.Find.Text = oldWord;
                            footerRange1.Find.Replacement.Text = newWord;
                            footerRange1.Find.Execute(ref _missing, ref _missing, ref _missing,
                                                   ref _missing, ref _missing, ref _missing,
                                                   ref _missing, ref _missing, ref _missing,
                                                   ref _missing, ref repalceTypObj, ref _missing,
                                                   ref _missing, ref _missing, ref _missing);
                            //footerRange1.Fields.Update();//更新域代码
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }

                    break;
                case 3://页眉
                    try
                    {
                        for (int i = 1; i <= _wordApp.Selection.Sections.Count; i++)
                        {
                            Section wordSection = _wordApp.Selection.Sections[i];
                            Microsoft.Office.Interop.Word.Range headerRange = wordSection.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

                            //headerRange.Fields.ToggleShowCodes();
                            headerRange.Find.ClearFormatting();
                            headerRange.Find.Replacement.ClearFormatting();
                            headerRange.Find.Text = oldWord;
                            headerRange.Find.Replacement.Text = newWord;
                            headerRange.Find.Execute(ref _missing, ref _missing, ref _missing,
                                                   ref _missing, ref _missing, ref _missing,
                                                   ref _missing, ref _missing, ref _missing,
                                                   ref _missing, ref repalceTypObj, ref _missing,
                                                   ref _missing, ref _missing, ref _missing);
                            //headerRange.Fields.Update();

                            Microsoft.Office.Interop.Word.Range headerRange1 = wordSection.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                            //footerRange1.Fields.ToggleShowCodes();//显示域代码
                            headerRange1.Find.ClearFormatting();
                            headerRange1.Find.Replacement.ClearFormatting();
                            headerRange1.Find.Text = oldWord;
                            headerRange1.Find.Replacement.Text = newWord;
                            headerRange1.Find.Execute(ref _missing, ref _missing, ref _missing,
                                                   ref _missing, ref _missing, ref _missing,
                                                   ref _missing, ref _missing, ref _missing,
                                                   ref _missing, ref repalceTypObj, ref _missing,
                                                   ref _missing, ref _missing, ref _missing);
                            //footerRange1.Fields.Update();//更新域代码
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }

                    break;

            }
        }

        /// <summary>
        /// table合并单元格
        /// </summary>
        /// <param name="table">表格</param>
        /// <param name="startRow">首单元格的行</param>
        /// <param name="startColumn">首单元格的列</param>
        /// <param name="endRow">尾单元格的行</param>
        /// <param name="endColumn">尾单元格的列</param>
        private void MergeCell(Table table, int startRow, int startColumn, int endRow, int endColumn)
        {
            MergeCell(table.Cell(startRow, startColumn), table.Cell(endRow, endColumn));
        }

        /// <summary>
        /// 重写table合并单元格
        /// </summary>
        /// <param name="startCell">首单元格</param>
        /// <param name="endCell">尾单元格</param>
        private void MergeCell(Cell startCell, Cell endCell)
        {
            startCell.Merge(endCell);
        }

        /// <summary>
        /// 当前选中文字加粗居左
        /// </summary>
        private void FontBoldLeft()
        {
            _wordApp.Selection.Font.Bold = (int)WdConstants.wdToggle;
            _wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        }


        /// <summary>
        /// 设置表格添加边框
        /// </summary>
        /// <param name="table"></param>
        private void SetTabelFormat(Table table)
        {
            table.Select();
            table.Borders.Enable = (int)WdLineStyle.wdLineStyleSingle;
        }

        /// <summary>
        /// 为table添加序号
        /// </summary>
        /// <param name="table">表格</param>
        /// <param name="columnNumber">添加序号的列</param>
        /// <param name="isTitle">是否有表头</param>
        private void AddTableNumber(Table table, int columnNumber, bool isTitle = true)
        {
            table.Select();
            table.Cell(1, columnNumber).Select();
            _wordApp.Selection.SelectColumn();
            int intCell = 0;
            foreach (Cell item in _wordApp.Selection.Cells)
            {
                if (isTitle && intCell == 0)
                {
                    intCell++;
                    continue;
                }

                if (!isTitle)
                {
                    intCell++;
                }

                item.Range.Text = (intCell).ToString();
                intCell++;
            }
        }

       
        /// <summary>
        /// 创建并移动到下一个段落
        /// </summary>
        private void CreateAndGoToNextParagraph(Range range, bool isCreateParagraph, bool isMove)
        {
            range.Select();
            if (isCreateParagraph)
            {
                object unite = WdUnits.wdLine;
                _wordApp.Selection.EndKey(ref unite, ref _missing);
                _wordApp.Selection.TypeParagraph();
            }
            if (isMove)
            {
                object move = 1;
                range.Move(WdUnits.wdParagraph, ref move);
            }
        }
        /// <summary>
        /// 插入段落
        /// </summary>
        /// <param name="range"></param>
        private void InsertParagraph(Range range)
        {
            object contLine = 1;
            object WdLine = WdUnits.wdLine;//换一行;
            range.Document.ActiveWindow.Selection.MoveDown(ref WdLine, ref contLine, ref _missing);//移动焦点
            range.Document.ActiveWindow.Selection.TypeParagraph();//插入段落
        }

        /// <summary>
        /// 设置table格式
        /// </summary>
        /// <param name="table"></param>
        private void ClearFormatTable(Table table)
        {
            table.Select();
            if (table.Rows.Count > 0)
            {
                table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
                _wordApp.Selection.ClearFormatting();
                _wordApp.Options.DefaultHighlightColorIndex = WdColorIndex.wdNoHighlight;
                _wordApp.Selection.Range.HighlightColorIndex = WdColorIndex.wdNoHighlight;
                _wordApp.Selection.Shading.Texture = WdTextureIndex.wdTextureNone;
                _wordApp.Selection.Shading.ForegroundPatternColor = WdColor.wdColorAutomatic;
                _wordApp.Selection.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic;
                table.Select();
                _wordApp.Selection.SelectCell();
                _wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                _wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                //table.Range.Application.Selection.Cells.DistributeHeight();
                //table.Range.Application.Selection.Cells.DistributeWidth();

                SetDistributeTable(table);
            }

        }
        /// <summary>
        /// 设置table除表头之外的单元格等高
        /// </summary>
        /// <param name="table"></param>
        private void SetDistributeTable(Table table)
        {
            int tableRows = table.Rows.Count;
            if (tableRows >= 2)
            {
                int tableColumns = table.Columns.Count;
                object extend = WdMovementType.wdExtend;//选中拓展格式

                table.Cell(2, 1).Select();
                int t1 = table.Cell(2, 1).Range.Application.Selection.MoveDown(WdUnits.wdLine, tableRows - 2, ref extend);
                _wordApp.Selection.Cells.DistributeHeight();
            }

        }

        /// <summary>
        /// 根据windows大小设置表格大小
        /// </summary>
        /// <param name="table"></param>
        private void SetAutoFitContentForTable(Table table)
        {
            table.Select();
            _wordApp.Selection.Tables[1].AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
            _wordApp.Selection.Tables[1].AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
        }

        /// <summary>
        /// 根据文件名称结束word进程
        /// </summary>
        /// <param name="fileName">word文件名称</param>
        public void KillWordProcess(string fileName)
        {
            Process myProcess = new Process();
            Process[] wordProcess = Process.GetProcessesByName("winword");
            foreach (Process pro in wordProcess) //这里是找到那些没有界面的Word进程
            {
                //IntPtr ip = pro.MainWindowHandle;

                string str = pro.MainWindowTitle; //发现程序中打开跟用户自己打开的区别就在这个属性
                                                  //用户打开的str 是文件的名称，程序中打开的就是空字符串
                if (str == fileName)
                {
                    pro.Kill();
                }
            }
        }

        /// <summary>
        /// 删除word进程
        /// </summary>
        public void KillWordProcess()
        {
            Process myProcess = new Process();
            Process[] wordProcess = Process.GetProcessesByName("winword");
            foreach (Process pro in wordProcess) //这里是找到那些没有界面的Word进程
            {
                pro.Kill();
            }
        }

        #endregion

        #region 资源回收
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    if (_needWrite)
                    {
                        if (!_outFilePath.Equals("") && _isSaveAs)
                        {
                            SaveAsWord(_currentWord, _outFilePath);
                        }
                        else if (!_isSaveAs)
                        {
                            SaveWord(_currentWord);
                        }
                        _currentWord = null;
                    }

                    if (_wordApp != null)
                        CloseApp();

                    if (_fileDic != null)
                        _fileDic = null;
                }
                _disposed = true;
            }
        }
        #endregion
    }
}
