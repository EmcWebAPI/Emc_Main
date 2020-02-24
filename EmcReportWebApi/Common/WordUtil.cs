using Microsoft.Office.Interop.Word;
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

        public string CopyHtmlContentToTemplate(string htmlFilePath, string TemplateFilePath, string bookmark,bool isNeedBreak, bool isCloseTheFile,bool isCloseTemplateFile) {
            Document htmlDoc = this.OpenWord(htmlFilePath);
            htmlDoc.Select();
            htmlDoc.Content.Copy();

            Document templateDoc = this.OpenWord(TemplateFilePath);
            templateDoc.Select();
            Range range = this.GetBookmarkRank(templateDoc, bookmark);
            range.Select();

            object unite = WdUnits.wdStory;
            _wordApp.Selection.EndKey(ref unite, ref _missing);

            object breakType = WdBreakType.wdLineBreak;//换行符
            _wordApp.ActiveWindow.Selection.InsertBreak(breakType);

            range = _wordApp.Selection.Paragraphs.Last.Range;
            range.Select();
            CreateAndGoToNextParagraph(range, true, true);
            range.Paste();

            foreach (Table item in range.Tables)
            {
                item.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
            }
            //if (isNeedBreak)
            //{
            //    InsertBreakPage(false);
            //}
            if (isCloseTemplateFile)
            {
                CloseWord(templateDoc, TemplateFilePath);
            }
            if (isCloseTheFile)
                CloseWord(htmlDoc, htmlFilePath);
            return "创建成功";
        }

        public string CopyTableToWord(string otherFilePath, string bookmark, int tableIndex, bool isCloseTheFile)
        {
            try
            {
                Document otherFile = OpenWord(otherFilePath);
                otherFile.Tables[tableIndex].Range.Copy();
                Range range = GetBookmarkRank(_currentWord, bookmark);
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

        public string CopyImageToWord(string otherFilePath, string bookmark, bool isCloseTheFile)
        {
            try
            {

                Document otherFile = OpenWord(otherFilePath);
                otherFile.Select();

                Range bookmarkPic = GetBookmarkRank(_currentWord, bookmark);
                ShapeRange shapeRange = otherFile.Shapes.Range(1);
                InlineShape inlineShape = shapeRange.ConvertToInlineShape();
                inlineShape.Select();
                _wordApp.Selection.Copy();
                bookmarkPic.Select();
                _wordApp.Selection.Paste();
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


        /// <summary>
        /// 根据书签向word中插入内容
        /// </summary>
        /// <param name="content"></param>
        /// <param name="bookmark"></param>
        /// <returns></returns>
        public string InsertContentToWordByBookmark(string content, string bookmark)
        {
            try
            {
                Range range = GetBookmarkRank(_currentWord, bookmark);
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

                for (int i = 0; i < listCount; i++)
                {
                    string[] arrStr = list[i].Split(',');
                    string fileName = arrStr[0];
                    string content = arrStr[1];
                    table.Select();
                    Range cellRange = _wordApp.Selection.Cells[i + 1].Range;
                    cellRange.Select();
                    //CreateAndGoToNextParagraph(cellRange, true, true);
                    AddPicture(fileName, doc, cellRange);
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
                InsertBreakPage(false);
                table = _wordApp.Selection.Range.Sections.Last.Range;
                CreateAndGoToNextParagraph(table, true, true);
                CreateAndGoToNextParagraph(table, true, true);
            }
            int numRows = 0;
            int numColumns = 0;
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
        public string CopyOtherFileTableForColByTableIndex(string copyFileFullPath, int copyFileTableStartIndex, Dictionary<int, string> copyTableColDic, string wordBookmark, bool isNeedBreak, bool isCloseTheFile = true)
        {
            string result = "创建成功";
            try
            {
                Document rtfDoc = OpenWord(copyFileFullPath, true);
                result = CopyOtherFileTableForColByTableIndex(_currentWord, rtfDoc, copyFileTableStartIndex, copyTableColDic, wordBookmark, isNeedBreak);
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

        public string CopyOtherFileTableForColByTableIndex(string templateFullPath, string copyFileFullPath, int copyFileTableStartIndex, Dictionary<int, string> copyTableColDic, string wordBookmark, bool isCloseTemplateFile, bool isNeedBreak, bool isCloseTheFile = true)
        {
            string result = "创建成功";
            try
            {
                Document templateDoc = OpenWord(templateFullPath);
                Document rtfDoc = OpenWord(copyFileFullPath, true);
                result = CopyOtherFileTableForColByTableIndex(templateDoc, rtfDoc, copyFileTableStartIndex, copyTableColDic, wordBookmark, isNeedBreak);
                if (isCloseTemplateFile)
                    CloseWord(templateDoc, templateFullPath);
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


        private string CopyOtherFileTableForColByTableIndex(Document templateDoc, Document rtfDoc, int copyFileTableStartIndex, Dictionary<int, string> copyTableColDic, string wordBookmark, bool isNeedBreak)
        {
            try
            {
                //Document rtfDoc = OpenWord(copyFileFullPath, true);

                int rtfTableCount = rtfDoc.Tables.Count;

                Range wordTable = GetBookmarkRank(templateDoc, wordBookmark);
                wordTable.Select();
                if (isNeedBreak)
                {
                    InsertBreakPage(false);
                    wordTable = _wordApp.Selection.Range.Sections.Last.Range;
                }

                for (int i = copyFileTableStartIndex; i <= rtfTableCount; i++)
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
                            copyTable.Cell(1, j).Delete(ref wdDeleteCellsCol);
                        }
                        else
                        {
                            copyTable.Cell(1, j).Range.Text = copyTableColDic[j];
                        }
                    }

                    copyTable.Select();
                    copyTable.Range.Copy();

                    CreateAndGoToNextParagraph(wordTable, (i != copyFileTableStartIndex) || isNeedBreak, (i != copyFileTableStartIndex) || isNeedBreak);//获取下一个range
                    CreateAndGoToNextParagraph(wordTable, (i != copyFileTableStartIndex) || isNeedBreak, (i != copyFileTableStartIndex) || isNeedBreak);//InsertBR(wordTable, i <= rtfTableCount);//添加回车
                    wordTable.Paste();

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

        public string CopyOtherFileContentToWordReturnBookmark(string filePath, string bookmark, bool isCloseTheFile = true)
        {
            string newBookmark = "bookmark" + DateTime.Now.ToString("yyyyMMddhhmmss");
            try
            {
                Document htmldoc = OpenWord(filePath);
                Range rangeContent = htmldoc.Content;
                rangeContent.Select();
                InsertBreakPage(true);
                rangeContent = rangeContent.Sections.Last.Range;
                CreateAndGoToNextParagraph(rangeContent, false, true);
                rangeContent.Select();

                _wordApp.Selection.Bookmarks.Add(newBookmark, rangeContent);
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

        public string CopyOtherFileContentToWord(string firstFilePath, string secondFilePath, string bookmark, bool isCloseTheFile = true)
        {
            try
            {
                Document htmldoc = OpenWord(firstFilePath);
                htmldoc.Content.Copy();
                Document secondFile = OpenWord(secondFilePath);
                Range range = GetBookmarkRank(secondFile, bookmark);
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

        public int InsertFiles(List<string> fileFullNamelist)
        {
            try
            {
                _currentWord.Activate();
                object unite = WdUnits.wdStory;
                object breakType = WdBreakType.wdSectionBreakContinuous;//分节符
                _wordApp.Selection.EndKey(ref unite, ref _missing);
                _wordApp.ActiveWindow.Selection.InsertBreak(breakType);
                foreach (var item in fileFullNamelist)
                {
                    InsertWord(item);
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

        private int ClearWordCodeFormat()
        {
            _currentWord.Select();
            ClearCode();
            return 1;
        }

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

        public object FilterExtendName(string fileFullName)
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

        public string FilterFileName(string fileFullName)
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

        //插入图片
        private InlineShape AddPicture(string picFileName, Document doc, Range range,int width=0,int height=0)
        {
            InlineShape image = doc.InlineShapes.AddPicture(picFileName, ref _missing, ref _missing, range);
            if (width != 0&&height!=0) {
                image.Width = width;
                image.Height = height;
            }
            return image;
        }

        //获取bookmark的位置
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

        //域代码转文本
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

        //将域代码替换为文本域值
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

        private void MergeCell(Table table, int startRow, int startColumn, int endRow, int endColumn)
        {
            MergeCell(table.Cell(startRow, startColumn), table.Cell(endRow, endColumn));
        }

        private void MergeCell(Cell startCell, Cell endCell)
        {
            startCell.Merge(endCell);
        }

        private void FontBoldLeft()
        {
            _wordApp.Selection.Font.Bold = (int)WdConstants.wdToggle;
            _wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        }

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

        public void KillWordProcess()
        {
            Process myProcess = new Process();
            Process[] wordProcess = Process.GetProcessesByName("winword");
            foreach (Process pro in wordProcess) //这里是找到那些没有界面的Word进程
            {
                pro.Kill();
            }
        }

        #region emc
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

        private void SetAutoFitContentForTable(Table table)
        {
            table.Select();
            _wordApp.Selection.Tables[1].AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
            _wordApp.Selection.Tables[1].AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
        }

        private Dictionary<int, string> DicNum(Dictionary<int, string> dic)
        {
            List<int> list = new List<int>();
            foreach (var item in dic)
            {
                list.Add(item.Key);
            }
            //对list进行冒泡排序
            int temp = 0;
            for (int i = 0; i < list.Count - 1; i++)
            {
                for (int j = 0; j < list.Count - 1 - i; j++)
                {
                    if (list[j] > list[j + 1])
                    {
                        temp = list[j + 1];
                        list[j + 1] = list[j];
                        list[j] = temp;
                    }
                }
            }

            //新的dic
            Dictionary<int, string> newDic = new Dictionary<int, string>();
            for (int i = list.Count - 1; i >= 0; i--)
            {
                newDic.Add(list[i], dic[list[i]]);
            }
            return newDic;
        }
        #endregion

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
