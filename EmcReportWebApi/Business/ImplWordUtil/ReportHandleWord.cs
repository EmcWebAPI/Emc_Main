﻿using EmcReportWebApi.Utils;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using EmcReportWebApi.ReportComponent.Experiment;
using EmcReportWebApi.ReportComponent.Image;
using Newtonsoft.Json.Linq;

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
        public virtual string CopyHtmlContentToTemplate(string htmlFilePath, string templateFilePath, string bookmark, bool isNeedBreak, bool isCloseTheFile, bool isCloseTemplateFile)
        {
            try
            {
                Document htmlDoc = OpenWord(htmlFilePath);
                htmlDoc.Select();
                htmlDoc.Content.Copy();

                Document templateDoc = OpenWord(templateFilePath);

                templateDoc.Content.Select();
                _wordApp.Selection.MoveDown(WdUnits.wdLine, _wordApp.Selection.Paragraphs.Count, WdMovementType.wdMove);
                _wordApp.Selection.TypeParagraph();
                Range tableRange = _wordApp.Selection.Range;
                tableRange.Paste();

                foreach (Table item in templateDoc.Content.Tables)
                {
                    item.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
                }

                templateDoc.Content.Select();
                _wordApp.Selection.MoveDown(WdUnits.wdLine, _wordApp.Selection.Paragraphs.Count, WdMovementType.wdMove);
                object breakPage = WdBreakType.wdPageBreak;//分页符
                _wordApp.ActiveWindow.Selection.InsertBreak(breakPage);

                if (isCloseTemplateFile)
                {
                    CloseWord(templateDoc, templateFilePath);
                }
                if (isCloseTheFile)
                    CloseWord(htmlDoc, htmlFilePath);
            }
            catch (Exception ex)
            {

                _needWrite = false;
                Dispose();
                throw new Exception(message: $"错误信息:{ex.StackTrace}.{ex.Message}");
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
                throw new Exception($"错误信息:{ex.StackTrace.ToString()},{ex.Message}");
            }

            return "插入成功";
        }

        /// <summary>
        /// 向table中插入list(不需要合并单元格)
        /// </summary>
        /// <param name="list"></param>
        /// <param name="bookmark"></param>
        /// <returns></returns>
        public string InsertListToTable(JArray list, string bookmark)
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

                JObject item = (JObject)list[i];
                int j = 1;
                foreach (var tpItem in item)
                {
                    table.Cell(tableRowIndex, j).Range.Text = tpItem.Value.ToString();
                    j++;
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
        /// 样品图片
        /// </summary>
        /// <param name="list"></param>
        /// <param name="bookmark"></param>
        /// <returns></returns>
        public virtual string InsertImageToWordSample(IList<ImageInfoAbstract> list, string bookmark)
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
                    var arrStr = list[i];
                    string fileName = arrStr.ImageFileFullName;
                    string content = arrStr.Content;
                    table.Select();
                    Range cellRange = _wordApp.Selection.Cells[i + 1].Range;
                    cellRange.Select();

                    if (!fileName.Equals(""))
                    {
                        AddSamplePicture(fileName, _currentWord, cellRange, tableWidth - 40, tableWidth - 240);
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

        private void AddSamplePicture(string picFileName, Document doc, Range range, float width = 0,
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
        /// 实验连接图
        /// </summary>
        public virtual string InsertConnectionImageToTemplate(string fileFullPath, IList<ExperimentImage> list, string bookmark, bool isCloseTheFile = true)
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
                    string fileName = list[i].ImageFileFullName;
                    string content = list[i].Content;
                    table.Select();
                    Range cellRange = _wordApp.Selection.Cells[i + 1].Range;
                    cellRange.Select();

                    if (!fileName.Equals(""))
                    {
                        InlineShape image = AddPicture(fileName, doc, cellRange);
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
        /// 实验布置图
        /// </summary>
        /// <returns></returns>
        public virtual string InsertImageToTemplate(string fileFullPath, IList<ExperimentImage> list, string bookmark, bool isCloseTheFile = true)
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
                    string fileName = list[i].ImageFileFullName;
                    string content = list[i].Content;
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

        /// <summary>
        /// 复制第二个文件内容到第一个文件
        /// </summary>
        /// <returns></returns>
        public override string CopyOtherFileContentToWord(string firstFilePath, string secondFilePath, string bookmark, bool isCloseTheFile = true)
        {
            try
            {
                Document htmldoc = OpenWord(firstFilePath);
                Document secondFile = OpenWord(secondFilePath);
                Range range = GetBookmarkRank(secondFile, bookmark);

                htmldoc.Content.Copy();
                range.Select();
                range.PasteAndFormat(WdRecoveryType.wdUseDestinationStylesRecovery);

                Range nextRange;
                switch (bookmark)
                {
                    case "sysj1":
                        nextRange = GetBookmarkReturnNull(secondFile, "sysj2");
                        if (nextRange != null)
                        {
                            nextRange.Select();
                            _wordApp.Selection.MoveUp(WdUnits.wdLine, 2, WdMovementType.wdMove);
                            _wordApp.Selection.TypeBackspace();
                            _wordApp.Selection.TypeBackspace();
                        }
                        break;
                    case "sysj2":
                    case "sysj":
                        nextRange = GetBookmarkReturnNull(secondFile, "syljt");
                        if (nextRange != null)
                        {
                            nextRange.Select();
                            _wordApp.Selection.MoveUp(WdUnits.wdLine, 3, WdMovementType.wdMove);
                            _wordApp.Selection.TypeBackspace();
                            _wordApp.Selection.TypeBackspace();
                        }
                        break;
                }


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
                throw new Exception($"错误信息:{ex.StackTrace},{ex.Message}");
            }

            return "保存成功";
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
        public string CreateTableToCurrentWord(List<string> contentList, string bookmark, bool isNeedBreak)
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
        public string CreateTableToWord(string otherFileFullName, IList<string> contentList, string bookmark, bool isCloseTemplateFile, bool isNeedBreak, bool isCloseTheFile = true)
        {
            Document otherFileDoc = OpenWord(otherFileFullName);
            if (isCloseTemplateFile)
                CloseWord(otherFileDoc);
            return CreateTableToWord(otherFileDoc, contentList, bookmark, isNeedBreak, isCloseTheFile);
        }

        /// <summary>
        /// 
        /// </summary>
        private string CreateTableToWord(Document doc, IList<string> contentList, string bookmark, bool isNeedBreak, bool isCloseTheFile = true)
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
            }
            int numRows = 1;
            int numColumns = 2;
            switch (contentList.Count % 2)
            {
                case 1:
                    numRows = contentList.Count / 2 + 1;
                    break;
                default:
                    numRows = contentList.Count / 2==0?1: contentList.Count / 2;
                    break;

            }
            table.Select();
            _wordApp.Selection.Tables.Add(table, numRows, numColumns, ref _missing, ref _missing);
            _wordApp.Selection.Tables[1].Select();
            _wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
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
                int rtfTableCount = rtfDoc.Tables.Count;
                templateDoc.Content.Select();
                if (isNeedBreak)
                {
                    templateDoc.Content.Select();
                    _wordApp.Selection.MoveDown(WdUnits.wdLine, _wordApp.Selection.Paragraphs.Count, WdMovementType.wdMove);
                }

                Range wordTable = _wordApp.Selection.Range;

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
                    if (mainTitleArray != null)
                    {
                        copyTable.Cell(1, 1).Range.Text = mainTitleArray[m];
                        m++;
                    }

                    if (i != copyFileTableStartIndex)
                        CreateAndGoToNextParagraph(wordTable, (i != copyFileTableStartIndex) || isNeedBreak, (i != copyFileTableStartIndex) || isNeedBreak);//获取下一个range
                    CreateAndGoToNextParagraph(wordTable, (i != copyFileTableStartIndex) || isNeedBreak, (i != copyFileTableStartIndex) || isNeedBreak);//InsertBR(wordTable, i <= rtfTableCount);//添加回车
                    copyTable.Range.Copy();
                    wordTable.Paste();

                    ClearFormatTable(wordTable.Tables[1]);
                    wordTable.Tables[1].Rows.SetHeight(16f, WdRowHeightRule.wdRowHeightAtLeast);
                    wordTable.Tables[1].Select();
                    _wordApp.Selection.Cells.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
                    _wordApp.Selection.Cells.PreferredWidth = 20f;
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
        /// 从传导发射文件取内容到word
        /// </summary>
        public string CopyReExperimentFileTableForColByTableIndex(string templateFullPath, string copyFileFullPath, int copyFileTableStartIndex, int copyFileTableEndIndex, Dictionary<int, string> copyTableColDic, string wordBookmark, int titleRow, string mainTitle, bool isCloseTemplateFile, bool isNeedBreak, bool isCloseTheFile = true)
        {
            try
            {
                Document templateDoc = OpenWord(templateFullPath);
                Document rtfDoc = OpenWord(copyFileFullPath, true);
                CopyReExperimentFileTableForColByTableIndex(templateDoc, rtfDoc, copyFileTableStartIndex, copyFileTableEndIndex, copyTableColDic, wordBookmark, titleRow, mainTitle, isNeedBreak);
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
        private string CopyReExperimentFileTableForColByTableIndex(Document templateDoc, Document rtfDoc, int copyFileTableStartIndex, int copyFileTableEndIndex, Dictionary<int, string> copyTableColDic, string wordBookmark, int titleRow, string mainTitle, bool isNeedBreak)
        {
            try
            {
                int rtfTableCount = rtfDoc.Tables.Count;
                templateDoc.Content.Select();
                if (isNeedBreak)
                {
                    templateDoc.Content.Select();
                    _wordApp.Selection.MoveDown(WdUnits.wdLine, _wordApp.Selection.Paragraphs.Count, WdMovementType.wdMove);
                }

                Range wordTable = _wordApp.Selection.Range;

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
                    if (mainTitleArray != null)
                    {
                        copyTable.Cell(1, 1).Range.Text = mainTitleArray[m];
                        m++;
                    }

                    //给表格排序
                    List<Tuple<int, int>> sortColumn = new List<Tuple<int, int>>
                    {
                        new Tuple<int, int>(5,3),
                        new Tuple<int, int>(6,4),
                        new Tuple<int, int>(7,5),
                        new Tuple<int, int>(7,6)
                    };
                    foreach (var item in sortColumn)
                    {
                        copyTable.Select();
                        var cutColumn = copyTable.Columns[item.Item1];
                        cutColumn.Select();
                        _wordApp.Selection.Cut();
                        var pasteCell = copyTable.Columns[item.Item2];
                        pasteCell.Select();
                        _wordApp.Selection.PasteAndFormat(WdRecoveryType.wdFormatOriginalFormatting);
                    }

                    if (i != copyFileTableStartIndex)
                        CreateAndGoToNextParagraph(wordTable, (i != copyFileTableStartIndex) || isNeedBreak, (i != copyFileTableStartIndex) || isNeedBreak);//获取下一个range
                    CreateAndGoToNextParagraph(wordTable, (i != copyFileTableStartIndex) || isNeedBreak, (i != copyFileTableStartIndex) || isNeedBreak);//InsertBR(wordTable, i <= rtfTableCount);//添加回车
                    copyTable.Range.Copy();
                    wordTable.Paste();

                    ClearFormatTable(wordTable.Tables[1]);
                    wordTable.Tables[1].Rows.SetHeight(14f, WdRowHeightRule.wdRowHeightAtLeast);
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
        /// 从谐波失真文件取内容到word
        /// </summary>
        public string CopyHarmonicOtherFileTableForColByTableIndex(string templateFullPath, string copyFileFullPath, int copyFileTableStartIndex, int copyFileTableEndIndex, Dictionary<int, string> copyTableColDic, string wordBookmark, int titleRow, string mainTitle, bool isCloseTemplateFile, bool isNeedBreak, bool isCloseTheFile = true)
        {
            try
            {
                Document templateDoc = OpenWord(templateFullPath);
                Document rtfDoc = OpenWord(copyFileFullPath, true);
                CopyHarmonicOtherFileTableForColByTableIndex(templateDoc, rtfDoc, copyFileTableStartIndex, copyFileTableEndIndex, copyTableColDic, wordBookmark, titleRow, mainTitle, isNeedBreak);
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
        private string CopyHarmonicOtherFileTableForColByTableIndex(Document templateDoc, Document rtfDoc, int copyFileTableStartIndex, int copyFileTableEndIndex, Dictionary<int, string> copyTableColDic, string wordBookmark, int titleRow, string mainTitle, bool isNeedBreak)
        {
            try
            {
                int rtfTableCount = rtfDoc.Tables.Count;
                templateDoc.Content.Select();
                if (isNeedBreak)
                {
                    templateDoc.Content.Select();
                    _wordApp.Selection.MoveDown(WdUnits.wdLine, _wordApp.Selection.Paragraphs.Count, WdMovementType.wdMove);
                }

                Range wordTable = _wordApp.Selection.Range;

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
                    if (mainTitleArray != null)
                    {
                        copyTable.Cell(1, 1).Range.Text = mainTitleArray[m];
                        m++;
                    }

                    templateDoc.Content.Select();
                    _wordApp.Selection.MoveDown(WdUnits.wdLine, _wordApp.Selection.Paragraphs.Count, WdMovementType.wdMove);
                    _wordApp.Selection.TypeParagraph();

                    copyTable.Range.Copy();
                    _wordApp.Selection.Paste();
                    wordTable = _wordApp.Selection.Range;
                    ClearFormatTable(wordTable.Tables[1]);
                    wordTable.Tables[1].Rows.SetHeight(14f, WdRowHeightRule.wdRowHeightAtLeast);

                    //谐波失真最后一列变符合
                    for (int j = 3; j <= wordTable.Tables[1].Rows.Count; j++)
                    {
                        wordTable.Tables[1].Cell(j, 5).Range.Text = "符合";
                    }

                    if (i != copyFileTableStartIndex)
                    {
                        templateDoc.Content.Select();
                        _wordApp.Selection.MoveDown(WdUnits.wdLine, _wordApp.Selection.Paragraphs.Count, WdMovementType.wdMove);
                        object breakPage = WdBreakType.wdPageBreak;//分页符
                        _wordApp.ActiveWindow.Selection.InsertBreak(breakPage);
                    }
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
        /// 从电压波动文件取内容到word
        /// </summary>
        public string CopyFluctuationFileTableForColByTableIndex(string templateFullPath,
            string copyFileFullPath,
            int copyFileTableStartIndex,
            int copyFileTableEndIndex,
            Dictionary<int, string> copyTableColDic,
            string wordBookmark,
            int titleRow,
            string mainTitle,
            bool isCloseTemplateFile,
            bool isNeedBreak,
            bool isPage,
            bool isCloseTheFile = true)
        {
            try
            {
                Document templateDoc = OpenWord(templateFullPath);
                Document rtfDoc = OpenWord(copyFileFullPath, true);
                CopyFluctuationFileTableForColByTableIndex(templateDoc, rtfDoc, copyFileTableStartIndex, copyFileTableEndIndex, copyTableColDic, wordBookmark, titleRow, mainTitle, isNeedBreak, isPage);
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
        private string CopyFluctuationFileTableForColByTableIndex(Document templateDoc,
            Document rtfDoc,
            int copyFileTableStartIndex,
            int copyFileTableEndIndex,
            Dictionary<int, string> copyTableColDic,
            string wordBookmark,
            int titleRow,
            string mainTitle,
            bool isNeedBreak,
            bool isPage)
        {
            try
            {
                int rtfTableCount = rtfDoc.Tables.Count;
                templateDoc.Content.Select();
                if (isNeedBreak)
                {
                    templateDoc.Content.Select();
                    _wordApp.Selection.MoveDown(WdUnits.wdLine, _wordApp.Selection.Paragraphs.Count, WdMovementType.wdMove);
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
                    if (mainTitleArray != null)
                    {
                        copyTable.Cell(1, 1).Range.Text = mainTitleArray[m];
                        m++;
                    }

                    templateDoc.Content.Select();
                    _wordApp.Selection.MoveDown(WdUnits.wdLine, _wordApp.Selection.Paragraphs.Count, WdMovementType.wdMove);
                    _wordApp.Selection.TypeParagraph();

                    copyTable.Range.Copy();
                    _wordApp.Selection.Paste();
                    var wordTable = _wordApp.Selection.Range;
                    var table1 = wordTable.Tables[1];
                    //电压波动最后一列变符合

                    ClearFormatTable(table1);
                    for (int j = 2; j <= table1.Rows.Count; j++)
                    {
                        switch (j)
                        {
                            case 2:
                                table1.Cell(j, 1).Range.Text = "短时间闪烁指数Pst";
                                break;
                            case 3:
                                table1.Cell(j, 1).Range.Text = "长时间闪烁指数Plt";
                                break;
                            case 4:
                                table1.Cell(j, 1).Range.Text = "相对稳态电压变化dc（%）";
                                break;
                            case 5:
                                table1.Cell(j, 1).Range.Text = "最大相对电压变化dmax（%）";
                                break;
                            case 6:
                                table1.Cell(j, 1).Range.Text = "t（d（t）>3.3%的时间）（s）";
                                break;
                        }
                        wordTable.Tables[1].Cell(j, 4).Range.Text = "符合";
                    }
                    table1.Columns[1].SetWidth(149f, WdRulerStyle.wdAdjustSameWidth);
                    table1.Rows.SetHeight(14f, WdRowHeightRule.wdRowHeightAtLeast);
                }

                if (isPage)
                {
                    templateDoc.Content.Select();
                    _wordApp.Selection.MoveDown(WdUnits.wdLine, _wordApp.Selection.Paragraphs.Count, WdMovementType.wdMove);
                    object breakPage = WdBreakType.wdPageBreak;//分页符
                    _wordApp.ActiveWindow.Selection.InsertBreak(breakPage);
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
        public virtual string CopyOtherFilePictureToWord(string templateFileFullName, string copyFileFullPath, int copyFilePictureStartIndex, string workBookmark, bool isCloseTemplateFile, bool isNeedBreak, bool isPage, bool isCloseTheFile = true)
        {
            string result;
            try
            {
                Document templateDoc = OpenWord(templateFileFullName);
                Document copyFileDoc = OpenWord(copyFileFullPath, true);
                result = CopyOtherFilePictureToWord(templateDoc,
                    copyFileDoc,
                    copyFilePictureStartIndex,
                    workBookmark,
                    isNeedBreak,
                    isPage);
                if (isCloseTemplateFile)
                {
                    CloseWord(templateDoc, templateFileFullName);
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
                Range bookmarkPic = fileDoc.Content;
                bookmarkPic.Select();

                if (isNeedBreak)
                {
                    fileDoc.Content.Select();
                    _wordApp.Selection.MoveDown(WdUnits.wdLine, _wordApp.Selection.Paragraphs.Count, WdMovementType.wdMove);
                    _wordApp.Selection.TypeParagraph();
                    bookmarkPic = _wordApp.Selection.Range;
                }

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
                            bookmarkPic.Paste();
                            foreach (InlineShape inlineShape in bookmarkPic.InlineShapes)
                            {
                                inlineShape.Select();
                                _wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            }

                            CreateAndGoToNextParagraph(bookmarkPic, true, true);
                        }
                        i++;
                    }
                }
                if (isPage)
                {
                    fileDoc.Content.Select();
                    _wordApp.Selection.MoveDown(WdUnits.wdLine, _wordApp.Selection.Paragraphs.Count, WdMovementType.wdMove);
                    object breakPage = WdBreakType.wdPageBreak;//分页符
                    _wordApp.ActiveWindow.Selection.InsertBreak(breakPage);
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