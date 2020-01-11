﻿using Microsoft.Office.Interop.Word;
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


        public string InsertContentInBookmark(string content, string bookmark)
        {
            Range range = GetBookmarkRank(_currentWord, bookmark);
            range.InsertAfter(content);
            return "插入成功";
        }
        /// <summary>
        /// 向table中插入list
        /// </summary>
        /// <param name="list">内容集合</param>
        /// <param name="bookmark">插入内容的位置 用bookmark获取</param>
        /// <param name="mergeColumn">需要合并的列</param>
        /// <returns></returns>
        public string InsertListToTable(List<string> list, string bookmark, int mergeColumn)
        {
            if (mergeColumn < 1)
            {
                return "合并列不能小于1";
            }
            //获取bookmark位置的table
            Range range = GetBookmarkRank(_currentWord, bookmark);
            range.Select();
            Table table = range.Tables[1];
            int rowCount = table.Rows.Count;
            //设置合并第二列相邻的相同内容

            int startRow = 0;
            int endRow = 0;
            string mergeContent = "";

            foreach (var item in list)
            {
                string[] arrStr = item.Split(',');
                if (table.Columns.Count != arrStr.Length)
                {
                    return "列和list集合不匹配";
                }

                table.Rows.Add(ref _missing);
                rowCount++;
                for (int i = 0; i < arrStr.Length; i++)
                {
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
                                MergeCell(table, startRow, i + 1, endRow, i + 1);
                                endRow = 0;
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

            table.Select();
            rowCount = table.Rows.Count;
            for (int i = rowCount; i >= 1; i--)
            {
                if (table.Cell(i, mergeColumn + 1).Range.Text.Equals("")|| table.Cell(i, mergeColumn + 1).Range.Text.Equals("\r\a"))
                {
                    MergeCell(table, i, mergeColumn, i, mergeColumn + 1);
                }
            }
            SetAutoFitContentForTable(table);

            return "保存成功";
        }

        public string InsertContentToWord(string content, string bookmark)
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
                    CloseWord(rtfDoc);
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
        /// 插入其他文件的表格 保留表格列
        /// </summary>
        /// <param name="copyFileFullPath">其他文件的路径</param>
        /// <param name="copyFileTableStartIndex">从第几个开始获取表格</param>
        /// <param name="copyTableColDic">保留表格的列并替换表头</param>
        /// <param name="wordBookmark">需要插入内容的书签</param>
        /// <param name="isCloseTheFile">是否关闭新打开的文件</param>
        /// <returns></returns>
        public string CopyOtherFileTableForColByTableIndex(string copyFileFullPath, int copyFileTableStartIndex, Dictionary<int, string> copyTableColDic, string wordBookmark, bool isCloseTheFile = true)
        {

            try
            {
                Document rtfDoc = OpenWord(copyFileFullPath, true);

                int rtfTableCount = rtfDoc.Tables.Count;

                Range wordTable = GetBookmarkRank(_currentWord, wordBookmark);

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

                    CreateAndGoToNextParagraph(wordTable, i != copyFileTableStartIndex, i != copyFileTableStartIndex);//获取下一个range
                    CreateAndGoToNextParagraph(wordTable, i != copyFileTableStartIndex, i != copyFileTableStartIndex);//InsertBR(wordTable, i <= rtfTableCount);//添加回车
                    wordTable.Paste();

                    ClearFormatTable(wordTable.Tables[1]);
                }

                if (isCloseTheFile)
                    CloseWord(rtfDoc);
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
        public string CopyOtherFilePictureToWord(string copyFileFullPath, int copyFilePictureStartIndex, string workBookmark, bool isCloseTheFile = true)
        {
            try
            {
                Range bookmarkPic = GetBookmarkRank(_currentWord, workBookmark);
                Document copyFileDoc = OpenWord(copyFileFullPath, true);
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
                if (isCloseTheFile)
                    CloseWord(copyFileDoc);
            }
            catch (Exception ex)
            {
                _needWrite = false;
                Dispose();
                throw new Exception(string.Format("错误信息:{0}.{1}", ex.StackTrace.ToString(), ex.Message));
            }
            return "创建成功";
        }

        public string CopyContentToWord(string filePth, string bookmark, bool isCloseTheFile = true)
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
                    CloseWord(htmldoc);
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

        public int ClearWordCodeFormat()
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
        private void AddPicture(string picFileName, Document doc, Range range)
        {
            doc.InlineShapes.AddPicture(picFileName, ref _missing, ref _missing, range);
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
            table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
            table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
            table.Select();
            table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
            table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
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
