using ArticleArray.Functions;
using ArticleArray.MyForms;
using ArticleArray.Verify;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;
using Word = Microsoft.Office.Interop.Word;

namespace ArticleArray
{
    public partial class Ribbon1
    {
        //private Document _document;
        //private StylesForm stylesForm = null;
        private InsertCaptionForm insertCaptionForm = null;
        //private WidthAndHeightSetForm widthAndHeightSetForm = null;
        private string folderPath = @"C:/WordStyleConfig/"; //保存和打开默认文档的路径
        private string licPath = @"C:/WordStyleConfig/"; //保存和打开默认文档的路径

        private readonly string[] styleNames = { "标题 1", "标题 2", "标题 3", "标题 4", "标题 5", "标题", "摘要", "参考文献", "参考文献列表样式", "题注", "目录", "页眉", "页脚", "正文 1" };//"正文 1",
        #region 初始化与加载
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)//加载事件
        {
            if (!Directory.Exists(folderPath)) Directory.CreateDirectory(folderPath);
            if (!Directory.Exists(licPath)) Directory.CreateDirectory(licPath);

            var resourceName = licPath + "Licence";
            if (!File.Exists(resourceName))
            {
                string nowTime = DateTime.Today.AddDays(2).ToString("yyyyMMdd");
                char[] nums = nowTime.ToArray();
                string lic1 = $"ALMPN{nums[0]}HAUUL{nums[1]}QFJWT{nums[2]}CCUQR{nums[3]}VWSCQ{nums[4]}HALPU{nums[5]}HVYNU{nums[6]}SCESV{nums[7]}ZIKTM";
                File.WriteAllText(resourceName, lic1);
            }
            File.SetAttributes(resourceName, File.GetAttributes(resourceName) | FileAttributes.Hidden);//设置文件隐藏

            var verClass = new VerifyLicence(folderPath);
            int verify = verClass.UserVerifyLicence();
            if (verify <= -1)
            {
                BtnSetStyle.Enabled = false;
                BtnTitileConfig.Enabled = false;
                BtnQuickFit.Enabled = false;
                BtnFormatBulletsAndNumbering.Enabled = false;
                BtnCaption.Enabled = false;
                BtnQuoteSuper.Enabled = false;
                BtnSelectedPicSize.Enabled = false;
                BtnSelectedPicOutput.Enabled = false;
                BtnPicOutput.Enabled = false;
                BtnPicSize.Enabled = false;
                BtnConvertToThreeLineTable1.Enabled = false;
                BtnSelectedTableAutoFit.Enabled = false;
                BtnSelectedTableWidth.Enabled = false;
                BtnSelectedToThreeLineTable2.Enabled = false;
                BtnTableAutoFit.Enabled = false;
                BtnTableWidth.Enabled = false;
                BtnEnToChSymbol.Enabled = false;
                BtnRemoveWhitespace.Enabled = false;
                BtnDelEmptyLine.Enabled = false;
                BtnDelSpace.Enabled = false;
                BtnReference.Enabled = false;
                splitButton1.Enabled = false;
                BtnHeader2.Enabled = false;
                BtnRmoveHeaderLine.Enabled = false;
            }

        }
        #endregion

        #region 图片处理
        private void BtnPicSize_Click(object sender, RibbonControlEventArgs e)//所有图片尺寸批量设置
        {
            //if (widthAndHeightSetForm == null)
            //{
            var widthAndHeightSetForm = new WidthAndHeightSetForm();
            //}
            var result = widthAndHeightSetForm.ShowDialog();
            float width, height;
            bool isCenter;
            if (result != DialogResult.None)
            {
                // 获取用户输入的两个值，并进行处理
                width = widthAndHeightSetForm.width / 2.54f * 72f;
                height = widthAndHeightSetForm.height / 2.54f * 72f;
                isCenter = widthAndHeightSetForm.isCenter;
            }
            else
            {
                return;
            }

            Range range = Globals.ThisAddIn.Application.ActiveDocument.Range();
            InlineShapes inlineShapes = range.InlineShapes;
            Word.ShapeRange shapeRange = range.ShapeRange;

            if (inlineShapes.Count > 0) { CommonMethods.SetInlineShapes(width, height, isCenter, inlineShapes); }
            if (shapeRange.Count > 0) { CommonMethods.SetLineShapes(width, height, isCenter, shapeRange); }
            widthAndHeightSetForm.Dispose();
        }
        private void BtnSelectedPicSize_Click(object sender, RibbonControlEventArgs e)//范围图片尺寸设置
        {
            //if (widthAndHeightSetForm == null)
            //{
            var widthAndHeightSetForm = new WidthAndHeightSetForm();
            //}
            var result = widthAndHeightSetForm.ShowDialog();
            float width, height;
            bool isCenter;
            if (result != DialogResult.None)
            {
                // 获取用户输入的两个值，并进行处理
                width = widthAndHeightSetForm.width / 2.54f * 72f;
                height = widthAndHeightSetForm.height / 2.54f * 72f;
                isCenter = widthAndHeightSetForm.isCenter;
            }
            else
            {
                return;
            }

            Selection selection = Globals.ThisAddIn.Application.Selection;
            InlineShapes inlineShapes = selection.InlineShapes;
            Word.ShapeRange shapeRange = selection.ShapeRange;
            if (inlineShapes.Count > 0) { CommonMethods.SetInlineShapes(width, height, isCenter, inlineShapes); }
            if (shapeRange.Count > 0) { CommonMethods.SetLineShapes(width, height, isCenter, shapeRange); }
            widthAndHeightSetForm.Dispose();
        }
        private void BtnPicOutput_Click(object sender, RibbonControlEventArgs e)//所有图片导出
        {
            Range range = Globals.ThisAddIn.Application.ActiveDocument.Range();
            InlineShapes inlineShapes = range.InlineShapes;
            if (inlineShapes.Count == 0) return;
            var dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() != DialogResult.OK) return;
            var exportFolder = dialog.SelectedPath;
            CommonMethods.ExportImagePng(exportFolder, inlineShapes);
        }
        private void BtnSelectedPicOutput_Click(object sender, RibbonControlEventArgs e)//范围导出图片
        {
            Selection selection = Globals.ThisAddIn.Application.Selection;
            InlineShapes inlineShapes = selection.InlineShapes;
            if (inlineShapes.Count == 0) return;
            var dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() != DialogResult.OK) return;
            var exportFolder = dialog.SelectedPath;
            CommonMethods.ExportImagePng(exportFolder, inlineShapes);
        }
        private void BtnCaption_Click(object sender, RibbonControlEventArgs e) //插入图片与表格题注
        {
            if (insertCaptionForm == null)
            {
                insertCaptionForm = new InsertCaptionForm(styleNames);
            }
            insertCaptionForm.TopMost = true;
            insertCaptionForm.ShowDialog();


        }
        #endregion

        #region 表格处理
        private void BtnConvertToThreeLineTable_Click(object sender, RibbonControlEventArgs e) //所有表格转三线表
        {
            var _document = Globals.ThisAddIn.Application.ActiveDocument;
            Tables tables = _document.Tables;
            if (tables.Count > 0) CommonMethods.TableToThreeLineTableStyle1(tables);
        }
        private void BtnSelectedToThreeLineTable_Click(object sender, RibbonControlEventArgs e)//选中表格转三线表
        {
            Selection selection = Globals.ThisAddIn.Application.Selection;
            if (selection.Text == null) return;
            Tables tables = selection.Tables;
            if (tables.Count > 0) CommonMethods.TableToThreeLineTableStyle1(tables);
        }

        private void BtnConvertToThreeLineTableStyle2_Click(object sender, RibbonControlEventArgs e) //所有表格转三线表
        {
            var _document = Globals.ThisAddIn.Application.ActiveDocument;
            Tables tables = _document.Tables;
            if (tables.Count > 0) CommonMethods.TableToThreeLineTableStyle2(tables);
        }
        private void BtnSelectedToThreeLineTableStyle2_Click(object sender, RibbonControlEventArgs e)//选中表格转三线表
        {
            Selection selection = Globals.ThisAddIn.Application.Selection;
            if (selection.Text == null) return;
            Tables tables = selection.Tables;
            if (tables.Count > 0) CommonMethods.TableToThreeLineTableStyle2(tables);
        }

        private void BtnTableAutoFit_Click(object sender, RibbonControlEventArgs e)//表格内容自适应
        {
            var _document = Globals.ThisAddIn.Application.ActiveDocument;
            foreach (Word.Table tbl in _document.Tables)
            {
                tbl.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
                tbl.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
            }
            MessageBox.Show("表格自适应设置完成", "提示");
        }
        private void BtnSelectedTableAutoFit_Click(object sender, RibbonControlEventArgs e)//选中表格内容自适应
        {
            Selection selection = Globals.ThisAddIn.Application.Selection;
            if (selection.Text == null) return;
            Tables tables = selection.Tables;
            foreach (Word.Table tbl in tables)
            {
                tbl.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
                tbl.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
            }
            MessageBox.Show("选中表格自适应设置完成", "提示");
        }
        private void BtnTableWidth_Click(object sender, RibbonControlEventArgs e)//设置表格宽度与页面宽度相同
        {
            // 获取当前文档对象
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            // 获取文档页面宽度
            float pageWidth = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin;

            List<Task> tasks = new List<Task>();
            // 遍历文档中所有表格
            foreach (Word.Table tbl in doc.Tables)
            {
                Task task = Task.Run(() =>
                {
                    // 设置表格的宽度
                    tbl.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints;
                    tbl.PreferredWidth = pageWidth;
                });
                tasks.Add(task);
            }
            Task.WaitAll(tasks.ToArray());
            if (tasks.Count > 0) MessageBox.Show("表格宽度设置完成！", "提示");
        }
        private void BtnSelectedTableWidth_Click(object sender, RibbonControlEventArgs e)//设置表格宽度与页面宽度相同
        {
            // 获取当前文档对象
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Range range = Globals.ThisAddIn.Application.Selection.Range;

            // 获取文档页面宽度
            float pageWidth = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin;

            List<Task> tasks = new List<Task>();
            // 遍历文档中所有表格
            foreach (Word.Table tbl in range.Tables)
            {
                Task task = Task.Run(() =>
                {
                    // 设置表格的宽度
                    tbl.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints;
                    tbl.PreferredWidth = pageWidth;
                });
                tasks.Add(task);
            }
            Task.WaitAll(tasks.ToArray());
            if (tasks.Count > 0) MessageBox.Show("表格宽度设置完成！", "提示");
        }
        #endregion

        #region 目录处理
        private void BtnCatalog3Level_Click(object sender, RibbonControlEventArgs e) //三级目录
        {
            var _document = Globals.ThisAddIn.Application.ActiveDocument;
            Range range = Globals.ThisAddIn.Application.Selection.Range;


            if (_document.TablesOfContents.Count > 0)
            {
                TableOfContents toc = _document.TablesOfContents[1];
                toc.Update();
            }
            else
            {
                range.InsertAfter("目  录\r");
                range.set_Style("目录");
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                // 插入分页符，确保目录插入到新页上
                TableOfContents toc = _document.TablesOfContents.Add(
                                Range: range, // 将目录插入到文档开头
                                UseHeadingStyles: true, // 使用内置标题样式
                                UpperHeadingLevel: 1, // 指定目录的起始级别
                                LowerHeadingLevel: 3, //目录的结束标题级别
                                UseFields: false,//如果使用目录项 (TC) 域创建的目录
                                "目录",  //Type.Missing, // 标题样式
                                IncludePageNumbers: true, // 显示页码
                                UseOutlineLevels: true
                                ); ;
                toc.Range.Paragraphs.SpaceBefore = 0; // 目录前段落的间距
                toc.Range.Paragraphs.SpaceAfter = 0; // 目录后段落的间距

                toc.UpdatePageNumbers(); // 更新目录页码

                //range = Globals.ThisAddIn.Application.Selection.Range;
                //range.Collapse(WdCollapseDirection.wdCollapseEnd);

                //range.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
            }
        }
        private void BtnInsertCatalog_Click(object sender, RibbonControlEventArgs e) //插入自定义目录
        {
            try
            {
                Word.Dialog dialog = Globals.ThisAddIn.Application.Dialogs[Word.WdWordDialog.wdDialogInsertTableOfContents];
                dialog.Show();
            }
            catch (Exception)
            {
                return;
            }
        }
        #endregion

        #region 文本处理
        private void BtnTxtToTable_Click(object sender, RibbonControlEventArgs e) //文本转表格
        {
            try
            {
                Word.Dialog dialog = Globals.ThisAddIn.Application.Dialogs[Word.WdWordDialog.wdDialogTextToTable];
                dialog.Show();
            }
            catch (Exception)
            {
                return;
            }
        }
        private void BtnDelEmptyLine_Click(object sender, RibbonControlEventArgs e)//全文删除空行
        {
            Word.Range range = Globals.ThisAddIn.Application.ActiveDocument.Range();

            CommonMethods.DeleteEmptyParagraphs(range);

        }
        private void BtnSelectedDelEmptyLine_Click(object sender, RibbonControlEventArgs e)//范围删除空行
        {
            Word.Range range = Globals.ThisAddIn.Application.Selection.Range;

            CommonMethods.DeleteEmptyParagraphs(range);
        }
        private void BtnRemoveWhitespace_Click(object sender, RibbonControlEventArgs e)//合并范围内文本
        {
            Word.Range range = Globals.ThisAddIn.Application.Selection.Range;
            // 刷新段落样式
            if (range.Text == null) return;
            object style = range.get_Style();
            string text = range.Text;
            // 移除中文之间的空格
            text = Regex.Replace(text, @"(\S)\s+([\u4E00-\u9FFF])", "$1$2");

            // 移除英文之间的多余空格
            text = Regex.Replace(text, @"(\S)\s+(\S)", "$1 $2");

            // 移除中文和英文之间的空格
            text = Regex.Replace(text, @"([\u4E00-\u9FFF])\s+(\S)", "$1$2");
            text = Regex.Replace(text, @"(\S)\s+([\u4E00-\u9FFF])", "$1$2");

            // 移除段落末尾空格，但保留段落末尾的换行符
            //text = Regex.Replace(text, @"(?<=[^\r])\s+$", "");
            //text = text.TrimStart();
            if (text.EndsWith("\r"))
            {
                text = text.Trim();
                text = Regex.Replace(text, @"[\r\n\t]+", " ") + "\r";
            }
            else
            {
                text = text.Trim();
                // 移除换行符、段落符、制表符
                text = Regex.Replace(text, @"[\r\n\t]+", " ");
            }
            range.Text = text;
            if (style == null) return;
            range.set_Style(ref style);

        }
        private void BtnDelSpace_Click(object sender, RibbonControlEventArgs e) //删除中文之间空格
        {
            Word.Range range = Globals.ThisAddIn.Application.Selection.Range;
            CommonMethods.DelSelectedSpace(range);
        }
        private void BtnEnToChSymbol_Click(object sender, RibbonControlEventArgs e) //英文符号转中文
        {
            try
            {
                Range selectionRange = Globals.ThisAddIn.Application.Selection.Range;
                var doc = selectionRange.Document;
                if (selectionRange.Text == null) return;
                //string pattern = @"[\p{P}]";//匹配所有的英文符号
                string pattern = "[\"'.,;:)(-]";//匹配所有的符号
                // 定义替换符号映射表
                var symbolMap2 = new Dictionary<string, string>()
                {
                    { "\"", "“" },
                    { "'" , "‘"},
                    { "." , "。"},
                    { "," , "，"},
                    { ";" , "；"},
                    { ":" , "："},
                    { "!" , "！"},
                    { "?" , "？"},
                    { "(" , "（"},
                    { ")" , "）"},
                    { "[" , "【"},
                    { "]" , "】"},
                    { "-" , "—"}
                };
                Word.Find find = selectionRange.Find;
                find.ClearFormatting();
                find.MatchWildcards = true;
                find.Text = pattern;
                int end = selectionRange.End;
                while (find.Execute(Forward: true) && selectionRange.InRange(doc.Content))
                {
                    if (end <= selectionRange.End) break;
                    // 获取包含符号的文本
                    string symbolText = selectionRange.Text;
                    foreach (var symbol in symbolMap2)
                    {
                        symbolText = symbolText.Replace(symbol.Key, symbol.Value);
                    }
                    // 替换后的文本赋值回 Range.Text
                    selectionRange.Text = symbolText;
                    selectionRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                }
            }
            catch (Exception)
            {
                return;
            }
        }
        private void BtnChToEnSymbol_Click(object sender, RibbonControlEventArgs e) //中文符号转英文
        {
            CommonMethods.SymbolChToEn();
        }
        private void BtnQuoteSuper_Click(object sender, RibbonControlEventArgs e) //区域引文变上标
        {
            var _document = Globals.ThisAddIn.Application.ActiveDocument;

            Range range = Globals.ThisAddIn.Application.Selection.Range;

            if (string.IsNullOrEmpty(range.Text))
            {
                //选中文本为空则跳出
                return;
            }

            string pattern = @"\[(.*?)\]";
            Regex regex = new Regex(pattern);
            // 遍历选中范围内所有匹配的方括号
            foreach (Match match in regex.Matches(range.Text))
            {
                Word.Range subRange = _document.Range(match.Index + range.Start, match.Index + range.Start + match.Length);
                subRange.Font.Superscript = 1;
            }
        }
        private void BtnDelFormat_Click(object sender, RibbonControlEventArgs e) //移除所有格式
        {
            CommonMethods.DelSelectedFormat();
        }
        #endregion

        #region 样式处理
        private void BtnSetStyle_Click(object sender, RibbonControlEventArgs e)//设置样式
        {
            CommonMethods.CreateStyles(styleNames);
            var stylesForm = new StylesForm(styleNames);
            stylesForm.Show();
        }
        private void BtnFormatBulletsAndNumbering_Click(object sender, RibbonControlEventArgs e)//自定义符号与编号
        {
            Word.Dialog dialog = Globals.ThisAddIn.Application.Dialogs[Word.WdWordDialog.wdDialogFormatBulletsAndNumbering];
            dialog.Show();
        }
        private void BtnStyleManage_Click(object sender, RibbonControlEventArgs e) //样式管理
        {
            Word.Dialog styleManageDialog = Globals.ThisAddIn.Application.Dialogs[Word.WdWordDialog.wdDialogStyleManagement];  //样式管理窗口
            styleManageDialog.Show();
        }
        private void BtnTitileConfig_Click(object sender, RibbonControlEventArgs e)//标题快速识别
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            BtnTitileConfigForm btnTitileConfigForm = new BtnTitileConfigForm();
            btnTitileConfigForm.ShowDialog();
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
        #endregion

        #region 模板插入
        private void BtnPageNumber_Click(object sender, RibbonControlEventArgs e)//页码设置与插入
        {
            Word.Dialog dialog = Globals.ThisAddIn.Application.Dialogs[Word.WdWordDialog.wdDialogInsertPageNumbers];
            dialog.Show();
        }
        private void BtnCusBlockManage_Click(object sender, RibbonControlEventArgs e)//自定义块管理
        {
            Word.Dialog dialog = Globals.ThisAddIn.Application.Dialogs[Word.WdWordDialog.wdDialogBuildingBlockOrganizer];
            dialog.Show();
        }
        private void BtnSymbol_Click(object sender, RibbonControlEventArgs e) //特殊符号
        {
            Word.Dialog dialog = Globals.ThisAddIn.Application.Dialogs[Word.WdWordDialog.wdDialogInsertSymbol];
            dialog.Show();
        }
        private void BtnListCommands_Click(object sender, RibbonControlEventArgs e) //快捷键
        {
            Word.Dialog dialog = Globals.ThisAddIn.Application.Dialogs[Word.WdWordDialog.wdDialogListCommands];
            dialog.Show();
        }
        #endregion

        #region 页面设置
        private void BtnDivisionHeading1_Click(object sender, RibbonControlEventArgs e) //一级标题分页
        {

            // 获取当前活动文档
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            // 获取一级标题样式名称
            string heading1Style = "标题 1";
            // 获取分节符的 Unicode 码值
            int sectionBreakUnicode = 7;
            foreach (Word.Paragraph para in doc.Paragraphs)
            {
                try
                {
                    // 判断段落是否为一级标题
                    if (para.get_Style().NameLocal == heading1Style)
                    {
                        // 获取段落的范围
                        Word.Range range = para.Range;

                        int start = range.Start;
                        int end = range.End;
                        // 判断该范围内是否存在分节符
                        for (int i = start; i < end; i++)
                        {
                            Word.Range r = doc.Range(i, i + 1);
                            if (r.Text == "\f")
                            {
                                // 将分页符替换为分节符
                                r.Text = ((char)sectionBreakUnicode).ToString();
                            }
                        }
                        if (range.Text != null)
                        {
                            // 判断段落末尾是否已经存在分节符
                            if (range.Text.EndsWith(((char)sectionBreakUnicode).ToString()))
                            {
                                continue;
                            }
                            range.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.ToString());
                    continue;
                }
            }

        }
        private void BtnPageSetup_Click(object sender, RibbonControlEventArgs e) //页面设置
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            Word.Dialog dialog = Globals.ThisAddIn.Application.Dialogs[Word.WdWordDialog.wdDialogFilePageSetup];
            dialog.Show();
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
        private void BtnHeader_Click(object sender, RibbonControlEventArgs e)//页眉设置
        {
            // 获取当前文档
            var _document = Globals.ThisAddIn.Application.ActiveDocument;

            CommonMethods.CreateStyles(styleNames);//创建样式

            //开启奇偶页不同
            _document.PageSetup.OddAndEvenPagesHeaderFooter = -1;
            _document.PageSetup.DifferentFirstPageHeaderFooter = 0;
            foreach (Word.Section section in _document.Sections)
            {
                // 设置奇数页页眉
                section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
                Word.Range oddHeaderRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                //oddHeaderRange.Delete();
                Word.Paragraph para = section.Range.Paragraphs[1];
                try
                {
                    var text = para.Range.Text.Trim();
                    var text2 = int.Parse(para.Range.ListFormat.ListString.Trim());
                    text = text.Replace("\r\n", "");
                    oddHeaderRange.Text = "第" + CommonMethods.ConvertToChineseNumber(text2) + "章 " + text;
                }
                catch (Exception)
                {
                }
                oddHeaderRange.set_Style("页眉");
                oddHeaderRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                //oddHeaderRange.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                oddHeaderRange.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                oddHeaderRange.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                oddHeaderRange.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                oddHeaderRange.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;

                section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = true;
                // 设置偶数页页眉
                Word.Range evenHeaderRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range;
                evenHeaderRange.Text = "双击此处设置您的偶数页眉文字";
                evenHeaderRange.set_Style("页脚");

                //evenHeaderRange.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                evenHeaderRange.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                evenHeaderRange.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                evenHeaderRange.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                evenHeaderRange.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;

                evenHeaderRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }
        }


        private void BtnRmoveHeaderLine_Click(object sender, RibbonControlEventArgs e) //移除页眉线条
        {
            CommonMethods.RmHeaderLine();
        }
        #endregion

        #region 多级列表设置
        private void BtnToSpace_Click(object sender, RibbonControlEventArgs e)//转空格
        {
            // 获取当前文档
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            // 获取多级列表对象
            WdTrailingCharacter wdTrailingCharacter = WdTrailingCharacter.wdTrailingSpace;
            CommonMethods.SetListTrailingCharacters(doc, wdTrailingCharacter);
        }
        private void BtnToNone_Click(object sender, RibbonControlEventArgs e)//移除缩进
        {
            // 获取当前文档
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            // 获取多级列表对象
            WdTrailingCharacter wdTrailingCharacter = WdTrailingCharacter.wdTrailingNone;
            //ListTemplate UsinglistTemplate = doc.Range().ListFormat.ListTemplate;

            CommonMethods.SetListTrailingCharacters(doc, wdTrailingCharacter);
        }
        private void BtnToTab_Click(object sender, RibbonControlEventArgs e)//转制表符
        {
            // 获取当前文档
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            // 获取多级列表对象
            WdTrailingCharacter wdTrailingCharacter = WdTrailingCharacter.wdTrailingTab;
            //ListTemplate UsinglistTemplate = doc.Range().ListFormat.ListTemplate;

            CommonMethods.SetListTrailingCharacters(doc, wdTrailingCharacter);
        }
        #endregion

        private void BtnCrossReference_Click(object sender, RibbonControlEventArgs e)//交叉引用
        {
            // 获取交叉引用对话框对象
            Word.Dialog crossRefDialog = Globals.ThisAddIn.Application.Dialogs[Word.WdWordDialog.wdDialogInsertCrossReference];
            crossRefDialog.Show();
        }

        private void BtnQuickFit_Click(object sender, RibbonControlEventArgs e)//一键排版
        {
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            var filename = folderPath + "DefaultStyles.dotx";
            var newPath = doc.Path + "\\(新样式)" + doc.Name;

            Word.Document newDoc = Globals.ThisAddIn.Application.Documents.Open(filename);
            newDoc.Content.FormattedText = doc.Content.FormattedText;

            newDoc.SaveAs2(newPath, Word.WdSaveFormat.wdFormatXMLDocument);
            doc.Close(SaveChanges: true);
            Clipboard.Clear();
            Range range = Globals.ThisAddIn.Application.ActiveDocument.Range();
            CommonMethods.RefreshTitle(range);

            CommonMethods.AutoSetNormal1(newDoc);
        }

        private void BtnRefreshParaStyle_Click(object sender, RibbonControlEventArgs e)//刷新选中段落样式
        {
            Range range = Globals.ThisAddIn.Application.Selection.Range;
            CommonMethods.RefreshSelectedRange(range);
        }

        private void BtnFormulaIndex_Click(object sender, RibbonControlEventArgs e)
        {// TODO 此功能还有问题，以后再解决
            Word.Document activeDocument = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Paragraphs paragraphs = activeDocument.Paragraphs;

            // 创建样式
            Word.Style equationStyle;
            try
            {
                equationStyle = activeDocument.Styles["公式"];
            }
            catch (Exception)
            {
                equationStyle = activeDocument.Styles.Add("公式");
            }
            equationStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            Word.TabStops tabStops = equationStyle.ParagraphFormat.TabStops;
            //    tabStops.ClearAll();

            // 设置样式属性
            equationStyle.set_BaseStyle("正文");

            foreach (Word.Paragraph para in paragraphs)
            {
                if (para.Range.OMaths.Count > 0) // 判断段落是否包含公式
                {
                    //if (para.Format.Alignment == WdParagraphAlignment.wdAlignParagraphRight) { para.Range.Fields.Update();continue; }//如果居右跳过
                    // 将插入点移到段落的末尾
                    Word.Range selectionRange = para.Range;

                    //selectionRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    selectionRange.MoveEnd(Word.WdUnits.wdCharacter, Count: -1);
                    selectionRange.Move(Word.WdUnits.wdCharacter, Count: 1);

                    //创建一个新的range
                    //selectionRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    Word.Field autoNumField = activeDocument.Fields.Add(selectionRange, Word.WdFieldType.wdFieldEmpty, "AUTONUM", false);

                    autoNumField.Update();
                    autoNumField.Result.InsertAfter(")");

                    Word.Field styleRefField = activeDocument.Fields.Add(selectionRange, Word.WdFieldType.wdFieldEmpty, "STYLEREF \"标题 1\" \\n", true);
                    styleRefField.Update();
                    styleRefField.Result.InsertBefore("(");
                    styleRefField.Result.InsertAfter(".");

                    para.set_Style(equationStyle);
                }
            }
        }

        //    private void BtnFormulaIndex_Click(object sender, RibbonControlEventArgs e)//公式自动序号
        //{
        //    // 获取文档对象
        //    Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

        //// 创建样式
        //Word.Style equationStyle;
        //try
        //{
        //    equationStyle = doc.Styles["公式"];
        //}
        //catch (Exception)
        //{
        //    equationStyle = doc.Styles.Add("公式");
        //}

        //// 设置样式属性
        //equationStyle.set_BaseStyle("正文");
        //    equationStyle.NoSpaceBetweenParagraphsOfSameStyle = false;
        //    equationStyle.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
        //    equationStyle.ParagraphFormat.LineUnitBefore = 0;
        //    equationStyle.ParagraphFormat.LineUnitAfter = 0;
        //    equationStyle.ParagraphFormat.SpaceAfter = 0;

        //    // 设置制表符
        //    Word.TabStops tabStops = equationStyle.ParagraphFormat.TabStops;
        //    tabStops.ClearAll();

        //    Word.Paragraphs paragraphs = Globals.ThisAddIn.Application.ActiveDocument.Paragraphs;
        //    Word.ListTemplate listTemplate = Globals.ThisAddIn.Application.ActiveDocument.ListTemplates[WdListGalleryType.wdNumberGallery];

        //    foreach (Word.Paragraph para in paragraphs)
        //    {
        //        if (para.Range.OMaths.Count > 0) // 判断段落是否包含公式
        //        {
        //            // 将插入点移到公式结尾
        //            Word.Range insertRange = para.Range.OMaths[1].Range;
        //            insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
        //            // 在公式后插入序号
        //            insertRange.Fields.Add(insertRange, Word.WdFieldType.wdFieldEmpty, "AUTONUM ", false);
        //            insertRange.InsertAfter("#()\n");
        //        }
        //    }
        //}

        private void BtnRmFooterLine_Click(object sender, RibbonControlEventArgs e)//移除页脚线
        {
            CommonMethods.RmFooterLine();
        }

        private void BtnGetLicence_Click(object sender, RibbonControlEventArgs e)//验证
        {
            var verifyForm = new VerifyForm(licPath);
            verifyForm.Show();
        }

        private void BtnMultilevelList_Click(object sender, RibbonControlEventArgs e)//多级列表
        {
            Word.Application wordApp = Globals.ThisAddIn.Application;

            var doc = wordApp.ActiveDocument;

            // 获取默认的多级列表模板
            ListLevel listLevel1 = doc.ListTemplates[1].ListLevels[1];

            listLevel1.NumberFormat = "第%1章";
            listLevel1.TrailingCharacter = WdTrailingCharacter.wdTrailingSpace;//可设置后面的tab或者空格或者无
            listLevel1.NumberStyle = WdListNumberStyle.wdListNumberStyleSimpChinNum3;//简体中文,可选
            listLevel1.NumberPosition = 0f;
            listLevel1.Alignment = WdListLevelAlignment.wdListLevelAlignLeft;//可选
            listLevel1.TabPosition = 0f;
            listLevel1.StartAt = 1; //可设置
            listLevel1.LinkedStyle = "标题 1";//自定义关联样式

        }

        private void BtnRefreshTitle_Click(object sender, RibbonControlEventArgs e)//刷新标题
        {
            Range range = Globals.ThisAddIn.Application.ActiveDocument.Range();
            CommonMethods.RefreshTitle(range);
        }

        private void BtnSetStyleNormal1_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            CommonMethods.AutoSetNormal1(doc);
        }

        private void BtnReference_Click(object sender, RibbonControlEventArgs e)//参考文献
        {
            var selectedRange = Globals.ThisAddIn.Application.Selection.Range;
            if (selectedRange.Text == "" || selectedRange.Text == "\r" || selectedRange.Text == null) return;
            selectedRange.Text = Regex.Replace(selectedRange.Text, @"[ ]+", " ");
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            Style referenceStyle;
            try
            {
                referenceStyle = doc.Styles.Add("参考文献列表样式", Word.WdStyleType.wdStyleTypeParagraph);
            }
            catch (Exception) { referenceStyle = doc.Styles["参考文献列表样式"]; }

            // 设置段落格式
            referenceStyle.ParagraphFormat.CharacterUnitLeftIndent = 0;
            referenceStyle.ParagraphFormat.CharacterUnitFirstLineIndent = -2;
            ListTemplate listTemplate;
            // 设置列表符号样式
            try { listTemplate = doc.ListTemplates.Add(false, "参考文献列表模板"); } catch { listTemplate = doc.ListTemplates["参考文献列表模板"]; }
            listTemplate.ListLevels[1].NumberFormat = @"[%1]"; // 第一级编号格式为 "[%1]"
            listTemplate.ListLevels[1].TrailingCharacter = WdTrailingCharacter.wdTrailingTab;//要求是tab
            referenceStyle.LinkToListTemplate(listTemplate, 1);
            selectedRange.set_Style(referenceStyle);
            CommonMethods.SymbolChToEn();//符号转为英文符号
        }

        private void BtnTitleUp_Click(object sender, RibbonControlEventArgs e)//选中区域标题升级
        {
            Range range = Globals.ThisAddIn.Application.Selection.Range;
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            CommonMethods.TitleUpgrade(doc, range);
        }

        private void BtnTitleDown_Click(object sender, RibbonControlEventArgs e)//选中标题降级
        {
            Range range = Globals.ThisAddIn.Application.Selection.Range;
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            CommonMethods.TitleDowngrade(doc, range);
        }

        private void BtnDocTitleUp_Click(object sender, RibbonControlEventArgs e)//全部标题升级
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Range range = doc.Range();

            CommonMethods.TitleUpgrade(doc, range);
        }

        private void BtnDocTitleDown_Click(object sender, RibbonControlEventArgs e)//全部标记降级
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Range range = doc.Range();

            CommonMethods.TitleDowngrade(doc, range);
        }

        private void BtnHeader2_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前文档
            var _document = Globals.ThisAddIn.Application.ActiveDocument;

            CommonMethods.CreateStyles(styleNames);//创建样式

            //开启奇偶页不同
            _document.PageSetup.OddAndEvenPagesHeaderFooter = -1;
            _document.PageSetup.DifferentFirstPageHeaderFooter = 0;

            foreach (Word.Section section in _document.Sections)
            {
                // 设置奇数页页眉
                section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = true;
                Word.Range oddHeaderRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

                oddHeaderRange.Text = "双击输入奇数页眉";
                oddHeaderRange.set_Style("页眉");
                oddHeaderRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                //oddHeaderRange.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                oddHeaderRange.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                oddHeaderRange.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                oddHeaderRange.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                oddHeaderRange.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;

                // 设置偶数页页眉
                Word.Range evenHeaderRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range;
                evenHeaderRange.Text = "双击设置您的偶数页眉";
                evenHeaderRange.set_Style("页脚");
                section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = true;
                //evenHeaderRange.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                evenHeaderRange.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                evenHeaderRange.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                evenHeaderRange.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                evenHeaderRange.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;

                evenHeaderRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }

        }

        private void BtnOneKeyEditorSetting_Click(object sender, RibbonControlEventArgs e)
        {

            var oneKeyEditor = new OneKeyEditorForm(folderPath);
            oneKeyEditor.Show();
        }
    }
}