using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Document = Microsoft.Office.Interop.Word.Document;
using Task = System.Threading.Tasks.Task;
using Word = Microsoft.Office.Interop.Word;

namespace ArticleArray.Functions
{
    public static class CommonMethods
    {
        public static void ReplaceCharacters(Range range, string findText, string replaceText)
        {
            //Word.Range range = document.Content;
            Word.Find find = range.Find;

            // 设置查找参数
            find.ClearFormatting();
            find.Text = findText;
            int end = range.End;
            // 执行查找和替换操作
            while (find.Execute())
            {
                range.Text = replaceText;
                if (range.End >= end) break;
            }
        }
        public static void TitleDowngrade(Word.Document doc, Range range)//标题降级
        {
            foreach (Paragraph paragraph in range.Paragraphs)
            {
                var style = paragraph.get_Style() as Word.Style;
                if (style != null && style.NameLocal.Contains("标题"))
                {
                    int j = 1;
                    while (j <= 5)
                    {
                        if (style != null && style.NameLocal == $"标题 {j}")
                        {
                            var newStyle = doc.Styles[$"标题 {j + 1}"];

                            if (newStyle != null)
                            {
                                // 应用新的样式
                                paragraph.set_Style(newStyle);
                            }
                        }
                        j++;
                    }
                }
            }
        }

        public static void TitleUpgrade(Document doc, Range range)//标题升级
        {
            foreach (Paragraph paragraph in range.Paragraphs)
            {
                var style = paragraph.get_Style() as Word.Style;
                if (style != null && style.NameLocal.Contains("标题"))
                {
                    int j = 5;
                    while (j > 1)
                    {
                        if (style != null && style.NameLocal == $"标题 {j}")
                        {
                            var newStyle = doc.Styles[$"标题 {j - 1}"];

                            if (newStyle != null)
                            {
                                // 应用新的样式
                                paragraph.set_Style(newStyle);
                            }
                        }
                        j--;
                    }
                }
            }
        }

        public static void SymbolChToEn()
        {
            try
            {
                Range selectionRange = Globals.ThisAddIn.Application.Selection.Range;
                if (selectionRange.Text == null) return;
                // 定义正则表达式匹配模式
                string pattern = "[“”‘’。，；：（）【】—、]";//匹配所有的符号
                // 定义替换符号映射表
                var symbolMap = new Dictionary<string, string>()
                {
                    { "“", "\"" },
                    { "”", "\"" },
                    { "‘", "'" },
                    { "’", "'" },
                    { "。", "." },
                    { "，", "," },
                    { "；", ";" },
                    { "：", ":" },
                    { "！", "!" },
                    { "？", "?" },
                    { "（", "(" },
                    { "）", ")" },
                    { "【", "[" },
                    { "】", "]" },
                    { "—", "-" },
                    { "、", "," },
                };
                // 将所有的中文符号替换为英文符号
                Word.Find find = selectionRange.Find;
                find.ClearFormatting();
                find.MatchWildcards = true;
                find.Text = pattern;
                int end = selectionRange.End;
                while (find.Execute())
                {
                    // 获取包含符号的文本
                    if (end <= selectionRange.End) break;
                    string symbolText = selectionRange.Text;
                    foreach (var symbol in symbolMap)
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

        public static void AutoSetNormal1(Document doc)
        {
            Word.Range documentRange = doc.Content;
            int startPosition = 1;
            foreach (Word.Paragraph paragraph in documentRange.Paragraphs)
            {
                string nameLocal = paragraph.get_Style().NameLocal;
                if (nameLocal == null || nameLocal.Length == 0) { continue; }
                if (!nameLocal.Contains("正文"))
                {
                    startPosition = paragraph.Range.Start;
                    break;
                }
            }
            Word.Range newRange = doc.Range(startPosition, doc.Range().End);

            CommonMethods.SetAfterTitleStyleNormal1(doc, newRange);
        }

        public static void SetAfterTitleStyleNormal1(Document doc, Range range)//范围设置为正文转正文1
        {
            if (range == null)
            {
                range = doc.Range();
            }
            // 开始遍历文档的每个段落
            foreach (Word.Paragraph paragraph in range.Paragraphs)
            {
                // 检查段落的样式名
                if (paragraph.get_Style().NameLocal == "正文") paragraph.set_Style(doc.Styles["正文 1"]);
            }
        }

        /// <summary>
        /// 删除选中范围的中文空格
        /// </summary>
        /// <param name="range"></param>
        public static void DelSelectedSpace(Range range)
        {
            if (range.Text == null) return;
            for (int i = 1; i <= range.Paragraphs.Count; i++)
            {
                object style = range.Paragraphs[i].get_Style();
                string text = range.Paragraphs[i].Range.Text;
                if (text == "\r" || text == "\n" || text == "\r\n") continue;
                text = Regex.Replace(text, @"(\S)[ ]+(\S)", "$1 $2"); // 移除英文之间的多余空格
                text = Regex.Replace(text, @"(?<![a-zA-Z])[ ]+(?![a-zA-Z])", "");//移除非英文字符间的空格
                text = text.TrimStart(); //移除段前空格
                if (text.EndsWith("\r") && i != range.Paragraphs.Count) // 检查段末是否有空格
                {
                    text = text.TrimEnd() + "\r"; // 移除段末空格，但不移除换行符
                }
                else if (i == 1) { }
                else text = text.TrimEnd();
                range.Paragraphs[i].Range.Text = text;
                range.Paragraphs[i].set_Style(ref style);
            }
        }
        public static string ConvertToChineseNumber(int number)//数字转中文
        {
            string[] chineseDigits = { "零", "一", "二", "三", "四", "五", "六", "七", "八", "九" };
            string[] chineseUnits = { "", "十", "百", "千", "万", "亿" };

            if (number == 0)
            {
                return chineseDigits[0];
            }

            string result = "";
            int unitIndex = 0;

            while (number > 0)
            {
                int digit = number % 10;
                if (digit != 0)
                {
                    string digitStr = chineseDigits[digit];
                    string unitStr = chineseUnits[unitIndex];
                    result = digitStr + unitStr + result;
                }
                else if (result.Length > 0 && result[0] != chineseDigits[0][0])
                {
                    result = chineseDigits[0] + result;
                }
                unitIndex++;
                number /= 10;
            }

            return result;
        }
        /// <summary>
        /// 创建样式，正文1继承自正文
        /// </summary>
        /// <param name="styleNames"></param>
        public static void CreateStyles(string[] styleNames)
        {
            // 获取当前文档
            var _document = Globals.ThisAddIn.Application.ActiveDocument;
            //添加正文 1
            try
            {
                var st = _document.Styles.Add("正文 1");
                st.set_BaseStyle(_document.Styles["正文"]);
                st.set_NextParagraphStyle("正文 1");
            }
            catch (Exception) { }
            foreach (string styleName in styleNames)
            {
                Word.Style style;
                try { style = _document.Styles[styleName]; }
                catch (Exception) { style = _document.Styles.Add(styleName); }

                if (styleName != "正文 1" && styleName != "正文")
                {
                    style.set_BaseStyle(_document.Styles["正文"]);
                    style.set_NextParagraphStyle("正文 1");
                }
            }
        }


        /// <summary>
        /// 刷新选中区域标题样式
        /// </summary>
        public static void RefreshTitle(Range range)
        {
            //Range range = Globals.ThisAddIn.Application.ActiveDocument.Range();
            // 遍历文档中的所有段落
            if (range.Text == null) return;
            foreach (Word.Paragraph paragraph in range.Paragraphs)
            {
                // 刷新段落样式
                Style style = paragraph.get_Style();
                if (style.NameLocal.StartsWith("标题"))
                {
                    object tempstyle = paragraph.get_Style();
                    paragraph.set_Style(ref tempstyle);
                }
            }
        }
        /// <summary>
        /// 刷新选中区域段落样式
        /// </summary>
        public static void RefreshSelectedRange(Range range)
        {
            //Range range = Globals.ThisAddIn.Application.Selection.Range;
            // 遍历文档中的所有段落
            if (range.Text == null) return;
            foreach (Word.Paragraph paragraph in range.Paragraphs)
            {
                // 刷新段落样式
                object style = paragraph.get_Style();
                paragraph.set_Style(ref style);
            }
        }
        /// <summary>
        /// 清除选中区域格式
        /// </summary>
        public static void DelSelectedFormat()
        {
            Range selectionRange = Globals.ThisAddIn.Application.Selection.Range;
            selectionRange.Font.Reset(); // 重置字体
            selectionRange.ParagraphFormat.Reset(); // 重置段落格式
            selectionRange.Shading.BackgroundPatternColorIndex = WdColorIndex.wdAuto;
            selectionRange.HighlightColorIndex = WdColorIndex.wdAuto;
            // 获取正文样式
            Word.Style normalStyle = Globals.ThisAddIn.Application.ActiveDocument.Styles[WdBuiltinStyle.wdStyleNormal];
            // 设置选择区域的段落样式为正文
            selectionRange.set_Style(normalStyle);
        }
        /// <summary>
        /// 移除页脚横线
        /// </summary>
        public static void RmFooterLine()
        {
            var _document = Globals.ThisAddIn.Application.ActiveDocument;

            //页脚横线移除
            foreach (Section section in _document.Sections)
            {
                Word.Range oddFooterRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                foreach (Border border in oddFooterRange.Borders) { border.LineStyle = WdLineStyle.wdLineStyleNone; }
            }
        }
        /// <summary>
        /// 移除页眉横线
        /// </summary>
        public static void RmHeaderLine()
        {
            var _document = Globals.ThisAddIn.Application.ActiveDocument;
            // 设置首页页眉
            foreach (Section section in _document.Sections)
            {
                Word.Range HeaderRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                foreach (Border border in HeaderRange.Borders) { border.LineStyle = WdLineStyle.wdLineStyleNone; }
                // 设置奇数页页眉
                Word.Range oddHeaderRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                foreach (Border border in oddHeaderRange.Borders) { border.LineStyle = WdLineStyle.wdLineStyleNone; }
                // 设置偶数页页眉
                Word.Range evenHeaderRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range;
                foreach (Border border in evenHeaderRange.Borders) { border.LineStyle = WdLineStyle.wdLineStyleNone; }
            }

        }
        /// <summary>
        /// 设置多级列表缩进
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="wdTrailingCharacter"></param>
        public static void SetListTrailingCharacters(Word.Document doc, WdTrailingCharacter wdTrailingCharacter)//设置多级列表的缩进样式
        {
            try
            {
                foreach (ListTemplate listTemplate in doc.ListTemplates)
                {
                    //循环遍历多级列表中的所有级别
                    foreach (Word.ListLevel listLevel in listTemplate.ListLevels) { listLevel.TrailingCharacter = wdTrailingCharacter; }
                }
            }
            catch (Exception) { MessageBox.Show("发生错误，请重试", "提示"); }
        }

        private static Word.Document StyleResave(Word.Document doc, string[] styleNames)//文档保存方法
        {
            //获取文档中的样式集合
            var styles = doc.Styles;
            var app = Globals.ThisAddIn.Application;
            Word.Document styleDoc = app.Documents.Add();
            foreach (string styleName in styleNames)
            {
                Style style = styles[styleName];
                if (style == null) continue;
                Style newStyle;
                try { newStyle = styleDoc.Styles[styleName]; }
                catch (Exception) { newStyle = styleDoc.Styles.Add(styleName); }
                newStyle.AutomaticallyUpdate = false;
                newStyle.Font = style.Font;
                newStyle.NameLocal = style.NameLocal;
                object basestyle = style.get_BaseStyle();
                newStyle.set_BaseStyle(ref basestyle);
                object linkstyle = style.get_LinkStyle();
                newStyle.set_LinkStyle(ref linkstyle);
                newStyle.NoProofing = style.NoProofing;
                newStyle.ParagraphFormat = style.ParagraphFormat;
            }
            return styleDoc;
        }

        public static void ApplyMultiLevelListStyle(Word.Document sourceDoc, Word.Document targetDoc)
        {
            // 获取多级列表样式
            var sourceListTemplate = sourceDoc.ListTemplates.Cast<Word.ListTemplate>().FirstOrDefault(t => t.Name == "多级列表样式");

            // 如果在源文档中不存在多级列表样式，则不做任何操作
            if (sourceListTemplate == null)
            {
                return;
            }

            // 获取目标文档中的多级列表样式
            var targetListTemplate = targetDoc.ListTemplates.Cast<Word.ListTemplate>().FirstOrDefault(t => t.Name == "多级列表样式");

            // 如果在目标文档中不存在多级列表样式，则复制源文档中的多级列表样式到目标文档
            if (targetListTemplate == null)
            {
                targetListTemplate = targetDoc.ListTemplates.Add(sourceListTemplate.OutlineNumbered);
                targetListTemplate.Name = "多级列表样式";
            }

            // 遍历目标文档中的所有标题样式
            foreach (Word.Style style in targetDoc.Styles)
            {
                if (style.Type != Word.WdStyleType.wdStyleTypeParagraph)
                {
                    continue;
                }

                if (style.NameLocal.StartsWith("标题 "))
                {
                    //        // 获取当前标题样式的多级列表样式
                    var sourceListLevel = style.ListLevelNumber;

                    if (sourceListLevel != 0)
                    {
                        // 设置目标文档中当前标题样式的多级列表样式
                        style.LinkToListTemplate(targetListTemplate, sourceListLevel);
                    }
                }
            }
        }

        public static void DeleteEmptyParagraphs(Word.Range range)//删除空行
        {
            int count = 0;
            List<Task> tasks = new List<Task>();
            foreach (Word.Paragraph paragraph in range.Paragraphs)
            {
                if (string.IsNullOrWhiteSpace(paragraph.Range.Text))
                {
                    Task task = Task.Run(() => { paragraph.Range.Delete(); });
                    tasks.Add(task);
                    count++;
                }
                // 每隔1000个段落执行一次垃圾回收
                if (count % 1000 == 0) GC.Collect();
            }
            // 等待所有任务完成
            Task.WhenAll(tasks);
            if (count > 0) MessageBox.Show($"共删除{count}空行", "提示");
            else MessageBox.Show("无空行", "提示");
        }
        /// <summary>
        /// 表格转三线方法
        /// </summary>
        /// <param name="tables"></param>
        public static void TableToThreeLineTableStyle2(Tables tables)
        {
            List<Task> tasks = new List<Task>();
            foreach (Word.Table tbl in tables)
            {
                Task task = Task.Run(() =>
                {
                    // 设置表格的背景色为白色
                    tbl.Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;
                    // 设置表格中所有字体的颜色为黑色
                    tbl.Range.Font.Color = Word.WdColor.wdColorBlack;
                    Word.Borders borders = tbl.Borders;
                    borders.Enable = 0;
                    borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;//设置外部线条为单线
                    borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth100pt;//设置外部线条宽度为 1.5 磅
                    borders.OutsideColor = Word.WdColor.wdColorBlack;
                    tbl.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
                    // 将所有单元格中的内容垂直居中
                    foreach (Word.Cell cell in tbl.Range.Cells)
                    {
                        if (cell.RowIndex == 1)
                        {
                            cell.Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                            cell.Range.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth100pt;
                            cell.Range.Borders.OutsideColor = Word.WdColor.wdColorBlack;
                        }
                        cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        // 设置单元格竖向线条为白色，实现不显示竖向线条的效果
                        cell.Range.Borders[Word.WdBorderType.wdBorderLeft].Color = Word.WdColor.wdColorWhite;
                        cell.Range.Borders[Word.WdBorderType.wdBorderRight].Color = Word.WdColor.wdColorWhite;
                    }
                });
                tasks.Add(task);
            }
            Task.WaitAll(tasks.ToArray());
            if (tasks.Count > 0) MessageBox.Show("三线表转换完成", "提示");
        }

        public static void TableToThreeLineTableStyle1(Tables tables)
        {
            List<Task> tasks = new List<Task>();
            foreach (Word.Table tbl in tables)
            {
                Task task = Task.Run(() =>
                {
                    // 设置表格的背景色为白色
                    tbl.Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;
                    // 设置表格中所有字体的颜色为黑色
                    tbl.Range.Font.Color = Word.WdColor.wdColorBlack;
                    Word.Borders borders = tbl.Borders;
                    borders.Enable = 0;
                    borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;//设置外部线条为单线
                    borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
                    borders.InsideLineWidth = Word.WdLineWidth.wdLineWidth050pt;//设置外部线条宽度为 1.5 磅

                    borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth100pt;//设置外部线条宽度为 1.5 磅
                    borders.OutsideColor = Word.WdColor.wdColorBlack;
                    tbl.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
                    // 将所有单元格中的内容垂直居中
                    foreach (Word.Cell cell in tbl.Range.Cells)
                    {
                        if (cell.RowIndex == 1)
                        {
                            cell.Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                            cell.Range.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth100pt;
                            cell.Range.Borders.OutsideColor = Word.WdColor.wdColorBlack;
                            cell.Range.Borders[Word.WdBorderType.wdBorderLeft].LineWidth = Word.WdLineWidth.wdLineWidth050pt;
                            cell.Range.Borders[Word.WdBorderType.wdBorderRight].LineWidth = Word.WdLineWidth.wdLineWidth050pt;
                        }
                        if (cell.ColumnIndex == 1)
                        {
                            cell.Range.Borders[Word.WdBorderType.wdBorderLeft].Color = Word.WdColor.wdColorWhite;
                        }
                        if (cell.ColumnIndex == tbl.Range.Columns.Count)
                        {
                            cell.Range.Borders[Word.WdBorderType.wdBorderRight].Color = Word.WdColor.wdColorWhite;
                        }
                        //cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        // 设置单元格竖向线条为白色，实现不显示竖向线条的效果
                    }
                });
                tasks.Add(task);
            }
            Task.WaitAll(tasks.ToArray());
            if (tasks.Count > 0) MessageBox.Show("三线表转换完成", "提示");
        }

        public static void SetLineShapes(float width, float height, bool isCenter, Word.ShapeRange shapeRange)//设置图片尺寸
        {
            var _document = Globals.ThisAddIn.Application.ActiveDocument;
            foreach (Word.Shape shape in shapeRange)
            {
                if (isCenter) shape.Left = (float)_document.PageSetup.PageWidth / 2 - (float)shape.Width / 2; //水平居中

                // 设置图片尺寸
                shape.LockAspectRatio = MsoTriState.msoTrue; //开启图片比例锁定
                if (width != 0 && height != 0)
                {
                    shape.LockAspectRatio = MsoTriState.msoFalse;
                    shape.Width = width; shape.Height = height;
                }
                if (width != 0 && height == 0) { shape.Height = width * (shape.Height / shape.Width); shape.Width = width; }
                if (width == 0 && height != 0) { shape.Width = height * (shape.Width / shape.Height); shape.Height = height; }
            }
        }
        public static void SetLineShapes(float width, float height, bool isCenter, Word.Shapes shapeRange)//设置图片尺寸
        {
            var _document = Globals.ThisAddIn.Application.ActiveDocument;
            foreach (Word.Shape shape in shapeRange)
            {
                if (isCenter) shape.Left = (float)_document.PageSetup.PageWidth / 2 - (float)shape.Width / 2; //水平居中

                // 设置图片尺寸
                shape.LockAspectRatio = MsoTriState.msoTrue; //开启图片比例锁定
                if (width != 0 && height != 0)
                {
                    shape.LockAspectRatio = MsoTriState.msoFalse;
                    shape.Width = width; shape.Height = height;
                }
                if (width != 0 && height == 0) { shape.Height = width * (shape.Height / shape.Width); shape.Width = width; }
                if (width == 0 && height != 0) { shape.Width = height * (shape.Width / shape.Height); shape.Height = height; }
            }
        }
        public static void SetInlineShapes(float width, float height, bool isCenter, InlineShapes inlineShapes)//设置图片尺寸
        {
            foreach (InlineShape shape in inlineShapes.Cast<InlineShape>().Where(s => s.Type == WdInlineShapeType.wdInlineShapePicture))
            {
                if (isCenter)
                {
                    shape.Range.Paragraphs[1].set_Style("正文");
                    shape.Range.ParagraphFormat.LeftIndent = 0;
                    shape.Range.ParagraphFormat.FirstLineIndent = 0;
                    shape.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0.0f;
                    shape.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                }

                shape.LockAspectRatio = MsoTriState.msoTrue; //开启图片比例锁定
                if (width != 0 && height != 0)
                {
                    shape.LockAspectRatio = MsoTriState.msoFalse;
                    shape.Width = width; shape.Height = height;
                }
                if (width != 0 && height == 0) { shape.Height = width * (shape.Height / shape.Width); shape.Width = width; }
                if (width == 0 && height != 0) { shape.Width = height * (shape.Width / shape.Height); shape.Height = height; }
            }
        }
        public static ImageCodecInfo GetEncoder(ImageFormat format)// 根据指定的 ImageFormat 获取对应的编码器信息
        {
            var codecs = ImageCodecInfo.GetImageDecoders();
            foreach (var codec in codecs)
            {
                if (codec.FormatID == format.Guid) return codec;
            }
            return null;
        }

        public static void ExportImagePng(string exportFolder, InlineShapes inlineShapes)//导出图片方法
        {
            object imageIndexLock = new object();
            int numThreads = Environment.ProcessorCount; // 使用 CPU 核心数作为线程数
            var tasks = new List<System.Threading.Tasks.Task>();
            ConcurrentBag<string> results = new ConcurrentBag<string>(); // 用于收集导出结果
            // 将图片分成多个子任务，每个子任务处理一部分图片
            var imageChunks = inlineShapes.Cast<InlineShape>()
                .Where(s => s.Type == WdInlineShapeType.wdInlineShapePicture)
                .Select((s, i) => new { Index = i, Shape = s })
                .GroupBy(x => x.Index % numThreads)
                .Select(g => g.Select(x => x.Shape).ToList())
                .ToList();
            int imageIndex = 1;

            // 执行子任务
            foreach (var chunk in imageChunks)
            {
                var task = System.Threading.Tasks.Task.Run(() =>
                {
                    var chunkResults = new List<string>();
                    foreach (InlineShape inlineShape in chunk)
                    {
                        int index;
                        lock (imageIndexLock) { index = imageIndex++; }
                        // 获取图片的字节数据
                        var imageData = inlineShape.Range.EnhMetaFileBits;
                        // 构造保存图片的文件路径
                        var imagePath = Path.Combine(exportFolder, $"image{index}.png");

                        // 从字节数据创建图片对象
                        using (var stream = new MemoryStream(imageData))
                        {
                            // 使用一个独立的 Bitmap 对象来保存图片
                            using (var image = new Bitmap(stream))
                            {
                                // 设置图片质量选项
                                using (var encoderParameters = new EncoderParameters(1))
                                {
                                    encoderParameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 100L);

                                    var pngEncoder = CommonMethods.GetEncoder(ImageFormat.Png);
                                    int i = 1;
                                    while (File.Exists(imagePath))
                                    {
                                        string newName = Path.GetFileNameWithoutExtension(imagePath) + $" ({i})" + Path.GetExtension(imagePath);
                                        imagePath = Path.Combine(Path.GetDirectoryName(imagePath), newName);
                                        i++;
                                    }
                                    image.Save(imagePath, pngEncoder, encoderParameters);
                                }
                            }
                        }
                        chunkResults.Add(imagePath);
                        imageIndex++;
                    }
                });
                tasks.Add(task);
            }
            // 等待所有子任务完成
            System.Threading.Tasks.Task.WaitAll(tasks.ToArray());
            if (tasks.Count > 0) MessageBox.Show("图片导出完成", "提示"); ;
        }


    }
}
