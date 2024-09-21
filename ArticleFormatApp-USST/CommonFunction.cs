using ArticleFormatApp_USST.Models;
using ArticleFormatApp_USST.Properties;
using CsvHelper;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Xml.Serialization;


namespace ArticleFormatApp_USST
{
    public class CommonFunction
    {
        /// <summary>
        /// 设置样式
        /// </summary>
        /// <param name="title"></param>
        public static bool SetStyle(string title)
        {
            Range range = Globals.ThisAddIn.Application.Selection.Range;
            try
            {
                range.set_Style(title);
                return true;
            }
            catch (Exception)
            {
                MessageBox.Show($"样式 {title} 未创建，请单击【生成论文样式】创建样式！");
                return false;
            }
        }

        /// <summary>
        /// 设置样式
        /// </summary>
        /// <param name="title"></param>
        public static bool SetStyle(WdBuiltinStyle title)
        {
            Range range = Globals.ThisAddIn.Application.Selection.Range;
            try
            {
                range.set_Style(title); return true;
            }
            catch (Exception)
            {
                MessageBox.Show($"样式 {title} 未创建，请单击【生成论文样式】创建样式！");
                return false;
            }
        }



        /// <summary>
        /// 创建所有的目标样式
        /// </summary>
        /// <param name="doc"></param>
        public static void CreateStyles(ref Document doc)
        {
            List<string> oldstyleNames = doc.Styles.Cast<Style>().Select(t => t.NameLocal).ToList();

            CreateCaptionLabels(ref doc, ref oldstyleNames);
            CreateThreelinesTableStyle(ref doc, ref oldstyleNames);
            CreateListTemplates(ref doc);

            //获取资源文件
            Type type = MethodBase.GetCurrentMethod().DeclaringType;
            string _namespace = type.Namespace; //获取命名空间
            string basePath = _namespace + Resources.StylesPath;
            List<StyleModel> styleModels = new List<StyleModel>();
            Assembly _assembly = Assembly.GetExecutingAssembly();
            Stream stream = _assembly.GetManifestResourceStream(basePath);
            using (StreamReader reader1 = new StreamReader(stream))
            {
                using (var csv = new CsvReader(reader1, CultureInfo.InvariantCulture))
                { styleModels = csv.GetRecords<StyleModel>().ToList(); }
            }

            List<string> oldlistTemplateNames = doc.ListTemplates.Cast<ListTemplate>().Select(t => t.Name).ToList();
            var left = doc.PageSetup.LeftMargin;
            var right = doc.PageSetup.RightMargin;
            var gutter = doc.PageSetup.Gutter;
            var width = doc.PageSetup.PageWidth;
            var center = (width - gutter - left - right) / 2;

            foreach (StyleModel styleModel in styleModels)
            {
                Style newStyle;
                if (oldstyleNames.Contains(styleModel.NameLocal)) { newStyle = doc.Styles[styleModel.NameLocal]; }
                else { newStyle = doc.Styles.Add(styleModel.NameLocal); }
                //看是否需要链接到lt且lt样式存在
                if (styleModel.LinkToListTemplate != string.Empty && oldlistTemplateNames.Contains(styleModel.LinkToListTemplate))
                {
                    //样式存在
                    var lt = doc.ListTemplates[styleModel.LinkToListTemplate];
                    if (newStyle.ListTemplate == null)
                    {
                        newStyle.LinkToListTemplate(lt);
                    }
                    else
                    {
                        if (!newStyle.ListTemplate.Name.Contains(styleModel.LinkToListTemplate))
                        {
                            newStyle.LinkToListTemplate(lt);
                        }
                    }
                }
                else { newStyle.LinkToListTemplate(null); }

                CreateStyle(styleModel, ref newStyle, ref doc);
            }

        }

        /// <summary>
        /// 创建公式样式
        /// </summary>
        /// <param name="doc"></param>
        public static void CreateFormulaStyle(ref Document doc)
        {
            List<string> oldstyleNames = doc.Styles.Cast<Style>().Select(t => t.NameLocal).ToList();

            CreateCaptionLabels(ref doc, ref oldstyleNames);

            //获取资源文件
            Type type = MethodBase.GetCurrentMethod().DeclaringType;
            string _namespace = type.Namespace; //获取命名空间
            string basePath = _namespace + Resources.StylesPath;
            List<StyleModel> styleModels;
            Assembly _assembly = Assembly.GetExecutingAssembly();
            Stream stream = _assembly.GetManifestResourceStream(basePath);
            using (StreamReader reader1 = new StreamReader(stream))
            {
                using (var csv = new CsvReader(reader1, CultureInfo.InvariantCulture))
                {
                    styleModels = csv.GetRecords<StyleModel>()
                        .Where(t => t.NameLocal.Contains(Resources.FormulaLabelName)).ToList();
                }
            }

            foreach (StyleModel styleModel in styleModels)
            {
                Style newStyle = doc.Styles.OfType<Style>().Where(t => t.NameLocal == Resources.FormulaLabelName).FirstOrDefault();
                if (newStyle == null) newStyle = doc.Styles.Add(Resources.FormulaLabelName);
                CreateStyle(styleModel, ref newStyle, ref doc);
            }

        }

        /// <summary>
        /// 创建图表样式
        /// </summary>
        /// <param name="doc"></param>
        public static void CreatePicTableStyle(ref Document doc)
        {
            List<string> oldstyleNames = doc.Styles.Cast<Style>().Select(t => t.NameLocal).ToList();

            CreateCaptionLabels(ref doc, ref oldstyleNames);
            CreateThreelinesTableStyle(ref doc, ref oldstyleNames);
            //获取资源文件
            Type type = MethodBase.GetCurrentMethod().DeclaringType;
            string _namespace = type.Namespace; //获取命名空间
            string basePath = _namespace + Resources.StylesPath;
            List<StyleModel> styleModels;
            Assembly _assembly = Assembly.GetExecutingAssembly();
            Stream stream = _assembly.GetManifestResourceStream(basePath);
            using (StreamReader reader1 = new StreamReader(stream))
            {
                using (var csv = new CsvReader(reader1, CultureInfo.InvariantCulture))
                {
                    styleModels = csv.GetRecords<StyleModel>()
                        .Where(t => t.NameLocal.Contains(Resources.PicLabelName)).ToList();

                    styleModels.AddRange(csv.GetRecords<StyleModel>()
                        .Where(t => t.NameLocal.Contains(Resources.TableLabelName)).ToList());
                }
            }

            foreach (StyleModel styleModel in styleModels)
            {
                Style newStyle = doc.Styles.OfType<Style>().Where(t => t.NameLocal == Resources.FormulaLabelName).FirstOrDefault();
                if (newStyle == null) newStyle = doc.Styles[Resources.FormulaLabelName];
                CreateStyle(styleModel, ref newStyle, ref doc);
            }

        }

        /// <summary>
        /// 插入题注
        /// </summary>
        public static void InsertCaption()
        {
            InlineShapes inlineShapes = Globals.ThisAddIn.Application.Selection.InlineShapes;
            foreach (InlineShape inlineShape in inlineShapes.OfType<InlineShape>().Where(s => s.Type == WdInlineShapeType.wdInlineShapePicture))
            {
                Range rng = inlineShape.Range;
                rng.set_Style("图片内容");
                Paragraph para = rng.Paragraphs.First.Next();
                if (para == null) inlineShape.Range.InsertCaption("图");
                else
                {
                    if (rng.Paragraphs.First.Next().Range.Fields.Count == 0)
                    {// 插入新题注
                        string text = para.Range.Text.Trim();

                        inlineShape.Range.InsertCaption("图");
                        // 移除题注和序号之间的空格
                    }
                    else
                    {// 刷新题注
                        para.Range.Fields.Update();
                    }
                }
                para = rng.Paragraphs.First.Next();
                para.set_Style("图题注");
            }
            Tables tables = Globals.ThisAddIn.Application.Selection.Tables;
            foreach (Table table in tables)
            {
                Range rng = table.Range;
                rng.set_Style("表内容");
                Paragraph para = rng.Paragraphs.First.Previous();
                if (para == null) table.Range.InsertCaption("表");
                else
                {
                    if (rng.Paragraphs.First.Previous().Range.Fields.Count == 0)
                    {
                        table.Range.InsertCaption("表");
                    }
                    else
                    {// 刷新题注
                        rng.Paragraphs.First.Previous().Range.Fields.Update();
                    }
                    para.set_Style("表题注");
                }
            }
        }



        /// <summary>
        /// 创建style样式
        /// </summary>
        /// <param name="styleModel"></param>
        /// <param name="newStyle"></param>
        /// <param name="doc"></param>
        private static void CreateStyle(StyleModel styleModel, ref Style newStyle, ref Document doc)
        {
            newStyle.set_BaseStyle(styleModel.BaseStyle);
            if (styleModel.NextParagraphStyle != string.Empty)
                newStyle.set_NextParagraphStyle(styleModel.NextParagraphStyle);
            newStyle.Font.Name = styleModel.FontName;
            newStyle.Font.NameFarEast = styleModel.FontNameFarEast;
            newStyle.Font.Size = styleModel.FontSize;
            newStyle.Font.Bold = styleModel.FontBold;
            newStyle.Font.Color = styleModel.FontColor;

            newStyle.ParagraphFormat.Alignment = styleModel.ParaAlignment;
            newStyle.ParagraphFormat.CharacterUnitFirstLineIndent = styleModel.ParaCharacterUnitFirstLineIndent;
            newStyle.ParagraphFormat.CharacterUnitRightIndent = styleModel.ParaCharacterUnitRightIndent;
            //下面两个同时设置才有用
            newStyle.ParagraphFormat.CharacterUnitLeftIndent = styleModel.ParaCharacterUnitLeftIndent;
            newStyle.ParagraphFormat.FirstLineIndent = styleModel.ParaFirstLineIndent;
            //newStyle.ParagraphFormat.LeftIndent = styleModel.ParaLeftIndent;
            newStyle.ParagraphFormat.SpaceBefore = styleModel.ParaSpaceBefore;
            newStyle.ParagraphFormat.SpaceAfter = styleModel.ParaSpaceAfter;
            newStyle.ParagraphFormat.LineSpacingRule = styleModel.ParaLineSpacingRule;

            if (styleModel.ParaLineSpacingRule != WdLineSpacing.wdLineSpaceSingle)
                newStyle.ParagraphFormat.LineSpacing = styleModel.ParaLineSpacing;
            newStyle.ParagraphFormat.OutlineLevel = styleModel.OutlineLevel;

            if (styleModel.NameLocal == Resources.FormulaLabelName)
            {
                var left = doc.PageSetup.LeftMargin;
                var right = doc.PageSetup.RightMargin;
                var gutter = doc.PageSetup.Gutter;
                var width = doc.PageSetup.PageWidth;
                var center = (width - gutter - left - right) / 2;
                newStyle.ParagraphFormat.TabStops.ClearAll();
                newStyle.ParagraphFormat.TabStops.Add(
                    Position: center,   // 制表位位置，以磅为单位
                    Alignment: WdTabAlignment.wdAlignTabCenter,      // 制表位对齐方式
                    Leader: WdTabLeader.wdTabLeaderSpaces             // 制表位前导符
                );
                newStyle.ParagraphFormat.TabStops.Add(
                    Position: center * 2,   // doc.Application.CentimetersToPoints(13f)制表位位置，以磅为单位
                    Alignment: WdTabAlignment.wdAlignTabRight,      // 制表位对齐方式
                    Leader: WdTabLeader.wdTabLeaderSpaces             // 制表位前导符
                );
            }
            //目录的制表符
            if (styleModel.NameLocal.Contains("TOC "))
            {
                var left = doc.PageSetup.LeftMargin;
                var right = doc.PageSetup.RightMargin;
                var gutter = doc.PageSetup.Gutter;
                var width = doc.PageSetup.PageWidth;
                var center = (width - gutter - left - right) / 2;

                newStyle.ParagraphFormat.TabStops.ClearAll();
                newStyle.ParagraphFormat.TabStops.Add(
                    Position: center * 2,   // 制表位位置，以磅为单位
                    Alignment: WdTabAlignment.wdAlignTabRight,      // 制表位对齐方式
                    Leader: WdTabLeader.wdTabLeaderDots             // 制表位前导符
                );
            }
        }


        /// <summary>
        /// 创建题注标签
        /// </summary>
        /// <param name="doc"></param>
        private static void CreateCaptionLabels(ref Document doc, ref List<string> styleNamelist)
        {
            CaptionLabels captionLabels = doc.InlineShapes.Application.CaptionLabels;//获取标签

            CaptionLabel picLabel = captionLabels.Add(Resources.PicLabelName); //可直接添加，重复不会影响
            CaptionLabel tableLabel = captionLabels.Add(Resources.TableLabelName);
            CaptionLabel formulaLabel = captionLabels.Add(Resources.FormulaLabelName);//公式
            //分隔符
            picLabel.Separator = (WdSeparatorType)Enum.Parse(typeof(WdSeparatorType), Resources.PicSeparator);// WdSeparatorType.wdSeparatorPeriod;//wdSeparatorPeriod
            tableLabel.Separator = (WdSeparatorType)Enum.Parse(typeof(WdSeparatorType), Resources.TableSeparator);//WdSeparatorType.wdSeparatorPeriod;
            formulaLabel.Separator = (WdSeparatorType)Enum.Parse(typeof(WdSeparatorType), Resources.FormulaSeparator);//WdSeparatorType.wdSeparatorPeriod;
            //包含章节号
            picLabel.IncludeChapterNumber = bool.Parse(Resources.IncludeChapterNumber);
            tableLabel.IncludeChapterNumber = bool.Parse(Resources.IncludeChapterNumber);
            formulaLabel.IncludeChapterNumber = bool.Parse(Resources.FormulaIncludeChapterNumber);

            picLabel.ChapterStyleLevel = 1;
            tableLabel.ChapterStyleLevel = 1;
            formulaLabel.ChapterStyleLevel = 1;
            //位置
            picLabel.Position = (WdCaptionPosition)Enum.Parse(typeof(WdCaptionPosition), Resources.PicLabelPostion);//WdCaptionPosition.wdCaptionPositionBelow;
            tableLabel.Position = (WdCaptionPosition)Enum.Parse(typeof(WdCaptionPosition), Resources.TableLabelPostion);// WdCaptionPosition.wdCaptionPositionAbove;
            formulaLabel.Position = (WdCaptionPosition)Enum.Parse(typeof(WdCaptionPosition), Resources.FormulaLabelPostion);//WdCaptionPosition.wdCaptionPositionBelow;
        }


        /// <summary>
        /// 创建三线表样式
        /// </summary>
        /// <param name="doc"></param>
        private static void CreateThreelinesTableStyle(ref Document doc, ref List<string> styleNamelist)
        {
            //三线表样式
            string threeLineTableHeadlineTitle = "三线表";
            Style threeLineTableHeadlineStyle;
            if (!styleNamelist.Contains(threeLineTableHeadlineTitle))
            {
                threeLineTableHeadlineStyle = doc.Styles.Add(threeLineTableHeadlineTitle, WdStyleType.wdStyleTypeTable);
            }
            else { threeLineTableHeadlineStyle = doc.Styles[threeLineTableHeadlineTitle]; }

            // 设置边框样式
            threeLineTableHeadlineStyle.Table.Borders.Enable = 0;
            threeLineTableHeadlineStyle.Table.Borders[WdBorderType.wdBorderVertical].Visible = false;//纵向不显示
            threeLineTableHeadlineStyle.Table.Borders[WdBorderType.wdBorderHorizontal].Visible = false;//横向不显示
            threeLineTableHeadlineStyle.Table.Borders[WdBorderType.wdBorderLeft].Visible = false;//左侧不显示
            threeLineTableHeadlineStyle.Table.Borders[WdBorderType.wdBorderRight].Visible = false;//右侧不显示

            threeLineTableHeadlineStyle.Table.Borders[WdBorderType.wdBorderTop].Visible = true;//顶端显示
            threeLineTableHeadlineStyle.Table.Borders[WdBorderType.wdBorderTop].LineWidth = WdLineWidth.wdLineWidth150pt;//宽度为 1.5 磅;
            threeLineTableHeadlineStyle.Table.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;//单线;

            threeLineTableHeadlineStyle.Table.Borders[WdBorderType.wdBorderBottom].Visible = true;//底端显示
            threeLineTableHeadlineStyle.Table.Borders[WdBorderType.wdBorderBottom].LineWidth = WdLineWidth.wdLineWidth150pt;//宽度为 150 磅;
            threeLineTableHeadlineStyle.Table.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;//单线;
            threeLineTableHeadlineStyle.Table.TableDirection = WdTableDirection.wdTableDirectionLtr;
        }


        /// <summary>
        /// 创建列表样式
        /// </summary>
        /// <param name="doc"></param>
        private static void CreateListTemplates(ref Document doc)
        {
            OtherConfigModel otherConfig;

            //获取配置文件
            Type type = MethodBase.GetCurrentMethod().DeclaringType;
            string _namespace = type.Namespace; //获取命名空间
            string basePath = _namespace + Resources.OtherConfigPath;

            Assembly _assembly = Assembly.GetExecutingAssembly();
            Stream stream = _assembly.GetManifestResourceStream(basePath);
            using (StreamReader reader = new StreamReader(stream)) //str +
            {
                XmlSerializer xs = new XmlSerializer(typeof(OtherConfigModel));
                otherConfig = (OtherConfigModel)xs.Deserialize(reader);
            }

            HeadingListTemplateModel headingLtConf = otherConfig.HeadingListTemplate;
            ReferencesListTemplateModel referenceLtConf = otherConfig.ReferencesListTemplate;

            List<string> oldlistTemplateNames = doc.ListTemplates.OfType<ListTemplate>().Select(t => t.Name).ToList();
            ListTemplate newHeadingLt = null;
            //ListTemplate newHeadingLt = Globals.ThisAddIn.Application.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[5];
            if (oldlistTemplateNames.Contains(headingLtConf.Name))
                newHeadingLt = doc.ListTemplates[headingLtConf.Name];
            else
            {
                newHeadingLt = doc.ListTemplates.Add(OutlineNumbered: true);
                newHeadingLt.Name = headingLtConf.Name;

            }
            for (int i = 1; i < 10; i++)
            {
                newHeadingLt.ListLevels[i].LinkedStyle = $"标题 {i}";//普通数字
                newHeadingLt.ListLevels[i].NumberStyle = headingLtConf.NumberStyle;//普通数字
                newHeadingLt.ListLevels[i].TextPosition = headingLtConf.TextPosition;
                newHeadingLt.ListLevels[i].NumberPosition = headingLtConf.NumberPosition;
                newHeadingLt.ListLevels[i].TrailingCharacter = headingLtConf.TrailingCharacter;
                //不变的参数
                newHeadingLt.ListLevels[i].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
                newHeadingLt.ListLevels[i].StartAt = 1;
                newHeadingLt.ListLevels[i].ResetOnHigher = i - 1;
            }
            newHeadingLt.ListLevels[1].NumberFormat = headingLtConf.NumberFormat1;
            newHeadingLt.ListLevels[2].NumberFormat = headingLtConf.NumberFormat2;
            newHeadingLt.ListLevels[3].NumberFormat = headingLtConf.NumberFormat3;
            newHeadingLt.ListLevels[4].NumberFormat = headingLtConf.NumberFormat4;
            newHeadingLt.ListLevels[5].NumberFormat = headingLtConf.NumberFormat5;
            newHeadingLt.ListLevels[6].NumberFormat = headingLtConf.NumberFormat6;
            newHeadingLt.ListLevels[7].NumberFormat = headingLtConf.NumberFormat7;
            newHeadingLt.ListLevels[8].NumberFormat = headingLtConf.NumberFormat8;
            newHeadingLt.ListLevels[9].NumberFormat = headingLtConf.NumberFormat9;

            ////绑定一级标题
            //ListTemplate heading1Lt = null;
            ////ListTemplate newHeadingLt = Globals.ThisAddIn.Application.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[5];
            //if (oldlistTemplateNames.Contains("一级标题大写列表"))
            //    heading1Lt = doc.ListTemplates[headingLtConf.Name];
            //else
            //{
            //    heading1Lt = doc.ListTemplates.Add(OutlineNumbered: true);
            //    heading1Lt.Name = "一级标题大写列表";
            //}
            //heading1Lt.ListLevels[1].LinkedStyle = $"标题 1";//普通数字
            //heading1Lt.ListLevels[1].NumberStyle = WdListNumberStyle.wdListNumberStyleSimpChinNum1;//普通数字
            //heading1Lt.ListLevels[1].TextPosition = headingLtConf.TextPosition;
            //heading1Lt.ListLevels[1].NumberPosition = headingLtConf.NumberPosition;
            //heading1Lt.ListLevels[1].TrailingCharacter = headingLtConf.TrailingCharacter;
            ////不变的参数
            //heading1Lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            //heading1Lt.ListLevels[1].StartAt = 1;
            //heading1Lt.ListLevels[1].ResetOnHigher = 0;

            //heading1Lt.ListLevels[1].NumberFormat = headingLtConf.NumberFormat1;

            //参考文献列表
            ListTemplate newReferenceLt = null;
            if (oldlistTemplateNames.Contains(referenceLtConf.Name))
                newReferenceLt = doc.ListTemplates[referenceLtConf.Name];
            else
            {
                newReferenceLt = doc.ListTemplates.Add(OutlineNumbered: false);
                newReferenceLt.Name = referenceLtConf.Name;
            }
            newReferenceLt.ListLevels[1].NumberFormat = referenceLtConf.NumberFormat;// 设置自动序号格式
            newReferenceLt.ListLevels[1].TrailingCharacter = referenceLtConf.TrailingCharacter;
            //不变参数
            newReferenceLt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft; // 设置自动序号对齐方式（可选）
            newReferenceLt.ListLevels[1].NumberPosition = 0f;
            newReferenceLt.ListLevels[1].StartAt = 1;
            newReferenceLt.ListLevels[1].ResetOnHigher = 0;
        }

    }

}
