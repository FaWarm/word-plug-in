using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;


namespace 沈体论文格式调整word插件
{
    public class CommonFunction
    {
        /// <summary>
        /// 设置样式
        /// </summary>
        /// <param name="title"></param>
        public static void SetStyle(string title)
        {
            Range range = Globals.ThisAddIn.Application.Selection.Range;
            try
            {
                range.set_Style(title);
            }
            catch (Exception)
            {
                MessageBox.Show($"样式 {title} 未创建，请单击【生成论文样式】创建样式！");
            }
        }

        /// <summary>
        /// 设置样式
        /// </summary>
        /// <param name="title"></param>
        public static void SetStyle(WdBuiltinStyle title)
        {
            Range range = Globals.ThisAddIn.Application.Selection.Range;
            try
            {
                range.set_Style(title);
            }
            catch (Exception)
            {
                MessageBox.Show($"样式 {title} 未创建，请单击【生成论文样式】创建样式！");
            }
        }

        /// <summary>
        /// 应用多级列表样式
        /// </summary>
        /// <param name="sourceDoc"></param>
        /// <param name="targetDoc"></param>
        public static void ApplyMultiLevelListStyle(Document doc)
        {
            Style heading1Style = doc.Styles[WdBuiltinStyle.wdStyleHeading1];
            Style heading2Style = doc.Styles[WdBuiltinStyle.wdStyleHeading2];
            Style heading3Style = doc.Styles[WdBuiltinStyle.wdStyleHeading3];
            Style heading4Style = doc.Styles[WdBuiltinStyle.wdStyleHeading4];
            Style heading5Style = doc.Styles[WdBuiltinStyle.wdStyleHeading5];
            Style heading6Style = doc.Styles[WdBuiltinStyle.wdStyleHeading6];
            Style heading7Style = doc.Styles[WdBuiltinStyle.wdStyleHeading7];
            Style heading8Style = doc.Styles[WdBuiltinStyle.wdStyleHeading8];
            Style heading9Style = doc.Styles[WdBuiltinStyle.wdStyleHeading9];
            ListTemplate listTemplate = Globals.ThisAddIn.Application.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[5];

            for (int i = 1; i < 10; i++)
            {
                listTemplate.ListLevels[i].Font.Name = "黑体";
                listTemplate.ListLevels[i].NumberStyle = WdListNumberStyle.wdListNumberStyleArabic;//普通数字样式
                listTemplate.ListLevels[i].TextPosition = 0f;
                listTemplate.ListLevels[i].TrailingCharacter = WdTrailingCharacter.wdTrailingSpace;
                listTemplate.ListLevels[i].NumberPosition = 0f;
                listTemplate.ListLevels[i].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
                listTemplate.ListLevels[i].StartAt = 1;
                listTemplate.ListLevels[i].ResetOnHigher = i - 1;
            }

            listTemplate.ListLevels[1].NumberFormat = "%1";
            listTemplate.ListLevels[2].NumberFormat = "%1.%2";
            listTemplate.ListLevels[3].NumberFormat = "%1.%2.%3";
            listTemplate.ListLevels[4].NumberFormat = "%1.%2.%3.%4";
            listTemplate.ListLevels[5].NumberFormat = "%1.%2.%3.%4.%5";
            listTemplate.ListLevels[6].NumberFormat = "%1.%2.%3.%4.%5.%6";
            listTemplate.ListLevels[7].NumberFormat = "%1.%2.%3.%4.%5.%6.%7";
            listTemplate.ListLevels[8].NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8";
            listTemplate.ListLevels[9].NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8.%9";

            heading1Style.LinkToListTemplate(listTemplate); // 将标题1的样式链接到数字序号样式
            heading2Style.LinkToListTemplate(listTemplate); // 将标题2的样式链接到数字序号样式
            heading3Style.LinkToListTemplate(listTemplate); // 将标题3的样式链接到数字序号样式
            heading4Style.LinkToListTemplate(listTemplate); // 将标题4的样式链接到数字序号样式
            heading5Style.LinkToListTemplate(listTemplate); // 将标题5的样式链接到数字序号样式
            heading6Style.LinkToListTemplate(listTemplate); // 将标题6的样式链接到数字序号样式
            heading7Style.LinkToListTemplate(listTemplate); // 将标题7的样式链接到数字序号样式
            heading8Style.LinkToListTemplate(listTemplate); // 将标题8的样式链接到数字序号样式
            heading9Style.LinkToListTemplate(listTemplate); // 将标题9的样式链接到数字序号样式

            ListTemplate customListTemplate = null;
            foreach (ListTemplate lt in doc.ListTemplates)
            {
                if (lt.Name == "参考文献编号")
                {
                    customListTemplate = lt;
                    break;
                }
            }
            if (customListTemplate == null)
            {
                customListTemplate = doc.ListTemplates.Add(OutlineNumbered: true);
                customListTemplate.Name = "参考文献编号";
            }
            customListTemplate.ListLevels[1].NumberFormat = "[%1]";// 设置自动序号格式
            customListTemplate.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft; // 设置自动序号对齐方式（可选）
            customListTemplate.ListLevels[1].TrailingCharacter = WdTrailingCharacter.wdTrailingSpace;
            customListTemplate.ListLevels[1].NumberPosition = 0f;
            customListTemplate.ListLevels[1].StartAt = 1;
            customListTemplate.ListLevels[1].ResetOnHigher = 0;
        }

        /// <summary>
        /// 样式初始化
        /// </summary>
        /// <param name="doc"></param>
        public static void InitStyles(Document doc)
        {
            List<string> styleNamelist = new List<string>();
            foreach (Style style in doc.Styles)
            { styleNamelist.Add(style.NameLocal); }

            #region 题注初始化
            CaptionLabels captionLabels = doc.InlineShapes.Application.CaptionLabels;//获取标签
            CaptionLabel picLabel = null;
            CaptionLabel tableLabel = null;
            foreach (CaptionLabel label in captionLabels)
            {
                if (label.Name.Equals("图")) picLabel = captionLabels["图"];
                if (label.Name.Equals("表")) picLabel = captionLabels["表"];
            }
            picLabel = captionLabels.Add("图");
            picLabel.Separator = WdSeparatorType.wdSeparatorHyphen;
            picLabel.IncludeChapterNumber = false;//包含章节号
            picLabel.ChapterStyleLevel = 1;
            picLabel.Position = WdCaptionPosition.wdCaptionPositionBelow;
            tableLabel = captionLabels.Add("表");
            tableLabel.Separator = WdSeparatorType.wdSeparatorHyphen;
            tableLabel.IncludeChapterNumber = false;//包含章节号
            tableLabel.ChapterStyleLevel = 1;
            tableLabel.Position = WdCaptionPosition.wdCaptionPositionAbove;
            #endregion

            #region 样式初始化
            //段落正文
            string title = "段落正文";
            Style customStyle;
            if (!styleNamelist.Contains(title))
            {
                customStyle = doc.Styles.Add(title);
            }
            else { customStyle = doc.Styles[title]; }
            customStyle.set_BaseStyle(WdBuiltinStyle.wdStyleNormal); // 设置基础样式为普通文本
            customStyle.set_NextParagraphStyle(title);
            customStyle.Font.NameFarEast = "宋体";
            customStyle.Font.Name = "Times New Roman";
            customStyle.Font.Size = 12;
            customStyle.Font.Bold = 0;
            customStyle.Font.Color = WdColor.wdColorBlack;
            customStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
            customStyle.ParagraphFormat.FirstLineIndent = 2; // 2个汉字的左缩进
            customStyle.ParagraphFormat.LeftIndent = 0;
            customStyle.ParagraphFormat.SpaceBefore = 0;
            customStyle.ParagraphFormat.SpaceAfter = 0;
            customStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            customStyle.ParagraphFormat.LineSpacing = 22;

            //摘要正文
            string abstractBodyTitle = "摘要正文";
            Style abstractBodyStyle;
            if (!styleNamelist.Contains(abstractBodyTitle))
            {
                abstractBodyStyle = doc.Styles.Add(abstractBodyTitle);
            }
            else { abstractBodyStyle = doc.Styles[abstractBodyTitle]; }
            abstractBodyStyle.set_BaseStyle(WdBuiltinStyle.wdStyleNormal); // 设置基础样式为普通文本
            abstractBodyStyle.Font.NameFarEast = "宋体";
            abstractBodyStyle.Font.Size = 12;
            abstractBodyStyle.Font.Bold = 0;
            abstractBodyStyle.Font.Name = "Times New Roman";
            abstractBodyStyle.Font.Color = WdColor.wdColorBlack;
            abstractBodyStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
            abstractBodyStyle.ParagraphFormat.FirstLineIndent = 2; // 2个汉字的左缩进
            abstractBodyStyle.ParagraphFormat.LeftIndent = 0;
            abstractBodyStyle.ParagraphFormat.SpaceBefore = 0;
            abstractBodyStyle.ParagraphFormat.SpaceAfter = 0;
            abstractBodyStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            abstractBodyStyle.ParagraphFormat.LineSpacing = 22;

            //参考文献正文
            string referenceBodyTitle = "参考文献列表";
            Style referenceBodyStyle;
            if (!styleNamelist.Contains(referenceBodyTitle))
            {
                referenceBodyStyle = doc.Styles.Add(referenceBodyTitle);
                referenceBodyStyle.LinkToListTemplate(doc.ListTemplates["参考文献编号"]);
                referenceBodyStyle.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText;
            }
            else { referenceBodyStyle = doc.Styles[referenceBodyTitle]; }
            referenceBodyStyle.set_BaseStyle(WdBuiltinStyle.wdStyleNormal); // 设置基础样式为普通文本
            referenceBodyStyle.Font.NameFarEast = "宋体";
            referenceBodyStyle.Font.Size = 10.5f;
            referenceBodyStyle.Font.Bold = 0;
            referenceBodyStyle.Font.Name = "Times New Roman";
            referenceBodyStyle.Font.Color = WdColor.wdColorBlack;
            referenceBodyStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
            referenceBodyStyle.ParagraphFormat.FirstLineIndent = 0f; // 0个汉字的左缩进
            referenceBodyStyle.ParagraphFormat.CharacterUnitFirstLineIndent = 0f; // 0个汉字的左缩进
            referenceBodyStyle.ParagraphFormat.LeftIndent = 0;
            referenceBodyStyle.ParagraphFormat.SpaceBefore = 0;
            referenceBodyStyle.ParagraphFormat.SpaceAfter = 0;
            referenceBodyStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            referenceBodyStyle.ParagraphFormat.LineSpacing = 22;

            //致谢正文
            string acknowledgementBodyTitle = "致谢正文";
            Style acknowledgementBodyStyle;
            if (!styleNamelist.Contains(acknowledgementBodyTitle))
            {
                acknowledgementBodyStyle = doc.Styles.Add(acknowledgementBodyTitle);
            }
            else { acknowledgementBodyStyle = doc.Styles[acknowledgementBodyTitle]; }
            acknowledgementBodyStyle.set_BaseStyle(WdBuiltinStyle.wdStyleNormal); // 设置基础样式为普通文本
            acknowledgementBodyStyle.Font.NameFarEast = "宋体";
            acknowledgementBodyStyle.Font.Size = 12;
            acknowledgementBodyStyle.Font.Bold = 0;
            acknowledgementBodyStyle.Font.Name = "Times New Roman";
            acknowledgementBodyStyle.Font.Color = WdColor.wdColorBlack;
            acknowledgementBodyStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
            acknowledgementBodyStyle.ParagraphFormat.FirstLineIndent = 2; // 2个汉字的左缩进
            acknowledgementBodyStyle.ParagraphFormat.SpaceBefore = 0;
            acknowledgementBodyStyle.ParagraphFormat.SpaceAfter = 0;
            acknowledgementBodyStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            acknowledgementBodyStyle.ParagraphFormat.LineSpacing = 22;

            //学术成果正文
            string achievementsBodyTitle = "学术成果正文";
            Style achievementsBodyStyle;
            if (!styleNamelist.Contains(achievementsBodyTitle))
            {
                achievementsBodyStyle = doc.Styles.Add(achievementsBodyTitle);
                achievementsBodyStyle.set_BaseStyle(WdBuiltinStyle.wdStyleNormal); // 设置基础样式为普通文本
            }
            else { achievementsBodyStyle = doc.Styles[achievementsBodyTitle]; }
            achievementsBodyStyle.Font.NameFarEast = "宋体";
            achievementsBodyStyle.Font.Size = 12;
            achievementsBodyStyle.Font.Bold = 0;
            achievementsBodyStyle.Font.Name = "Times New Roman";
            achievementsBodyStyle.Font.Color = WdColor.wdColorBlack;
            achievementsBodyStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
            achievementsBodyStyle.ParagraphFormat.FirstLineIndent = 2; // 2个汉字的左缩进
            achievementsBodyStyle.ParagraphFormat.SpaceBefore = 0;
            achievementsBodyStyle.ParagraphFormat.SpaceAfter = 0;
            achievementsBodyStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            achievementsBodyStyle.ParagraphFormat.LineSpacing = 22;

            //附录正文
            string appendixBodyTitle = "附录正文";
            Style appendixBodyStyle;
            if (!styleNamelist.Contains(appendixBodyTitle))
            {
                appendixBodyStyle = doc.Styles.Add(appendixBodyTitle);
            }
            else { appendixBodyStyle = doc.Styles[appendixBodyTitle]; }
            appendixBodyStyle.set_BaseStyle(WdBuiltinStyle.wdStyleNormal); // 设置基础样式为普通文本
            appendixBodyStyle.Font.NameFarEast = "宋体";
            appendixBodyStyle.Font.Size = 10.5f;
            appendixBodyStyle.Font.Bold = 0;
            appendixBodyStyle.Font.Name = "Times New Roman";
            appendixBodyStyle.Font.Color = WdColor.wdColorBlack;
            appendixBodyStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
            appendixBodyStyle.ParagraphFormat.FirstLineIndent = 2; // 2个汉字的左缩进
            appendixBodyStyle.ParagraphFormat.SpaceBefore = 0;
            appendixBodyStyle.ParagraphFormat.SpaceAfter = 0;
            appendixBodyStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            appendixBodyStyle.ParagraphFormat.LineSpacing = 22;

            //标题1
            Style firstHeadingStyle = doc.Styles[WdBuiltinStyle.wdStyleHeading1];
            firstHeadingStyle.set_BaseStyle(WdBuiltinStyle.wdStyleNormal); // 设置基础样式为普通文本
            firstHeadingStyle.set_NextParagraphStyle("段落正文");
            firstHeadingStyle.Font.NameFarEast = "黑体";
            firstHeadingStyle.Font.Size = 22;
            firstHeadingStyle.Font.Bold = 0;
            firstHeadingStyle.Font.Color = WdColor.wdColorBlack;
            firstHeadingStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            firstHeadingStyle.ParagraphFormat.FirstLineIndent = 0; // 2个汉字的左缩进
            firstHeadingStyle.ParagraphFormat.SpaceBefore = 24;
            firstHeadingStyle.ParagraphFormat.SpaceAfter = 12;
            firstHeadingStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            firstHeadingStyle.ParagraphFormat.LineSpacing = 22;
            firstHeadingStyle.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;

            //标题2
            Style secondHeadingStyle = doc.Styles[WdBuiltinStyle.wdStyleHeading2];
            secondHeadingStyle.set_BaseStyle(WdBuiltinStyle.wdStyleNormal); // 设置基础样式为普通文本
            secondHeadingStyle.set_NextParagraphStyle("段落正文");
            secondHeadingStyle.Font.NameFarEast = "黑体";
            secondHeadingStyle.Font.Size = 16;
            secondHeadingStyle.Font.Bold = 0;
            secondHeadingStyle.Font.Color = WdColor.wdColorBlack;
            secondHeadingStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            secondHeadingStyle.ParagraphFormat.FirstLineIndent = 0; // 2个汉字的左缩进
            secondHeadingStyle.ParagraphFormat.SpaceBefore = 8;
            secondHeadingStyle.ParagraphFormat.SpaceAfter = 4;
            secondHeadingStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            secondHeadingStyle.ParagraphFormat.LineSpacing = 22;

            //标题3
            Style thirdHeadingStyle = doc.Styles[WdBuiltinStyle.wdStyleHeading3];
            thirdHeadingStyle.set_BaseStyle(WdBuiltinStyle.wdStyleNormal); // 设置基础样式为普通文本
            thirdHeadingStyle.set_NextParagraphStyle("段落正文");
            thirdHeadingStyle.Font.NameFarEast = "黑体";
            thirdHeadingStyle.Font.Size = 14;
            thirdHeadingStyle.Font.Bold = 0;
            thirdHeadingStyle.Font.Color = WdColor.wdColorBlack;
            thirdHeadingStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            thirdHeadingStyle.ParagraphFormat.FirstLineIndent = 0; // 2个汉字的左缩进
            thirdHeadingStyle.ParagraphFormat.SpaceBefore = 0;
            thirdHeadingStyle.ParagraphFormat.SpaceAfter = 0;
            thirdHeadingStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            thirdHeadingStyle.ParagraphFormat.LineSpacing = 22;

            //标题4
            Style fourthHeadingStyle = doc.Styles[WdBuiltinStyle.wdStyleHeading4];
            fourthHeadingStyle.set_BaseStyle(WdBuiltinStyle.wdStyleNormal); // 设置基础样式为普通文本
            fourthHeadingStyle.set_NextParagraphStyle("段落正文");
            fourthHeadingStyle.Font.NameFarEast = "黑体";
            fourthHeadingStyle.Font.Size = 12;
            fourthHeadingStyle.Font.Bold = 0;
            fourthHeadingStyle.Font.Color = WdColor.wdColorBlack;
            fourthHeadingStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            fourthHeadingStyle.ParagraphFormat.FirstLineIndent = 0; // 2个汉字的左缩进
            fourthHeadingStyle.ParagraphFormat.SpaceBefore = 0;
            fourthHeadingStyle.ParagraphFormat.SpaceAfter = 0;
            fourthHeadingStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            fourthHeadingStyle.ParagraphFormat.LineSpacing = 22;

            //其他标题
            string otherHeadingTitle = "其他标题";
            Style otherHeadingStyle;
            if (!styleNamelist.Contains(otherHeadingTitle))
            {
                otherHeadingStyle = doc.Styles.Add(otherHeadingTitle);
            }
            else { otherHeadingStyle = doc.Styles[otherHeadingTitle]; }
            otherHeadingStyle.set_BaseStyle(WdBuiltinStyle.wdStyleHeading1); // 设置基础样式为标题1
                                                                             //if (!CheckStyle("段落正文")) { doc.Styles.Add("段落正文"); }
            otherHeadingStyle.set_NextParagraphStyle("段落正文");
            otherHeadingStyle.Font.Name = "Times New Roman";
            otherHeadingStyle.Font.NameFarEast = "黑体";
            otherHeadingStyle.Font.Size = 22;
            otherHeadingStyle.Font.Bold = 1;
            otherHeadingStyle.Font.Color = WdColor.wdColorBlack;
            otherHeadingStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            otherHeadingStyle.ParagraphFormat.FirstLineIndent = 0; // 2个汉字的左缩进
            otherHeadingStyle.ParagraphFormat.SpaceBefore = 0;
            otherHeadingStyle.ParagraphFormat.SpaceAfter = 0;
            otherHeadingStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            otherHeadingStyle.ParagraphFormat.LineSpacing = 22;
            otherHeadingStyle.LinkToListTemplate(null);
            otherHeadingStyle.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
            //题注
            string pictureHeadingTitle = "题注";
            Style pictureHeadingStyle;
            if (!styleNamelist.Contains(pictureHeadingTitle))
            {
                pictureHeadingStyle = doc.Styles.Add(pictureHeadingTitle);
                pictureHeadingStyle.set_BaseStyle(WdBuiltinStyle.wdStyleNormal); // 设置基础样式为普通文本
                                                                                 //if (!CheckStyle("段落正文")) { doc.Styles.Add("段落正文"); }
                pictureHeadingStyle.set_NextParagraphStyle("段落正文");
            }
            else { pictureHeadingStyle = doc.Styles[pictureHeadingTitle]; }
            pictureHeadingStyle.Font.NameFarEast = "黑体";
            pictureHeadingStyle.Font.Size = 10.5f;
            pictureHeadingStyle.Font.Color = WdColor.wdColorBlack;
            pictureHeadingStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            pictureHeadingStyle.ParagraphFormat.FirstLineIndent = 0; // 2个汉字的左缩进
            pictureHeadingStyle.ParagraphFormat.SpaceBefore = 0;
            pictureHeadingStyle.ParagraphFormat.SpaceAfter = 0;
            pictureHeadingStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            pictureHeadingStyle.ParagraphFormat.LineSpacing = 22;

            //摘要标题--标题
            Style abstractHeadingStyle = doc.Styles[WdBuiltinStyle.wdStyleTitle];
            abstractHeadingStyle.set_NextParagraphStyle("摘要正文");
            abstractHeadingStyle.Font.Name = "Times New Roman";
            abstractHeadingStyle.Font.NameFarEast = "黑体";
            abstractHeadingStyle.Font.Size = 22;
            abstractHeadingStyle.Font.Bold = 1;
            abstractHeadingStyle.Font.Color = WdColor.wdColorBlack;
            abstractHeadingStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            abstractHeadingStyle.ParagraphFormat.CharacterUnitFirstLineIndent = 0; // 0个汉字的左缩进
            abstractHeadingStyle.ParagraphFormat.FirstLineIndent = 0;
            abstractHeadingStyle.ParagraphFormat.SpaceBefore = 24;
            abstractHeadingStyle.ParagraphFormat.SpaceAfter = 12;
            abstractHeadingStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            abstractHeadingStyle.ParagraphFormat.LineSpacing = 22;
            abstractHeadingStyle.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText;

            //图片样式
            string pictureBodyTitle = "图片内容";
            Style pictureBodyStyle;
            if (!styleNamelist.Contains(pictureBodyTitle))
            {
                pictureBodyStyle = doc.Styles.Add(pictureBodyTitle);
                pictureBodyStyle.set_BaseStyle(WdBuiltinStyle.wdStyleNormal); // 设置基础样式为普通文本
                pictureBodyStyle.set_NextParagraphStyle("段落正文");
            }
            else { pictureBodyStyle = doc.Styles[pictureBodyTitle]; }
            pictureBodyStyle.Font.NameFarEast = "黑体";
            pictureBodyStyle.Font.Size = 10.5f;
            pictureBodyStyle.Font.Color = WdColor.wdColorBlack;
            pictureBodyStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            pictureBodyStyle.ParagraphFormat.CharacterUnitFirstLineIndent = 0f; // 0个汉字的左缩进
            pictureBodyStyle.ParagraphFormat.FirstLineIndent = 0f;
            pictureBodyStyle.ParagraphFormat.SpaceBefore = 0;
            pictureBodyStyle.ParagraphFormat.SpaceAfter = 0;
            pictureBodyStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;

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

            // 设置表格字体
            string tableHeadlineTitle = "表内容";
            Style tableHeadlineStyle;
            if (!styleNamelist.Contains(tableHeadlineTitle))
            {
                tableHeadlineStyle = doc.Styles.Add(tableHeadlineTitle);
            }
            else { tableHeadlineStyle = doc.Styles[tableHeadlineTitle]; }

            tableHeadlineStyle.Font.Name = "Times New Roman";
            tableHeadlineStyle.Font.NameFarEast = "宋体";
            tableHeadlineStyle.Font.Size = 10.5f;
            tableHeadlineStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            #endregion
            SetCatalogStyles(doc);
        }

        /// <summary>
        /// 设置目录/页眉页脚 页码格式
        /// </summary>
        /// <param name="doc"></param>
        public static void SetCatalogStyles(Document doc)
        {
            //图片样式
            Style toc1Style = doc.Styles[WdBuiltinStyle.wdStyleTOC1];
            toc1Style.Font.NameFarEast = "宋体";
            toc1Style.Font.Name = "Times New Roman";
            toc1Style.Font.Size = 14f;
            toc1Style.Font.Color = WdColor.wdColorBlack;
            toc1Style.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
            toc1Style.ParagraphFormat.FirstLineIndent = 0; // 0个汉字的左缩进
            toc1Style.ParagraphFormat.LeftIndent = 0;
            toc1Style.ParagraphFormat.SpaceBefore = 0;
            toc1Style.ParagraphFormat.SpaceAfter = 0;
            toc1Style.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            toc1Style.ParagraphFormat.LineSpacing = 22;
            toc1Style.ParagraphFormat.TabStops.ClearAll();
            toc1Style.ParagraphFormat.TabStops.Add(
                Position: 450f,   // 制表位位置，以磅为单位
                Alignment: WdTabAlignment.wdAlignTabRight,      // 制表位对齐方式
                Leader: WdTabLeader.wdTabLeaderDots             // 制表位前导符
            );


            Style toc2Style = doc.Styles[WdBuiltinStyle.wdStyleTOC2];
            toc2Style.Font.NameFarEast = "宋体";
            toc2Style.Font.Name = "Times New Roman";
            toc2Style.Font.Size = 12f;
            toc2Style.Font.Color = WdColor.wdColorBlack;
            toc2Style.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
            toc2Style.ParagraphFormat.FirstLineIndent = 0; // 0个汉字的左缩进
            toc2Style.ParagraphFormat.LeftIndent = 2f;
            toc2Style.ParagraphFormat.SpaceBefore = 0;
            toc2Style.ParagraphFormat.SpaceAfter = 0;
            toc2Style.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            toc2Style.ParagraphFormat.LineSpacing = 22;
            toc2Style.ParagraphFormat.TabStops.ClearAll();
            toc2Style.ParagraphFormat.TabStops.Add(
                Position: 450f,   // 制表位位置，以磅为单位
                Alignment: WdTabAlignment.wdAlignTabRight,      // 制表位对齐方式
                Leader: WdTabLeader.wdTabLeaderDots             // 制表位前导符
            );

            Style toc3Style = doc.Styles[WdBuiltinStyle.wdStyleTOC3];
            toc3Style.Font.NameFarEast = "宋体";
            toc3Style.Font.Name = "Times New Roman";
            toc3Style.Font.Size = 12f;
            toc3Style.Font.Color = WdColor.wdColorBlack;
            toc3Style.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
            toc3Style.ParagraphFormat.FirstLineIndent = 0; // 0个汉字的左缩进
            toc3Style.ParagraphFormat.LeftIndent = 2f; // 0个汉字的左缩进
            toc3Style.ParagraphFormat.SpaceBefore = 0;
            toc3Style.ParagraphFormat.SpaceAfter = 0;
            toc3Style.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            toc3Style.ParagraphFormat.LineSpacing = 22;
            toc3Style.ParagraphFormat.TabStops.ClearAll();
            toc3Style.ParagraphFormat.TabStops.Add(
                Position: 450f,   // 制表位位置，以磅为单位
                Alignment: WdTabAlignment.wdAlignTabRight,      // 制表位对齐方式
                Leader: WdTabLeader.wdTabLeaderDots             // 制表位前导符
            );

            //页眉
            Style headerStyle = doc.Styles[WdBuiltinStyle.wdStyleHeader];
            headerStyle.Font.Name = "Times New Roman";
            headerStyle.Font.NameFarEast = "楷体";
            headerStyle.Font.Size = 9f;
            headerStyle.Font.Color = WdColor.wdColorBlack;
            headerStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            headerStyle.ParagraphFormat.FirstLineIndent = 0; // 0个汉字的左缩进
            headerStyle.ParagraphFormat.LeftIndent = 0f; // 0个汉字的左缩进
            headerStyle.ParagraphFormat.SpaceBefore = 0;
            headerStyle.ParagraphFormat.SpaceAfter = 0;
            headerStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            headerStyle.ParagraphFormat.TabStops.ClearAll();
            headerStyle.ParagraphFormat.TabStops.Add(
                Position: 500f,   // 制表位位置，以磅为单位
                Alignment: WdTabAlignment.wdAlignTabRight,      // 制表位对齐方式
                Leader: WdTabLeader.wdTabLeaderSpaces             // 制表位前导符
            );

            Style footerStyle = doc.Styles[WdBuiltinStyle.wdStyleFooter];
            footerStyle.Font.Name = "Times New Roman";
            footerStyle.Font.NameFarEast = "楷体";
            footerStyle.Font.Size = 9f;
            footerStyle.Font.Color = WdColor.wdColorBlack;
            footerStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            footerStyle.ParagraphFormat.FirstLineIndent = 0; // 0个汉字的左缩进
            footerStyle.ParagraphFormat.LeftIndent = 0f; // 0个汉字的左缩进
            footerStyle.ParagraphFormat.SpaceBefore = 0;
            footerStyle.ParagraphFormat.SpaceAfter = 0;
            footerStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
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
                para.set_Style(WdBuiltinStyle.wdStyleCaption);
            }
            Tables tables = Globals.ThisAddIn.Application.Selection.Tables;
            foreach (Table table in tables)
            {
                Range rng = table.Range;
                //rng.set_Style("表格内容");
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
                    para.set_Style(WdBuiltinStyle.wdStyleCaption);
                }
            }
        }
    }

}
