using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows.Forms;

namespace 沈体论文格式调整word插件
{
    public partial class MainRibbon
    {
        private void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }
        private void btn_CreateStyles_Click(object sender, RibbonControlEventArgs e)
        {
            var mbx = MessageBox.Show("是否初始化样式！", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (mbx.Equals(DialogResult.OK))
            {
                try
                {
                    var doc = Globals.ThisAddIn.Application.ActiveDocument;
                    CommonFunction.ApplyMultiLevelListStyle(doc);
                    CommonFunction.InitStyles(doc);
                    MessageBox.Show("样式初始化完成！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        #region 标题部分
        private void btn_FirstHeading_Click(object sender, RibbonControlEventArgs e)
        {
            CommonFunction.SetStyle(WdBuiltinStyle.wdStyleHeading1);
        }

        private void btn_SecondHeading_Click(object sender, RibbonControlEventArgs e)
        {
            CommonFunction.SetStyle(WdBuiltinStyle.wdStyleHeading2);
        }

        private void btn_ThirdHeading_Click(object sender, RibbonControlEventArgs e)
        {
            CommonFunction.SetStyle(WdBuiltinStyle.wdStyleHeading3);
        }

        private void btn_FourthHeading_Click(object sender, RibbonControlEventArgs e)
        {
            CommonFunction.SetStyle(WdBuiltinStyle.wdStyleHeading4);
        }

        private void btn_OtherHeading_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string otherHeadingTitle = "其他标题";
                CommonFunction.SetStyle(otherHeadingTitle);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn_PictureHeading_Click(object sender, RibbonControlEventArgs e)
        {
            CommonFunction.SetStyle(WdBuiltinStyle.wdStyleCaption);
        }

        private void btn_TableHeading_Click(object sender, RibbonControlEventArgs e)
        {
            CommonFunction.SetStyle(WdBuiltinStyle.wdStyleCaption);
        }

        private void btn_AbstractHeading_Click(object sender, RibbonControlEventArgs e)
        {
            CommonFunction.SetStyle(WdBuiltinStyle.wdStyleTitle);
        }
        #endregion

        #region 正文部分

        private void btn_ParaBody_Click(object sender, RibbonControlEventArgs e)
        {
            string paraBodyTitle = "段落正文";
            CommonFunction.SetStyle(paraBodyTitle);
        }

        private void btn_AbstractBody_Click(object sender, RibbonControlEventArgs e)
        {
            string abstractBodyTitle = "摘要正文";
            CommonFunction.SetStyle(abstractBodyTitle);
        }

        private void btn_ReferenceBody_Click(object sender, RibbonControlEventArgs e)
        {
            string referenceBodyTitle = "参考文献列表";
            CommonFunction.SetStyle(referenceBodyTitle);
        }

        private void btn_AcknowledgementBody_Click(object sender, RibbonControlEventArgs e)
        {
            string acknowledgementBodyTitle = "致谢正文";
            CommonFunction.SetStyle(acknowledgementBodyTitle);
        }

        private void btn_AchievementsBody_Click(object sender, RibbonControlEventArgs e)
        {
            string achievementsBodyTitle = "学术成果正文";
            CommonFunction.SetStyle(achievementsBodyTitle);
        }

        private void btn_AppendixBody_Click(object sender, RibbonControlEventArgs e)
        {
            string appendixBodyTitle = "附录正文";
            CommonFunction.SetStyle(appendixBodyTitle);
        }

        #endregion

        private void btn_PictureBody_Click(object sender, RibbonControlEventArgs e)
        {
            string pictureBodyTitle = "图片内容";
            CommonFunction.SetStyle(pictureBodyTitle);
        }

        private void btn_TableBody_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //CommonFunction.ThreeLineTable(Globals.ThisAddIn.Application.ActiveDocument);
                string title = "三线表";
                Table table = Globals.ThisAddIn.Application.Selection.Tables[1];
                table.set_Style(title);
                table.Rows[1].Range.Font.Bold = 1;
                table.Rows[1].Borders[WdBorderType.wdBorderBottom].Visible = true;
                table.Rows[1].Borders[WdBorderType.wdBorderBottom].LineWidth = WdLineWidth.wdLineWidth075pt;
                table.Rows[1].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
                table.Range.set_Style("表内容");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn_CatalogHeading_Click(object sender, RibbonControlEventArgs e)
        {
            CommonFunction.SetStyle(WdBuiltinStyle.wdStyleTitle);
        }

        private void btn_InsertCaption_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                CommonFunction.InsertCaption();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn_RefrushIndex_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Document doc = Globals.ThisAddIn.Application.ActiveDocument;

                // 遍历文档中的所有字段
                foreach (Field field in doc.Fields)
                {
                    if (field.Type == WdFieldType.wdFieldRef || field.Type == WdFieldType.wdFieldTOC)
                    {
                        field.Update();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            MessageBox.Show("交叉引用及目录刷新完成！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btn_PageSetting_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Document doc = Globals.ThisAddIn.Application.ActiveDocument;

                // 获取文档的页边距集合
                PageSetup pageSetup = doc.PageSetup;

                // 设置页边距
                pageSetup.TopMargin = doc.Application.CentimetersToPoints(2.5f);
                pageSetup.BottomMargin = doc.Application.CentimetersToPoints(2.0f);
                pageSetup.LeftMargin = doc.Application.CentimetersToPoints(2.5f);
                pageSetup.RightMargin = doc.Application.CentimetersToPoints(2.0f);

                // 设置装订线
                pageSetup.Gutter = doc.Application.CentimetersToPoints(0.5f);

                // 设置页眉页脚距离
                pageSetup.HeaderDistance = doc.Application.CentimetersToPoints(1.5f);
                pageSetup.FooterDistance = doc.Application.CentimetersToPoints(1.75f);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn_SetHeaderAndFooter_Click(object sender, RibbonControlEventArgs e)
        {
            var mbx = MessageBox.Show("是否重新设置页眉页脚！", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (mbx.Equals(DialogResult.OK))
            {
                try
                {
                    Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                    //doc.PageSetup.OddAndEvenPagesHeaderFooter = 0;
                    foreach (Section section in doc.Sections)
                    {
                        Paragraph para = section.Range.Paragraphs[1];
                        Range headerRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;//当前节的所有页眉
                        Style style = para.get_Style();
                        if (style.NameLocal == "标题 1" ||
                            style.NameLocal == "标题" ||
                            style.NameLocal == "其他标题")
                        {
                            section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
                            // 设置页眉
                            headerRange.Delete();
                            var text = para.Range.Text.Trim();
                            var text2 = para.Range.ListFormat.ListString.Trim();
                            text = text.Replace("\r\n", "");
                            headerRange.Text = "沈阳体育学院硕士学位论文\t" + text2 + text;
                            headerRange.set_Style(WdBuiltinStyle.wdStyleHeader);

                            headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;//居中
                            headerRange.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;//底线
                            headerRange.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;//无线
                            headerRange.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleNone;//无线
                            headerRange.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleNone;//无线
                            section.Headers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = true;//链接到上一节
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                MessageBox.Show("页眉页脚设置完成！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btn_SeparateSection_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                Range currentRange = doc.ActiveWindow.Selection.Range;

                // 插入分节符
                currentRange.InsertBreak(WdBreakType.wdSectionBreakNextPage);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn_quote_Click(object sender, RibbonControlEventArgs e)
        {
            Dialog crossRefDialog = Globals.ThisAddIn.Application.ActiveDocument.Application.Dialogs[WdWordDialog.wdDialogInsertCrossReference];
            // 获取交叉引用对话框对象
            //Dialog crossRefDialog = Globals.ThisAddIn.Application.Dialogs[WdWordDialog.wdDialogInsertCrossReference];
            crossRefDialog.Show();
        }
    }
}
