using ArticleFormatApp_USST.Properties;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows.Forms;


namespace ArticleFormatApp_USST
{
    public partial class MainRibbon
    {
        private void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            this.tab2.Label = Resources.AppName;
        }

        #region 生成样式
        /// <summary>
        /// 初始化样式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_CreateStyles_Click(object sender, RibbonControlEventArgs e)
        {
            var mbx = MessageBox.Show("是否初始化样式！", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (mbx.Equals(DialogResult.OK))
            {
                try
                {
                    var doc = Globals.ThisAddIn.Application.ActiveDocument;
                    CommonFunction.CreateStyles(ref doc);
                    MessageBox.Show("样式初始化完成！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// 生成公式样式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_CreateFormulaStyle1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var doc = Globals.ThisAddIn.Application.ActiveDocument;
                CommonFunction.CreateFormulaStyle(ref doc);
                MessageBox.Show("公式样式初始化完成！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        /// <summary>
        /// 生成图表样式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_CreatePicTableStyle_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var doc = Globals.ThisAddIn.Application.ActiveDocument;
                CommonFunction.CreatePicTableStyle(ref doc);
                MessageBox.Show("图表样式初始化完成！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region 标题部分
        /// <summary>
        /// 目录
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_CatalogHeading_Click(object sender, RibbonControlEventArgs e)
        {
            CommonFunction.SetStyle(Resources.Heading);
        }

        /// <summary>
        /// 一级标题
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_FirstHeading_Click(object sender, RibbonControlEventArgs e)
        {
            CommonFunction.SetStyle(Resources.Heading1);
        }

        /// <summary>
        /// 二级标题
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_SecondHeading_Click(object sender, RibbonControlEventArgs e)
        {
            CommonFunction.SetStyle(Resources.Heading2);
        }

        /// <summary>
        /// 三级标题
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_ThirdHeading_Click(object sender, RibbonControlEventArgs e)
        {
            CommonFunction.SetStyle(Resources.Heading3);
        }

        /// <summary>
        /// 四级标题
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_FourthHeading_Click(object sender, RibbonControlEventArgs e)
        {
            CommonFunction.SetStyle(Resources.Heading4);
        }


        /// <summary>
        /// 其他标题
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_OtherHeading_Click(object sender, RibbonControlEventArgs e)
        {
            CommonFunction.SetStyle(Resources.OtherHeading);
        }

        private void btn_PictureHeading_Click(object sender, RibbonControlEventArgs e)
        {
            CommonFunction.SetStyle("图题注");
        }

        private void btn_TableHeading_Click(object sender, RibbonControlEventArgs e)
        {
            CommonFunction.SetStyle("表题注");
        }

        private void btn_AbstractHeading_Click(object sender, RibbonControlEventArgs e)
        {
            CommonFunction.SetStyle(Resources.OtherHeading);
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



        private void btn_PictureBody_Click(object sender, RibbonControlEventArgs e)
        {
            CommonFunction.SetStyle("图片内容");
        }

        private void btn_TableBody_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Table table = Globals.ThisAddIn.Application.Selection.Tables[1];
                table.Range.set_Style("表内容");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        /// <summary>
        /// 关键词
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_Keywords_Click(object sender, RibbonControlEventArgs e)
        {
            string title = "关键词";
            CommonFunction.SetStyle(title);
        }

        /// <summary>
        /// 表注释
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_tableNote_Click(object sender, RibbonControlEventArgs e)
        {
            string title = "表注释";
            CommonFunction.SetStyle(title);
        }

        /// <summary>
        /// 公式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_formula_Click(object sender, RibbonControlEventArgs e)
        {
            string title = "公式";
            CommonFunction.SetStyle(title);
        }

        /// <summary>
        /// 带公式正文
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_ContainsFormula_Click(object sender, RibbonControlEventArgs e)
        {
            string title = "带公式正文";
            CommonFunction.SetStyle(title);
        }

        /// <summary>
        /// 三线表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_tableThreeline_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Table table = Globals.ThisAddIn.Application.Selection.Tables[1];
                table.set_Style("三线表");
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
        #endregion


        #region 其他
        /// <summary>
        /// 插入题注
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        /// <summary>
        /// 刷新引用序号
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

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


        /// <summary>
        /// 页面格式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_PageSetting_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Document doc = Globals.ThisAddIn.Application.ActiveDocument;

                // 获取文档的页边距集合
                PageSetup pageSetup = doc.PageSetup;

                // 设置页边距
                pageSetup.TopMargin = doc.Application.CentimetersToPoints(3.5f);
                pageSetup.BottomMargin = doc.Application.CentimetersToPoints(2.5f);
                pageSetup.LeftMargin = doc.Application.CentimetersToPoints(3.0f);
                pageSetup.RightMargin = doc.Application.CentimetersToPoints(3.0f);

                // 设置装订线
                pageSetup.Gutter = doc.Application.CentimetersToPoints(0f);

                // 设置页眉页脚距离
                pageSetup.HeaderDistance = doc.Application.CentimetersToPoints(2.0f);
                pageSetup.FooterDistance = doc.Application.CentimetersToPoints(2.0f);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        /// <summary>
        /// 设置页眉
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_SetHeaderAndFooter_Click(object sender, RibbonControlEventArgs e)
        {
            var mbx = MessageBox.Show("是否重新设置页眉！", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (mbx.Equals(DialogResult.OK))
            {
                try
                {
                    Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                    //doc.PageSetup.OddAndEvenPagesHeaderFooter = 0; //-1(功能开启)和0（功能关闭）
                    foreach (Section section in doc.Sections)
                    {
                        section.PageSetup.OddAndEvenPagesHeaderFooter = -1; //-1(功能开启)和0（功能关闭）
                        Paragraph para = section.Range.Paragraphs[1]; //获取节的第一段
                        Style style = para.get_Style();

                        section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
                        section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
                        section.Headers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = false;//链接到上一节
                        section.Footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = false;//链接到上一节
                        #region 设置页眉
                        //奇数页区域
                        Range headerRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;//奇，偶 wdHeaderFooterEvenPages
                        if (style.NameLocal.Equals(Resources.Heading1))
                        {
                            // 移除页眉
                            headerRange.Delete();
                            Paragraph p1 = headerRange.Paragraphs[1];

                            //添加标题级别
                            Paragraph rightP = p1.Range.Paragraphs.Add(p1.Range);
                            rightP.Range.Text = para.Range.ListFormat.ListString.Trim() + " " + para.Range.Text.Trim();

                            //headerRange.Fields.Add(rightP.Range, WdFieldType.wdFieldEmpty, $" STYLEREF  \"{Resources.Heading1}\" ", true);
                            ////添加分隔内容
                            //headerRange.Text += " ";
                            ////标题内容
                            //Paragraph rightP3 = p1.Range.Paragraphs.Add(p1.Range);
                            //headerRange.Fields.Add(rightP3.Range, WdFieldType.wdFieldEmpty, $" STYLEREF  \"{Resources.Heading1}\" \\n ", true);
                            headerRange.set_Style(WdBuiltinStyle.wdStyleHeader);
                        }
                        else if (style.NameLocal.Equals(Resources.OtherHeading)) //其他标题
                        {
                            // 移除页眉
                            headerRange.Delete();
                            Paragraph p1 = headerRange.Paragraphs[1];
                            //标题内容
                            Paragraph rightP3 = p1.Range.Paragraphs.Add(p1.Range);
                            rightP3.Range.Text = para.Range.Text.Trim();

                            //headerRange.Fields.Add(rightP3.Range, WdFieldType.wdFieldEmpty, $" STYLEREF  \"{Resources.OtherHeading}\" ", true);
                            //headerRange.set_Style(WdBuiltinStyle.wdStyleHeader);
                        }
                        else if (style.NameLocal.Equals(Resources.Heading)) //标题
                        {
                            // 移除页眉
                            headerRange.Delete();
                            Paragraph p1 = headerRange.Paragraphs[1];
                            //标题内容
                            Paragraph rightP3 = p1.Range.Paragraphs.Add(p1.Range);
                            rightP3.Range.Text = para.Range.Text.Trim();
                            //headerRange.Fields.Add(rightP3.Range, WdFieldType.wdFieldEmpty, $" STYLEREF  \"{Resources.Heading}\" ", true);
                            //headerRange.set_Style(WdBuiltinStyle.wdStyleHeader);
                        }
                        //headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;//居中
                        headerRange.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;//底线
                        headerRange.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;//无线
                        headerRange.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleNone;//无线
                        headerRange.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleNone;//无线

                        //偶数页
                        Range headerEvenRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range;
                        section.Headers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = false;
                        if (style.NameLocal == Resources.Heading1 ||
                            style.NameLocal == Resources.Heading ||
                            style.NameLocal == Resources.OtherHeading)
                        {
                            // 设置页眉
                            headerEvenRange.Delete();
                            headerEvenRange.Text = "上海理工大学硕士学位论文";
                            headerEvenRange.set_Style(WdBuiltinStyle.wdStyleHeader);

                            //headerEvenRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;//居中
                            headerEvenRange.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;//底线
                            headerEvenRange.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;//无线
                            headerEvenRange.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleNone;//无线
                            headerEvenRange.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleNone;//无线
                            //section.Headers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = true;//链接到上一节
                        }
                        #endregion

                        #region 设置页脚
                        #endregion
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                MessageBox.Show("页眉设置完成！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// 分节
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_SeparateSection_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                Range currentRange = doc.ActiveWindow.Selection.Range;

                // 插入分节符
                currentRange.InsertBreak(WdBreakType.wdSectionBreakNextPage);//wdSectionBreakNextPage
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 奇数页分节
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_SepOdd_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                Range currentRange = doc.ActiveWindow.Selection.Range;

                // 插入分节符
                currentRange.InsertBreak(WdBreakType.wdSectionBreakOddPage);//wdSectionBreakNextPage
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 交叉引用
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_quote_Click(object sender, RibbonControlEventArgs e)
        {
            Dialog crossRefDialog = Globals.ThisAddIn.Application.ActiveDocument.Application.Dialogs[WdWordDialog.wdDialogInsertCrossReference];
            crossRefDialog.Show();
        }




        private void btn_InsertFormulaCaption_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (CommonFunction.SetStyle(Resources.FormulaLabelName))
                {
                    Selection selection = Globals.ThisAddIn.Application.Selection;
                    //selection = Globals.ThisAddIn.Application.Selection;
                    selection.TypeText("\t(");
                    selection.InsertCaption(Resources.FormulaLabelName,
                        ExcludeLabel: bool.Parse(Resources.TableExclueLabel));
                    selection.TypeText(")");
                    CommonFunction.SetStyle(Resources.FormulaLabelName);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void btn_ListToText_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var lst = Globals.ThisAddIn.Application.Selection.Range.ListFormat.List;
                if (lst != null)
                    lst.ConvertNumbersToText();
                else
                {
                    var mbx = MessageBox.Show("是否将全文的列表转为文本！", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                    if (mbx.Equals(DialogResult.OK))
                    {
                        var doc = Globals.ThisAddIn.Application.ActiveDocument;
                        doc.ConvertNumbersToText();
                        MessageBox.Show("转换完成！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        #endregion


    }
}
