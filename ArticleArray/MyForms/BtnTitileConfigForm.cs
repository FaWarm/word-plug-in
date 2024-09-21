using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using static System.Collections.Specialized.BitVector32;
using Section = Microsoft.Office.Interop.Word.Section;

namespace ArticleArray.MyForms
{
    public partial class BtnTitileConfigForm : Form
    {
        public BtnTitileConfigForm()
        {
            InitializeComponent();
            this.ChbxRmIndex.Checked = true;
        }
        private Document _document;
        private void BtnSet_Click(object sender, EventArgs e)
        {
            _document = Globals.ThisAddIn.Application.ActiveDocument;
            Regex regexH1 = new Regex(@"^([ ]*\d+[\t ]*)", RegexOptions.Compiled);//一级标题 1
            Regex regexH1_2 = new Regex(@"^[ ]*第.{0,4}章[\t ]*", RegexOptions.Compiled);//第x章
            Regex regexH1_3 = new Regex(@"^[ ]*第.{0,4}节[\t ]*", RegexOptions.Compiled);//第x节
            Regex regexH1_4 = new Regex(@"^[ ]*\(.{0,4}\)[\t ]*", RegexOptions.Compiled);//(1)
            Regex regexH1_5 = new Regex(@"^[ ]*\（.{0,4}\）[\t ]*", RegexOptions.Compiled);//（1）
            Regex regexH1_6 = new Regex(@"^[ ]*\d{1,4}\)[\t ]*", RegexOptions.Compiled);//1)
            Regex regexH1_7 = new Regex(@"^[ ]*([一二三四五六七八九十]{1,3}、[\t ]*)", RegexOptions.Compiled);//一、
            Regex regexH1_8 = new Regex(@"^[ ]*([一二三四五六七八九十]{1,3}\.[\t ]*)", RegexOptions.Compiled);//一.
            Regex regexH1_9 = new Regex(@"^[ ]*(\d+\.[\t ]*)", RegexOptions.Compiled);//1.
            Regex regexH1_10 = new Regex(@"^[ ]*(\d+、[\t ]*)", RegexOptions.Compiled);//1、

            Regex regexH2 = new Regex(@"^([ ]*\d+\.\d+[\t ]*)", RegexOptions.Compiled);//二级标题1.x
            Regex regexH3 = new Regex(@"^([ ]*\d+\.\d+\.\d+[\t ]*)", RegexOptions.Compiled);//三级标题1.x.x
            string patternH3 = "<([0-9]@[.][0-9]@[.][0-9]@)";//三级标题1.x.x
            Regex regexH4 = new Regex(@"^([ ]*\d+\.\d+\.\d+\.\d+[\t ]*)", RegexOptions.Compiled);//四级标题1.x.x.x
            Regex regexH5 = new Regex(@"^([ ]*\d+\.\d+\.\d+\.\d+\.\d+[\t ]*)", RegexOptions.Compiled);//五级标题1.x.x.x.x

            Regex regexHZy = new Regex(@"\b[ ]*摘[ ]*要\r", RegexOptions.Compiled);//摘  要
            //Regex regexHZy1 = new Regex(@"\b摘要[:：]+\s*");//摘要:
            Regex regexHZy2 = new Regex(@"(?i)\b[ ]*abstract[:：]+\s*", RegexOptions.Compiled);//摘要
            //Regex regexHZy3 = new Regex(@"\b关键字[:：]+\s*");//关键字
            //Regex regexHZy4 = new Regex(@"(?i)\bkeywords[:：]+\s*");//关键字
            Regex regexHZy5 = new Regex(@"(?i)\b(?:[ ]*摘要[:：]+|[ ]*摘\s+要[:：]+|[ ]*abstract[:：]+|[ ]*关键字[:：]+|[ ]*关键词[:：]+|[ ]*key[ ]*words[:：]+)\s*", RegexOptions.Compiled);//摘要和关键字
            Regex regexHCkwx = new Regex(@"\b[ ]*参考文献\b", RegexOptions.Compiled);//参考文献

            switch (CbxZhaiyao.Text)
            {
                case "":
                    break;
                case "摘  要":
                    ChosingHeading(regexHZy, "摘要");
                    break;
                case "abstract":
                    ChosingHeading(regexHZy2, "摘要");
                    break;
                case "摘要,abstract,关键字,关键词,keywords":
                    ChosingHeading(regexHZy5, "摘要");
                    break;
                default: break;
            }

            switch (CbxCankao.Text)
            {
                case "":
                    break;
                case "参考文献":
                    ChosingHeading(regexHCkwx, "参考文献");
                    break;
                default: break;
            }


            switch (CbxHeading5.Text)
            {
                case "":
                    break;
                case "1.x.x.x.x":
                    ChosingHeading(regexH5, WdBuiltinStyle.wdStyleHeading5);
                    break;
                case "1.":
                    ChosingHeading(regexH1_9, WdBuiltinStyle.wdStyleHeading5);
                    break;
                case "1、":
                    ChosingHeading(regexH1_10, WdBuiltinStyle.wdStyleHeading5);
                    break;
                case "第x章":
                    ChosingHeading(regexH1_2, WdBuiltinStyle.wdStyleHeading5);
                    break;
                case "第x节":
                    ChosingHeading(regexH1_3, WdBuiltinStyle.wdStyleHeading5);
                    break;
                case "(1)":
                    ChosingHeading(regexH1_4, WdBuiltinStyle.wdStyleHeading5);
                    break;
                case "（1）":
                    ChosingHeading(regexH1_5, WdBuiltinStyle.wdStyleHeading5);
                    break;
                case "1)":
                    ChosingHeading(regexH1_6, WdBuiltinStyle.wdStyleHeading5);
                    break;
                case "一、":
                    ChosingHeading(regexH1_7, WdBuiltinStyle.wdStyleHeading5);
                    break;
                case "一.":
                    ChosingHeading(regexH1_8, WdBuiltinStyle.wdStyleHeading5);
                    break;
                default:
                    break;
            }

            switch (CbxHeading4.Text)
            {
                case "":
                    break;
                case "1、":
                    ChosingHeading(regexH1_10, WdBuiltinStyle.wdStyleHeading4);
                    break;
                case "1":
                    ChosingHeading(regexH1, WdBuiltinStyle.wdStyleHeading4);
                    break;
                case "1.":
                    ChosingHeading(regexH1_9, WdBuiltinStyle.wdStyleHeading4);
                    break;
                case "1.x.x.x":
                    ChosingHeading(regexH4, WdBuiltinStyle.wdStyleHeading4);
                    break;
                case "第x章":
                    ChosingHeading(regexH1_2, WdBuiltinStyle.wdStyleHeading4);
                    break;
                case "第x节":
                    ChosingHeading(regexH1_3, WdBuiltinStyle.wdStyleHeading4);
                    break;
                case "(1)":
                    ChosingHeading(regexH1_4, WdBuiltinStyle.wdStyleHeading4);
                    break;
                case "（1）":
                    ChosingHeading(regexH1_5, WdBuiltinStyle.wdStyleHeading4);
                    break;
                case "1)":
                    ChosingHeading(regexH1_6, WdBuiltinStyle.wdStyleHeading4);
                    break;
                case "一、":
                    ChosingHeading(regexH1_7, WdBuiltinStyle.wdStyleHeading4);
                    break;
                case "一.":
                    ChosingHeading(regexH1_8, WdBuiltinStyle.wdStyleHeading4);
                    break;
                default:
                    break;
            }

            switch (CbxHeading3.Text)
            {
                case "":
                    break;
                case "1、":
                    ChosingHeading(regexH1_10, WdBuiltinStyle.wdStyleHeading3);
                    break;
                case "1":
                    ChosingHeading(regexH1, WdBuiltinStyle.wdStyleHeading3);
                    break;
                case "1.":
                    ChosingHeading(regexH1_9, WdBuiltinStyle.wdStyleHeading3);
                    break;
                case "1.x.x":
                    findHeading(patternH3, WdBuiltinStyle.wdStyleHeading3);
                    //ChosingHeading(regexH3, WdBuiltinStyle.wdStyleHeading3);
                    break;
                case "第x章":
                    ChosingHeading(regexH1_2, WdBuiltinStyle.wdStyleHeading3);
                    break;
                case "第x节":
                    ChosingHeading(regexH1_3, WdBuiltinStyle.wdStyleHeading3);
                    break;
                case "(1)":
                    ChosingHeading(regexH1_4, WdBuiltinStyle.wdStyleHeading3);
                    break;
                case "（1）":
                    ChosingHeading(regexH1_5, WdBuiltinStyle.wdStyleHeading3);
                    break;
                case "1)":
                    ChosingHeading(regexH1_6, WdBuiltinStyle.wdStyleHeading3);
                    break;
                case "一、":
                    ChosingHeading(regexH1_7, WdBuiltinStyle.wdStyleHeading3);
                    break;
                case "一.":
                    ChosingHeading(regexH1_8, WdBuiltinStyle.wdStyleHeading3);
                    break;
                default:
                    break;
            }

            switch (CbxHeading2.Text)
            {
                case "":
                    break;
                case "1、":
                    ChosingHeading(regexH1_10, WdBuiltinStyle.wdStyleHeading2);
                    break;
                case "1":
                    ChosingHeading(regexH1, WdBuiltinStyle.wdStyleHeading2);
                    break;
                case "1.":
                    ChosingHeading(regexH1_9, WdBuiltinStyle.wdStyleHeading2);
                    break;
                case "1.x":
                    ChosingHeading(regexH2, WdBuiltinStyle.wdStyleHeading2);
                    break;
                case "第x章":
                    ChosingHeading(regexH1_2, WdBuiltinStyle.wdStyleHeading2);
                    break;
                case "第x节":
                    ChosingHeading(regexH1_3, WdBuiltinStyle.wdStyleHeading2);
                    break;
                case "(1)":
                    ChosingHeading(regexH1_4, WdBuiltinStyle.wdStyleHeading2);
                    break;
                case "（1）":
                    ChosingHeading(regexH1_5, WdBuiltinStyle.wdStyleHeading2);
                    break;
                case "1)":
                    ChosingHeading(regexH1_6, WdBuiltinStyle.wdStyleHeading2);
                    break;
                case "一、":
                    ChosingHeading(regexH1_7, WdBuiltinStyle.wdStyleHeading2);
                    break;
                case "一.":
                    ChosingHeading(regexH1_8, WdBuiltinStyle.wdStyleHeading2);
                    break;
                default:
                    break;
            }

            switch (CbxHeading1.Text)
            {
                case "":
                    break;
                case "1、":
                    ChosingHeading(regexH1_10, WdBuiltinStyle.wdStyleHeading1);
                    break;
                case "1":
                    ChosingHeading(regexH1, WdBuiltinStyle.wdStyleHeading1);
                    break;
                case "1.":
                    ChosingHeading(regexH1_9, WdBuiltinStyle.wdStyleHeading1);
                    break;
                case "第x章":
                    ChosingHeading(regexH1_2, WdBuiltinStyle.wdStyleHeading1);
                    break;
                case "第x节":
                    ChosingHeading(regexH1_3, WdBuiltinStyle.wdStyleHeading1);
                    break;
                case "(1)":
                    ChosingHeading(regexH1_4, WdBuiltinStyle.wdStyleHeading1);
                    break;
                case "（1）":
                    ChosingHeading(regexH1_5, WdBuiltinStyle.wdStyleHeading1);
                    break;
                case "1)":
                    ChosingHeading(regexH1_6, WdBuiltinStyle.wdStyleHeading1);
                    break;
                case "一、":
                    ChosingHeading(regexH1_7, WdBuiltinStyle.wdStyleHeading1);
                    break;
                case "一.":
                    ChosingHeading(regexH1_8, WdBuiltinStyle.wdStyleHeading1);
                    break;
                default:
                    break;
            }

            this.Close();
        }
        private void findHeading(string pattern, WdBuiltinStyle wdBuiltinStyle)//设置系统样式
        {
            foreach (Paragraph para in _document.Paragraphs)
            {
                Range selectionRange = para.Range;
                Find find = selectionRange.Find;

                find.ClearFormatting();
                find.MatchWildcards = true;
                find.Text = pattern;
                int end = selectionRange.End;

                while (find.Execute(Forward: true) && selectionRange.InRange(_document.Content))
                {
                    if (end <= selectionRange.End) break;

                    // 匹配到的文本
                    string symbolText = selectionRange.Text;
                    if (ChbxRmIndex.Checked)
                    {
                        symbolText = "";

                    }
                    para.set_Style(wdBuiltinStyle);
                    // 替换后的文本赋值回 Range.Text
                    selectionRange.Text = symbolText;
                    selectionRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                }
            }


        }


        private void ChosingHeading(Regex regex, WdBuiltinStyle wdBuiltinStyle)//设置系统样式
        {
            foreach (Section section in _document.Sections)
            {
                foreach (Paragraph para in section.Range.Paragraphs)
                {
                    if (para.Range.Information[WdInformation.wdWithInTable]) continue;
                    TitleConfig(para, wdBuiltinStyle, regex);
                }
            }
        }
        private void ChosingHeading(Regex regex, string wdBuiltinStyle)//设置自定义样式
        {
            foreach (Section section in _document.Sections)
            {
                foreach (Paragraph para in section.Range.Paragraphs)
                {
                    if (para.Range.Information[WdInformation.wdWithInTable]) continue;
                    TitleConfig(para, wdBuiltinStyle, regex);
                }
            }
        }
        private void TitleConfig(Paragraph para, WdBuiltinStyle wdBuiltinStyle, Regex regex)//设置段落样式
        {
            if (regex.IsMatch(para.Range.Text))
            {
                if (para.Range.Fields.Count == 0)
                {
                    if (ChbxRmIndex.Checked)
                    {
                        string text = regex.Replace(para.Range.Text, "");
                        para.Range.Text = text;
                        para = para.Previous();

                    }
                    para.Range.set_Style(wdBuiltinStyle);
                }
                else
                {
                    foreach (Field field in para.Range.Fields)
                    {
                        // 检查域代码是否包含目录代码
                        if (field.Code.Text.Contains("TOC"))
                        { continue; }
                        else
                        {// 设置为标题1样式
                            if (ChbxRmIndex.Checked)
                            {
                                string text = regex.Replace(para.Range.Text, "");
                                para.Range.Text = text;
                                para = para.Previous();
                            }
                            para.Range.set_Style(wdBuiltinStyle);
                        }
                    }
                }
            }
        }
        private void TitleConfig(Paragraph para, string styleName, Regex regex)//设置自定义段落样式
        {
            if (regex.IsMatch(para.Range.Text))
            {
                // 遍历段落中的每一个匹配到的文本


                //Debug.WriteLine();
                if (para.Range.Fields.Count == 0)
                {
                    foreach (Match match in regex.Matches(para.Range.Text))
                    {
                        Range matchRange = para.Range;
                        matchRange.Start += match.Index;
                        matchRange.End = matchRange.Start + match.Length;
                        matchRange.Select();
                        matchRange.set_Style(styleName);
                        // 对选中的文本进行样式设置
                        if (styleName == "摘要") matchRange.Font.Bold = 1;
                        // 取消选中状态
                        para.Range.Select();
                    }
                }
                else
                {
                    foreach (Field field in para.Range.Fields)
                    {
                        // 检查域代码是否包含目录代码
                        if (field.Code.Text.Contains("TOC"))
                        { continue; }
                        else
                        {
                            foreach (Match match in regex.Matches(para.Range.Text))
                            {
                                Range matchRange = para.Range;
                                matchRange.Start += match.Index;
                                matchRange.End = matchRange.Start + match.Length;
                                matchRange.Select();
                                matchRange.set_Style(styleName);
                                // 对选中的文本进行样式设置
                                if (styleName == "摘要") matchRange.Font.Bold = 1;
                                // 取消选中状态
                                para.Range.Select();
                            }
                        }
                    }
                }
            }
        }
        private void BtnTitileConfigForm_Load(object sender, EventArgs e)//加载
        {
            List<string> list = new List<string>
            {
                "",
                "1",
                "1.",
                "1、",
                "第x章",
                "第x节",
                "(1)",
                "（1）",
                "1)",
                "一、",
                "一."
            };

            List<string> list1 = list.ToList();
            CbxHeading1.DataSource = list1;

            List<string> list2 = list.ToList();
            list2.Insert(1, "1.x");
            CbxHeading2.DataSource = list2;

            List<string> list3 = list.ToList();
            list3.Insert(1, "1.x.x");
            CbxHeading3.DataSource = list3;

            List<string> list4 = list.ToList();
            list4.Insert(1, "1.x.x.x");
            CbxHeading4.DataSource = list4;

            List<string> list5 = list.ToList();
            list5.Insert(1, "1.x.x.x.x");
            CbxHeading5.DataSource = list5;

            List<string> list6 = new List<string> {
                "",
                "摘  要",
                "摘要,abstract,关键字,关键词,keywords"
            };
            CbxZhaiyao.DataSource = list6;

            List<string> list7 = new List<string> {
                "",
                "参考文献",
            };
            CbxCankao.DataSource = list7;

        }

        private void BtnTitileConfigForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Dispose();
        }
    }
}
