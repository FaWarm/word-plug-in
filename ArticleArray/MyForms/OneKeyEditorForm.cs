using ArticleArray.Functions;
using ArticleArray.Models;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using Task = System.Threading.Tasks.Task;

namespace ArticleArray.MyForms
{
    public partial class OneKeyEditorForm : Form
    {
        private string _folderPath;
        private string _settingPath;
        private Document _document = Globals.ThisAddIn.Application.ActiveDocument;
        public OneKeyEditorForm(string folderPath)
        {
            _folderPath = folderPath;
            _settingPath = folderPath + "settings";
            InitializeComponent();
        }

        #region 加载事件
        private void OneKeyEditorForm_Load(object sender, EventArgs e)
        {
            // 排版设置文件列表
            if (!Directory.Exists(_settingPath)) Directory.CreateDirectory(_settingPath);
            string[] settingPaths = Directory.GetFiles(_settingPath, "*.xml");
            List<string> settingNames = new List<string>();
            foreach (string filePath in settingPaths)
            {
                settingNames.Add(Path.GetFileName(filePath));
            }

            CbxSettingChose.DataSource = settingNames;
            //CbxSettingChose.SelectedIndex = 0;

            //样式模板文件列表
            if (!Directory.Exists(_folderPath)) Directory.CreateDirectory(_folderPath);
            string[] templatePaths = Directory.GetFiles(_folderPath, "*.dotx");//样式模板列表
            CbxChoseTemplate.DataSource = templatePaths;
            //CbxChoseTemplate.SelectedIndex = 0;

            //标题识别下拉列表
            #region 标题识别列表
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
            CbxAbstract.DataSource = list6;

            List<string> list7 = new List<string> {
                "",
                "参考文献",
            };
            CbxReference.DataSource = list7;
            #endregion

            //表格宽度下拉列表
            ChbxTableAutoFit.Items.Add("自适应");
            ChbxTableAutoFit.Items.Add("同页面宽");

            //表格样式下拉表
            ChbxThreeLineStyle.Items.Add("三线表样式1");
            ChbxThreeLineStyle.Items.Add("三线表样式2");

            if (ChbxParityHeader.Checked) { TbxOddHeader.Enabled = true; TbxEvenHeader.Enabled = true; }
            else { TbxOddHeader.Enabled = false; TbxEvenHeader.Enabled = false; }
        }
        #endregion

        private void BtnApply_Click(object sender, EventArgs e)
        {

            #region 正则
            Regex regexH1 = new Regex(@"^([ ]*\d+[\t ]*)");//一级标题 1
            Regex regexH1_2 = new Regex(@"^[ ]*第.{0,4}章[\t ]*");//第x章
            Regex regexH1_3 = new Regex(@"^[ ]*第.{0,4}节[\t ]*");//第x节
            Regex regexH1_4 = new Regex(@"^[ ]*\(.{0,4}\)[\t ]*");//(1)
            Regex regexH1_5 = new Regex(@"^[ ]*\（.{0,4}\）[\t ]*");//（1）
            Regex regexH1_6 = new Regex(@"^[ ]*\d{1,4}\)[\t ]*");//1)
            Regex regexH1_7 = new Regex(@"^[ ]*([一二三四五六七八九十]{1,3}、[\t ]*)");//一、
            Regex regexH1_8 = new Regex(@"^[ ]*([一二三四五六七八九十]{1,3}\.[\t ]*)");//一.
            Regex regexH1_9 = new Regex(@"^[ ]*(\d+\.[\t ]*)");//1.
            Regex regexH1_10 = new Regex(@"^[ ]*(\d+、[\t ]*)");//1、

            Regex regexH2 = new Regex(@"^([ ]*\d+\.\d+[\t ]*)");//二级标题1.x
            Regex regexH3 = new Regex(@"^([ ]*\d+\.\d+\.\d+[\t ]*)");//三级标题1.x.x
            Regex regexH4 = new Regex(@"^([ ]*\d+\.\d+\.\d+\.\d+[\t ]*)");//四级标题1.x.x.x
            Regex regexH5 = new Regex(@"^([ ]*\d+\.\d+\.\d+\.\d+\.\d+[\t ]*)");//五级标题1.x.x.x.x

            Regex regexHZy = new Regex(@"\b[ ]*摘[ ]*要\r");//摘  要
            //Regex regexHZy1 = new Regex(@"\b摘要[:：]+\s*");//摘要:
            Regex regexHZy2 = new Regex(@"(?i)\b[ ]*abstract[:：]+\s*");//摘要
            //Regex regexHZy3 = new Regex(@"\b关键字[:：]+\s*");//关键字
            //Regex regexHZy4 = new Regex(@"(?i)\bkeywords[:：]+\s*");//关键字
            Regex regexHZy5 = new Regex(@"(?i)\b(?:[ ]*摘要[:：]+|[ ]*摘\s+要[:：]+|[ ]*abstract[:：]+|[ ]*关键字[:：]+|[ ]*关键词[:：]+|[ ]*key[ ]*words[:：]+)\s*");//摘要和关键字
            Regex regexHCkwx = new Regex(@"\b[ ]*参考文献\b");//参考文献
            #endregion
            #region 标题识别
            switch (CbxAbstract.Text)
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

            switch (CbxReference.Text)
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
                    ChosingHeading(regexH3, WdBuiltinStyle.wdStyleHeading3);
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
            #endregion
            #region 图片
            if (ChbxPicCenter.Checked || TbxPicWidth.Text != "" || TbxPicWidth.Text != null || TbxPicHeight.Text != "" || TbxPicHeight.Text != null)
            {
                float width = float.Parse(TbxPicWidth.Text);
                float height = float.Parse(TbxPicHeight.Text);
                var isCenter = ChbxPicCenter.Checked;

                Range range = _document.Range();
                InlineShapes inlineShapes = range.InlineShapes;
                Word.ShapeRange shapeRange = range.ShapeRange;

                if (inlineShapes.Count > 0) { CommonMethods.SetInlineShapes(width, height, isCenter, inlineShapes); }
                if (shapeRange.Count > 0) { CommonMethods.SetLineShapes(width, height, isCenter, shapeRange); }
            }
            #endregion
            #region 表格
            switch (ChbxTableAutoFit.Text)
            {
                case "自适应":
                    foreach (Word.Table tbl in _document.Tables)
                    {
                        tbl.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
                        tbl.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
                    }
                    break;
                case "同页面宽":
                    // 获取文档页面宽度
                    float pageWidth = _document.PageSetup.PageWidth - _document.PageSetup.LeftMargin - _document.PageSetup.RightMargin;
                    List<Task> tasks = new List<Task>();
                    // 遍历文档中所有表格
                    foreach (Word.Table tbl in _document.Tables)
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
                    break;
                default: break;
            }

            switch (ChbxThreeLineStyle.Text)
            {
                case "三线表样式1":

                    CommonMethods.TableToThreeLineTableStyle1(_document.Tables);
                    break;
                case "三线表样式2":
                    CommonMethods.TableToThreeLineTableStyle2(_document.Tables);
                    break;
                default: break;
            }
            #endregion

            #region 页眉
            if (ChbxParityHeader.Checked)
            {

            }

            #endregion

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

        private void BtnSave_Click(object sender, EventArgs e)//保存设置
        {
            var values = new OneKeySettingModel();
            values.SettingChose = CbxSettingChose.Text;
            values.ChoseTemplate = CbxChoseTemplate.Text;
            values.Heading1 = CbxHeading1.Text;
            values.Heading2 = CbxHeading2.Text;
            values.Heading3 = CbxHeading3.Text;
            values.Heading4 = CbxHeading4.Text;
            values.Heading5 = CbxHeading5.Text;
            values.Abstract = CbxAbstract.Text;
            values.Reference = CbxReference.Text;
            values.PicCenter = ChbxPicCenter.Checked;
            values.PicWidth = TbxPicWidth.Text;
            values.PicHeight = TbxPicHeight.Text;
            values.TableAutoFit = ChbxTableAutoFit.Text;
            values.ThreeLineStyle = ChbxThreeLineStyle.Text;
            values.ParityHeader = ChbxParityHeader.Checked;
            values.OddHeader = TbxOddHeader.Text;
            values.EvenHeader = TbxEvenHeader.Text;
            values.Pagination = ChbxPagination.Checked;
            values.PaginationCenter = ChbxPaginationCenter.Checked;
            values.PaginationContinue = ChbxPaginationContinue.Checked;
            values.TableCaption = ChbxTableCaption.Checked;
            values.PicCaption = ChbxPicCaption.Checked;
            values.Seg = ChbxSeg.Checked;
            values.Normal1 = ChbxNormal1.Checked;
            values.RmSpace = ChbxRmSpace.Checked;

            //存为xml
            string filename;
            if (values.SettingChose.EndsWith(".xml")) filename = _settingPath + "/" + values.SettingChose;
            else filename = _settingPath + "/" + values.SettingChose + ".xml";
            SaveToXml(filename, values);
        }

        /// <summary>
        /// 写入xml
        /// </summary>
        /// <param name="fileName">文件名</param>
        /// <param name="Values">内容模型</param>
        public void SaveToXml(string fileName, OneKeySettingModel Values)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(OneKeySettingModel));
            // 创建一个文件流，用于将对象序列化为 XML 文件
            using (FileStream stream = new FileStream(fileName, FileMode.Create))
            {
                // 将对象序列化为 XML 文件
                serializer.Serialize(stream, Values);
            }
            MessageBox.Show("保存成功！", "提示");
        }
    }
}
