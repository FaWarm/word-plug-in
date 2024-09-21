using ArticleArray.Functions;
using Microsoft.Office.Interop.Word;
using System;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace ArticleArray.MyForms
{
    public partial class InsertCaptionForm : Form
    {
        private string[] _styleNames;
        private bool useChapter = false;
        private CaptionLabel picLabel;
        private CaptionLabel tableLabel;
        public InsertCaptionForm(string[] styleNames)
        {
            InitializeComponent();
            var _document = Globals.ThisAddIn.Application.ActiveDocument;

            _styleNames = styleNames;
            ChbxChapter.Checked = false;

            CbxChapterStyle.DataSource = _styleNames;//章节起始样式列表，默认标题 1
            string[] Separators = new string[] { ".", "-", ":" };
            CbxSeparator.DataSource = Separators;

            //插入位置及初始化
            string[] TableInsertPosition = new string[] { "上侧", "下侧" };
            CbxTableInsertPosition.DataSource = TableInsertPosition;
            CbxTableInsertPosition.SelectedIndex = 0;

            string[] PictureInsertPosition = new string[] { "上侧", "下侧" };
            CbxPictureInsertPosition.DataSource = PictureInsertPosition;
            CbxPictureInsertPosition.SelectedIndex = 1;

            //前缀
            TbxPicturePrefix.Text = "图";
            TbxTablePrefix.Text = "表";
            //默认样式
            string[] pictureStyles = new string[] { "题注", "图" };
            string[] tableStyles = new string[] { "题注", "表" };
            CbxPictureStyle.DataSource = pictureStyles;
            CbxPictureStyle.SelectedIndex = 0;

            CbxTableStyle.DataSource = tableStyles;
            CbxTableStyle.SelectedIndex = 0;

            ChbxPictureUse.Checked = true;
            ChbxTableUse.Checked = true;

            //添加标签
            CaptionLabels captionLabels = _document.InlineShapes.Application.CaptionLabels;//获取标签
            picLabel = captionLabels.Add(TbxPicturePrefix.Text);
            tableLabel = captionLabels.Add(TbxTablePrefix.Text);
        }

        private void InsertCaptionForm_Load(object sender, EventArgs e)//加载事件
        {
            ChbxChapter.Checked = useChapter;
        }

        private void BtnInsertCaption_Click(object sender, EventArgs e)//插入题注
        {
            var _document = Globals.ThisAddIn.Application.ActiveDocument;

            if (ChbxPictureUse.Checked) SetCaption(picLabel, CbxPictureInsertPosition);
            if (ChbxTableUse.Checked) SetCaption(tableLabel, CbxTableInsertPosition);

            foreach (InlineShape inlineShape in _document.InlineShapes.Cast<InlineShape>().Where(s => s.Type == WdInlineShapeType.wdInlineShapePicture))
            {
                Word.Range rng = inlineShape.Range;
                Word.Paragraph para = rng.Paragraphs.First.Next();
                if (rng.Paragraphs.First.Next().Range.Fields.Count == 0)
                {// 插入新题注
                    string text = para.Range.Text;
                    Match match = Regex.Match(text, @"^图[ ]*(\d+)[ ]+");
                    string remainingText = " ";
                    if (match.Success)
                    {
                        Debug.WriteLine(match.Value);
                        string number = match.Groups[1].Value; // 获取捕获组中的数字
                        remainingText += text.Substring(match.Length); // 获取图+数字+空格后面的文本
                        para.Range.Delete();
                    }
                    inlineShape.Range.InsertCaption(TbxPicturePrefix.Text, Title: remainingText);
                }
                else
                {// 刷新题注
                    para.Range.Fields.Update();
                }
                para = rng.Paragraphs.First.Next();
                CommonMethods.ReplaceCharacters(para.Range, TbxPicturePrefix.Text + " ", TbxPicturePrefix.Text);//移除图后的空格
                para.set_Style(CbxPictureStyle.Text);
            }

            // 遍历文档中的所有表格
            foreach (Table table in _document.Tables)
            {
                Word.Range rng = table.Range;
                Word.Paragraph para = rng.Paragraphs.First.Previous();
                if (rng.Paragraphs.First.Previous().Range.Fields.Count == 0)
                {// 插入新题注
                    string text = para.Range.Text;
                    Match match = Regex.Match(text, @"^表[ ]*(\d+)[ ]+");
                    string remainingText = " ";
                    if (match.Success)
                    {
                        string number = match.Groups[1].Value; // 获取捕获组中的数字
                        remainingText += text.Substring(match.Length).Trim(); // 获取图+数字+空格后面的文本
                        para.Range.Delete();
                    }
                    table.Range.InsertCaption(TbxTablePrefix.Text, Title: remainingText);
                }
                else
                {// 刷新题注
                    rng.Paragraphs.First.Previous().Range.Fields.Update();
                }
                para = rng.Paragraphs.First.Previous();
                CommonMethods.ReplaceCharacters(para.Range, TbxTablePrefix.Text + " ", TbxTablePrefix.Text);
                para.set_Style(CbxTableStyle.Text);
            }
            _document.Range().Fields.Update();

            useChapter = ChbxChapter.Checked;
            this.Close();

        }

        private void SetCaption(CaptionLabel Selectedlabel, ComboBox CbxInsertPosition)//设置内容
        {
            try
            {
                if (ChbxChapter.Checked)//包含章节号
                {
                    Selectedlabel.IncludeChapterNumber = true;
                    switch (CbxChapterStyle.SelectedIndex)//起始样式
                    {
                        case 0:
                            Selectedlabel.ChapterStyleLevel = 1;
                            break;
                        case 1:
                            Selectedlabel.ChapterStyleLevel = 2;
                            break;
                        case 2:
                            Selectedlabel.ChapterStyleLevel = 3;
                            break;
                        default:
                            Selectedlabel.ChapterStyleLevel = 1;
                            break;
                    }
                    switch (CbxSeparator.SelectedItem)//分隔符
                    {
                        case ".":
                            Selectedlabel.Separator = WdSeparatorType.wdSeparatorPeriod;
                            break;
                        case ":":
                            Selectedlabel.Separator = WdSeparatorType.wdSeparatorColon;
                            break;
                        case "-":
                            Selectedlabel.Separator = WdSeparatorType.wdSeparatorHyphen;
                            break;
                        default:
                            Selectedlabel.Separator = WdSeparatorType.wdSeparatorColon;
                            break;
                    }
                }
                else
                {
                    Selectedlabel.IncludeChapterNumber = false;
                }

                switch (CbxInsertPosition.SelectedIndex)//图片插入位置
                {
                    case 0:
                        Selectedlabel.Position = WdCaptionPosition.wdCaptionPositionAbove;
                        break;
                    case 1:
                        Selectedlabel.Position = WdCaptionPosition.wdCaptionPositionBelow;
                        break;
                    default:
                        Selectedlabel.Position = WdCaptionPosition.wdCaptionPositionAbove;
                        break;
                }
            }
            catch (Exception)
            {
                return;
            }
        }
    }
}

