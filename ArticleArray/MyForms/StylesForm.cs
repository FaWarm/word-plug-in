using ArticleArray.Functions;
using Microsoft.Office.Interop.Word;
using System;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using DataTable = System.Data.DataTable;
using Word = Microsoft.Office.Interop.Word;

namespace ArticleArray.MyForms
{
    public partial class StylesForm : Form
    {
        #region 通用字段
        private Word.Application _app;
        private Word.Document _docUsing = null;//当前使用的文档
        private Word.Document _tempdoc = null;//临时文档
        private string folderPath = @"C:/WordStyleConfig"; //保存和打开默认文档的路径
        //private string[] styleNames = { "标题 1", "标题 2", "标题 3", "标题 4", "标题 5", "正文 1", "摘要", "题注", "目录" };
        private string[] styleNames;
        private ListTemplate _template;//多级列表缓存
        #endregion

        #region 加载页面
        public StylesForm(string[] styles)
        {
            InitializeComponent();
            styleNames = styles;
            dataGridView1.DefaultCellStyle.Font = new System.Drawing.Font("宋体", 12);

            CbxEnFontName.Items.Add("");
            CbxEnFontName.Items.Add("Times New Roman");

            var EnFontName = Globals.ThisAddIn.Application.ActiveDocument.Styles[WdBuiltinStyle.wdStyleNormal].Font.Name;
            // 设置默认字体
            if (EnFontName == "Times New Roman") CbxEnFontName.Text = "Times New Roman";


        }
        private void StylesForm_Load(object sender, EventArgs e)
        {
            CommonMethods.CreateStyles(styleNames: styleNames);
            _app = Globals.ThisAddIn.Application;
            _docUsing = _app.ActiveDocument;
            var styles = _docUsing.Styles;
            dataGridView1.Columns.Clear();
            DataGridView dataGridView = SetDataTable(styles);
            dataGridView1.DataSource = dataGridView.DataSource;
            SetDataGridViewStyle();
        }
        #endregion

        #region 应用样式
        private void BtnApply_Click(object sender, EventArgs e)//应用样式事件
        {
            ApplyStyles();
            Range range = _docUsing.Range();
            CommonMethods.RefreshTitle(range);//刷新选中区域标题样式
            MessageBox.Show("样式已成功应用到当前文档！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            _docUsing = Globals.ThisAddIn.Application.ActiveDocument;
            // 刷新数据
            ((DataTable)dataGridView1.DataSource).AcceptChanges();
            this.Close();
        }
        private void ApplyStyles()//应用样式
        {
            //CommonMethods.CreateStyles(styleNames: styleNames);

            var _docUsing = Globals.ThisAddIn.Application.ActiveDocument;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells["样式名称"].Value != null)
                {
                    string styleName = dataGridView1.Rows[i].Cells["样式名称"].Value.ToString();
                    Word.Style style;
                    try { style = _docUsing.Styles[styleName]; }
                    catch (Exception ex) { Debug.WriteLine(ex); style = _docUsing.Styles.Add(styleName); }

                    if (CbxEnFontName.Text == "Times New Roman") style.Font.Name = "Times New Roman";
                    //设置字体
                    style.Font.NameFarEast = dataGridView1.Rows[i].Cells["字体"].Value.ToString();
                    style.Font.Size = float.Parse(dataGridView1.Rows[i].Cells["字体大小"].Value.ToString());

                    // 获取下拉框的值，转换为Color类型，再转换为OLE颜色值
                    Color fontColor = (Color)dataGridView1.Rows[i].Cells["字体颜色"].Value;
                    int oleColor = ColorTranslator.ToOle(fontColor);
                    // 设置字体颜色
                    style.Font.Color = (Word.WdColor)oleColor;

                    //设置加粗、斜体、下划线
                    style.Font.Bold = (bool)dataGridView1.Rows[i].Cells["加粗"].Value ? -1 : 0;
                    style.Font.Italic = (bool)dataGridView1.Rows[i].Cells["斜体"].Value ? -1 : 0;
                    style.Font.Underline = ((bool)dataGridView1.Rows[i].Cells["下划线"].Value) ? WdUnderline.wdUnderlineSingle : WdUnderline.wdUnderlineNone;

                    //设置水平居中、首行缩进、行间距、段前间距、段后间距
                    style.ParagraphFormat.Alignment = ((bool)dataGridView1.Rows[i].Cells["水平居中"].Value) ? WdParagraphAlignment.wdAlignParagraphCenter : WdParagraphAlignment.wdAlignParagraphJustify;
                    if ((bool)dataGridView1.Rows[i].Cells["首行缩进"].Value) style.ParagraphFormat.CharacterUnitFirstLineIndent = 2;
                    else
                    {
                        style.ParagraphFormat.FirstLineIndent = 0f;
                        style.ParagraphFormat.CharacterUnitFirstLineIndent = 0f;
                        style.ParagraphFormat.LeftIndent = 0f;
                    }
                    style.ParagraphFormat.AddSpaceBetweenFarEastAndAlpha = 0;
                    style.ParagraphFormat.AddSpaceBetweenFarEastAndDigit = 0;
                    if (styleName.Contains("正文")) style.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;//非标题分散对齐
                    switch (float.Parse(dataGridView1.Rows[i].Cells["行间距"].Value.ToString()))
                    {
                        case 1:
                            style.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                            break;
                        case (float)1.5:
                            style.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpace1pt5;
                            break;
                        case 2:
                            style.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceDouble;
                            break;
                        default:
                            if (float.Parse(dataGridView1.Rows[i].Cells["行间距"].Value.ToString()) < 12)
                            {
                                style.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;
                                style.ParagraphFormat.LineSpacing = (float)(float.Parse(dataGridView1.Rows[i].Cells["行间距"].Value.ToString()) * 12);
                            }
                            else
                            {
                                style.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                                style.ParagraphFormat.LineSpacing = float.Parse(dataGridView1.Rows[i].Cells["行间距"].Value.ToString());
                            }
                            break;
                    }
                    style.ParagraphFormat.LineUnitBefore = float.Parse(dataGridView1.Rows[i].Cells["段前行距"].Value.ToString());
                    if ((float)dataGridView1.Rows[i].Cells["段前行距"].Value == 0) style.ParagraphFormat.SpaceBefore = 0;
                    style.ParagraphFormat.LineUnitAfter = float.Parse(dataGridView1.Rows[i].Cells["段后行距"].Value.ToString());
                    if ((float)dataGridView1.Rows[i].Cells["段后行距"].Value == 0) style.ParagraphFormat.SpaceAfter = 0;
                }
            }
        }
        private DataGridView SetDataTable(Word.Styles styles)//设置数据
        {
            DataGridView dataGridView = new DataGridView();
            DataTable dtStyles = new DataTable();
            dtStyles.Columns.Add("样式名称", typeof(string));
            dtStyles.Columns.Add("字体", typeof(string));
            dtStyles.Columns.Add("字体大小", typeof(float));
            dtStyles.Columns.Add("字体颜色", typeof(Color));
            dtStyles.Columns.Add("加粗", typeof(bool));
            dtStyles.Columns.Add("斜体", typeof(bool));
            dtStyles.Columns.Add("下划线", typeof(bool));

            dtStyles.Columns.Add("水平居中", typeof(bool));
            dtStyles.Columns.Add("首行缩进", typeof(bool));
            dtStyles.Columns.Add("行间距", typeof(float));
            //dtStyles.Columns.Add("左侧缩进", typeof(float));
            dtStyles.Columns.Add("段前行距", typeof(float));
            dtStyles.Columns.Add("段后行距", typeof(float));
            foreach (string styleName in styleNames)
            {
                Word.Style style;
                try { style = styles[styleName]; }
                catch (Exception) { style = styles.Add(styleName); }

                DataRow row = dtStyles.NewRow();
                row["样式名称"] = style.NameLocal;
                row["字体"] = style.Font.NameFarEast;
                row["字体大小"] = style.Font.Size;
                row["字体颜色"] = ColorTranslator.FromOle((int)style.Font.Color);
                row["加粗"] = style.Font.Bold;
                row["斜体"] = style.Font.Italic;
                row["下划线"] = style.Font.Underline == WdUnderline.wdUnderlineSingle;
                row["水平居中"] = style.ParagraphFormat.Alignment == Word.WdParagraphAlignment.wdAlignParagraphCenter;
                row["首行缩进"] = style.ParagraphFormat.CharacterUnitFirstLineIndent == 2;

                switch (style.ParagraphFormat.LineSpacingRule)
                {
                    case WdLineSpacing.wdLineSpaceSingle:
                        row["行间距"] = 1;
                        break;
                    case WdLineSpacing.wdLineSpace1pt5:
                        row["行间距"] = 1.5;
                        break;
                    case WdLineSpacing.wdLineSpaceDouble:
                        row["行间距"] = 2;
                        break;
                    case WdLineSpacing.wdLineSpaceExactly:
                        row["行间距"] = style.ParagraphFormat.LineSpacing;
                        break;
                    case WdLineSpacing.wdLineSpaceMultiple:
                        row["行间距"] = ((float)style.ParagraphFormat.LineSpacing / 12).ToString("F2");
                        break;
                    default:
                        row["行间距"] = style.ParagraphFormat.LineSpacing;
                        break;
                }
                row["段前行距"] = style.ParagraphFormat.LineUnitBefore;
                row["段后行距"] = style.ParagraphFormat.LineUnitAfter;
                dtStyles.Rows.Add(row);
            }
            dataGridView.DataSource = dtStyles;
            return dataGridView;
        }
        private void SetColor(int rowIndex)//设置颜色
        {
            var co = (Color)dataGridView1.Rows[rowIndex].Cells["字体颜色"].Value;
            ColorDialog colorDialog = new ColorDialog();
            //允许使用该对话框的自定义颜色
            colorDialog.Color = co;
            colorDialog.AllowFullOpen = true;
            colorDialog.FullOpen = true;
            colorDialog.ShowHelp = false;
            //初始化当前文本框中的字体颜色，
            //当用户在ColorDialog对话框中点击"取消"按钮
            if (colorDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                dataGridView1.Rows[rowIndex].Cells["字体颜色"].Value = colorDialog.Color;
                dataGridView1.Rows[rowIndex].Cells["字体颜色"].Style.BackColor = colorDialog.Color;
                dataGridView1.Rows[rowIndex].Cells["字体颜色"].Style.ForeColor = colorDialog.Color;
            }
        }
        private void SetFont(int rowIndex)//设置字体
        {
            var name = dataGridView1.Rows[rowIndex].Cells["字体"].Value.ToString();
            if (name.Contains(" Light")) { name = name.Replace(" Light", ""); }

            var size = float.Parse(dataGridView1.Rows[rowIndex].Cells["字体大小"].Value.ToString());
            var bold = bool.Parse(dataGridView1.Rows[rowIndex].Cells["加粗"].Value.ToString());
            var italic = bool.Parse(dataGridView1.Rows[rowIndex].Cells["斜体"].Value.ToString());
            var underline = bool.Parse(dataGridView1.Rows[rowIndex].Cells["下划线"].Value.ToString());
            System.Drawing.Font newFont;
            if (bold)
            {
                if (italic)
                {
                    if (underline) newFont = new System.Drawing.Font(name, size, FontStyle.Italic | FontStyle.Bold | FontStyle.Underline);
                    else newFont = new System.Drawing.Font(name, size, FontStyle.Italic | FontStyle.Bold);
                }
                else
                {
                    if (underline) newFont = new System.Drawing.Font(name, size, FontStyle.Bold | FontStyle.Underline);
                    else newFont = new System.Drawing.Font(name, size, FontStyle.Bold);
                }
            }
            else
            {
                if (italic)
                {
                    if (underline) newFont = new System.Drawing.Font(name, size, FontStyle.Italic | FontStyle.Underline);
                    else newFont = new System.Drawing.Font(name, size, FontStyle.Italic);
                }
                else
                {
                    if (underline) newFont = new System.Drawing.Font(name, size, FontStyle.Underline);
                    else newFont = new System.Drawing.Font(name, size, FontStyle.Regular);
                }
            }
            FontDialog fontDialog = new FontDialog();
            fontDialog.Font = newFont;
            if (fontDialog.ShowDialog() == DialogResult.OK)
            {
                newFont = fontDialog.Font;
                dataGridView1.Rows[rowIndex].Cells["字体大小"].Value = newFont.Size;
                dataGridView1.Rows[rowIndex].Cells["加粗"].Value = newFont.Bold;
                dataGridView1.Rows[rowIndex].Cells["斜体"].Value = newFont.Italic;
                dataGridView1.Rows[rowIndex].Cells["下划线"].Value = newFont.Underline;
                dataGridView1.Rows[rowIndex].Cells["字体"].Value = newFont.Name;
            }
        }
        private void SetDataGridViewStyle()//设置datagridview样式
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells["字体颜色"].Value != null)
                {
                    dataGridView1.Rows[i].Cells["字体颜色"].Style.BackColor = (Color)dataGridView1.Rows[i].Cells["字体颜色"].Value;
                    dataGridView1.Rows[i].Cells["字体颜色"].Style.ForeColor = (Color)dataGridView1.Rows[i].Cells["字体颜色"].Value;
                }
            }
            dataGridView1.AutoSize = true;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }
        private void StylesFormReLoad(Word.Styles styles)//再加载
        {
            DataGridView dataGridView = SetDataTable(styles);
            dataGridView1.Columns.Clear();
            dataGridView1.DataSource = dataGridView.DataSource;
            SetDataGridViewStyle();
        }
        #endregion

        #region 保存样式

        private void BtnSaveStyles_Click(object sender, EventArgs e)//保存样式配置
        {
            try
            {
                var doc = Globals.ThisAddIn.Application.ActiveDocument;
                doc.Save();
                var oldName = doc.FullName;
                if (oldName == null) return;
                if (!Directory.Exists(folderPath)) Directory.CreateDirectory(folderPath);
                // 设置默认保存地址
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Title = "保存样式配置";
                saveFileDialog.Filter = "Word template(模板)|*.dotx|Word document|*.docx";
                saveFileDialog.FilterIndex = 1;
                saveFileDialog.FileName = "我的配置1";
                //saveFileDialog.DefaultExt = "docx";
                saveFileDialog.InitialDirectory = folderPath;
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // 将文档保存到用户选择的路径
                    string filePath = saveFileDialog.FileName;
                    doc.SaveAs2(filePath, Word.WdSaveFormat.wdFormatXMLTemplate);
                    // 移除中间的内容
                    Word.Range contentRange = doc.Range(doc.Content.Start, doc.Content.End);
                    contentRange.Delete();
                    doc.Close(SaveChanges: true);
                    doc = null;
                    Globals.ThisAddIn.Application.Documents.Open(oldName);
                }
            }
            catch (Exception) { }
        }
        #endregion

        #region 表格事件
        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e) //单元个双击更改 字体
        {
            if (e.ColumnIndex == 3) SetColor(e.RowIndex);

            if (e.ColumnIndex == 2 || e.ColumnIndex == 1) SetFont(e.RowIndex);

        }
        private void dataGridView1_BindingContextChanged(object sender, EventArgs e)//数据刷新时，刷新样式设置
        {
            SetDataGridViewStyle();
        }
        #endregion


        #region 获取样式
        private void BtnReadStyles_Click(object sender, EventArgs e)//读取样式
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = folderPath;
            openFileDialog.Filter = "Word Template (模板)|*.dotx|Word Documents (*.docx)|*.docx|Word Documents (*.doc)|*.doc";
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var filename = openFileDialog.FileName;
                if (filename.EndsWith(".docx") || filename.EndsWith(".doc"))
                {
                    Word.Styles styles = GetDocStyles(filename);
                    if (styles != null) { StylesFormReLoad(styles); }
                    if (_tempdoc != null) { _tempdoc.Close(false); }
                }
                else
                {//打开
                    var doc = Globals.ThisAddIn.Application.ActiveDocument;
                    var newPath = doc.Path + "\\(新样式)" + doc.Name;
                    doc.Content.Copy();
                    Word.Document newDoc = Globals.ThisAddIn.Application.Documents.Open(filename);
                    newDoc.Range().Paste();
                    newDoc.SaveAs2(newPath, Word.WdSaveFormat.wdFormatXMLDocument);
                    doc.Close(SaveChanges: true);
                    Clipboard.Clear();

                    CommonMethods.AutoSetNormal1(newDoc);

                    this.Close();
                }
            }
        }
        private Styles GetDocStyles(string filename)//获取样式
        {
            try
            {
                Styles styles = null;
                _tempdoc = _app.Documents.Open(filename, ReadOnly: true);
                styles = _tempdoc.Styles;
                _template = _tempdoc.Range().ListFormat.ListTemplate;
                return styles;
            }
            catch (Exception ex)
            {
                MessageBox.Show("错误: " + ex.Message);
                return null;
            }
        }

        #endregion

        private void BtnSetDefaultStyles_Click(object sender, EventArgs e)//设置默认样式
        {
            try
            {
                var doc = Globals.ThisAddIn.Application.ActiveDocument;
                doc.Save();
                var oldName = doc.FullName;
                if (oldName == null) return;
                // 移除中间的内容
                if (!Directory.Exists(folderPath)) Directory.CreateDirectory(folderPath);

                var filePath = folderPath + "/DefaultStyles.dotx";

                doc.SaveAs2(filePath, Word.WdSaveFormat.wdFormatXMLTemplate);
                Word.Range contentRange = doc.Range(doc.Content.Start, doc.Content.End);
                contentRange.Delete();
                doc.Close(SaveChanges: true);

                Globals.ThisAddIn.Application.Documents.Open(oldName);
                this.Close();
                MessageBox.Show("默认样式设置成功！");
            }
            catch (Exception) { }


        }

        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)//数据绑定后事件
        {
            foreach (DataGridViewColumn column in dataGridView1.Columns)//禁用排序
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        private void StylesForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Dispose();//释放
        }

        private Word.Document GetDocument(string filename)//判断文档是否已经打开
        {
            Word.Application app = Globals.ThisAddIn.Application;

            Word.Document doc = null;
            try
            {
                // 遍历当前已打开的文档，寻找指定名称的文档
                foreach (Word.Document d in app.Documents)
                {
                    if (d.FullName == filename)
                    {
                        // 如果找到，则直接返回已打开的文档
                        doc = d;
                        break;
                    }
                }
                // 如果没有找到，则打开指定名称的文档
                if (doc == null) doc = app.Documents.Open(filename, ReadOnly: false, Visible: true);
                return doc;
            }
            catch (Exception) { return null; }
        }
    }
}
