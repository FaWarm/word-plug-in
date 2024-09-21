namespace ArticleFormatApp_USST
{
    partial class MainRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MainRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.tab2 = this.Factory.CreateRibbonTab();
            this.group6 = this.Factory.CreateRibbonGroup();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.group8 = this.Factory.CreateRibbonGroup();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.lbl_formatNote = this.Factory.CreateRibbonLabel();
            this.btn_AbstractHeading = this.Factory.CreateRibbonButton();
            this.btn_CatalogHeading = this.Factory.CreateRibbonButton();
            this.btn_FirstHeading = this.Factory.CreateRibbonButton();
            this.btn_SecondHeading = this.Factory.CreateRibbonButton();
            this.btn_ThirdHeading = this.Factory.CreateRibbonButton();
            this.btn_FourthHeading = this.Factory.CreateRibbonButton();
            this.btn_OtherHeading = this.Factory.CreateRibbonButton();
            this.btn_PictureHeading = this.Factory.CreateRibbonButton();
            this.btn_TableHeading = this.Factory.CreateRibbonButton();
            this.btn_ParaBody = this.Factory.CreateRibbonButton();
            this.btn_ReferenceBody = this.Factory.CreateRibbonButton();
            this.btn_AbstractBody = this.Factory.CreateRibbonButton();
            this.btn_Keywords = this.Factory.CreateRibbonButton();
            this.btn_AcknowledgementBody = this.Factory.CreateRibbonButton();
            this.btn_AppendixBody = this.Factory.CreateRibbonButton();
            this.btn_AchievementsBody = this.Factory.CreateRibbonButton();
            this.btn_PictureBody = this.Factory.CreateRibbonButton();
            this.btn_tableThreeline = this.Factory.CreateRibbonButton();
            this.btn_TableBody = this.Factory.CreateRibbonButton();
            this.btn_tableNote = this.Factory.CreateRibbonButton();
            this.btn_InsertCaption = this.Factory.CreateRibbonButton();
            this.btn_formula = this.Factory.CreateRibbonButton();
            this.btn_ContainsFormula = this.Factory.CreateRibbonButton();
            this.btn_InsertFormulaCaption = this.Factory.CreateRibbonButton();
            this.btn_quote = this.Factory.CreateRibbonButton();
            this.btn_SeparateSection = this.Factory.CreateRibbonButton();
            this.btn_RefrushIndex = this.Factory.CreateRibbonButton();
            this.btn_SepOdd = this.Factory.CreateRibbonButton();
            this.btn_SetHeaderAndFooter = this.Factory.CreateRibbonButton();
            this.btn_PageSetting = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btn_CreateStyles = this.Factory.CreateRibbonSplitButton();
            this.btn_CreateFormulaStyle1 = this.Factory.CreateRibbonButton();
            this.btn_CreatePicTableStyle = this.Factory.CreateRibbonButton();
            this.btn_ListToText = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.tab2.SuspendLayout();
            this.group6.SuspendLayout();
            this.group2.SuspendLayout();
            this.group4.SuspendLayout();
            this.group8.SuspendLayout();
            this.group1.SuspendLayout();
            this.group5.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "上理研插件";
            this.tab1.Name = "tab1";
            // 
            // tab2
            // 
            this.tab2.Groups.Add(this.group6);
            this.tab2.Groups.Add(this.group2);
            this.tab2.Groups.Add(this.group4);
            this.tab2.Groups.Add(this.group8);
            this.tab2.Groups.Add(this.group1);
            this.tab2.Groups.Add(this.group5);
            this.tab2.Groups.Add(this.group3);
            this.tab2.Label = "论文插件";
            this.tab2.Name = "tab2";
            // 
            // group6
            // 
            this.group6.Items.Add(this.btn_CreateStyles);
            this.group6.Label = "生成样式";
            this.group6.Name = "group6";
            // 
            // group2
            // 
            this.group2.Items.Add(this.btn_AbstractHeading);
            this.group2.Items.Add(this.btn_CatalogHeading);
            this.group2.Items.Add(this.btn_FirstHeading);
            this.group2.Items.Add(this.btn_SecondHeading);
            this.group2.Items.Add(this.btn_ThirdHeading);
            this.group2.Items.Add(this.btn_FourthHeading);
            this.group2.Items.Add(this.btn_OtherHeading);
            this.group2.Items.Add(this.btn_PictureHeading);
            this.group2.Items.Add(this.btn_TableHeading);
            this.group2.Label = "标题";
            this.group2.Name = "group2";
            // 
            // group4
            // 
            this.group4.Items.Add(this.btn_ParaBody);
            this.group4.Items.Add(this.btn_ReferenceBody);
            this.group4.Items.Add(this.btn_AbstractBody);
            this.group4.Items.Add(this.btn_Keywords);
            this.group4.Items.Add(this.btn_AcknowledgementBody);
            this.group4.Items.Add(this.btn_AppendixBody);
            this.group4.Items.Add(this.btn_AchievementsBody);
            this.group4.Label = "正文";
            this.group4.Name = "group4";
            // 
            // group8
            // 
            this.group8.Items.Add(this.btn_PictureBody);
            this.group8.Items.Add(this.btn_tableThreeline);
            this.group8.Items.Add(this.btn_TableBody);
            this.group8.Items.Add(this.btn_tableNote);
            this.group8.Items.Add(this.btn_InsertCaption);
            this.group8.Label = "图表";
            this.group8.Name = "group8";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btn_formula);
            this.group1.Items.Add(this.btn_ContainsFormula);
            this.group1.Items.Add(this.btn_InsertFormulaCaption);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.lbl_formatNote);
            this.group1.Label = "公式";
            this.group1.Name = "group1";
            // 
            // group5
            // 
            this.group5.Items.Add(this.btn_quote);
            this.group5.Items.Add(this.btn_SeparateSection);
            this.group5.Items.Add(this.btn_RefrushIndex);
            this.group5.Items.Add(this.btn_SepOdd);
            this.group5.Items.Add(this.btn_ListToText);
            this.group5.Label = "其他";
            this.group5.Name = "group5";
            // 
            // group3
            // 
            this.group3.Items.Add(this.btn_SetHeaderAndFooter);
            this.group3.Items.Add(this.btn_PageSetting);
            this.group3.Label = "格式";
            this.group3.Name = "group3";
            // 
            // lbl_formatNote
            // 
            this.lbl_formatNote.Label = "对齐请在公式前按【tab】";
            this.lbl_formatNote.Name = "lbl_formatNote";
            // 
            // btn_AbstractHeading
            // 
            this.btn_AbstractHeading.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_AbstractHeading.Image = global::ArticleFormatApp_USST.Properties.Resources.zhaiyao;
            this.btn_AbstractHeading.Label = "摘要";
            this.btn_AbstractHeading.Name = "btn_AbstractHeading";
            this.btn_AbstractHeading.ShowImage = true;
            this.btn_AbstractHeading.SuperTip = "不显示在目录中";
            this.btn_AbstractHeading.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_AbstractHeading_Click);
            // 
            // btn_CatalogHeading
            // 
            this.btn_CatalogHeading.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_CatalogHeading.Image = global::ArticleFormatApp_USST.Properties.Resources.mulu;
            this.btn_CatalogHeading.Label = "目录";
            this.btn_CatalogHeading.Name = "btn_CatalogHeading";
            this.btn_CatalogHeading.ShowImage = true;
            this.btn_CatalogHeading.SuperTip = "不显示在目录中";
            this.btn_CatalogHeading.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_CatalogHeading_Click);
            // 
            // btn_FirstHeading
            // 
            this.btn_FirstHeading.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_FirstHeading.Image = global::ArticleFormatApp_USST.Properties.Resources.yijibiaoti;
            this.btn_FirstHeading.Label = "正文1级标题";
            this.btn_FirstHeading.Name = "btn_FirstHeading";
            this.btn_FirstHeading.ShowImage = true;
            this.btn_FirstHeading.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_FirstHeading_Click);
            // 
            // btn_SecondHeading
            // 
            this.btn_SecondHeading.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_SecondHeading.Image = global::ArticleFormatApp_USST.Properties.Resources.erjibiaoti;
            this.btn_SecondHeading.Label = "2级标题";
            this.btn_SecondHeading.Name = "btn_SecondHeading";
            this.btn_SecondHeading.ShowImage = true;
            this.btn_SecondHeading.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_SecondHeading_Click);
            // 
            // btn_ThirdHeading
            // 
            this.btn_ThirdHeading.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_ThirdHeading.Image = global::ArticleFormatApp_USST.Properties.Resources.sanjibiaoti;
            this.btn_ThirdHeading.Label = "3级标题";
            this.btn_ThirdHeading.Name = "btn_ThirdHeading";
            this.btn_ThirdHeading.ShowImage = true;
            this.btn_ThirdHeading.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_ThirdHeading_Click);
            // 
            // btn_FourthHeading
            // 
            this.btn_FourthHeading.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_FourthHeading.Image = global::ArticleFormatApp_USST.Properties.Resources.sijibiaoti;
            this.btn_FourthHeading.Label = "4级标题";
            this.btn_FourthHeading.Name = "btn_FourthHeading";
            this.btn_FourthHeading.ShowImage = true;
            this.btn_FourthHeading.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_FourthHeading_Click);
            // 
            // btn_OtherHeading
            // 
            this.btn_OtherHeading.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_OtherHeading.Image = global::ArticleFormatApp_USST.Properties.Resources.qitabiaoti;
            this.btn_OtherHeading.Label = "其他1级标题";
            this.btn_OtherHeading.Name = "btn_OtherHeading";
            this.btn_OtherHeading.ShowImage = true;
            this.btn_OtherHeading.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_OtherHeading_Click);
            // 
            // btn_PictureHeading
            // 
            this.btn_PictureHeading.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_PictureHeading.Image = global::ArticleFormatApp_USST.Properties.Resources.tupianzhengwen;
            this.btn_PictureHeading.Label = "图片题注";
            this.btn_PictureHeading.Name = "btn_PictureHeading";
            this.btn_PictureHeading.ShowImage = true;
            this.btn_PictureHeading.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_PictureHeading_Click);
            // 
            // btn_TableHeading
            // 
            this.btn_TableHeading.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_TableHeading.Image = global::ArticleFormatApp_USST.Properties.Resources.biaoge;
            this.btn_TableHeading.Label = "表格题注";
            this.btn_TableHeading.Name = "btn_TableHeading";
            this.btn_TableHeading.ShowImage = true;
            this.btn_TableHeading.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_TableHeading_Click);
            // 
            // btn_ParaBody
            // 
            this.btn_ParaBody.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_ParaBody.Image = global::ArticleFormatApp_USST.Properties.Resources.zhengwen;
            this.btn_ParaBody.Label = "段落正文";
            this.btn_ParaBody.Name = "btn_ParaBody";
            this.btn_ParaBody.ShowImage = true;
            this.btn_ParaBody.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_ParaBody_Click);
            // 
            // btn_ReferenceBody
            // 
            this.btn_ReferenceBody.Label = "参考文献列表";
            this.btn_ReferenceBody.Name = "btn_ReferenceBody";
            this.btn_ReferenceBody.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_ReferenceBody_Click);
            // 
            // btn_AbstractBody
            // 
            this.btn_AbstractBody.Label = "摘要正文";
            this.btn_AbstractBody.Name = "btn_AbstractBody";
            this.btn_AbstractBody.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_AbstractBody_Click);
            // 
            // btn_Keywords
            // 
            this.btn_Keywords.Label = "关键词";
            this.btn_Keywords.Name = "btn_Keywords";
            this.btn_Keywords.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Keywords_Click);
            // 
            // btn_AcknowledgementBody
            // 
            this.btn_AcknowledgementBody.Label = "致谢正文";
            this.btn_AcknowledgementBody.Name = "btn_AcknowledgementBody";
            this.btn_AcknowledgementBody.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_AcknowledgementBody_Click);
            // 
            // btn_AppendixBody
            // 
            this.btn_AppendixBody.Label = "附录正文";
            this.btn_AppendixBody.Name = "btn_AppendixBody";
            this.btn_AppendixBody.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_AppendixBody_Click);
            // 
            // btn_AchievementsBody
            // 
            this.btn_AchievementsBody.Label = "学术成果/作者正文";
            this.btn_AchievementsBody.Name = "btn_AchievementsBody";
            this.btn_AchievementsBody.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_AchievementsBody_Click);
            // 
            // btn_PictureBody
            // 
            this.btn_PictureBody.Label = "图片正文";
            this.btn_PictureBody.Name = "btn_PictureBody";
            this.btn_PictureBody.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_PictureBody_Click);
            // 
            // btn_tableThreeline
            // 
            this.btn_tableThreeline.Label = "表正文";
            this.btn_tableThreeline.Name = "btn_tableThreeline";
            this.btn_tableThreeline.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_TableBody_Click);
            // 
            // btn_TableBody
            // 
            this.btn_TableBody.Label = "三线表";
            this.btn_TableBody.Name = "btn_TableBody";
            this.btn_TableBody.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_tableThreeline_Click);
            // 
            // btn_tableNote
            // 
            this.btn_tableNote.Label = "表注释";
            this.btn_tableNote.Name = "btn_tableNote";
            this.btn_tableNote.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_tableNote_Click);
            // 
            // btn_InsertCaption
            // 
            this.btn_InsertCaption.Label = "插入图表题注";
            this.btn_InsertCaption.Name = "btn_InsertCaption";
            this.btn_InsertCaption.SuperTip = "详细样式可在【引用】-【插入题注】中设置";
            this.btn_InsertCaption.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_InsertCaption_Click);
            // 
            // btn_formula
            // 
            this.btn_formula.Label = "公式正文";
            this.btn_formula.Name = "btn_formula";
            this.btn_formula.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_formula_Click);
            // 
            // btn_ContainsFormula
            // 
            this.btn_ContainsFormula.Label = "带公式正文";
            this.btn_ContainsFormula.Name = "btn_ContainsFormula";
            this.btn_ContainsFormula.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_ContainsFormula_Click);
            // 
            // btn_InsertFormulaCaption
            // 
            this.btn_InsertFormulaCaption.Label = "插入公式编号";
            this.btn_InsertFormulaCaption.Name = "btn_InsertFormulaCaption";
            this.btn_InsertFormulaCaption.ScreenTip = "对齐请在公式前按【tab】";
            this.btn_InsertFormulaCaption.SuperTip = "对齐请在公式前按【tab】";
            this.btn_InsertFormulaCaption.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_InsertFormulaCaption_Click);
            // 
            // btn_quote
            // 
            this.btn_quote.Label = "交叉引用";
            this.btn_quote.Name = "btn_quote";
            this.btn_quote.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_quote_Click);
            // 
            // btn_SeparateSection
            // 
            this.btn_SeparateSection.Label = "分节/章";
            this.btn_SeparateSection.Name = "btn_SeparateSection";
            this.btn_SeparateSection.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_SeparateSection_Click);
            // 
            // btn_RefrushIndex
            // 
            this.btn_RefrushIndex.Label = "刷新序号";
            this.btn_RefrushIndex.Name = "btn_RefrushIndex";
            this.btn_RefrushIndex.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_RefrushIndex_Click);
            // 
            // btn_SepOdd
            // 
            this.btn_SepOdd.Label = "奇数页分节/章";
            this.btn_SepOdd.Name = "btn_SepOdd";
            this.btn_SepOdd.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_SepOdd_Click);
            // 
            // btn_SetHeaderAndFooter
            // 
            this.btn_SetHeaderAndFooter.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_SetHeaderAndFooter.Image = global::ArticleFormatApp_USST.Properties.Resources.yemeiyejiao;
            this.btn_SetHeaderAndFooter.Label = "设置页眉";
            this.btn_SetHeaderAndFooter.Name = "btn_SetHeaderAndFooter";
            this.btn_SetHeaderAndFooter.ShowImage = true;
            this.btn_SetHeaderAndFooter.SuperTip = "页眉会设置每一章/节的首行文本内容";
            this.btn_SetHeaderAndFooter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_SetHeaderAndFooter_Click);
            // 
            // btn_PageSetting
            // 
            this.btn_PageSetting.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_PageSetting.Label = "设置页面格式";
            this.btn_PageSetting.Name = "btn_PageSetting";
            this.btn_PageSetting.ShowImage = true;
            this.btn_PageSetting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_PageSetting_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // btn_CreateStyles
            // 
            this.btn_CreateStyles.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_CreateStyles.Image = global::ArticleFormatApp_USST.Properties.Resources.shengcheng;
            this.btn_CreateStyles.Items.Add(this.btn_CreateFormulaStyle1);
            this.btn_CreateStyles.Items.Add(this.btn_CreatePicTableStyle);
            this.btn_CreateStyles.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_CreateStyles.Label = "生成所有样式";
            this.btn_CreateStyles.Name = "btn_CreateStyles";
            this.btn_CreateStyles.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_CreateStyles_Click);
            // 
            // btn_CreateFormulaStyle1
            // 
            this.btn_CreateFormulaStyle1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_CreateFormulaStyle1.Label = "生成公式样式";
            this.btn_CreateFormulaStyle1.Name = "btn_CreateFormulaStyle1";
            this.btn_CreateFormulaStyle1.ShowImage = true;
            this.btn_CreateFormulaStyle1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_CreateFormulaStyle1_Click);
            // 
            // btn_CreatePicTableStyle
            // 
            this.btn_CreatePicTableStyle.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_CreatePicTableStyle.Label = "生成图表样式";
            this.btn_CreatePicTableStyle.Name = "btn_CreatePicTableStyle";
            this.btn_CreatePicTableStyle.ShowImage = true;
            this.btn_CreatePicTableStyle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_CreatePicTableStyle_Click);
            // 
            // btn_ListToText
            // 
            this.btn_ListToText.Label = "列表转文本";
            this.btn_ListToText.Name = "btn_ListToText";
            this.btn_ListToText.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_ListToText_Click);
            // 
            // MainRibbon
            // 
            this.Name = "MainRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tab2);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MainRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.group6.ResumeLayout(false);
            this.group6.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group8.ResumeLayout(false);
            this.group8.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_ParaBody;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_AbstractHeading;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_FirstHeading;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_SecondHeading;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_ThirdHeading;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_FourthHeading;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_PictureHeading;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_PictureBody;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_TableHeading;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_TableBody;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_ReferenceBody;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_PageSetting;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_AcknowledgementBody;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_AchievementsBody;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_AppendixBody;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_OtherHeading;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_AbstractBody;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_InsertCaption;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_RefrushIndex;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_CatalogHeading;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_SetHeaderAndFooter;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_SeparateSection;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_quote;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_ContainsFormula;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_tableNote;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_formula;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Keywords;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_SepOdd;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group8;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_tableThreeline;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_InsertFormulaCaption;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lbl_formatNote;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton btn_CreateStyles;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_CreateFormulaStyle1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_CreatePicTableStyle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_ListToText;
    }

    partial class ThisRibbonCollection
    {
        internal MainRibbon MainRibbon
        {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
