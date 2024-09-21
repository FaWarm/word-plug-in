namespace 沈体论文格式调整word插件
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
            this.btn_CreateStyles = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btn_AbstractHeading = this.Factory.CreateRibbonButton();
            this.btn_CatalogHeading = this.Factory.CreateRibbonButton();
            this.btn_FirstHeading = this.Factory.CreateRibbonButton();
            this.btn_SecondHeading = this.Factory.CreateRibbonButton();
            this.btn_ThirdHeading = this.Factory.CreateRibbonButton();
            this.btn_FourthHeading = this.Factory.CreateRibbonButton();
            this.btn_OtherHeading = this.Factory.CreateRibbonButton();
            this.btn_PictureHeading = this.Factory.CreateRibbonButton();
            this.btn_TableHeading = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.btn_ParaBody = this.Factory.CreateRibbonButton();
            this.btn_ReferenceBody = this.Factory.CreateRibbonButton();
            this.btn_AbstractBody = this.Factory.CreateRibbonButton();
            this.btn_AcknowledgementBody = this.Factory.CreateRibbonButton();
            this.btn_AchievementsBody = this.Factory.CreateRibbonButton();
            this.btn_AppendixBody = this.Factory.CreateRibbonButton();
            this.btn_PictureBody = this.Factory.CreateRibbonButton();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.btn_SetHeaderAndFooter = this.Factory.CreateRibbonButton();
            this.btn_TableBody = this.Factory.CreateRibbonButton();
            this.btn_InsertCaption = this.Factory.CreateRibbonButton();
            this.btn_quote = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btn_PageSetting = this.Factory.CreateRibbonButton();
            this.btn_RefrushIndex = this.Factory.CreateRibbonButton();
            this.btn_SeparateSection = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.tab2.SuspendLayout();
            this.group6.SuspendLayout();
            this.group2.SuspendLayout();
            this.group4.SuspendLayout();
            this.group5.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "沈体排版";
            this.tab1.Name = "tab1";
            // 
            // tab2
            // 
            this.tab2.Groups.Add(this.group6);
            this.tab2.Groups.Add(this.group2);
            this.tab2.Groups.Add(this.group4);
            this.tab2.Groups.Add(this.group5);
            this.tab2.Groups.Add(this.group3);
            this.tab2.Label = "沈体插件";
            this.tab2.Name = "tab2";
            // 
            // group6
            // 
            this.group6.Items.Add(this.btn_CreateStyles);
            this.group6.Label = "生成样式";
            this.group6.Name = "group6";
            // 
            // btn_CreateStyles
            // 
            this.btn_CreateStyles.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_CreateStyles.Image = global::沈体论文格式调整word插件.Properties.Resources.shengcheng;
            this.btn_CreateStyles.Label = "生成论文样式";
            this.btn_CreateStyles.Name = "btn_CreateStyles";
            this.btn_CreateStyles.ShowImage = true;
            this.btn_CreateStyles.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_CreateStyles_Click);
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
            // btn_AbstractHeading
            // 
            this.btn_AbstractHeading.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_AbstractHeading.Image = global::沈体论文格式调整word插件.Properties.Resources.zhaiyao;
            this.btn_AbstractHeading.Label = "摘要";
            this.btn_AbstractHeading.Name = "btn_AbstractHeading";
            this.btn_AbstractHeading.ShowImage = true;
            this.btn_AbstractHeading.SuperTip = "不显示在目录中";
            this.btn_AbstractHeading.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_AbstractHeading_Click);
            // 
            // btn_CatalogHeading
            // 
            this.btn_CatalogHeading.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_CatalogHeading.Image = global::沈体论文格式调整word插件.Properties.Resources.mulu;
            this.btn_CatalogHeading.Label = "目录";
            this.btn_CatalogHeading.Name = "btn_CatalogHeading";
            this.btn_CatalogHeading.ShowImage = true;
            this.btn_CatalogHeading.SuperTip = "不显示在目录中";
            this.btn_CatalogHeading.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_CatalogHeading_Click);
            // 
            // btn_FirstHeading
            // 
            this.btn_FirstHeading.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_FirstHeading.Image = global::沈体论文格式调整word插件.Properties.Resources.yijibiaoti;
            this.btn_FirstHeading.Label = "正文1级标题";
            this.btn_FirstHeading.Name = "btn_FirstHeading";
            this.btn_FirstHeading.ShowImage = true;
            this.btn_FirstHeading.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_FirstHeading_Click);
            // 
            // btn_SecondHeading
            // 
            this.btn_SecondHeading.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_SecondHeading.Image = global::沈体论文格式调整word插件.Properties.Resources.erjibiaoti;
            this.btn_SecondHeading.Label = "2级标题";
            this.btn_SecondHeading.Name = "btn_SecondHeading";
            this.btn_SecondHeading.ShowImage = true;
            this.btn_SecondHeading.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_SecondHeading_Click);
            // 
            // btn_ThirdHeading
            // 
            this.btn_ThirdHeading.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_ThirdHeading.Image = global::沈体论文格式调整word插件.Properties.Resources.sanjibiaoti;
            this.btn_ThirdHeading.Label = "3级标题";
            this.btn_ThirdHeading.Name = "btn_ThirdHeading";
            this.btn_ThirdHeading.ShowImage = true;
            this.btn_ThirdHeading.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_ThirdHeading_Click);
            // 
            // btn_FourthHeading
            // 
            this.btn_FourthHeading.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_FourthHeading.Image = global::沈体论文格式调整word插件.Properties.Resources.sijibiaoti;
            this.btn_FourthHeading.Label = "4级标题";
            this.btn_FourthHeading.Name = "btn_FourthHeading";
            this.btn_FourthHeading.ShowImage = true;
            this.btn_FourthHeading.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_FourthHeading_Click);
            // 
            // btn_OtherHeading
            // 
            this.btn_OtherHeading.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_OtherHeading.Image = global::沈体论文格式调整word插件.Properties.Resources.qitabiaoti;
            this.btn_OtherHeading.Label = "其他1级标题";
            this.btn_OtherHeading.Name = "btn_OtherHeading";
            this.btn_OtherHeading.ShowImage = true;
            this.btn_OtherHeading.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_OtherHeading_Click);
            // 
            // btn_PictureHeading
            // 
            this.btn_PictureHeading.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_PictureHeading.Image = global::沈体论文格式调整word插件.Properties.Resources.tupianzhengwen;
            this.btn_PictureHeading.Label = "图片题注";
            this.btn_PictureHeading.Name = "btn_PictureHeading";
            this.btn_PictureHeading.ShowImage = true;
            this.btn_PictureHeading.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_PictureHeading_Click);
            // 
            // btn_TableHeading
            // 
            this.btn_TableHeading.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_TableHeading.Image = global::沈体论文格式调整word插件.Properties.Resources.biaoge;
            this.btn_TableHeading.Label = "表格题注";
            this.btn_TableHeading.Name = "btn_TableHeading";
            this.btn_TableHeading.ShowImage = true;
            this.btn_TableHeading.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_TableHeading_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.btn_ParaBody);
            this.group4.Items.Add(this.btn_ReferenceBody);
            this.group4.Items.Add(this.btn_AbstractBody);
            this.group4.Items.Add(this.btn_AcknowledgementBody);
            this.group4.Items.Add(this.btn_AchievementsBody);
            this.group4.Items.Add(this.btn_AppendixBody);
            this.group4.Items.Add(this.btn_PictureBody);
            this.group4.Label = "正文";
            this.group4.Name = "group4";
            // 
            // btn_ParaBody
            // 
            this.btn_ParaBody.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_ParaBody.Image = global::沈体论文格式调整word插件.Properties.Resources.zhengwen;
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
            // btn_AcknowledgementBody
            // 
            this.btn_AcknowledgementBody.Label = "致谢正文";
            this.btn_AcknowledgementBody.Name = "btn_AcknowledgementBody";
            this.btn_AcknowledgementBody.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_AcknowledgementBody_Click);
            // 
            // btn_AchievementsBody
            // 
            this.btn_AchievementsBody.Label = "学术成果/作者正文";
            this.btn_AchievementsBody.Name = "btn_AchievementsBody";
            this.btn_AchievementsBody.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_AchievementsBody_Click);
            // 
            // btn_AppendixBody
            // 
            this.btn_AppendixBody.Label = "附录正文";
            this.btn_AppendixBody.Name = "btn_AppendixBody";
            this.btn_AppendixBody.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_AppendixBody_Click);
            // 
            // btn_PictureBody
            // 
            this.btn_PictureBody.Label = "图片正文";
            this.btn_PictureBody.Name = "btn_PictureBody";
            this.btn_PictureBody.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_PictureBody_Click);
            // 
            // group5
            // 
            this.group5.Items.Add(this.btn_SetHeaderAndFooter);
            this.group5.Items.Add(this.btn_TableBody);
            this.group5.Items.Add(this.btn_InsertCaption);
            this.group5.Items.Add(this.btn_quote);
            this.group5.Label = "图表公式";
            this.group5.Name = "group5";
            // 
            // btn_SetHeaderAndFooter
            // 
            this.btn_SetHeaderAndFooter.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_SetHeaderAndFooter.Image = global::沈体论文格式调整word插件.Properties.Resources.yemeiyejiao;
            this.btn_SetHeaderAndFooter.Label = "设置页眉页脚";
            this.btn_SetHeaderAndFooter.Name = "btn_SetHeaderAndFooter";
            this.btn_SetHeaderAndFooter.ShowImage = true;
            this.btn_SetHeaderAndFooter.SuperTip = "页眉会设置每一章/节的首行文本内容";
            this.btn_SetHeaderAndFooter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_SetHeaderAndFooter_Click);
            // 
            // btn_TableBody
            // 
            this.btn_TableBody.Label = "三线表";
            this.btn_TableBody.Name = "btn_TableBody";
            this.btn_TableBody.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_TableBody_Click);
            // 
            // btn_InsertCaption
            // 
            this.btn_InsertCaption.Label = "插入题注";
            this.btn_InsertCaption.Name = "btn_InsertCaption";
            this.btn_InsertCaption.SuperTip = "详细样式可在【引用】-【插入题注】中设置";
            this.btn_InsertCaption.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_InsertCaption_Click);
            // 
            // btn_quote
            // 
            this.btn_quote.Label = "交叉引用";
            this.btn_quote.Name = "btn_quote";
            this.btn_quote.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_quote_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.btn_PageSetting);
            this.group3.Items.Add(this.btn_RefrushIndex);
            this.group3.Items.Add(this.btn_SeparateSection);
            this.group3.Label = "其他功能";
            this.group3.Name = "group3";
            // 
            // btn_PageSetting
            // 
            this.btn_PageSetting.Label = "页面格式规范";
            this.btn_PageSetting.Name = "btn_PageSetting";
            this.btn_PageSetting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_PageSetting_Click);
            // 
            // btn_RefrushIndex
            // 
            this.btn_RefrushIndex.Label = "刷新序号";
            this.btn_RefrushIndex.Name = "btn_RefrushIndex";
            this.btn_RefrushIndex.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_RefrushIndex_Click);
            // 
            // btn_SeparateSection
            // 
            this.btn_SeparateSection.Label = "分节/章";
            this.btn_SeparateSection.Name = "btn_SeparateSection";
            this.btn_SeparateSection.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_SeparateSection_Click);
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_CreateStyles;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_InsertCaption;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_RefrushIndex;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_CatalogHeading;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_SetHeaderAndFooter;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_SeparateSection;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_quote;
    }

    partial class ThisRibbonCollection
    {
        internal MainRibbon MainRibbon
        {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
