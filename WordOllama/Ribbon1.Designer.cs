namespace WordOllama
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl7 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl8 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.CreateGroup = this.Factory.CreateRibbonGroup();
            this.CreateButton = this.Factory.CreateRibbonButton();
            this.BtnModifyText = this.Factory.CreateRibbonButton();
            this.IMGAnalyze = this.Factory.CreateRibbonButton();
            this.Text2Table = this.Factory.CreateRibbonButton();
            this.CreateHTML = this.Factory.CreateRibbonButton();
            this.BtnMarkdown = this.Factory.CreateRibbonButton();
            this.AgentBtn = this.Factory.CreateRibbonButton();
            this.ReviewWorkspaceBtn = this.Factory.CreateRibbonButton();
            this.optimise = this.Factory.CreateRibbonGroup();
            this.优化 = this.Factory.CreateRibbonButton();
            this.扩写 = this.Factory.CreateRibbonButton();
            this.简化 = this.Factory.CreateRibbonButton();
            this.ContinueWriting = this.Factory.CreateRibbonButton();
            this.归纳 = this.Factory.CreateRibbonButton();
            this.TypoFixBtn = this.Factory.CreateRibbonButton();
            this.DiffCheckBtn = this.Factory.CreateRibbonCheckBox();
            this.translate = this.Factory.CreateRibbonGroup();
            this.TranslatePane = this.Factory.CreateRibbonButton();
            this.toCN = this.Factory.CreateRibbonButton();
            this.toEN = this.Factory.CreateRibbonButton();
            this.TranslateTo = this.Factory.CreateRibbonDropDown();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.RiskAssessment = this.Factory.CreateRibbonButton();
            this.FairnessBtn = this.Factory.CreateRibbonButton();
            this.PleadingReview = this.Factory.CreateRibbonButton();
            this.OpenCompareForm = this.Factory.CreateRibbonButton();
            this.LawSearchBtn = this.Factory.CreateRibbonButton();
            this.CustomizedFunctionsGroup = this.Factory.CreateRibbonGroup();
            this.CFBtn1 = this.Factory.CreateRibbonButton();
            this.CFBtn2 = this.Factory.CreateRibbonButton();
            this.CFBtn3 = this.Factory.CreateRibbonButton();
            this.CFBtn4 = this.Factory.CreateRibbonButton();
            this.CustomPromptList = this.Factory.CreateRibbonComboBox();
            this.OpenCusFunForm = this.Factory.CreateRibbonButton();
            this.SettingGroup = this.Factory.CreateRibbonGroup();
            this.SettingButton = this.Factory.CreateRibbonButton();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.AIModeNow = this.Factory.CreateRibbonLabel();
            this.AIModelNow = this.Factory.CreateRibbonLabel();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.AIStatus = this.Factory.CreateRibbonLabel();
            this.showThinkingPane = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.tab1.SuspendLayout();
            this.CreateGroup.SuspendLayout();
            this.optimise.SuspendLayout();
            this.translate.SuspendLayout();
            this.group1.SuspendLayout();
            this.CustomizedFunctionsGroup.SuspendLayout();
            this.SettingGroup.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.CreateGroup);
            this.tab1.Groups.Add(this.optimise);
            this.tab1.Groups.Add(this.translate);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.CustomizedFunctionsGroup);
            this.tab1.Groups.Add(this.SettingGroup);
            this.tab1.Groups.Add(this.group2);
            resources.ApplyResources(this.tab1, "tab1");
            this.tab1.Name = "tab1";
            // 
            // CreateGroup
            // 
            this.CreateGroup.Items.Add(this.CreateButton);
            this.CreateGroup.Items.Add(this.BtnModifyText);
            this.CreateGroup.Items.Add(this.IMGAnalyze);
            this.CreateGroup.Items.Add(this.Text2Table);
            this.CreateGroup.Items.Add(this.CreateHTML);
            this.CreateGroup.Items.Add(this.BtnMarkdown);
            this.CreateGroup.Label = global::WordOllama.Properties.Resources.创作;
            this.CreateGroup.Name = "CreateGroup";
            // 
            // CreateButton
            // 
            this.CreateButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.CreateButton.Image = global::WordOllama.Properties.Resources.wordollama_new_document_50;
            this.CreateButton.Label = global::WordOllama.Properties.Resources.自由创作;
            this.CreateButton.Name = "CreateButton";
            this.CreateButton.ShowImage = true;
            this.CreateButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CreateButton_Click);
            // 
            // BtnModifyText
            // 
            this.BtnModifyText.Label = global::WordOllama.Properties.Resources.按需修改;
            this.BtnModifyText.Name = "BtnModifyText";
            this.BtnModifyText.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnModifyText_Click);
            // 
            // IMGAnalyze
            // 
            this.IMGAnalyze.Label = global::WordOllama.Properties.Resources.图片理解;
            this.IMGAnalyze.Name = "IMGAnalyze";
            this.IMGAnalyze.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click_4);
            // 
            // Text2Table
            // 
            this.Text2Table.Label = global::WordOllama.Properties.Resources.生成表格;
            this.Text2Table.Name = "Text2Table";
            this.Text2Table.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click_5);
            // 
            // CreateHTML
            // 
            this.CreateHTML.Label = global::WordOllama.Properties.Resources.HTML应用;
            this.CreateHTML.Name = "CreateHTML";
            this.CreateHTML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click_6);
            //
            // BtnMarkdown
            //
            this.BtnMarkdown.Label = global::WordOllama.Properties.Resources.Markdown转换;
            this.BtnMarkdown.Name = "BtnMarkdown";
            this.BtnMarkdown.ScreenTip = global::WordOllama.lib.Loc.T("I18n_RibbonMarkdownScreenTip");
            this.BtnMarkdown.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnMarkdown_Click);
            //
            // AgentBtn
            //
            resources.ApplyResources(this.AgentBtn, "AgentBtn");
            this.AgentBtn.Name = "AgentBtn";
            this.AgentBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AgentBtn_Click);
            // 
            // optimise
            // 
            this.optimise.Items.Add(this.优化);
            this.optimise.Items.Add(this.扩写);
            this.optimise.Items.Add(this.简化);
            this.optimise.Items.Add(this.ContinueWriting);
            this.optimise.Items.Add(this.归纳);
            this.optimise.Items.Add(this.TypoFixBtn);
            this.optimise.Items.Add(this.DiffCheckBtn);
            this.optimise.Label = global::WordOllama.Properties.Resources.编辑;
            this.optimise.Name = "optimise";
            // 
            // 优化
            // 
            this.优化.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.优化.Image = global::WordOllama.Properties.Resources.wordollama_assistant_48;
            this.优化.Label = global::WordOllama.Properties.Resources.润色文本;
            this.优化.Name = "优化";
            this.优化.ScreenTip = global::WordOllama.Properties.Resources.对本文进行润色;
            this.优化.ShowImage = true;
            this.优化.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.优化_Click);
            // 
            // 扩写
            // 
            this.扩写.Label = global::WordOllama.Properties.Resources.扩写;
            this.扩写.Name = "扩写";
            this.扩写.ScreenTip = global::WordOllama.Properties.Resources.对文本进行扩写;
            this.扩写.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.扩写_Click);
            // 
            // 简化
            // 
            this.简化.Label = global::WordOllama.Properties.Resources.简化;
            this.简化.Name = "简化";
            this.简化.ScreenTip = global::WordOllama.Properties.Resources.对文本进行简化;
            this.简化.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.简化_Click);
            // 
            // ContinueWriting
            // 
            this.ContinueWriting.Label = global::WordOllama.Properties.Resources.续写;
            this.ContinueWriting.Name = "ContinueWriting";
            this.ContinueWriting.ScreenTip = global::WordOllama.Properties.Resources.对文本进行续写;
            this.ContinueWriting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ContinueWriting_Click);
            // 
            // 归纳
            // 
            this.归纳.Label = global::WordOllama.Properties.Resources.归纳;
            this.归纳.Name = "归纳";
            this.归纳.ScreenTip = global::WordOllama.Properties.Resources.对文本内容进行归纳;
            this.归纳.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click_2);
            // 
            // TypoFixBtn
            // 
            this.TypoFixBtn.Label = global::WordOllama.Properties.Resources.改错;
            this.TypoFixBtn.Name = "TypoFixBtn";
            this.TypoFixBtn.ScreenTip = global::WordOllama.Properties.Resources.修改文段错别字与病句;
            this.TypoFixBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TypoFixBtn_Click);
            // 
            // DiffCheckBtn
            // 
            this.DiffCheckBtn.Label = global::WordOllama.Properties.Resources.对比结果;
            this.DiffCheckBtn.Name = "DiffCheckBtn";
            // 
            // translate
            // 
            this.translate.Items.Add(this.TranslatePane);
            this.translate.Items.Add(this.toCN);
            this.translate.Items.Add(this.toEN);
            this.translate.Items.Add(this.TranslateTo);
            this.translate.Label = global::WordOllama.Properties.Resources.翻译;
            this.translate.Name = "translate";
            // 
            // TranslatePane
            // 
            this.TranslatePane.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.TranslatePane.Image = global::WordOllama.Properties.Resources.wordollama_translate_50;
            this.TranslatePane.Label = global::WordOllama.Properties.Resources.自由翻译;
            this.TranslatePane.Name = "TranslatePane";
            this.TranslatePane.ShowImage = true;
            this.TranslatePane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TranslatePane_Click);
            // 
            // toCN
            // 
            this.toCN.Label = global::WordOllama.Properties.Resources.翻译到中文;
            this.toCN.Name = "toCN";
            this.toCN.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toCN_Click);
            // 
            // toEN
            // 
            this.toEN.Label = global::WordOllama.Properties.Resources.翻译到英文;
            this.toEN.Name = "toEN";
            this.toEN.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toEN_Click);
            // 
            // TranslateTo
            // 
            ribbonDropDownItemImpl2.Label = global::WordOllama.Properties.Resources.日语;
            ribbonDropDownItemImpl3.Label = global::WordOllama.Properties.Resources.韩语;
            ribbonDropDownItemImpl4.Label = global::WordOllama.Properties.Resources.德语;
            ribbonDropDownItemImpl5.Label = global::WordOllama.Properties.Resources.西班牙语;
            ribbonDropDownItemImpl6.Label = global::WordOllama.Properties.Resources.意大利语;
            ribbonDropDownItemImpl7.Label = global::WordOllama.Properties.Resources.法语;
            ribbonDropDownItemImpl8.Label = global::WordOllama.Properties.Resources.俄罗斯语;
            this.TranslateTo.Items.Add(ribbonDropDownItemImpl1);
            this.TranslateTo.Items.Add(ribbonDropDownItemImpl2);
            this.TranslateTo.Items.Add(ribbonDropDownItemImpl3);
            this.TranslateTo.Items.Add(ribbonDropDownItemImpl4);
            this.TranslateTo.Items.Add(ribbonDropDownItemImpl5);
            this.TranslateTo.Items.Add(ribbonDropDownItemImpl6);
            this.TranslateTo.Items.Add(ribbonDropDownItemImpl7);
            this.TranslateTo.Items.Add(ribbonDropDownItemImpl8);
            this.TranslateTo.Label = global::WordOllama.Properties.Resources.翻译到;
            this.TranslateTo.Name = "TranslateTo";
            this.TranslateTo.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TranslateTo_SelectionChanged);
            // 
            // group1
            // 
            this.group1.Items.Add(this.RiskAssessment);
            this.group1.Items.Add(this.FairnessBtn);
            this.group1.Items.Add(this.PleadingReview);
            this.group1.Items.Add(this.OpenCompareForm);
            this.group1.Items.Add(this.LawSearchBtn);
            this.group1.Label = global::WordOllama.Properties.Resources.法务;
            this.group1.Name = "group1";
            // 
            // RiskAssessment
            // 
            this.RiskAssessment.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.RiskAssessment.Image = global::WordOllama.Properties.Resources.wordollama_legal_50;
            this.RiskAssessment.Label = global::WordOllama.Properties.Resources.批注风险;
            this.RiskAssessment.Name = "RiskAssessment";
            this.RiskAssessment.ShowImage = true;
            this.RiskAssessment.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click_1);
            // 
            // FairnessBtn
            // 
            this.FairnessBtn.Label = global::WordOllama.Properties.Resources.公平性改写;
            this.FairnessBtn.Name = "FairnessBtn";
            this.FairnessBtn.ScreenTip = global::WordOllama.Properties.Resources.把文本修改得更公平;
            this.FairnessBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // PleadingReview
            // 
            this.PleadingReview.Label = global::WordOllama.Properties.Resources.模拟法庭调查;
            this.PleadingReview.Name = "PleadingReview";
            this.PleadingReview.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.PleadingReviewForm_Click);
            // 
            // OpenCompareForm
            // 
            this.OpenCompareForm.Label = global::WordOllama.Properties.Resources.合同修订评估;
            this.OpenCompareForm.Name = "OpenCompareForm";
            this.OpenCompareForm.ScreenTip = global::WordOllama.Properties.Resources.对两份合同文件的修订内容进行风险评估;
            this.OpenCompareForm.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OpenCompareForm_Click);
            // 
            // LawSearchBtn
            // 
            resources.ApplyResources(this.LawSearchBtn, "LawSearchBtn");
            this.LawSearchBtn.Name = "LawSearchBtn";
            this.LawSearchBtn.OfficeImageId = "ZoomPrintPreviewExcel";
            this.LawSearchBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LawSearchBtn_Click);
            // 
            // CustomizedFunctionsGroup
            // 
            this.CustomizedFunctionsGroup.Items.Add(this.CFBtn1);
            this.CustomizedFunctionsGroup.Items.Add(this.CFBtn2);
            this.CustomizedFunctionsGroup.Items.Add(this.CFBtn3);
            this.CustomizedFunctionsGroup.Items.Add(this.CFBtn4);
            this.CustomizedFunctionsGroup.Items.Add(this.CustomPromptList);
            this.CustomizedFunctionsGroup.Items.Add(this.OpenCusFunForm);
            this.CustomizedFunctionsGroup.Label = global::WordOllama.Properties.Resources.自定义功能;
            this.CustomizedFunctionsGroup.Name = "CustomizedFunctionsGroup";
            // 
            // CFBtn1
            // 
            resources.ApplyResources(this.CFBtn1, "CFBtn1");
            this.CFBtn1.Name = "CFBtn1";
            this.CFBtn1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CFBtn1_Click);
            // 
            // CFBtn2
            // 
            resources.ApplyResources(this.CFBtn2, "CFBtn2");
            this.CFBtn2.Name = "CFBtn2";
            this.CFBtn2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CFBtn2_Click);
            // 
            // CFBtn3
            // 
            resources.ApplyResources(this.CFBtn3, "CFBtn3");
            this.CFBtn3.Name = "CFBtn3";
            this.CFBtn3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CFBtn3_Click);
            // 
            // CFBtn4
            // 
            resources.ApplyResources(this.CFBtn4, "CFBtn4");
            this.CFBtn4.Name = "CFBtn4";
            this.CFBtn4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CFBtn4_Click);
            // 
            // CustomPromptList
            // 
            this.CustomPromptList.Label = global::WordOllama.Properties.Resources.更多;
            this.CustomPromptList.Name = "CustomPromptList";
            resources.ApplyResources(this.CustomPromptList, "CustomPromptList");
            this.CustomPromptList.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CustomPromptList_TextChanged);
            // 
            // OpenCusFunForm
            // 
            this.OpenCusFunForm.Image = global::WordOllama.Properties.Resources.wordollama_custom_50;
            this.OpenCusFunForm.Label = global::WordOllama.Properties.Resources.配置;
            this.OpenCusFunForm.Name = "OpenCusFunForm";
            this.OpenCusFunForm.ShowImage = true;
            this.OpenCusFunForm.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OpenCusFunForm_Click);
            // 
            // SettingGroup
            // 
            this.SettingGroup.Items.Add(this.SettingButton);
            this.SettingGroup.Items.Add(this.label1);
            this.SettingGroup.Items.Add(this.AIModeNow);
            this.SettingGroup.Items.Add(this.AIModelNow);
            this.SettingGroup.Label = global::WordOllama.Properties.Resources.设置;
            this.SettingGroup.Name = "SettingGroup";
            // 
            // SettingButton
            // 
            this.SettingButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.SettingButton.Image = global::WordOllama.Properties.Resources.wordollama_settings_50;
            this.SettingButton.Label = global::WordOllama.Properties.Resources.设置;
            this.SettingButton.Name = "SettingButton";
            this.SettingButton.ShowImage = true;
            this.SettingButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SettingButton_Click);
            // 
            // label1
            // 
            this.label1.Label = global::WordOllama.Properties.Resources.当前模式模型;
            this.label1.Name = "label1";
            // 
            // AIModeNow
            // 
            this.AIModeNow.Label = global::WordOllama.Properties.Resources.当前模式;
            this.AIModeNow.Name = "AIModeNow";
            // 
            // AIModelNow
            // 
            this.AIModelNow.Label = global::WordOllama.Properties.Resources.LABEL_CURRENT_MODEL;
            this.AIModelNow.Name = "AIModelNow";
            // 
            // group2
            // 
            this.group2.Items.Add(this.AIStatus);
            this.group2.Items.Add(this.showThinkingPane);
            this.group2.Items.Add(this.button1);
            this.group2.Label = global::WordOllama.Properties.Resources.当前状态;
            this.group2.Name = "group2";
            // 
            // AIStatus
            // 
            this.AIStatus.Label = global::WordOllama.Properties.Resources.AI空闲中;
            this.AIStatus.Name = "AIStatus";
            // 
            // showThinkingPane
            // 
            this.showThinkingPane.Label = global::WordOllama.Properties.Resources.查看思维链;
            this.showThinkingPane.Name = "showThinkingPane";
            this.showThinkingPane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click_7);
            // 
            // button1
            // 
            resources.ApplyResources(this.button1, "button1");
            this.button1.Name = "button1";
            // 
            // group3
            // 
            this.group3.Items.Add(this.AgentBtn);
            this.group3.Items.Add(this.ReviewWorkspaceBtn);
            resources.ApplyResources(this.group3, "group3");
            this.group3.Name = "group3";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.CreateGroup.ResumeLayout(false);
            this.CreateGroup.PerformLayout();
            this.optimise.ResumeLayout(false);
            this.optimise.PerformLayout();
            this.translate.ResumeLayout(false);
            this.translate.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.CustomizedFunctionsGroup.ResumeLayout(false);
            this.CustomizedFunctionsGroup.PerformLayout();
            this.SettingGroup.ResumeLayout(false);
            this.SettingGroup.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup translate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton toCN;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton toEN;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup optimise;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 扩写;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 优化;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 简化;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TranslatePane;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ContinueWriting;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup CreateGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CreateButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup SettingGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SettingButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RiskAssessment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton FairnessBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown TranslateTo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 归纳;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnModifyText;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton PleadingReview;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TypoFixBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox DiffCheckBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton OpenCompareForm;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel AIModeNow;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel AIModelNow;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup CustomizedFunctionsGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CFBtn1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CFBtn2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CFBtn3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CFBtn4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton OpenCusFunForm;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox CustomPromptList;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton IMGAnalyze;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Text2Table;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel AIStatus;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CreateHTML;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnMarkdown;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AgentBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ReviewWorkspaceBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton showThinkingPane;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton LawSearchBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        private System.Windows.Forms.ColorDialog colorDialog1;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
