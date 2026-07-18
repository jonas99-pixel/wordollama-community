using System;
using System.Collections.Generic;
using System.Deployment.Application;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
using Microsoft.Office.Tools.Ribbon;
using WordOllama.Forms;
using WordOllama.NewUI;
using WordOllama.lib;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics.Eventing.Reader;

namespace WordOllama
{

    public partial class Ribbon1
    {
        private bool _handlingCustomPrompt;
        private const int MIN_MINUTES_BETWEEN_CHECKS = 30;
        private static DateTime lastUpdateCheck = DateTime.MinValue;
        private readonly System.Windows.Forms.Timer _notificationTimer = new System.Windows.Forms.Timer();
        private RibbonButton _reviewWorkspaceBtn;


        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            _notificationTimer.Tick -= NotificationTimer_Tick;
            _notificationTimer.Tick += NotificationTimer_Tick;
            LocalizeRuntimeRibbonText();
            EnsureReviewWorkspaceButton();
            ChangeRibbonLabels();
            LoadCustomPrompts();
            CustomPromptList.TextChanged += ComboBox_TextChanged;
        }

        private void EnsureReviewWorkspaceButton()
        {
            if (group3 == null || ReviewWorkspaceBtn == null) return;

            // Ribbon controls must be added during InitializeComponent. Once the
            // Ribbon is loaded, its Items collection is read-only.
            _reviewWorkspaceBtn = ReviewWorkspaceBtn;
            _reviewWorkspaceBtn.Label = Loc.T("I18n_ReviewPaneButton");
            _reviewWorkspaceBtn.ScreenTip = Loc.T("I18n_ReviewPaneButtonTip");
            _reviewWorkspaceBtn.Click -= ReviewWorkspaceBtn_Click;
            _reviewWorkspaceBtn.Click += ReviewWorkspaceBtn_Click;
        }

        private void LocalizeRuntimeRibbonText()
        {
            try
            {
                BtnMarkdown.ScreenTip = Loc.T("I18n_MarkdownScreenTip");
            }
            catch { }
        }

        private void NotificationTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                _notificationTimer.Stop();
                AIStatus.Label = WordOllama.Properties.Resources.AI空闲中;
            }
            catch { }
        }

        public void DisposeNotifications()
        {
            try { _notificationTimer.Stop(); } catch { }
            try { _notificationTimer.Dispose(); } catch { }
        }

        public void LoadCustomPrompts()
        {
            CustomPromptList.Items.Clear();
            var CustomPromptsItem = CustomPromptManager.LoadPrompts();
            foreach (var prompt in CustomPromptsItem)
            {
                var item = this.Factory.CreateRibbonDropDownItem();
                item.Label = prompt.Name;
                CustomPromptList.Items.Add(item);
            }
        }




        private void ComboBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (_handlingCustomPrompt) return;
            var comboBox = sender as RibbonComboBox;
            if (comboBox != null && !string.IsNullOrEmpty(comboBox.Text))
            {
                try
                {
                    _handlingCustomPrompt = true;
                    CustomFuncAdv(comboBox.Text);
                }
                finally
                {
                    _handlingCustomPrompt = false;
                }
            }
        }



        private void toCN_Click(object sender, RibbonControlEventArgs e)
        {
            // 将 sender 转换为 RibbonButton 类型
            RibbonButton button = sender as RibbonButton;

            // 获取按钮的 ID
            string senderName = button.Id;

            // 输出按钮的 ID 到调试窗口
            Debug.WriteLine(senderName);
            translatePrompt("中文");
        }

        private void toEN_Click(object sender, RibbonControlEventArgs e)
        {
            translatePrompt("英文");
        }

        public void ShowNotification(string message, int duration = 1000, ToolTipIcon iconType = ToolTipIcon.None)
        {
            if (SharedSettings.SuppressAgentNotifications)
                return;
            try
            {
                AIStatus.Label = message;
                _notificationTimer.Stop();
                _notificationTimer.Interval = Math.Max(2000, duration + 500);
                _notificationTimer.Start();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Showing in-ribbon notification failed: {ex.Message}");
            }
        }


        private async void translatePrompt(string lang)
        {
            await RunSelectionPromptAsync(PromptTaskType.Translate, targetLanguage: lang);
        }

        private async System.Threading.Tasks.Task RunSelectionPromptAsync(
            PromptTaskType taskType,
            string targetLanguage = null,
            string customInstruction = null,
            PromptOutputMode? forcedOutputMode = null)
        {
            try
            {
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                Word.Range selectionRange = doc.Application.Selection.Range;
                string selectedText = selectionRange.Text;

                if (string.IsNullOrWhiteSpace(selectedText))
                {
                    System.Windows.MessageBox.Show(WordOllama.Properties.Resources.请先选择文本);
                    return;
                }

                PromptOutputMode outputMode = ResolveOutputMode(taskType, selectedText, forcedOutputMode);
                if (outputMode == PromptOutputMode.ReviewPane)
                {
                    OpenReviewPaneForSelection(taskType, targetLanguage, customInstruction);
                    return;
                }

                string prompt = PromptTemplateService.BuildPrompt(
                    taskType,
                    selectedText,
                    targetLanguage,
                    customInstruction);

                if (taskType != PromptTaskType.Translate)
                    await Globals.ThisAddIn.LanguageCheck(selectedText.Substring(0, Math.Min(selectedText.Length, 10)));

                string result = await Globals.ThisAddIn.GenerateModelSeletor(prompt);
                if (string.IsNullOrWhiteSpace(result))
                {
                    System.Windows.MessageBox.Show(WordOllama.Properties.Resources.未能返回结果请检查);
                    return;
                }

                ApplyPromptResult(selectionRange, result, outputMode);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(WordOllama.Properties.Resources.遇到错误 + ex.Message);
            }
        }

        private PromptOutputMode ResolveOutputMode(
            PromptTaskType taskType,
            string selectedText,
            PromptOutputMode? forcedOutputMode)
        {
            if (forcedOutputMode.HasValue)
                return forcedOutputMode.Value;

            PromptOutputMode? configuredOutputMode = ReadConfiguredOutputMode();
            if (configuredOutputMode.HasValue)
                return configuredOutputMode.Value;

            if (PromptTemplateService.ShouldUseReviewPane(taskType, selectedText))
                return PromptOutputMode.ReviewPane;

            return DiffCheckBtn.Checked
                ? PromptOutputMode.InsertBelowWithDiff
                : PromptOutputMode.InsertBelow;
        }

        private PromptOutputMode ResolveCustomOutputMode(string function, string selectedText)
        {
            if (function == "CM")
                return PromptOutputMode.Comment;

            if (function == "InDiff")
                return PromptOutputMode.InsertBelowWithDiff;

            return ResolveOutputMode(PromptTaskType.Custom, selectedText, null);
        }

        private PromptOutputMode? ReadConfiguredOutputMode()
        {
            switch ((SharedSettings.DefaultPromptOutputMode ?? "Auto").Trim())
            {
                case "InsertBelow":
                    return PromptOutputMode.InsertBelow;
                case "InsertBelowWithDiff":
                    return PromptOutputMode.InsertBelowWithDiff;
                case "ReplaceOriginal":
                    return PromptOutputMode.ReplaceOriginal;
                case "Comment":
                    return PromptOutputMode.Comment;
                case "ReviewPane":
                    return PromptOutputMode.ReviewPane;
                default:
                    return null;
            }
        }

        private void ApplyPromptResult(Word.Range selectionRange, string result, PromptOutputMode outputMode)
        {
            if (outputMode == PromptOutputMode.Comment)
            {
                Globals.ThisAddIn.AddCommentToSelection(selectionRange, result);
                return;
            }

            if (outputMode == PromptOutputMode.InsertBelowWithDiff)
            {
                Globals.ThisAddIn.InsertTextBelowSelectionWithDiff(selectionRange, result);
                return;
            }

            if (outputMode == PromptOutputMode.ReplaceOriginal)
            {
                Globals.ThisAddIn.ReplaceTextPrecisely(selectionRange, result);
                return;
            }

            Globals.ThisAddIn.InsertTextBelowSelection(selectionRange, result);
        }

        private void OpenReviewPaneForSelection(
            PromptTaskType taskType,
            string targetLanguage = null,
            string customInstruction = null)
        {
            string instruction = PromptTemplateService.BuildReviewInstruction(
                taskType,
                targetLanguage,
                customInstruction);
            Globals.ThisAddIn.OpenReviewWorkspace(
                loadSelection: true,
                instruction: instruction,
                replaceExisting: true);
        }

        private void 扩写_Click(object sender, RibbonControlEventArgs e)
        {
            _ = RunSelectionPromptAsync(PromptTaskType.Expand);
        }

        private void 优化_Click(object sender, RibbonControlEventArgs e)
        {
            _ = RunSelectionPromptAsync(PromptTaskType.Polish);
        }

        private void 简化_Click(object sender, RibbonControlEventArgs e)
        {
            _ = RunSelectionPromptAsync(PromptTaskType.Simplify);
        }

        private void ContinueWriting_Click(object sender, RibbonControlEventArgs e)
        {
            _ = RunSelectionPromptAsync(PromptTaskType.ContinueWriting);
        }

        private void TranslatePane_Click(object sender, RibbonControlEventArgs e)
        {
            HybridWindowManager.ShowWindow<TranslatorWindow>();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void SettingButton_Click(object sender, RibbonControlEventArgs e)
        {
            HybridWindowManager.ShowWindow<NewSetting>();
        }

        private void LawSearchBtn_Click(object sender, RibbonControlEventArgs e)
        {
            HybridWindowManager.ShowWindow<LawSearchWindow>();
        }

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            RiskAnylse();
        }

        private async void RiskAnylse()
        {
            await RunSelectionPromptAsync(PromptTaskType.LegalRisk, forcedOutputMode: PromptOutputMode.Comment);
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            _ = RunSelectionPromptAsync(PromptTaskType.FairRewrite);
        }

        private void AboutButton_Click(object sender, RibbonControlEventArgs e)
        {
            var aboutWindow = new WordOllama.NewUI.AboutWindow();
            aboutWindow.Show();
        }

        private void CreateButton_Click(object sender, RibbonControlEventArgs e)
        {
            HybridWindowManager.ShowWindow<WritingUI>();
        }

        private void TranslateTo_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            RibbonDropDownItem selecedItem = this.TranslateTo.SelectedItem;

            if (selecedItem != null)
            {
                string selectedValue = selecedItem.Label;
                if (selectedValue != null)
                {
                    translatePrompt(selectedValue);
                    this.TranslateTo.SelectedItemIndex = 0;
                }
            }


        }

        private async void button1_Click_2(object sender, RibbonControlEventArgs e)
        {
            await RunSelectionPromptAsync(PromptTaskType.Summarize);
        }

        private void BtnModifyText_Click(object sender, RibbonControlEventArgs e)
        {
            HybridWindowManager.ShowWindow<ModWritting>();
        }

        private void PleadingReviewForm_Click(object sender, RibbonControlEventArgs e)
        {
            HybridWindowManager.ShowWindow<MootCourtWindow>();
        }

        private void TypoFixBtn_Click(object sender, RibbonControlEventArgs e)
        {
            _ = RunSelectionPromptAsync(PromptTaskType.TypoFix);
        }

        private void fontDialog1_Apply(object sender, EventArgs e)
        {

        }

        private void WrodOllamaIcon_MouseDoubleClick(object sender, System.Windows.Forms.MouseEventArgs e)
        {

        }

        private void UpdateLogBtn_Click(object sender, RibbonControlEventArgs e)
        {
            var win = new UpdateLogWindow();
            try
            {
                var log = SharedSettings.UpdateLog;
                if (!string.IsNullOrEmpty(log))
                    win.SetContent(log);
            }
            catch { }
            win.Show();
        }

        private void OpenCompareForm_Click(object sender, RibbonControlEventArgs e)
        {
            HybridWindowManager.ShowWindow<DocxDiffAnalyze>();
        }

        private void button1_Click_3(object sender, RibbonControlEventArgs e)
        {

        }

        private void CheckForUpdatesManually()
        {
            // 检查距离上次检查的时间间隔
            if ((DateTime.Now - lastUpdateCheck).TotalMinutes < MIN_MINUTES_BETWEEN_CHECKS)
            {
                System.Windows.MessageBox.Show(Loc.F("I18n_ClickOnceTooSoon", MIN_MINUTES_BETWEEN_CHECKS), Loc.T("I18n_TooSoonTitle"));
                return;
            }

            lastUpdateCheck = DateTime.Now;
            UpdateCheckInfo info;

            if (ApplicationDeployment.IsNetworkDeployed)
            {
                ApplicationDeployment ad = ApplicationDeployment.CurrentDeployment;

                try
                {
                    info = ad.CheckForDetailedUpdate();
                    Debug.WriteLine(info);

                }
                catch (DeploymentDownloadException dde)
                {
                    System.Windows.Forms.MessageBox.Show(Loc.F("I18n_UpdateDownloadFailed", dde.Message));
                    return;
                }
                catch (InvalidDeploymentException ide)
                {
                    System.Windows.Forms.MessageBox.Show(Loc.F("I18n_UpdateCheckFailed", ide.Message));
                    return;
                }
                catch (InvalidOperationException ioe)
                {
                    System.Windows.Forms.MessageBox.Show(Loc.F("I18n_UpdateNotClickOnce", ioe.Message));
                    return;
                }
                if (info.UpdateAvailable)
                {
                    bool doUpdate = true;
                    if (!info.IsUpdateRequired)
                    {
                        DialogResult dr = System.Windows.Forms.MessageBox.Show(Loc.T("I18n_UpdateAvailableQuestion"), Loc.T("I18n_UpdateAvailableTitle"), MessageBoxButtons.OKCancel);
                        if (dr != DialogResult.OK)
                        {
                            doUpdate = false;
                        }
                    }
                    else
                    {
                        // 显示强制更新的消息,包括最低要求版本
                        System.Windows.Forms.MessageBox.Show(Loc.F("I18n_UpdateRequired", info.MinimumRequiredVersion.ToString()),
                            Loc.T("I18n_UpdateAvailableTitle"), MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                    if (doUpdate)
                    {
                        try
                        {
                            ad.Update();
                            System.Windows.Forms.MessageBox.Show(Loc.T("I18n_UpdateRestart"));
                            System.Windows.Forms.Application.Restart();
                        }
                        catch (DeploymentDownloadException dde)
                        {
                            System.Windows.Forms.MessageBox.Show(Loc.F("I18n_UpdateInstallFailed", dde));
                            return;
                        }
                    }
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show(Loc.T("I18n_UpdateUnavailableOffline"), Loc.T("I18n_UpdateUnavailableTitle"));
            }
        }

        private void UpdateCheckBtn_Click(object sender, RibbonControlEventArgs e)
        {
            // CheckForUpdatesManually();
        }

        public void ChangeRibbonLabels()
        {
            if (AppConfig.Instance == null) return;
            AIModeNow.Label = AppConfig.Instance.AIMode == "Ollama" ? AppConfig.Instance.AIMode : Loc.T("I18n_OnlineModelLabel");
            if (AppConfig.Instance.AIMode == "Ollama")
            {
                AIModelNow.Label = AppConfig.Instance.OllamaModel;
            }
            else if (AppConfig.Instance.AIMode == "ChatGPT")
            {
                AIModelNow.Label = AppConfig.Instance.ChatGPTModel;
            }

            CFBtn1.Label = AppConfig.Instance.C1_label;
            CFBtn2.Label = AppConfig.Instance.C2_label;
            CFBtn3.Label = AppConfig.Instance.C3_label;
            CFBtn4.Label = AppConfig.Instance.C4_label;

        }

        private void OpenCusFunForm_Click(object sender, RibbonControlEventArgs e)
        {
            HybridWindowManager.ShowWindow<CustomPromptUI>();
        }

        private void CFBtn1_Click(object sender, RibbonControlEventArgs e)
        {
            CustomFunc("C1");
        }


        private async void CustomFunc(string prefix)
        {
            string func = AppConfig.Instance.GetSetting($"{prefix}_Funt").ToString();
            string func_prompt = AppConfig.Instance.GetSetting($"{prefix}_Prompt").ToString();
            string func_name = AppConfig.Instance.GetSetting($"{prefix}_label").ToString();

            if (string.IsNullOrEmpty(func) || string.IsNullOrEmpty(func_prompt))
            {
                System.Windows.MessageBox.Show(Loc.T("I18n_CustomNotConfigured"));
                return;
            }

            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            string selectedText = doc.Application.Selection.Range.Text;

            if (!string.IsNullOrEmpty(selectedText))
            {
                Globals.Ribbons.Ribbon1.ShowNotification(Loc.F("I18n_StartExecuting", func_name), 1000);
                PromptOutputMode outputMode = ResolveCustomOutputMode(func, selectedText);
                await RunSelectionPromptAsync(
                    PromptTaskType.Custom,
                    customInstruction: func_prompt,
                    forcedOutputMode: outputMode);
            }
            else
            {
                System.Windows.MessageBox.Show(WordOllama.Properties.Resources.请先选择文本);
            }
        }

        private async void CustomFuncAdv(string promptName)
        {
            var prompt = CustomPromptManager.GetPromptByName(promptName);
            if (prompt == null)
            {
                System.Windows.MessageBox.Show(Loc.T("I18n_CustomPromptNotFound"));
                return;
            }

            string func = prompt.Function;
            string func_prompt = prompt.Prompt;

            if (string.IsNullOrEmpty(func) || string.IsNullOrEmpty(func_prompt))
            {
                System.Windows.MessageBox.Show(Loc.T("I18n_CustomPromptEmpty"));
                return;
            }

            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            string selectedText = doc.Application.Selection.Range.Text;
            if (!string.IsNullOrEmpty(selectedText))
            {
                Globals.Ribbons.Ribbon1.ShowNotification(Loc.F("I18n_StartExecuting", prompt.Name), 1000);
                PromptOutputMode outputMode = ResolveCustomOutputMode(func, selectedText);
                await RunSelectionPromptAsync(
                    PromptTaskType.Custom,
                    customInstruction: func_prompt,
                    forcedOutputMode: outputMode);
            }
            else
            {
                System.Windows.MessageBox.Show(WordOllama.Properties.Resources.请先选择文本);
            }
        }




        private void CFBtn2_Click(object sender, RibbonControlEventArgs e)
        {
            CustomFunc("C2");
        }

        private void CFBtn3_Click(object sender, RibbonControlEventArgs e)
        {
            CustomFunc("C3");
        }

        private void CFBtn4_Click(object sender, RibbonControlEventArgs e)
        {
            CustomFunc("C4");
        }

        private void button2_Click_1(object sender, RibbonControlEventArgs e)
        {
            
        }

        private void CustomPromptList_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void button1_Click_4(object sender, RibbonControlEventArgs e)
        {
            HybridWindowManager.ShowWindow<ChatWithIMG>();
        }

        private void button1_Click_5(object sender, RibbonControlEventArgs e)
        {
            HybridWindowManager.ShowWindow<TextToTableWindow>();
        }

        private void button1_Click_6(object sender, RibbonControlEventArgs e)
        {
            HybridWindowManager.ShowWindow<CreateHTML>();
        }

        private void button1_Click_7(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.EnsureThinkingPane();
            // 遍历 CTP 找到当前窗口的 ThinkingPane
            string thinkTitle = WordOllama.Properties.Resources.思维链浏览;
            try
            {
                var activeWindow = Globals.ThisAddIn.Application.ActiveWindow;
                foreach (Microsoft.Office.Tools.CustomTaskPane ctp in Globals.ThisAddIn.CustomTaskPanes)
                {
                    if (ctp.Window == activeWindow && ctp.Title == thinkTitle)
                    {
                        ctp.Visible = !ctp.Visible;
                        return;
                    }
                }
            }
            catch { }
        }

        private void AgentBtn_Click(object sender, RibbonControlEventArgs e)
        {
            var ctp = Globals.ThisAddIn.EnsureAgentTaskPane();
            if (ctp != null)
                ctp.Visible = !ctp.Visible;
        }

        private void ReviewWorkspaceBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.OpenReviewWorkspace(loadSelection: true);
        }

        private void BtnMarkdown_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var doc = Globals.ThisAddIn.Application.ActiveDocument;
                if (doc == null)
                {
                    System.Windows.Forms.MessageBox.Show(Loc.T("I18n_OpenWordDocument"));
                    return;
                }

                // Get markdown from clipboard
                string markdown = System.Windows.Forms.Clipboard.GetText();
                if (string.IsNullOrWhiteSpace(markdown))
                {
                    System.Windows.Forms.MessageBox.Show(Loc.T("I18n_NoClipboardText"));
                    return;
                }

                // Load style mappings from settings
                var mappings = new List<MarkdownStyleMapping>();
                string json = WordOllama.Forms.SharedSettings.MarkdownStyleMappingsJson;
                if (!string.IsNullOrEmpty(json))
                {
                    try { mappings = Newtonsoft.Json.JsonConvert.DeserializeObject<List<MarkdownStyleMapping>>(json) ?? mappings; }
                    catch { }
                }

                MarkdownToWord.InsertMarkdown(markdown, mappings);

                Globals.Ribbons.Ribbon1.ShowNotification(Loc.T("I18n_MarkdownConverted"), 2000);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(Loc.F("I18n_ConversionFailed", ex.Message));
            }
        }

    }
}
