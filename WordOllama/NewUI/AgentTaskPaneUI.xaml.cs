using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using MaterialDesignThemes.Wpf;
using Newtonsoft.Json.Linq;
using WordOllama.lib;
using WordOllama.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Threading;
using System.Windows.Interop;

namespace WordOllama.NewUI
{
    /// <summary>
    /// Agent 任务窗格 — 聊天式 UI + 实际调用 DocumentAgent。
    /// </summary>
    public partial class AgentTaskPaneUI : UserControl
    {
        private readonly ObservableCollection<ChatItem> _chatItems = new ObservableCollection<ChatItem>();
        private readonly ObservableCollection<ReviewIssue> _reviewIssues = new ObservableCollection<ReviewIssue>();
        private readonly ObservableCollection<EditSuggestionSession> _editSuggestions = new ObservableCollection<EditSuggestionSession>();
        private readonly ObservableCollection<AgentSource> _agentSources = new ObservableCollection<AgentSource>();
        private WritingProfile _writingProfile;
        private DocumentAgent _agent;
        private CancellationTokenSource _cts;
        private CancellationTokenSource _reviewCts;
        private bool _isRunning;
        private bool _isWaitingForResponse;
        private string _pendingRequirement;
        private string _pendingImageBase64;
        private readonly SkillManager _skillManager = new SkillManager();
        private int _selectedCommandIndex = -1;
        private const int ReviewMinimumParagraphCharacters = 20;
        private const int ReviewMergedParagraphMaxCharacters = 2000;
        private string _issueReviewCoverage;
        private bool _issueReviewIncomplete;
        private bool _themeSubscribed;
        private bool _isLoadingReviewScope;

        // ========== 图片粘贴 ==========
        /// <summary>用户粘贴的图片 Base64 数据（发送前暂存）。</summary>
        private string _pastedImageBase64;

        // ========== 步骤卡片缓存 ==========
        private string _bufferedThought = null;
        private string _bufferedAction = null;
        private ChatItem _executionProcessItem;
        private ChatItem _finalAnswerItem;
        private int _executionProcessEntryCount;
        private readonly bool _reviewOnly;

        private sealed class ReviewParagraphUnit
        {
            public int ParagraphIndex { get; set; }
            public int Start { get; set; }
            public int End { get; set; }
            public string Text { get; set; }
        }

        public AgentTaskPaneUI() : this(reviewOnly: false)
        {
        }

        public AgentTaskPaneUI(bool reviewOnly)
        {
            _reviewOnly = reviewOnly;
            InitializeComponent();
            _chatItems.CollectionChanged += (_, __) => UpdateEmptyChatState();
            txtPaneTitle.Text = _reviewOnly ? UiText.ReviewPaneTitle : UiText.AgentTitle;
            ConfigurePanelMode();
            chatPanel.ItemsSource = _chatItems;
            issueList.ItemsSource = _reviewIssues;
            suggestionList.ItemsSource = _editSuggestions;
            sourcesList.ItemsSource = _agentSources;
            txtReviewInstruction.Text = UiText.DefaultReviewInstruction;
            _writingProfile = WritingProfileManager.LoadForDocument(Globals.ThisAddIn?.Application?.ActiveDocument);
            txtWritingProfile.Text = _writingProfile.Instructions ?? "";
            InitializeReviewSessions();
            UpdateIssueSummary();
            UpdateSuggestionEmptyState();
            UpdateEmptyChatState();

            // 加载技能列表用于斜杠选择
            _skillManager.LoadSkills(SkillsRootPath);
            RefreshAgentApiDocumentation();

            ApplyTheme();
            Loaded += AgentTaskPaneUI_Loaded;
            Unloaded += AgentTaskPaneUI_Unloaded;
        }

        private void AgentTaskPaneUI_Loaded(object sender, RoutedEventArgs e)
        {
            if (_themeSubscribed) return;
            ThemeManager.ThemeChanged += ThemeManager_ThemeChanged;
            _themeSubscribed = true;
            ApplyTheme();
        }

        private void AgentTaskPaneUI_Unloaded(object sender, RoutedEventArgs e)
        {
            FlushReviewSuggestionEdits();
            if (!_themeSubscribed) return;
            ThemeManager.ThemeChanged -= ThemeManager_ThemeChanged;
            _themeSubscribed = false;
        }

        private void ThemeManager_ThemeChanged(object sender, EventArgs e)
        {
            ApplyTheme();
        }

        private void ConfigurePanelMode()
        {
            if (_reviewOnly)
            {
                ChatTabItem.Visibility = Visibility.Collapsed;
                IssuesTabItem.Visibility = Visibility.Collapsed;
                ApiTabItem.Visibility = Visibility.Collapsed;
                ReviewTabItem.Visibility = Visibility.Visible;
                mainTabs.SelectedItem = ReviewTabItem;
            }
            else
            {
                // Review has a separate task pane so the Agent surface stays focused on chat and tasks.
                ReviewTabItem.Visibility = Visibility.Collapsed;
            }
        }

        private void ApplyTheme()
        {
            try
            {
                if (!Dispatcher.CheckAccess())
                {
                    Dispatcher.Invoke(() => ApplyTheme());
                    return;
                }
                var theme = Resources.GetTheme();
                bool dark = SharedSettings.themeType == "Dark";
                theme.SetBaseTheme(dark ? BaseTheme.Dark : BaseTheme.Light);
                Resources.SetTheme(theme);
                Resources["AgentUserBubbleBrush"] = new SolidColorBrush(dark
                    ? Color.FromRgb(0x32, 0x36, 0x3F)
                    : Color.FromRgb(0xE9, 0xEC, 0xF3));
                Resources["AgentUserBubbleForegroundBrush"] = new SolidColorBrush(dark
                    ? Color.FromRgb(0xF4, 0xF4, 0xF5)
                    : Color.FromRgb(0x20, 0x21, 0x24));
                Resources["AgentProcessBackgroundBrush"] = new SolidColorBrush(dark
                    ? Color.FromRgb(0x25, 0x27, 0x2B)
                    : Color.FromRgb(0xF3, 0xF4, 0xF6));
            }
            catch { }
        }

        private void RefreshAgentApiDocumentation()
        {
            try
            {
                ToolRegistry registry = _agent?.ToolRegistry;
                if (registry == null)
                {
                    registry = new ToolRegistry();
                    registry.DiscoverAndRegister(SharedSettings.AgentExternalToolsEnabled);
                }

                bool english = !string.Equals(SharedSettings.LanguageMode, "Chinese", StringComparison.OrdinalIgnoreCase);
                txtApiDocumentation.Text = AgentToolCatalogService.BuildMarkdown(registry, english);
            }
            catch (Exception ex)
            {
                txtApiDocumentation.Text = string.Format(UiText.ApiBuildFailed, ex.Message);
            }
        }

        private void BtnRefreshApi_Click(object sender, RoutedEventArgs e)
        {
            RefreshAgentApiDocumentation();
        }

        private void BtnCopyApi_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Clipboard.SetText(txtApiDocumentation.Text ?? "");
                Globals.Ribbons.Ribbon1.ShowNotification(
                    UiText.ApiCopied, 1500);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, UiText.ApiTab, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void BtnGenerateSkill_Click(object sender, RoutedEventArgs e)
        {
            string name = txtSkillName.Text?.Trim();
            string goal = txtSkillGoal.Text?.Trim();
            if (string.IsNullOrWhiteSpace(name))
            {
                MessageBox.Show(UiText.ApiNameRequired, UiText.ApiTab, MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            if (string.IsNullOrWhiteSpace(goal))
            {
                MessageBox.Show(UiText.ApiGoalRequired, UiText.ApiTab, MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtApiDocumentation.Text))
                RefreshAgentApiDocumentation();

            string description = txtSkillDescription.Text?.Trim() ?? "";
            bool english = !string.Equals(SharedSettings.LanguageMode, "Chinese", StringComparison.OrdinalIgnoreCase);
            string prompt = AgentToolCatalogService.BuildSkillAuthoringPrompt(
                txtApiDocumentation.Text, name, description, goal, english);

            btnGenerateSkill.IsEnabled = false;
            txtSkillDraft.Text = UiText.ApiGenerating;
            try
            {
                string draft = await Globals.ThisAddIn.GenerateModelSeletor(prompt);
                if (string.IsNullOrWhiteSpace(draft) || draft.StartsWith("[API ", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(draft ?? UiText.ApiEmptyModel);
                txtSkillDraft.Text = draft.Trim();
            }
            catch (Exception ex)
            {
                txtSkillDraft.Text = "";
                MessageBox.Show(string.Format(UiText.ApiGenerationFailed, ex.Message), UiText.ApiTab, MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnGenerateSkill.IsEnabled = true;
            }
        }

        private void BtnSaveSkill_Click(object sender, RoutedEventArgs e)
        {
            string name = txtSkillName.Text?.Trim();
            string draft = txtSkillDraft.Text?.Trim();
            if (string.IsNullOrWhiteSpace(name))
            {
                MessageBox.Show(UiText.ApiNameRequired, UiText.ApiTab, MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            if (string.IsNullOrWhiteSpace(draft) || string.Equals(draft, UiText.ApiGenerating, StringComparison.Ordinal))
            {
                MessageBox.Show(UiText.ApiDraftRequired, UiText.ApiTab, MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (!AgentToolCatalogService.TryGetSkillPath(SkillsRootPath, name, out string skillDirectory, out _))
            {
                MessageBox.Show(UiText.ApiInvalidSkillName, UiText.ApiTab, MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string skillFile = System.IO.Path.Combine(skillDirectory, "SKILL.md");
            if (System.IO.File.Exists(skillFile) &&
                MessageBox.Show(UiText.ApiSkillExists, UiText.ApiTab, MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            try
            {
                string normalized = AgentToolCatalogService.NormalizeSkillDocument(
                    name, txtSkillDescription.Text?.Trim(), draft);
                System.IO.Directory.CreateDirectory(skillDirectory);
                System.IO.File.WriteAllText(skillFile, normalized, new UTF8Encoding(false));
                _skillManager.LoadSkills(SkillsRootPath);
                _agent?.RefreshSkills();
                MessageBox.Show(string.Format(UiText.ApiSkillSaved, skillFile), UiText.ApiTab, MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, UiText.ApiTab, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ========== 核心：用户输入 → 实际执行 Agent ==========

        /// <summary>
        /// Ctrl+Enter 发送，Enter 换行。
        /// </summary>
        private void TxtInput_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            // Ctrl+V 粘贴图片检测
            if (e.Key == System.Windows.Input.Key.V &&
                (System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Control) == System.Windows.Input.ModifierKeys.Control)
            {
                if (System.Windows.Clipboard.ContainsImage())
                {
                    var image = System.Windows.Clipboard.GetImage();
                    if (image != null)
                    {
                        ShowPastePreview(image);
                        e.Handled = true;
                        return;
                    }
                }
            }

            if (commandPopup.IsOpen)
            {
                if (e.Key == System.Windows.Input.Key.Down)
                {
                    SelectNext(1);
                    e.Handled = true;
                    return;
                }
                if (e.Key == System.Windows.Input.Key.Up)
                {
                    SelectNext(-1);
                    e.Handled = true;
                    return;
                }
                if (e.Key == System.Windows.Input.Key.Enter && _selectedCommandIndex >= 0)
                {
                    if (_selectedCommandIndex < commandList.Children.Count)
                    {
                        if (commandList.Children[_selectedCommandIndex] is Button btn)
                        {
                            btn.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                            e.Handled = true;
                            return;
                        }
                    }
                }
            }
        }
        private void TxtInput_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            // 弹出菜单打开时：上下键导航，Enter 确认
            if (commandPopup.IsOpen)
            {
                if (e.Key == System.Windows.Input.Key.Down)
                {
                    SelectNext(1);
                    e.Handled = true;
                    return;
                }
                if (e.Key == System.Windows.Input.Key.Up)
                {
                    SelectNext(-1);
                    e.Handled = true;
                    return;
                }
                if (e.Key == System.Windows.Input.Key.Enter && _selectedCommandIndex >= 0)
                {
                    if (_selectedCommandIndex < commandList.Children.Count)
                    {
                        if (commandList.Children[_selectedCommandIndex] is Button btn)
                        {
                            btn.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                            e.Handled = true;
                            return;
                        }
                    }
                }
            }

            // Ctrl+Enter 发送
            if (e.Key == System.Windows.Input.Key.Enter &&
                (System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Control) == System.Windows.Input.ModifierKeys.Control)
            {
                SendInput();
                e.Handled = true;
            }
        }

        private void SelectNext(int direction)
        {
            int count = commandList.Children.Count;
            if (count == 0) return;

            int newIndex = _selectedCommandIndex + direction;
            if (newIndex < 0) newIndex = count - 1;
            else if (newIndex >= count) newIndex = 0;

            _selectedCommandIndex = newIndex;
            UpdateCommandSelection();
        }

        private void UpdateCommandSelection()
        {
            for (int i = 0; i < commandList.Children.Count; i++)
            {
                if (commandList.Children[i] is Button btn)
                {
                    btn.Background = i == _selectedCommandIndex
                        ? System.Windows.Media.Brushes.LightGray
                        : System.Windows.Media.Brushes.Transparent;
                }
            }
        }

        private void BtnSend_Click(object sender, RoutedEventArgs e) => SendInput();

        // ========== 斜杠命令 ==========

        private static readonly (string Command, string Description, string Action)[] SlashCommands =
        {
            ("/clear",  UiText.T("I18n_UiClearPane"),   "clear"),
            ("/newchat", UiText.T("I18n_UiNewChat"),    "newchat"),
            ("/review", UiText.ReviewSelection,    "review"),
            ("/refreshskills", UiText.T("I18n_UiRefreshSkills"), "refreshskills"),
            ("/make-skill", UiText.T("I18n_MakeSkill"), "makeskill"),
        };

        private string SkillsRootPath => System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "WordAgentSkills");

        /// <summary>斜杠技能前缀，用于区分技能项和系统命令。</summary>
        private const string SkillPrefix = "/skill:";

        private void TxtInput_TextChanged(object sender, TextChangedEventArgs e)
        {
            string text = txtInput.Text?.TrimStart() ?? "";

            if (text.StartsWith("/") && !text.Contains("\n"))
            {
                commandList.Children.Clear();
                _selectedCommandIndex = -1;

                // 显示系统命令
                foreach (var (cmd, desc, _) in SlashCommands)
                {
                    if (cmd.StartsWith(text, StringComparison.OrdinalIgnoreCase))
                    {
                        AddCommandItem(cmd, desc, cmd);
                    }
                }

                // 显示可用技能（以 /skill:name 形式展示）
                foreach (var skill in _skillManager.GetAllSkills())
                {
                    string skillCmd = SkillPrefix + skill.Name;
                    if (skillCmd.StartsWith(text, StringComparison.OrdinalIgnoreCase) ||
                        ("/" + skill.Name).StartsWith(text, StringComparison.OrdinalIgnoreCase))
                    {
                        AddSkillItem(skill.Name, skill.Description);
                    }
                }

                commandPopup.IsOpen = commandList.Children.Count > 0;
                if (commandPopup.IsOpen && _selectedCommandIndex < 0 && commandList.Children.Count > 0)
                {
                    _selectedCommandIndex = 0;
                    UpdateCommandSelection();
                }
            }
            else
            {
                commandPopup.IsOpen = false;
            }
        }

        private void AddCommandItem(string display, string desc, string tag)
        {
            var item = new Button
            {
                Content = CreateCommandItemContent(display, desc),
                Style = (Style)FindResource("MaterialDesignFlatButton"),
                HorizontalContentAlignment = HorizontalAlignment.Stretch,
                Tag = tag,
                Cursor = System.Windows.Input.Cursors.Hand,
            };
            item.Click += CommandItem_Click;
            commandList.Children.Add(item);
        }

        private void AddSkillItem(string skillName, string _)
        {
            var item = new Button
            {
                Content = CreateCommandItemContent("📋 " + skillName, "skill"),
                Style = (Style)FindResource("MaterialDesignFlatButton"),
                HorizontalContentAlignment = HorizontalAlignment.Stretch,
                Tag = SkillPrefix + skillName,
                Cursor = System.Windows.Input.Cursors.Hand,
            };
            item.Click += SkillItem_Click;
            commandList.Children.Add(item);
        }

        private StackPanel CreateCommandItemContent(string cmd, string desc)
        {
            var sp = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(8, 4, 8, 4) };
            sp.Children.Add(new TextBlock { Text = cmd, FontWeight = FontWeights.SemiBold, FontSize = 13, Width = 90 });
            sp.Children.Add(new TextBlock { Text = desc, Opacity = 0.6, FontSize = 12, VerticalAlignment = VerticalAlignment.Center });
            return sp;
        }

        private void CommandItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string cmd)
            {
                commandPopup.IsOpen = false;
                txtInput.Text = "";
                ExecuteSlashCommand(cmd);
            }
        }

        /// <summary>技能选择 — 在输入框插入提示前缀，让用户继续输入具体需求。</summary>
        private void SkillItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string tag && tag.StartsWith(SkillPrefix))
            {
                commandPopup.IsOpen = false;
                string skillName = tag.Substring(SkillPrefix.Length);
                txtInput.Text = UiText.F("I18n_UiUseSkill", skillName);
                txtInput.CaretIndex = txtInput.Text.Length;
                txtInput.Focus();
            }
        }

        private void ExecuteSlashCommand(string cmd)
        {
            // 如果是技能命令（用户直接输入 /skill:xxx 回车）
            if (cmd.StartsWith(SkillPrefix))
            {
                string skillName = cmd.Substring(SkillPrefix.Length);
                txtInput.Text = UiText.F("I18n_UiUseSkill", skillName);
                txtInput.CaretIndex = txtInput.Text.Length;
                txtInput.Focus();
                return;
            }

            var match = System.Array.Find(SlashCommands, c =>
                string.Equals(c.Command, cmd, StringComparison.OrdinalIgnoreCase));
            if (match.Action == null) return;

            switch (match.Action)
            {
                case "clear":
                    // 清空窗口显示，但保留 Agent 内部对话上下文
                    _chatItems.Clear();
                    _bufferedThought = null;
                    _bufferedAction = null;
                    _executionProcessItem = null;
                    _finalAnswerItem = null;
                    _executionProcessEntryCount = 0;
                    break;

                case "newchat":
                    // 清空窗口 + 重置 Agent
                    _chatItems.Clear();
                    _bufferedThought = null;
                    _bufferedAction = null;
                    _executionProcessItem = null;
                    _finalAnswerItem = null;
                    _executionProcessEntryCount = 0;
                    _agent?.CancelPendingInteractions();
                    _cts?.Cancel();
                    _agent = null;
                    _pendingRequirement = null;
                    _pendingImageBase64 = null;
                    _isWaitingForResponse = false;
                    lib.ConfirmationManager.ClearExternalPermissionsForSession();
                    btnCancelTask.Visibility = Visibility.Collapsed;
                    statusBarCard.Visibility = Visibility.Collapsed;
                    typingIndicator.Visibility = Visibility.Collapsed;
                    break;

                case "review":
                    _ = RunManualReviewAsync(selectionOnly: true);
                    break;

                case "refreshskills":
                    // 重新加载技能
                    _skillManager.LoadSkills(SkillsRootPath);
                    _agent?.RefreshSkills();
                    AddAgentBubble(UiText.F("I18n_UiSkillsReloaded", _skillManager.Count));
                    break;
                case "makeskill":
                    PrepareMakeSkillRequest();
                    break;
            }
        }

        private void SendInput()
        {
            string text = txtInput.Text?.Trim();
            if (string.IsNullOrEmpty(text)) return;

            // 斜杠命令拦截（包括 /skill:xxx）
            if (text.StartsWith("/"))
            {
                commandPopup.IsOpen = false;

                // 技能选择格式
                if (text.StartsWith(SkillPrefix, StringComparison.OrdinalIgnoreCase))
                {
                    string skillName = text.Substring(SkillPrefix.Length).Trim();
                    if (_skillManager.HasSkill(skillName))
                    {
                        txtInput.Text = UiText.F("I18n_UiUseSkill", skillName);
                        txtInput.CaretIndex = txtInput.Text.Length;
                        txtInput.Focus();
                        return;
                    }
                }

                var match = System.Array.Find(SlashCommands, c =>
                    string.Equals(c.Command, text, StringComparison.OrdinalIgnoreCase));
                if (match.Action != null)
                {
                    txtInput.Text = "";
                    ExecuteSlashCommand(text);
                    return;
                }

                // 如果输入 /xxx 且是一个技能名，也插入前缀
                string bareName = text.TrimStart('/').Trim();
                if (_skillManager.HasSkill(bareName))
                {
                    txtInput.Text = UiText.F("I18n_UiUseSkill", bareName);
                    txtInput.CaretIndex = txtInput.Text.Length;
                    txtInput.Focus();
                    return;
                }
            }

            // [HITL] 如果正在等待用户回答，提交回答而非启动新 Agent
            if (_isWaitingForResponse && _agent != null)
            {
                // 显示用户回答气泡
                _chatItems.Add(new ChatItem { MessageType = ChatMessageType.User, Message = text });
                txtInput.Text = "";
                ScrollToEnd();

                _isWaitingForResponse = false;
                _agent.SetUserResponse(text);

                AppendExecutionProcess(Properties.Resources.Agent_AnswerReceived);
                return;
            }

            // 如果 Agent 正在运行（包括等待用户回答），用户输入新需求 → 打断当前任务
            if (_isRunning)
            {
                // 取消当前任务并释放可能挂起的 TCS
                _agent?.CancelPendingInteractions();
                _isWaitingForResponse = false;
                _cts?.Cancel();

                // 标记被中断
                AppendExecutionProcess(Properties.Resources.Agent_Interrupted);

                // 显示用户新需求气泡（含粘贴图片）
                var userItem = new ChatItem { MessageType = ChatMessageType.User, Message = text };
                if (!string.IsNullOrEmpty(_pastedImageBase64))
                {
                    userItem.ImageSource = pastePreviewImage.Source;
                    userItem.ImageBase64 = _pastedImageBase64;
                }
                _chatItems.Add(userItem);
                txtInput.Text = "";
                ScrollToEnd();

                // 设置待发送图片并清除预览
                string pendingImage = _pastedImageBase64;
                RemovePasteImage();

                // 由旧任务的 finally 在真正退出后启动新任务，避免复用已被取消的 CTS。
                _pendingRequirement = text;
                _pendingImageBase64 = pendingImage;
                return;
            }

            // 显示用户气泡（正常启动，含粘贴图片）
            var newUserItem = new ChatItem { MessageType = ChatMessageType.User, Message = text };
            if (!string.IsNullOrEmpty(_pastedImageBase64))
            {
                newUserItem.ImageSource = pastePreviewImage.Source;
                newUserItem.ImageBase64 = _pastedImageBase64;
                DocumentAgent.PendingImageBase64 = _pastedImageBase64;
                RemovePasteImage();
            }
            _chatItems.Add(newUserItem);
            txtInput.Text = "";
            ScrollToEnd();

            // 启动 Agent
            ExecuteAgent(text);
        }

        private async void ExecuteAgent(string requirement)
        {
            if (!Dispatcher.CheckAccess())
            {
                _ = Dispatcher.BeginInvoke(new Action(() => ExecuteAgent(requirement)));
                return;
            }

            if (_isRunning) return;
            _isRunning = true;
            _cts = new CancellationTokenSource();
            // Capture WPF-owned values before the first await. Provider and MCP continuations
            // are not required to return to this control's Dispatcher.
            string agentGoal = txtAgentGoal.Text;

            // 重置步骤卡片计数器
            _bufferedThought = null;
            _bufferedAction = null;
            _executionProcessItem = null;
            _finalAnswerItem = null;
            _executionProcessEntryCount = 0;
            _agentSources.Clear();
            sourcesExpander.Visibility = Visibility.Collapsed;
            btnMakeSkill.Visibility = Visibility.Collapsed;
            McpToolManager mcpManager = null;
            bool externalToolsEnabled = SharedSettings.AgentExternalToolsEnabled;

            if (!externalToolsEnabled)
            {
                externalToolsEnabled = Dispatcher.Invoke(() => MessageBox.Show(
                    UiText.T("I18n_UiAgentExternalToolsPrompt"),
                    UiText.T("I18n_UiAgentExternalToolsTitle"),
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question)) == MessageBoxResult.Yes;
            }

            Dispatcher.Invoke(() =>
            {
                btnCancelTask.Visibility = Visibility.Visible;
                statusBarCard.Visibility = Visibility.Visible;
                UpdateStatusIcon(true);
                lblStatus.Text = Properties.Resources.Agent_StatusThinking;
                lblIteration.Text = "";
                lblTaskSummary.Text = CompactTaskSummary(requirement);
                // 显示打字指示器
                typingIndicator.Visibility = Visibility.Visible;
            });

            try
            {
                // 仅在本任务获得外部工具授权后连接 MCP 服务器。
                if (externalToolsEnabled)
                    mcpManager = await Task.Run(() => McpToolManager.LoadConfigAndConnectAsync());

                _agent = new DocumentAgent(
                    (chatHistory, ct) => Globals.ThisAddIn.GenerateModelSeletorChat(chatHistory),
                    (chatHistory, tools, ct) => Globals.ThisAddIn.GenerateModelSeletorChatWithTools(chatHistory, tools, ct),
                    externalToolsEnabled);
                _agent.StateChanged += snapshot =>
                    Dispatcher.Invoke(() => SetWorkflowStatus(snapshot));
                _agent.SetReviewWorkflow(ExecuteReviewAgentToolAsync);
                _agent.MaxIterations = SharedSettings.AgentMaxIterations;
                if (mcpManager != null && mcpManager.ServerCount > 0)
                {
                    _agent.SetMcpManager(mcpManager);
                    Dispatcher.Invoke(RefreshAgentApiDocumentation);
                }

                // [Checkpoint] 断点恢复检测
                bool resumeCheckpoint = false;
                if (DocumentAgent.HasCheckpoint())
                {
                    bool resume = Dispatcher.Invoke(() => MessageBox.Show(
                        UiText.T("I18n_UiResumeCheckpointPrompt"),
                        UiText.T("I18n_UiResumeCheckpointTitle"),
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question)) == MessageBoxResult.Yes;

                    if (resume)
                    {
                        if (_agent.TryLoadCheckpoint(out var originalReq))
                        {
                            requirement = originalReq;
                            resumeCheckpoint = true;
                            AppendExecutionProcess(UiText.T("I18n_UiCheckpointResuming"));
                        }
                    }
                }

                Dispatcher.Invoke(() => lblTaskSummary.Text = CompactTaskSummary(requirement));

                // 每条日志 → 缓存合并为步骤卡片或直接显示
                _agent.LogEntry += msg =>
                    Dispatcher.Invoke(() =>
                    {
                        typingIndicator.Visibility = Visibility.Collapsed;

                        // === 步骤卡片缓存合并 ===
                        // 💭 思考 → 缓存，等待后续 🔧 和 👀
                        if (msg.StartsWith("💭"))
                        {
                            FlushStepBuffer(); // 前一步不完整则先冲刷
                            _bufferedThought = msg;
                            lblStatus.Text = Properties.Resources.Agent_StatusThinking;
                            return;
                        }
                        // 🔧 行动 → 缓存（必须紧跟在 💭 之后）
                        if (msg.StartsWith("🔧") && _bufferedThought != null)
                        {
                            _bufferedAction = msg;
                            lblStatus.Text = Properties.Resources.Agent_StatusActing;
                            return;
                        }
                        // 👀 观察 → 合并为一张步骤卡片
                        if (msg.StartsWith("👀") && _bufferedAction != null)
                        {
                            string combined = $"{_bufferedThought}\n{_bufferedAction}\n{msg}";
                            AppendExecutionProcess(combined);
                            _bufferedThought = null;
                            _bufferedAction = null;
                            UpdateStatusBarFromLog(msg);
                            ScrollToEnd();
                            return;
                        }

                        // === 非步骤消息：错误和提问单独显示，其余并入执行过程。 ===
                        FlushStepBuffer();
                        ChatMessageType messageType = DetectMessageType(msg);
                        if (messageType == ChatMessageType.Error || messageType == ChatMessageType.Question)
                            AddAgentBubble(msg, messageType);
                        else
                            AppendExecutionProcess(msg);
                        UpdateStatusBarFromLog(msg);
                        ScrollToEnd();
                    });

                // 最终结果 → 新气泡
                _agent.ResultReady += summary =>
                    Dispatcher.Invoke(() =>
                    {
                        SetFinalAnswer(summary);
                        AddReviewIssues(ParseIssuesFromText(summary, "Agent"));
                        UpdateStatusIcon(false);
                        lblStatus.Text = _agent.WorkflowStatus.Current.Detail;
                        btnMakeSkill.Visibility = Visibility.Visible;
                        ScrollToEnd();

                        // [快照对比] 有变更时自动弹出审查窗口
                        if (_agent.DiffChanges != null && _agent.DiffChanges.Count > 0)
                        {
                            AddDiffChangesAsIssues(_agent.DiffChanges);
                            // [免打扰] 检查用户是否设置了免审查弹窗
                            if (lib.ConfirmationManager.ShouldPrompt(lib.ConfirmationManager.KeyDiffReview))
                            {
                                var dlg = new ConfirmationDialog(
                                    UiText.F("I18n_UiAgentChangesDetected", _agent.DiffChanges.Count),
                                    UiText.T("I18n_UiDocumentChangeReview"),
                                    lib.ConfirmationManager.KeyDiffReview);
                                if (dlg.ShowDialog() == true)
                                {
                                    var win = new DiffReviewWindow(_agent, _agent.DiffChanges);
                                    win.ShowDialog();
                                }
                            }
                        }
                    });

                _agent.SourceDiscovered += source => Dispatcher.Invoke(() =>
                {
                    _agentSources.Add(source);
                    sourcesExpander.Visibility = Visibility.Visible;
                });
                _agent.MetricsChanged += metrics => Dispatcher.Invoke(() =>
                {
                    lblIteration.Text = string.Format("{0} · {1:mm\\:ss} · {2} tools · {3} sources",
                        metrics.Iteration, metrics.Elapsed, metrics.ToolCalls, metrics.Sources);
                });

                // [HITL] Agent 提问 → 新气泡
                _agent.QuestionAsked += question =>
                    Dispatcher.Invoke(() =>
                    {
                        _isWaitingForResponse = true;
                        AddAgentBubble(question, ChatMessageType.Question);
                        UpdateStatusIcon(null);
                        lblStatus.Text = Properties.Resources.Agent_StatusWaiting;
                        ScrollToEnd();
                    });

                _agent.ExternalPermissionRequested += request =>
                    Dispatcher.Invoke(() =>
                    {
                        string message = UiText.F(
                            "I18n_AgentPermissionMessage",
                            request.ToolName,
                            request.Operation,
                            request.EffectDescription);
                        if (!request.AllowSession &&
                            string.Equals(request.PermissionKey, "execute_command|dangerous", StringComparison.Ordinal))
                        {
                            message = UiText.F("I18n_AgentDangerousPermissionMessage", request.Operation);
                        }

                        var dialog = new AgentPermissionDialog(message, request.AllowSession);
                        var wordWindow = Process.GetCurrentProcess().MainWindowHandle;
                        if (wordWindow != IntPtr.Zero)
                            new WindowInteropHelper(dialog).Owner = wordWindow;
                        dialog.ShowDialog();
                        _agent?.SetExternalPermissionDecision(dialog.Decision);
                    });

                // [计划确认] 展示计划并等待用户确认
                _agent.PlanReady += plan =>
                    Dispatcher.Invoke(() =>
                    {
                        // [免打扰] 用户已设置计划免确认 → 自动确认
                        if (!lib.ConfirmationManager.ShouldPrompt(lib.ConfirmationManager.KeyPlanConfirmation))
                        {
                            _agent.ConfirmPlan();
                            return;
                        }
                        AddAgentBubble(plan, ChatMessageType.Plan);
                        ScrollToEnd();
                    });

                await _agent.ExecuteAsync(requirement, _cts.Token, resumeCheckpoint, agentGoal);
            }
            catch (OperationCanceledException)
            {
                SetWorkflowStatus(WorkflowState.Cancelled, Properties.Resources.Agent_Cancelled);
                Dispatcher.Invoke(() =>
                {
                    AddAgentBubble(Properties.Resources.Agent_Cancelled, ChatMessageType.System);
                });
                _agent?.CancelPendingInteractions();
            }
            catch (Exception ex)
            {
                DebugLogger.Log("Agent.Error", ex.ToString());
                System.Diagnostics.Debug.WriteLine(ex);
                SetWorkflowStatus(WorkflowState.Failed, string.Format(Properties.Resources.Agent_Error, ex.Message));
                Dispatcher.Invoke(() =>
                {
                    AddAgentBubble(string.Format(Properties.Resources.Agent_Error, ex.Message), ChatMessageType.Error);
                });
                // 立即隐藏状态栏，不等 finally（异常可能已破坏正常流程）
                Dispatcher.Invoke(() =>
                {
                    statusBarCard.Visibility = Visibility.Collapsed;
                    typingIndicator.Visibility = Visibility.Collapsed;
                });
            }
            finally
            {
                _isRunning = false;
                _isWaitingForResponse = false;
                DocumentAgent.PendingImageBase64 = null; // 清理待发送图片
                Dispatcher.Invoke(() =>
                {
                    btnCancelTask.Visibility = Visibility.Collapsed;
                    btnCancelTask.IsEnabled = true;
                    statusBarCard.Visibility = Visibility.Collapsed;
                    typingIndicator.Visibility = Visibility.Collapsed;
                    lblTaskSummary.Text = "";
                });
                _cts?.Dispose();
                _cts = null;
                mcpManager?.Dispose();

                if (!string.IsNullOrWhiteSpace(_pendingRequirement))
                {
                    string pendingRequirement = _pendingRequirement;
                    string pendingImage = _pendingImageBase64;
                    _pendingRequirement = null;
                    _pendingImageBase64 = null;
                    DocumentAgent.PendingImageBase64 = pendingImage;
                    _ = Dispatcher.BeginInvoke(new Action(() => ExecuteAgent(pendingRequirement)),
                        System.Windows.Threading.DispatcherPriority.Background);
                }
            }
        }

        private void BtnMakeSkill_Click(object sender, RoutedEventArgs e) => PrepareMakeSkillRequest();

        private void PrepareMakeSkillRequest()
        {
            txtInput.Text = UiText.T("I18n_MakeSkillRequest");
            txtInput.CaretIndex = txtInput.Text.Length;
            txtInput.Focus();
        }

        private void AgentSource_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string url && Uri.TryCreate(url, UriKind.Absolute, out Uri uri) && uri.Scheme == Uri.UriSchemeHttps)
            {
                try { Process.Start(new ProcessStartInfo { FileName = url, UseShellExecute = true }); } catch { }
            }
        }

        // ========== 状态栏辅助方法 ==========

        /// <summary>添加一个气泡到对话流，根据消息内容自动推断类型。</summary>
        private void AddAgentBubble(string message, ChatMessageType? type = null, string stepTitle = null)
        {
            ChatMessageType msgType = type ?? DetectMessageType(message);
            _chatItems.Add(new ChatItem
            {
                MessageType = msgType,
                Message = message,
                ShowAgentLabel = ShouldShowAgentLabel(),
                StepTitle = stepTitle ?? UiText.AgentExecutionDetails,
                IsCollapsible = msgType == ChatMessageType.StepGroup,
                IsExpanded = false
            });
        }

        /// <summary>
        /// 将过程日志汇入单个默认折叠的条目，避免每次状态变化都生成一张卡片。
        /// </summary>
        private void AppendExecutionProcess(string message)
        {
            if (string.IsNullOrWhiteSpace(message)) return;

            if (_executionProcessItem != null && !_chatItems.Contains(_executionProcessItem))
            {
                _executionProcessItem = null;
                _executionProcessEntryCount = 0;
            }

            _executionProcessEntryCount++;
            if (_executionProcessItem == null)
            {
                _executionProcessItem = new ChatItem
                {
                    MessageType = ChatMessageType.StepGroup,
                    Message = message.Trim(),
                    ShowAgentLabel = ShouldShowAgentLabel(),
                    StepTitle = $"{UiText.AgentExecutionDetails} · {_executionProcessEntryCount}",
                    IsCollapsible = true,
                    IsExpanded = false
                };
                _chatItems.Add(_executionProcessItem);
                return;
            }

            _executionProcessItem.Message = string.Concat(
                _executionProcessItem.Message?.TrimEnd(), "\n\n", message.Trim());
            _executionProcessItem.StepTitle = $"{UiText.AgentExecutionDetails} · {_executionProcessEntryCount}";
        }

        /// <summary>每次 Agent 运行只保留一个无边框的最终回答。</summary>
        private void SetFinalAnswer(string message)
        {
            if (string.IsNullOrWhiteSpace(message)) return;
            if (_finalAnswerItem == null)
            {
                _finalAnswerItem = new ChatItem
                {
                    MessageType = ChatMessageType.Result,
                    Message = message.Trim(),
                    ShowAgentLabel = ShouldShowAgentLabel()
                };
                _chatItems.Add(_finalAnswerItem);
            }
            else
            {
                _finalAnswerItem.Message = message.Trim();
            }
        }

        private bool ShouldShowAgentLabel()
        {
            return _chatItems.Count == 0 || _chatItems[_chatItems.Count - 1].IsUser;
        }

        private void UpdateEmptyChatState()
        {
            if (emptyChatState != null)
                emptyChatState.Visibility = _chatItems.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        private static string CompactTaskSummary(string requirement)
        {
            string summary = Regex.Replace(requirement ?? "", @"\s+", " ").Trim();
            return summary.Length > 160 ? summary.Substring(0, 157) + "..." : summary;
        }

        /// <summary>冲刷缓存的步骤消息（前一步不完整时，以独立气泡显示）。</summary>
        private void FlushStepBuffer()
        {
            if (_bufferedThought != null)
            {
                AppendExecutionProcess(_bufferedThought);
                _bufferedThought = null;
            }
            if (_bufferedAction != null)
            {
                AppendExecutionProcess(_bufferedAction);
                _bufferedAction = null;
            }
        }

        /// <summary>根据消息内容前缀推断气泡类型。</summary>
        private ChatMessageType DetectMessageType(string msg)
        {
            if (string.IsNullOrEmpty(msg)) return ChatMessageType.System;
            if (msg.StartsWith("💭")) return ChatMessageType.Thought;
            if (msg.StartsWith("🔧")) return ChatMessageType.Action;
            if (msg.StartsWith("⚠️") || msg.StartsWith("🚫") || msg.StartsWith("❌"))
                return ChatMessageType.Error;
            if (msg.StartsWith("异常", StringComparison.OrdinalIgnoreCase) ||
                msg.StartsWith("错误", StringComparison.OrdinalIgnoreCase) ||
                msg.StartsWith("Error", StringComparison.OrdinalIgnoreCase) ||
                msg.StartsWith("Exception", StringComparison.OrdinalIgnoreCase) ||
                msg.StartsWith("[API", StringComparison.OrdinalIgnoreCase))
                return ChatMessageType.Error;
            if (msg.StartsWith("🤖")) return ChatMessageType.Question;
            if (msg.StartsWith("🎉") || msg.StartsWith("✅")) return ChatMessageType.Result;
            if (msg.StartsWith("🚀") || msg.StartsWith("📦") || msg.StartsWith("📄") ||
                msg.StartsWith("📌") || msg.StartsWith("ℹ️"))
                return ChatMessageType.System;
            return ChatMessageType.Result; // 默认归为结果
        }

        private void UpdateStatusIcon(bool? running)
        {
            if (running == true)
            {
                statusIcon.Kind = MaterialDesignThemes.Wpf.PackIconKind.PlayCircle;
                var brush = System.Windows.Application.Current?.Resources["PrimaryHueMidBrush"] as System.Windows.Media.Brush;
                statusIcon.Foreground = brush ?? System.Windows.Media.Brushes.Indigo;
            }
            else if (running == false)
            {
                statusIcon.Kind = MaterialDesignThemes.Wpf.PackIconKind.CheckCircle;
                statusIcon.Foreground = System.Windows.Media.Brushes.Green;
            }
            else // null = waiting
            {
                statusIcon.Kind = MaterialDesignThemes.Wpf.PackIconKind.CommentQuestion;
                statusIcon.Foreground = System.Windows.Media.Brushes.Orange;
            }
        }

        /// <summary>根据日志内容推断 Agent 当前阶段，更新状态栏。</summary>
        private void UpdateStatusBarFromLog(string msg)
        {
            if (msg.StartsWith("💭"))
                lblStatus.Text = Properties.Resources.Agent_StatusThinking;
            else if (msg.StartsWith("🔧"))
                lblStatus.Text = Properties.Resources.Agent_StatusActing;
            else if (msg.StartsWith("👀"))
                lblStatus.Text = Properties.Resources.Agent_StatusObserving;
            else if (msg.StartsWith("🚀"))
                lblStatus.Text = Properties.Resources.Agent_StatusStarting;

            // 更新迭代计数
            if (msg.StartsWith("[") && msg.Contains("]"))
            {
                int bracketEnd = msg.IndexOf(']');
                if (bracketEnd > 1)
                {
                    string num = msg.Substring(1, bracketEnd - 1);
                    if (int.TryParse(num, out int iter))
                        lblIteration.Text = string.Format(Properties.Resources.Agent_Iteration, iter);
                }
            }
        }

        private void PlanButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string action && _agent != null)
            {
                if (action == "confirm")
                {
                    // 尚未设置免打扰时，弹出带免打扰选项的确认框
                    if (lib.ConfirmationManager.ShouldPrompt(lib.ConfirmationManager.KeyPlanConfirmation))
                    {
                        var dlg = new ConfirmationDialog(
                            UiText.T("I18n_UiPlanConfirmPrompt"),
                            UiText.T("I18n_UiPlanConfirmTitle"),
                            lib.ConfirmationManager.KeyPlanConfirmation);
                        if (dlg.ShowDialog() == true)
                            _agent.ConfirmPlan();
                    }
                    else
                    {
                        _agent.ConfirmPlan();
                    }
                }
                else
                    _agent.RejectPlan();
            }
        }

        private void BtnRemovePaste_Click(object sender, RoutedEventArgs e)
        {
            RemovePasteImage();
        }

        private void BtnCancelTask_Click(object sender, RoutedEventArgs e)
        {
            _isWaitingForResponse = false;
            _agent?.CancelPendingInteractions();
            _cts?.Cancel();
            btnCancelTask.IsEnabled = false;
        }

        // ========== 工具方法 ==========

        public void ClearChat()
        {
            Dispatcher.Invoke(() =>
            {
                _chatItems.Clear();
                _executionProcessItem = null;
                _finalAnswerItem = null;
                _executionProcessEntryCount = 0;
            });
        }

        /// <summary>供 SilentLinter 等外部组件调用 — 添加 Agent 气泡。</summary>
        public void StartAgentReply()
        {
            Dispatcher.Invoke(() =>
            {
                _executionProcessItem = null;
                _finalAnswerItem = null;
                _executionProcessEntryCount = 0;
            });
        }

        /// <summary>供外部组件调用 — 添加 Agent 气泡。</summary>
        public void AppendLog(string message)
        {
            Dispatcher.Invoke(() =>
            {
                ChatMessageType type = DetectMessageType(message);
                if (type == ChatMessageType.Error || type == ChatMessageType.Question)
                    AddAgentBubble(message, type);
                else
                    AppendExecutionProcess(message);
                ScrollToEnd();
            });
        }

        /// <summary>供外部组件调用 — 添加最终结论气泡。</summary>
        public void SetFinalMessage(string message)
        {
            Dispatcher.Invoke(() =>
            {
                SetFinalAnswer(message);
                ScrollToEnd();
            });
        }

        private void ScrollToEnd()
        {
            Dispatcher.Invoke(() =>
            {
                if (_chatItems.Count > 0)
                    scrollViewer.ScrollToEnd();
            });
        }

        // ========== 非破坏式审阅工作台 ==========

        public void OpenReviewWorkspace(bool loadSelection = true, string instruction = null, bool replaceExisting = false)
        {
            Dispatcher.Invoke(new Action(async () =>
            {
                if (_reviewCts != null || _isRunning || _isLoadingReviewScope) return;
                SelectReviewTab();
                if (!string.IsNullOrWhiteSpace(instruction))
                    txtReviewInstruction.Text = instruction.Trim();

                if (replaceExisting)
                    _editSuggestions.Clear();

                bool shouldGenerate = false;
                if (loadSelection && (_editSuggestions.Count == 0 || replaceExisting))
                {
                    await LoadSelectionForReview(splitParagraphs: true);
                    shouldGenerate = _editSuggestions.Count > 0;
                }

                if (shouldGenerate)
                    await GenerateAllSuggestionsAsync();
            }));
        }

        private void SelectReviewTab()
        {
            if (mainTabs != null)
                mainTabs.SelectedItem = ReviewTabItem;
        }

        private async void BtnLoadReviewSelection_Click(object sender, RoutedEventArgs e)
        {
            await LoadSelectionForReview(splitParagraphs: false);
        }

        private async void BtnLoadReviewParagraphs_Click(object sender, RoutedEventArgs e)
        {
            await LoadCurrentParagraphForReview();
        }

        private async void BtnLoadReviewDocument_Click(object sender, RoutedEventArgs e)
        {
            await LoadDocumentForReviewAsync();
        }

        private void BtnClearSuggestions_Click(object sender, RoutedEventArgs e)
        {
            if (_reviewCts != null || _isLoadingReviewScope) return;
            FlushReviewSuggestionEdits();
            foreach (var session in _editSuggestions.ToList())
            {
                RecordSuggestionAction(session, "cleared");
                DeleteSuggestionBookmark(session);
            }
            _editSuggestions.Clear();
            UpdateSuggestionEmptyState();
        }

        private async void OptimizeTextBox_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            await Utils.OptimizeTextBoxAsync(button?.Tag as TextBox, button);
        }

        private void BtnSaveWritingProfile_Click(object sender, RoutedEventArgs e)
        {
            _writingProfile = new WritingProfile { Instructions = txtWritingProfile.Text?.Trim() ?? "" };
            var document = Globals.ThisAddIn?.Application?.ActiveDocument;
            WritingProfileManager.SaveForDocument(document, _writingProfile);
            AddAgentBubble(UiText.T("I18n_UiWritingProfileSaved"), ChatMessageType.System);
        }

        private async Task LoadSelectionForReview(bool splitParagraphs)
        {
            if (_reviewCts != null || _isRunning || _isLoadingReviewScope) return;
            SetReviewScopeLoading(true);
            Word.Range range = null;
            try
            {
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;
                if (doc == null || app.Selection == null)
                    throw new InvalidOperationException(UiText.T("I18n_UiOpenDocumentFirst"));
                range = app.Selection.Range.Duplicate;
                await LoadReviewRangeAsync(doc, range, splitParagraphs);
            }
            catch (Exception ex)
            {
                _editSuggestions.Clear();
                SetReviewScopeStatus(UiText.F("I18n_UiLoadSelectionFailed", ex.Message));
            }
            finally { ComHelper.ReleaseCom(ref range); SetReviewScopeLoading(false); }
        }

        private async Task LoadDocumentForReviewAsync()
        {
            if (_isLoadingReviewScope || _reviewCts != null || _isRunning && !_agentReviewWorkflowActive) return;
            SetReviewScopeLoading(true);
            var stories = new List<ReviewStoryRange>();
            try
            {
                var doc = Globals.ThisAddIn.Application.ActiveDocument;
                if (doc == null) throw new InvalidOperationException(UiText.T("I18n_UiOpenDocumentFirst"));
                string stamp = ReviewDocumentHash(doc);
                FlushReviewSuggestionEdits(); _suggestionRecord = null; _editSuggestions.Clear();
                DeleteSuggestionBookmarks(doc);
                stories = ReviewDocumentScope.Capture(doc);
                foreach (var story in stories)
                {
                    if (_agentReviewWorkflowActive) _agentReviewToken.ThrowIfCancellationRequested();
                    await LoadReviewRangeAsync(doc, story.Range, true, append: true, persist: false, story: story.Address);
                    await Dispatcher.Yield(DispatcherPriority.Background);
                }
                if (stamp != ReviewDocumentHash(doc)) throw new InvalidOperationException(UiText.T("I18n_ReviewSourceChanged"));
                using (var body = new ReviewStoryRange { Range = doc.Content })
                    await CreateSuggestionRecordAsync(doc, body.Range, wholeDocument: true);
                UpdateSuggestionEmptyState();
                SetReviewScopeStatus(UiText.F("I18n_ReviewLoadedScope", _editSuggestions.Count, _editSuggestions.Sum(i => i.OriginalText.Length)));
            }
            catch (Exception ex)
            {
                _editSuggestions.Clear(); SetReviewScopeStatus(ex.Message);
                if (_agentReviewWorkflowActive) throw;
            }
            finally { foreach (var story in stories) story.Dispose(); SetReviewScopeLoading(false); }
        }

        private async Task LoadCurrentParagraphForReview()
        {
            if (_isLoadingReviewScope || _reviewCts != null || _isRunning) return;
            SetReviewScopeLoading(true);
            Word.Range range = null;
            try
            {
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;
                if (doc == null || app.Selection == null) throw new InvalidOperationException(UiText.T("I18n_UiOpenDocumentFirst"));
                range = app.Selection.Paragraphs[1].Range.Duplicate;
                await LoadReviewRangeAsync(doc, range, splitParagraphs: false);
            }
            catch (Exception ex)
            {
                _editSuggestions.Clear();
                SetReviewScopeStatus(UiText.F("I18n_UiLoadSelectionFailed", ex.Message));
            }
            finally { ComHelper.ReleaseCom(ref range); SetReviewScopeLoading(false); }
        }

        private async Task LoadReviewRangeAsync(Word.Document doc, Word.Range reviewRange, bool splitParagraphs, bool append = false, bool persist = true, ReviewStoryAddress story = null)
        {
            story = story ?? ReviewDocumentScope.AddressFor(doc, reviewRange);
            string documentName = doc.FullName;
            string sourceStamp = EditSuggestionSession.Hash(reviewRange.Text);
            string selectedText = EditSuggestionSession.NormalizeWordText(reviewRange?.Text);
            FlushReviewSuggestionEdits();
            if (!append)
            {
                _suggestionRecord = null; _editSuggestions.Clear(); DeleteSuggestionBookmarks(doc);
            }
            UpdateSuggestionEmptyState();
            if (string.IsNullOrWhiteSpace(selectedText))
            {
                SetReviewScopeStatus(UiText.ReviewScopeNone);
                return;
            }

            int index = _editSuggestions.Count + 1;

            int paragraphCount = reviewRange.Paragraphs.Count;
            if (splitParagraphs && paragraphCount > 1)
            {
                int paragraphIndex = FindParagraphIndexByRange(doc, reviewRange);
                bool createBookmarks = paragraphCount <= 200;
                var paragraphUnits = new List<ReviewParagraphUnit>();
                foreach (Word.Paragraph paragraph in reviewRange.Paragraphs)
                {
                    Word.Paragraph paragraphToRelease = paragraph;
                    Word.Range range = null;
                    try
                    {
                        if (_agentReviewWorkflowActive) _agentReviewToken.ThrowIfCancellationRequested();
                        range = paragraph.Range.Duplicate;
                        range.SetRange(Math.Max(reviewRange.Start, range.Start), Math.Min(reviewRange.End, range.End));
                        string text = EditSuggestionSession.NormalizeWordText(range.Text);
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            paragraphUnits.Add(new ReviewParagraphUnit
                            {
                                ParagraphIndex = paragraphIndex,
                                Start = range.Start,
                                End = range.End,
                                Text = text
                            });
                        }
                    }
                    finally
                    {
                        ComHelper.ReleaseCom(ref range);
                        ComHelper.ReleaseCom(ref paragraphToRelease);
                    }
                    paragraphIndex++;
                    if (paragraphIndex % 50 == 0)
                    {
                        SetReviewScopeStatus(UiText.F("I18n_ReviewReadingProgress", paragraphUnits.Count, paragraphCount));
                        await Dispatcher.Yield(DispatcherPriority.Background);
                    }
                }

                // Keep short paragraphs, but review them with the following paragraph so
                // headings and sentence fragments have enough context for the model.
                for (int unitIndex = 0; unitIndex < paragraphUnits.Count;)
                {
                    if (_agentReviewWorkflowActive) _agentReviewToken.ThrowIfCancellationRequested();
                    int firstIndex = unitIndex;
                    int lastIndex = unitIndex;
                    int mergedCharacters = paragraphUnits[firstIndex].Text.Length;
                    if (mergedCharacters < ReviewMinimumParagraphCharacters && lastIndex + 1 < paragraphUnits.Count)
                    {
                        while (lastIndex + 1 < paragraphUnits.Count
                            && (lastIndex == firstIndex
                                || mergedCharacters + paragraphUnits[lastIndex + 1].Text.Length + 1 <= ReviewMergedParagraphMaxCharacters))
                        {
                            lastIndex++;
                            mergedCharacters += paragraphUnits[lastIndex].Text.Length + 1;
                            if (mergedCharacters >= ReviewMinimumParagraphCharacters)
                                break;
                        }
                    }

                    Word.Range mergedRange = null;
                    try
                    {
                        mergedRange = reviewRange.Duplicate;
                        mergedRange.SetRange(paragraphUnits[firstIndex].Start, paragraphUnits[lastIndex].End);
                        string mergedText = EditSuggestionSession.NormalizeWordText(mergedRange.Text);
                        _editSuggestions.Add(CreateSuggestionSession(
                            doc,
                            mergedRange,
                            mergedText,
                            index++,
                            paragraphUnits[firstIndex].ParagraphIndex,
                            createBookmarks, story));
                    }
                    finally
                    {
                        ComHelper.ReleaseCom(ref mergedRange);
                    }

                    unitIndex = lastIndex + 1;
                    if (unitIndex % 50 == 0) await Dispatcher.Yield(DispatcherPriority.Background);
                }
            }
            else
            {
                Word.Range range = null;
                try
                {
                    range = reviewRange.Duplicate;
                    _editSuggestions.Add(CreateSuggestionSession(
                        doc,
                        range,
                        selectedText,
                        index,
                        FindParagraphIndexByRange(doc, reviewRange),
                        createBookmark: true, story: story));
                }
                finally
                {
                    ComHelper.ReleaseCom(ref range);
                }
            }

            if (sourceStamp != EditSuggestionSession.Hash(reviewRange.Text)
                || !string.Equals(documentName, Globals.ThisAddIn.Application.ActiveDocument?.FullName, StringComparison.OrdinalIgnoreCase))
            {
                DeleteSuggestionBookmarks(doc);
                _editSuggestions.Clear();
                throw new InvalidOperationException(UiText.T("I18n_ReviewSourceChanged"));
            }
            UpdateSuggestionEmptyState();
            SetReviewScopeStatus(UiText.F("I18n_ReviewLoadedScope", _editSuggestions.Count, selectedText.Length));
            if (persist) await CreateSuggestionRecordAsync(doc, reviewRange);
            SelectReviewTab();
        }

        private void SetReviewScopeStatus(string message)
        {
            if (txtReviewScopeStatus != null)
                txtReviewScopeStatus.Text = message ?? "";
        }

        private void SetReviewScopeLoading(bool isLoading)
        {
            _isLoadingReviewScope = isLoading;
            if (reviewLoadIndicator != null)
                reviewLoadIndicator.Visibility = isLoading ? Visibility.Visible : Visibility.Collapsed;
            if (reviewScopeActions != null)
                reviewScopeActions.IsEnabled = !isLoading;
            if (isLoading)
                SetReviewScopeStatus(UiText.LoadingReviewScope);
            SetReviewSessionActionsEnabled(!isLoading && _reviewCts == null && !_isRunning);
        }

        public void ShowExternalNotification(string message)
        {
            if (string.IsNullOrWhiteSpace(message)) return;
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.BeginInvoke(new Action(() => ShowExternalNotification(message)));
                return;
            }

            statusBarCard.Visibility = Visibility.Visible;
            UpdateStatusIcon(true);
            lblStatus.Text = message;
        }

        private EditSuggestionSession CreateSuggestionSession(
            Word.Document doc,
            Word.Range range,
            string text,
            int index,
            int paragraphIndex,
            bool createBookmark, ReviewStoryAddress story = null)
        {
            var session = new EditSuggestionSession
            {
                Index = index,
                ParagraphIndex = paragraphIndex, Story = story ?? ReviewDocumentScope.AddressFor(doc, range),
                RangeStart = range.Start,
                RangeEnd = range.End,
                OriginalText = text,
                OriginalHash = EditSuggestionSession.Hash(text),
                DocumentIdentity = doc.FullName,
                Status = UiText.T("I18n_UiReady")
            };
            if (!createBookmark)
                return session;

            session.BookmarkName = "WO_SUG_" + session.Id.Substring(0, 32);
            Word.Range bookmarkRange = null;
            try
            {
                if (doc.Bookmarks.Exists(session.BookmarkName))
                    doc.Bookmarks[session.BookmarkName].Delete();
                bookmarkRange = range.Duplicate;
                doc.Bookmarks.Add(session.BookmarkName, bookmarkRange);
            }
            catch
            {
                session.BookmarkName = null;
            }
            finally
            {
                ComHelper.ReleaseCom(ref bookmarkRange);
            }
            return session;
        }

        private async void BtnGenerateSuggestion_Click(object sender, RoutedEventArgs e)
        {
            if ((sender as Button)?.Tag is EditSuggestionSession session)
            {
                if (_reviewCts != null || _isLoadingReviewScope || _isRunning) return;
                var generationCts = StartReviewGeneration();
                ReviewSessionRecord record = null;
                bool cancelled = false;
                string error = null;
                try
                {
                    bool continueRun = _suggestionRecord?.Metadata.StartedUtc.HasValue == true
                        && _suggestionRecord.Metadata.ConfigurationFingerprint == ReviewConfigurationFingerprint("Suggestions");
                    await PrepareSuggestionRunAsync(continueRun);
                    record = _suggestionRecord;
                    if (!continueRun)
                        foreach (var card in _editSuggestions) ResetReviewSuggestion(card);
                    var item = record.FindItem(session.Id);
                    if (item.State == ReviewItemState.Succeeded) item.Fragments.Clear();
                    await GenerateSuggestionAsync(session, generationCts.Token);
                }
                catch (OperationCanceledException)
                {
                    cancelled = true;
                }
                catch (Exception ex) { error = ex.Message; AddAgentBubble(error, ChatMessageType.System); }
                finally
                {
                    FinishReviewRecord(record, cancelled, error);
                    CompleteReviewGeneration(generationCts);
                }
            }
        }

        private async void BtnGenerateAllSuggestions_Click(object sender, RoutedEventArgs e)
        {
            await GenerateAllSuggestionsAsync();
        }

        private static void ResetReviewSuggestion(EditSuggestionSession session)
        {
            session.GenerationSucceeded = false;
            session.SuggestedText = null;
            session.FailureReason = null;
            session.Status = UiText.T("I18n_UiReady");
        }

        private async Task GenerateAllSuggestionsAsync(bool resume = false)
        {
            if (_reviewCts != null || _isLoadingReviewScope || _isRunning && !_agentReviewWorkflowActive || _editSuggestions.Count == 0) return;
            var generationCts = StartReviewGeneration();
            var sessions = new List<EditSuggestionSession>();
            ReviewSessionRecord record = null;
            bool cancelled = false;
            string error = null;
            try
            {
                await PrepareSuggestionRunAsync(resume);
                record = _suggestionRecord;
                sessions = _editSuggestions.Where(s => !resume || record.FindItem(s.Id)?.NeedsProcessing == true).ToList();
                foreach (var session in sessions) ResetReviewSuggestion(session);
                var batches = CreateReviewBatches(sessions).ToList();
                statusBarCard.Visibility = Visibility.Visible;
                UpdateStatusIcon(true);
                lblStatus.Text = UiText.T("I18n_UiGeneratingReviewSuggestion");
                lblIteration.Text = $"0/{batches.Count}";
                using (var requestGate = new SemaphoreSlim(Math.Max(1, Math.Min(8, AppConfig.Instance.ReviewMaxConcurrentRequests))))
                {
                    int completedBatches = 0;
                    var tasks = batches.Select(async batch =>
                    {
                        await requestGate.WaitAsync(generationCts.Token);
                        try
                        {
                            generationCts.Token.ThrowIfCancellationRequested();
                            ValidateRunningReviewSource(record);
                            if (batch.Count == 1) await GenerateSuggestionAsync(batch[0], generationCts.Token, manageStatusBar: false);
                            else await GenerateSuggestionBatchAsync(batch, generationCts.Token);
                            int completed = Interlocked.Increment(ref completedBatches);
                            lblIteration.Text = $"{completed}/{batches.Count}";
                            SetReviewScopeStatus(UiText.F("I18n_ReviewSuggestionProgress", sessions.Count(s => s.GenerationSucceeded), sessions.Count));
                        }
                        finally { requestGate.Release(); }
                    }).ToList();
                    await Task.WhenAll(tasks);
                }
            }
            catch (OperationCanceledException)
            {
                cancelled = true;
                AddAgentBubble(UiText.T("I18n_UiReviewQueueCancelled"), ChatMessageType.System);
            }
            catch (Exception ex)
            {
                error = ex.Message;
                AddAgentBubble(UiText.F("I18n_UiReviewFailed", ex.Message), ChatMessageType.System);
            }
            finally
            {
                FinishReviewRecord(record, cancelled, error);
                int success = record?.Metadata.CompletedItems ?? 0;
                int failed = record?.Metadata.FailedItems ?? 0;
                string summary = UiText.F("I18n_ReviewSuggestionSummary", success, failed, record?.Metadata.PendingItems ?? _editSuggestions.Count);
                SetReviewScopeStatus(summary);
                AddAgentBubble(summary, ChatMessageType.System);
                var failures = sessions.Where(s => !s.GenerationSucceeded && !string.IsNullOrEmpty(s.FailureReason)).ToList();
                if (failures.Count > 0)
                    AddAgentBubble(string.Join("\n", failures.Take(20).Select(s =>
                        UiText.F("I18n_ReviewBatchFailed", s.ParagraphIndex, s.ParagraphIndex, s.FailureReason)))
                        + (failures.Count > 20 ? "\n" + UiText.F("I18n_ReviewMoreFailures", failures.Count - 20) : ""), ChatMessageType.System);
                statusBarCard.Visibility = Visibility.Collapsed;
                typingIndicator.Visibility = Visibility.Collapsed;
                CompleteReviewGeneration(generationCts);
            }
        }

        private IEnumerable<List<EditSuggestionSession>> CreateReviewBatches(IEnumerable<EditSuggestionSession> sessions)
        {
            var config = AppConfig.Instance;
            var budget = ReviewRequestBudget.Create(config, BuildSuggestionBatchPrompt(new List<EditSuggestionSession>()), true);
            int maxItems = Math.Max(1, Math.Min(100, config.ReviewSuggestionBatchMaxItems));
            int maxCharacters = Math.Max(128, Math.Min(1000000, config.ReviewSuggestionBatchMaxCharacters));
            var batch = new List<EditSuggestionSession>();
            foreach (var session in sessions)
            {
                if (_suggestionRecord?.FindItem(session.Id)?.Fragments.Count > 0)
                {
                    if (batch.Count > 0) { yield return batch; batch = new List<EditSuggestionSession>(); }
                    yield return new List<EditSuggestionSession> { session };
                    continue;
                }
                var candidate = batch.Concat(new[] { session }).ToList();
                string text = string.Concat(candidate.Select(item => item.OriginalText));
                if (batch.Count > 0 && (candidate.Count > maxItems || text.Length > maxCharacters
                    || !budget.Fits(BuildSuggestionBatchPrompt(candidate), text, candidate.Count)))
                {
                    yield return batch;
                    batch = new List<EditSuggestionSession>();
                }
                // Oversized individual items are split by GenerateSuggestionAsync.
                batch.Add(session);
            }
            if (batch.Count > 0) yield return batch;
        }

        private string BuildSuggestionBatchPrompt(List<EditSuggestionSession> batch)
        {
            string instruction = txtReviewInstruction.Text?.Trim();
            if (string.IsNullOrWhiteSpace(instruction)) instruction = UiText.DefaultReviewInstruction;
            return PromptTemplateService.BuildReviewSuggestionBatchPrompt(
                batch.Select(session => new KeyValuePair<string, string>(session.Index.ToString(), session.OriginalText)),
                instruction, txtWritingProfile.Text?.Trim());
        }

        private async Task GenerateSuggestionBatchAsync(
            List<EditSuggestionSession> batch,
            CancellationToken cancellationToken)
        {
            var unresolved = new List<EditSuggestionSession>();
            try
            {
                foreach (EditSuggestionSession session in batch)
                {
                    session.IsBusy = true;
                    ResetReviewSuggestion(session);
                    session.Status = UiText.T("I18n_UiGenerating");
                    RecordSuggestionState(session, ReviewItemState.Running);
                }
                cancellationToken.ThrowIfCancellationRequested();
                ValidateRunningReviewSource(_suggestionRecord);
                string instruction = txtReviewInstruction.Text?.Trim();
                if (string.IsNullOrWhiteSpace(instruction))
                    instruction = UiText.DefaultReviewInstruction;
                string profile = txtWritingProfile.Text?.Trim();
                var items = batch.Select(session =>
                    new KeyValuePair<string, string>(session.Index.ToString(), session.OriginalText));
                string prompt = PromptTemplateService.BuildReviewSuggestionBatchPrompt(items, instruction, profile);
                var budget = ReviewRequestBudget.Create(AppConfig.Instance, BuildSuggestionBatchPrompt(new List<EditSuggestionSession>()), true);
                if (!budget.Fits(prompt, string.Concat(batch.Select(s => s.OriginalText)), batch.Count))
                    throw new InvalidOperationException(UiText.T("I18n_ReviewContextTooSmall"));
                string result = await Globals.ThisAddIn.GenerateModelSeletor(
                    prompt,
                    cancellationToken: cancellationToken, ollamaOutputTokens: budget.OutputTokens, propagateErrors: true);
                cancellationToken.ThrowIfCancellationRequested();
                ValidateRunningReviewSource(_suggestionRecord);
                if (ReviewChunkPlanner.IsFailureResponse(result))
                    throw new InvalidOperationException(string.IsNullOrWhiteSpace(result) ? UiText.T("I18n_ReviewInvalidOutput") : result);

                Dictionary<string, string> suggestions = ReviewChunkPlanner.ReadSuggestions(result,
                    batch.Select(session => session.Index.ToString()).ToList());
                foreach (EditSuggestionSession session in batch)
                {
                    if (suggestions.TryGetValue(session.Index.ToString(), out string suggestion))
                    {
                        string cleaned = CleanModelSuggestion(suggestion);
                        session.SuggestedText = ReviewChunkPlanner.IsFailureResponse(cleaned) ? null
                            : ReviewChunkPlanner.PreserveFragmentWhitespace(session.OriginalText, cleaned);
                        session.GenerationSucceeded = !ReviewChunkPlanner.IsFailureResponse(session.SuggestedText);
                        session.Status = session.GenerationSucceeded ? UiText.T("I18n_UiGenerated") : UiText.T("I18n_UiGenerationFailed");
                        if (session.GenerationSucceeded) RecordSuggestionState(session, ReviewItemState.Succeeded);
                        else unresolved.Add(session);
                    }
                    else
                    {
                        unresolved.Add(session);
                    }
                }
                foreach (var session in unresolved)
                    RecordSuggestionState(session, ReviewItemState.Failed, UiText.T("I18n_ReviewInvalidOutput"));
            }
            catch (OperationCanceledException)
            {
                foreach (EditSuggestionSession session in batch)
                    if (!session.GenerationSucceeded)
                    {
                        session.Status = UiText.T("I18n_UiReady");
                        RecordSuggestionState(session, ReviewItemState.Pending);
                    }
                throw;
            }
            catch (Exception ex)
            {
                foreach (EditSuggestionSession session in batch)
                {
                    if (session.GenerationSucceeded) continue;
                    session.FailureReason = ex.Message;
                    session.Status = UiText.T("I18n_UiGenerationFailed");
                    RecordSuggestionState(session, ReviewItemState.Failed, ex.Message);
                }
                AddAgentBubble(UiText.F("I18n_UiGenerateSuggestionFailed", ex.Message), ChatMessageType.System);
            }
            finally
            {
                foreach (EditSuggestionSession session in batch)
                    session.IsBusy = false;
                UpdateSuggestionEmptyState();
            }

            // Some local models do not reliably return JSON. Retry only the unresolved
            // items with the proven single-item prompt instead of discarding the batch.
            foreach (EditSuggestionSession session in unresolved)
            {
                cancellationToken.ThrowIfCancellationRequested();
                await GenerateSuggestionAsync(session, cancellationToken, manageStatusBar: false);
            }
        }

        private CancellationTokenSource StartReviewGeneration()
        {
            var previous = _reviewCts;
            var current = _agentReviewWorkflowActive ? CancellationTokenSource.CreateLinkedTokenSource(_agentReviewToken) : new CancellationTokenSource();
            _reviewCts = current;
            previous?.Cancel();
            SetReviewGenerationUi(isRunning: true);
            return current;
        }

        private void CompleteReviewGeneration(CancellationTokenSource completed)
        {
            bool isCurrent = ReferenceEquals(_reviewCts, completed);
            if (isCurrent)
            {
                _reviewCts = null;
                SetReviewGenerationUi(isRunning: false);
            }
            completed.Dispose();
        }

        private bool IsCurrentReviewGeneration(CancellationToken cancellationToken)
        {
            return _reviewCts != null && _reviewCts.Token == cancellationToken;
        }

        private void BtnCancelReviewGeneration_Click(object sender, RoutedEventArgs e)
        {
            var generationCts = _reviewCts;
            if (generationCts == null) return;
            generationCts.Cancel();
            btnCancelReviewGeneration.IsEnabled = false;
            btnCancelIssueReview.IsEnabled = false;
        }

        private void SetReviewGenerationUi(bool isRunning)
        {
            if (btnGenerateAllSuggestions != null)
                btnGenerateAllSuggestions.IsEnabled = !isRunning;
            if (btnCancelReviewGeneration != null) btnCancelReviewGeneration.IsEnabled = isRunning;
            if (btnCancelIssueReview != null) btnCancelIssueReview.IsEnabled = isRunning;
            if (btnReviewDocument != null) btnReviewDocument.IsEnabled = !isRunning;
            if (btnReviewSelection != null) btnReviewSelection.IsEnabled = !isRunning;
            if (reviewScopeActions != null) reviewScopeActions.IsEnabled = !isRunning && !_isLoadingReviewScope;
            if (txtReviewInstruction != null) txtReviewInstruction.IsReadOnly = isRunning;
            if (txtWritingProfile != null) txtWritingProfile.IsReadOnly = isRunning;
            SetReviewSessionActionsEnabled(!isRunning && !_isLoadingReviewScope && !_isRunning);
        }

        private void BtnAcceptAllSuggestions_Click(object sender, RoutedEventArgs e)
        {
            ApplyAllSuggestions("replace");
        }

        private void BtnInsertAllSuggestions_Click(object sender, RoutedEventArgs e)
        {
            ApplyAllSuggestions("insert");
        }

        private void BtnCommentAllSuggestions_Click(object sender, RoutedEventArgs e)
        {
            ApplyAllSuggestions("comment");
        }

        private void BtnSkipAllSuggestions_Click(object sender, RoutedEventArgs e)
        {
            if (_reviewCts != null || _isLoadingReviewScope) return;
            foreach (var session in _editSuggestions.ToList())
            {
                RecordSuggestionAction(session, "skipped");
                DeleteSuggestionBookmark(session);
            }
            _editSuggestions.Clear();
            UpdateSuggestionEmptyState();
        }

        private void ApplyAllSuggestions(string mode)
        {
            if (_reviewCts != null || _isLoadingReviewScope) return;
            var app = Globals.ThisAddIn.Application;
            var doc = app?.ActiveDocument;
            if (doc == null) return;

            bool requiresTrackedChanges = mode == "replace" || mode == "insert";
            bool enabledTrackRevisions = requiresTrackedChanges && !doc.TrackRevisions;
            bool originalScreenUpdating = app.ScreenUpdating;
            bool originalDeferredSave = _deferReviewCopySave;
            _deferReviewCopySave = true;
            try
            {
                if (_suggestionRecord != null) ValidateReviewResume(_suggestionRecord, requireConfiguration: false);
                _bulkReviewSourceValidated = _suggestionRecord != null;
                if (enabledTrackRevisions)
                    doc.TrackRevisions = true;
                app.ScreenUpdating = false;

                // Work from the end of the document so earlier offsets and paragraph
                // locations stay stable while a whole review batch is applied.
                foreach (var session in _editSuggestions
                    .Where(item => item.HasSuggestion)
                    .OrderByDescending(item => item.RangeStart)
                    .ToList())
                {
                    if (_agentReviewWorkflowActive) _agentReviewToken.ThrowIfCancellationRequested();
                    ApplySuggestion(session, mode, manageTrackRevisions: false);
                }
            }
            catch (Exception ex)
            {
                if (_agentReviewWorkflowActive) throw;
                MessageBox.Show(ex.Message, UiText.ReviewTab, MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                try
                {
                    app.ScreenUpdating = originalScreenUpdating;
                    if (enabledTrackRevisions) doc.TrackRevisions = false;
                }
                finally
                {
                    try
                    {
                        if (_bulkReviewSourceValidated && _suggestionRecord != null)
                        {
                            _suggestionRecord.Metadata.ExpectedDocumentHash = ReviewDocumentHash(doc, _suggestionRecord.Metadata.SchemaVersion);
                            SaveReviewRecord(_suggestionRecord);
                        }
                    }
                    finally
                    {
                        _bulkReviewSourceValidated = false;
                        _deferReviewCopySave = originalDeferredSave;
                        FlushWorkingReviewCopy(doc);
                    }
                }
            }
        }

        private async Task GenerateSuggestionAsync(EditSuggestionSession session, CancellationToken cancellationToken, bool manageStatusBar = true)
        {
            if (session == null || session.IsBusy) return;
            var item = _suggestionRecord?.FindItem(session.Id);
            ReviewSessionFragment currentFragment = null;
            try
            {
                session.IsBusy = true;
                ResetReviewSuggestion(session);
                cancellationToken.ThrowIfCancellationRequested();
                if (item == null) throw new InvalidOperationException(UiText.T("I18n_ReviewSessionNotLoaded"));
                ValidateRunningReviewSource(_suggestionRecord);
                session.Status = UiText.T("I18n_UiGenerating");
                RecordSuggestionState(session, ReviewItemState.Running);
                if (manageStatusBar)
                {
                    statusBarCard.Visibility = Visibility.Visible;
                    UpdateStatusIcon(true);
                    lblStatus.Text = UiText.T("I18n_UiGeneratingReviewSuggestion");
                }
                var budget = ReviewRequestBudget.Create(AppConfig.Instance, BuildSuggestionTextPrompt(""), true);
                int maxCharacters = Math.Max(128, Math.Min(1000000, AppConfig.Instance.ReviewSuggestionBatchMaxCharacters));
                var parts = ReviewChunkPlanner.SplitText(session.OriginalText, maxCharacters, budget.TextTokens);
                if (item.Fragments.Count == 0)
                {
                    item.Fragments = parts.Select(p => new ReviewSessionFragment { Start = p.Start, Source = p.Text }).ToList();
                    SaveReviewRecord(_suggestionRecord, item);
                }
                if (parts.Count != item.Fragments.Count || parts.Where((p, i) => p.Start != item.Fragments[i].Start || p.Text != item.Fragments[i].Source).Any())
                    throw new InvalidOperationException(UiText.T("I18n_ReviewSessionConfigurationMismatch"));
                var revised = new StringBuilder();
                for (int i = 0; i < item.Fragments.Count; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var fragment = item.Fragments[i];
                    if (fragment.State == ReviewItemState.Succeeded) { revised.Append(fragment.Suggestion); continue; }
                    currentFragment = fragment;
                    ValidateRunningReviewSource(_suggestionRecord);
                    fragment.State = ReviewItemState.Running; fragment.Error = null; fragment.Attempts++;
                    item.BeginAttempt(fragment.Start);
                    SaveReviewRecord(_suggestionRecord, item);
                    if (string.IsNullOrWhiteSpace(fragment.Source)) fragment.Suggestion = fragment.Source;
                    else
                    {
                        string prompt = BuildSuggestionTextPrompt(fragment.Source);
                        if (!budget.Fits(prompt, fragment.Source)) throw new InvalidOperationException(UiText.T("I18n_ReviewContextTooSmall"));
                        session.Status = UiText.F("I18n_ReviewPartProgress", i + 1, item.Fragments.Count);
                        string result = await Globals.ThisAddIn.GenerateModelSeletor(prompt, cancellationToken: cancellationToken,
                            ollamaOutputTokens: budget.OutputTokens, propagateErrors: true);
                        cancellationToken.ThrowIfCancellationRequested();
                        ValidateRunningReviewSource(_suggestionRecord);
                        if (ReviewChunkPlanner.IsFailureResponse(result))
                            throw new InvalidOperationException(string.IsNullOrWhiteSpace(result) ? UiText.T("I18n_ReviewInvalidOutput") : result);
                        string suggestion = CleanModelSuggestion(result);
                        if (ReviewChunkPlanner.IsFailureResponse(suggestion)) throw new InvalidOperationException(UiText.T("I18n_ReviewInvalidOutput"));
                        fragment.Suggestion = ReviewChunkPlanner.PreserveFragmentWhitespace(fragment.Source, suggestion);
                    }
                    fragment.State = ReviewItemState.Succeeded;
                    item.FinishAttempt(ReviewItemState.Succeeded, fragmentStart: fragment.Start);
                    SaveReviewRecord(_suggestionRecord, item);
                    revised.Append(fragment.Suggestion);
                    currentFragment = null;
                }
                ValidateRunningReviewSource(_suggestionRecord);
                session.SuggestedText = revised.ToString();
                session.GenerationSucceeded = true;
                session.Status = UiText.T("I18n_UiGenerated");
                RecordSuggestionState(session, ReviewItemState.Succeeded);
            }
            catch (OperationCanceledException)
            {
                if (currentFragment != null)
                {
                    currentFragment.State = ReviewItemState.Pending;
                    item.FinishAttempt(ReviewItemState.Pending, fragmentStart: currentFragment.Start);
                }
                session.Status = UiText.T("I18n_UiReady");
                RecordSuggestionState(session, ReviewItemState.Pending);
                throw;
            }
            catch (Exception ex)
            {
                if (currentFragment != null)
                {
                    currentFragment.State = ReviewItemState.Failed; currentFragment.Error = ex.Message;
                    item.FinishAttempt(ReviewItemState.Failed, ex.Message, currentFragment.Start);
                }
                session.FailureReason = ex.Message;
                session.Status = UiText.T("I18n_UiGenerationFailed");
                RecordSuggestionState(session, ReviewItemState.Failed, ex.Message);
                if (manageStatusBar) AddAgentBubble(UiText.F("I18n_UiGenerateSuggestionFailed", ex.Message), ChatMessageType.System);
            }
            finally
            {
                session.IsBusy = false;
                if (manageStatusBar && IsCurrentReviewGeneration(cancellationToken))
                {
                    statusBarCard.Visibility = Visibility.Collapsed;
                    typingIndicator.Visibility = Visibility.Collapsed;
                }
                UpdateSuggestionEmptyState();
            }
        }

        private string BuildSuggestionTextPrompt(string text)
        {
            string instruction = txtReviewInstruction.Text?.Trim();
            if (string.IsNullOrWhiteSpace(instruction)) instruction = UiText.DefaultReviewInstruction;
            return PromptTemplateService.BuildReviewSuggestionPrompt(text, instruction, txtWritingProfile.Text?.Trim());
        }

        private static string CleanModelSuggestion(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return "";
            string text = raw.Trim();
            text = Regex.Replace(text, @"^```(?:\w+)?\s*", "", RegexOptions.Singleline).Trim();
            text = Regex.Replace(text, @"\s*```$", "", RegexOptions.Singleline).Trim();
            return text;
        }

        private void BtnFocusSuggestion_Click(object sender, RoutedEventArgs e)
        {
            if ((sender as Button)?.Tag is EditSuggestionSession session)
                FocusSuggestion(session);
        }

        private void BtnAcceptSuggestion_Click(object sender, RoutedEventArgs e)
        {
            if ((sender as Button)?.Tag is EditSuggestionSession session)
                ApplySuggestion(session, mode: "replace");
        }

        private void BtnInsertSuggestionBelow_Click(object sender, RoutedEventArgs e)
        {
            if ((sender as Button)?.Tag is EditSuggestionSession session)
                ApplySuggestion(session, mode: "insert");
        }

        private void BtnCommentSuggestion_Click(object sender, RoutedEventArgs e)
        {
            if ((sender as Button)?.Tag is EditSuggestionSession session)
                ApplySuggestion(session, mode: "comment");
        }

        private void BtnCopySuggestion_Click(object sender, RoutedEventArgs e)
        {
            if ((sender as Button)?.Tag is EditSuggestionSession session && !string.IsNullOrWhiteSpace(session.SuggestedText))
            {
                Clipboard.SetText(session.SuggestedText);
                session.Status = UiText.T("I18n_UiCopied");
            }
        }

        private void BtnSkipSuggestion_Click(object sender, RoutedEventArgs e)
        {
            if (_reviewCts != null || _isLoadingReviewScope) return;
            if ((sender as Button)?.Tag is EditSuggestionSession session)
            {
                RecordSuggestionAction(session, "skipped");
                session.Status = UiText.T("I18n_UiSkipped");
                DeleteSuggestionBookmark(session);
                _editSuggestions.Remove(session);
                UpdateSuggestionEmptyState();
            }
        }

        private void FocusSuggestion(EditSuggestionSession session)
        {
            try
            {
                var range = GetSuggestionRange(session);
                if (range != null)
                    DocumentAgent.FocusOnRange(range, Globals.ThisAddIn.Application);
            }
            catch { }
        }

        private void ApplySuggestion(
            EditSuggestionSession session,
            string mode,
            bool manageTrackRevisions = true)
        {
            if (_reviewCts != null || _isLoadingReviewScope) return;
            if (session == null || string.IsNullOrWhiteSpace(session.SuggestedText))
            {
                MessageBox.Show(UiText.T("I18n_UiSuggestionRequired"),
                    UiText.ReviewTab, MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                var range = GetSuggestionRange(session);
                if (range == null)
                {
                    if (_agentReviewWorkflowActive) throw new InvalidOperationException(UiText.T("I18n_UiOriginalNotLocated"));
                    MessageBox.Show(UiText.T("I18n_UiOriginalNotLocated"),
                        UiText.ReviewTab, MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                if (!ConfirmUnchanged(session, range)) return;

                var doc = range.Document;
                int oldEnd = range.End;
                var editedStory = ReviewDocumentScope.Resolve(doc, session.Story);
                int documentEnd;
                try { documentEnd = editedStory.End; } finally { ComHelper.ReleaseCom(ref editedStory); }
                bool sourceWasConsistent = _bulkReviewSourceValidated || _suggestionRecord != null && _suggestionRecord.Metadata.ExpectedDocumentHash == ReviewDocumentHash(doc, _suggestionRecord.Metadata.SchemaVersion);

                using (ReviewRevisionIdentity.Enter(Globals.ThisAddIn.Application, ReviewIdentity(_suggestionRecord)))
                {
                if (mode == "replace")
                {
                    ReplaceRangeWithSuggestion(range, session, manageTrackRevisions);
                    session.Status = UiText.T("I18n_UiReplaced");
                    DeleteSuggestionBookmark(session);
                    _editSuggestions.Remove(session);
                }
                else if (mode == "insert")
                {
                    Globals.ThisAddIn.InsertTextBelowSelection(
                        range.Duplicate,
                        session.SuggestedText,
                        manageTrackRevisions);
                    session.Status = UiText.T("I18n_UiInserted");
                }
                else if (mode == "comment")
                {
                    Globals.ThisAddIn.AddCommentToSelection(range.Duplicate, session.SuggestedText);
                    session.Status = UiText.T("I18n_UiCommented");
                }
                }
                if (mode == "replace" || mode == "insert")
                {
                    editedStory = ReviewDocumentScope.Resolve(doc, session.Story);
                    try { RebaseReviewSuggestionRanges(session.Id, oldEnd, editedStory.End - documentEnd, session.Story?.Key); }
                    finally { ComHelper.ReleaseCom(ref editedStory); }
                }
                RecordSuggestionAction(session, mode, mode == "replace" || mode == "insert", sourceWasConsistent, doc);
                UpdateSuggestionEmptyState();
            }
            catch (Exception ex)
            {
                if (_agentReviewWorkflowActive) throw;
                MessageBox.Show(UiText.F("I18n_UiApplySuggestionFailed", ex.Message),
                    UiText.ReviewTab, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private Word.Range GetSuggestionRange(EditSuggestionSession session)
        {
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            if (doc == null || session == null || !string.Equals(session.DocumentIdentity, doc.FullName, StringComparison.OrdinalIgnoreCase)) return null;
            Word.Range story = ReviewDocumentScope.Resolve(doc, session.Story);
            if (story == null) return null;
            try
            {
                if (!string.IsNullOrWhiteSpace(session.BookmarkName) && doc.Bookmarks.Exists(session.BookmarkName))
                {
                    Word.Range bookmark = doc.Bookmarks[session.BookmarkName].Range.Duplicate;
                    if (bookmark.InStory(story) && EditSuggestionSession.Hash(bookmark.Text) == session.OriginalHash) return bookmark;
                    ComHelper.ReleaseCom(ref bookmark);
                }
                if (session.RangeStart < story.Start || session.RangeEnd > story.End || session.RangeEnd <= session.RangeStart) return null;
                var range = story.Duplicate; range.SetRange(session.RangeStart, session.RangeEnd);
                if (EditSuggestionSession.Hash(range.Text) == session.OriginalHash) return range;
                ComHelper.ReleaseCom(ref range); return null;
            }
            finally { ComHelper.ReleaseCom(ref story); }
        }

        private void DeleteSuggestionBookmark(EditSuggestionSession session)
        {
            try
            {
                var doc = Globals.ThisAddIn.Application.ActiveDocument;
                if (doc != null && !string.IsNullOrWhiteSpace(session?.BookmarkName) && doc.Bookmarks.Exists(session.BookmarkName))
                    doc.Bookmarks[session.BookmarkName].Delete();
            }
            catch { }
        }

        private void DeleteSuggestionBookmarks(Word.Document doc)
        {
            try
            {
                var names = doc.Bookmarks.Cast<Word.Bookmark>()
                    .Where(bookmark => bookmark.Name.StartsWith("WO_SUG_", StringComparison.OrdinalIgnoreCase))
                    .Select(bookmark => bookmark.Name)
                    .ToList();
                foreach (string name in names)
                    doc.Bookmarks[name].Delete();
            }
            catch { }
        }

        private bool ConfirmUnchanged(EditSuggestionSession session, Word.Range range)
        {
            if (_agentReviewWorkflowActive && EditSuggestionSession.Hash(range.Text) != session.OriginalHash) throw new InvalidOperationException(UiText.T("I18n_ReviewSourceChanged"));
            string current = EditSuggestionSession.NormalizeWordText(range.Text);
            if (EditSuggestionSession.Hash(current) == session.OriginalHash)
                return true;

            var result = MessageBox.Show(
                UiText.T("I18n_UiOriginalChangedPrompt"),
                UiText.ReviewTab,
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);
            return result == MessageBoxResult.Yes;
        }

        private void ReplaceRangeWithSuggestion(
            Word.Range range,
            EditSuggestionSession session,
            bool manageTrackRevisions)
        {
            Globals.ThisAddIn.ReplaceTextPrecisely(
                range,
                session.SuggestedText ?? "",
                manageTrackRevisions);
            DocumentAgent.FocusOnRange(range, Globals.ThisAddIn.Application);
        }

        private void UpdateSuggestionEmptyState()
        {
            if (txtSuggestionEmpty != null)
                txtSuggestionEmpty.Visibility = _editSuggestions.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        // ========== 审查问题面板 ==========

        public void AddReviewIssuesFromExternal(IEnumerable<ReviewIssue> issues, bool switchToIssuesTab = true)
        {
            Dispatcher.Invoke(() =>
            {
                AddReviewIssues(issues, switchToIssuesTab);
            });
        }

        public void SetWorkflowStatus(WorkflowState state, string detail)
        {
            SetWorkflowStatus(new WorkflowStatusSnapshot(state, detail));
        }

        private void SetWorkflowStatus(WorkflowStatusSnapshot snapshot)
        {
            if (snapshot == null) return;
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => SetWorkflowStatus(snapshot));
                return;
            }
            lblStatus.Text = snapshot.Detail;
            if (snapshot.State == WorkflowState.Idle)
            {
                statusBarCard.Visibility = Visibility.Collapsed;
                return;
            }

            statusBarCard.Visibility = Visibility.Visible;
            lblIteration.Text = snapshot.Total > 0
                ? $"{snapshot.Current}/{snapshot.Total}"
                : "";
            UpdateStatusIcon(snapshot.State != WorkflowState.Failed && snapshot.State != WorkflowState.Cancelled);
        }

        private void AddReviewIssues(IEnumerable<ReviewIssue> issues, bool switchToIssuesTab = true)
        {
            if (issues == null) return;

            int added = 0;
            foreach (var issue in issues)
            {
                if (issue == null) continue;
                issue.RefreshComputed();
                if (IsDuplicateIssue(issue)) continue;
                _reviewIssues.Add(issue);
                added++;
            }

            UpdateIssueSummary();
            if (added > 0 && switchToIssuesTab)
                SelectIssuesTab();
        }

        private bool IsDuplicateIssue(ReviewIssue issue)
        {
            string key = NormalizeIssueKey(issue);
            return _reviewIssues.Any(existing => NormalizeIssueKey(existing) == key);
        }

        private static string NormalizeIssueKey(ReviewIssue issue)
        {
            return $"{issue.ParagraphIndex}|{issue.TitleText}|{issue.DescriptionText}"
                .Trim()
                .ToLowerInvariant();
        }

        private void UpdateIssueSummary()
        {
            int count = _reviewIssues.Count;
            string issueSummary = count == 0
                ? (_issueReviewIncomplete ? UiText.T("I18n_ReviewNoValidIssues") : UiText.NoIssues)
                : UiText.F("I18n_UiIssueCount", count);
            txtIssueSummary.Text = string.IsNullOrEmpty(_issueReviewCoverage) ? issueSummary : _issueReviewCoverage + "\n" + issueSummary;
            txtIssueEmpty.Text = issueSummary;
            txtIssueEmpty.Visibility = count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        private void SelectIssuesTab()
        {
            if (mainTabs != null)
                mainTabs.SelectedIndex = 1;
        }

        private void BtnReviewSelection_Click(object sender, RoutedEventArgs e)
        {
            _ = RunManualReviewAsync(selectionOnly: true);
        }

        private void BtnReviewDocument_Click(object sender, RoutedEventArgs e)
        {
            _ = RunManualReviewAsync(selectionOnly: false);
        }

        private void BtnClearIssues_Click(object sender, RoutedEventArgs e)
        {
            if (_reviewCts != null) return;
            if (_issueRecord != null)
                foreach (var item in _issueRecord.Items.Where(i => i.Findings.Any(f => f.UserAction != "ignored")))
                {
                    foreach (var issue in item.Findings) { issue.UserAction = "ignored"; issue.ActionUtc = DateTime.UtcNow; }
                    try { SaveReviewRecord(_issueRecord, item); }
                    catch (Exception ex) { AddAgentBubble(ex.Message, ChatMessageType.System); break; }
                }
            _issueReviewCoverage = null;
            _issueReviewIncomplete = false;
            _reviewIssues.Clear();
            UpdateIssueSummary();
        }

        private void IssueList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // 选择只负责查看详情；跳转由每条问题卡片上的明确按钮触发，避免列表滚动时文档跟着跳动。
        }

        private void BtnJumpIssue_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement fe && fe.Tag is ReviewIssue issue)
                JumpToIssue(issue);
        }

        private void BtnCommentIssue_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement fe && fe.Tag is ReviewIssue issue)
                AddIssueComment(issue);
        }

        private void BtnIgnoreIssue_Click(object sender, RoutedEventArgs e)
        {
            if (_reviewCts != null) return;
            if (sender is FrameworkElement fe && fe.Tag is ReviewIssue issue)
            {
                RecordIssueAction(issue, "ignored");
                _reviewIssues.Remove(issue);
                UpdateIssueSummary();
            }
        }

        private async Task RunManualReviewAsync(bool selectionOnly, bool resume = false)
        {
            if (!resume && !_agentReviewWorkflowActive) _issueReviewInstruction = null;
            if (_isRunning && !_agentReviewWorkflowActive || _reviewCts != null || _isLoadingReviewScope)
            {
                AddAgentBubble(UiText.T("I18n_UiAgentRunningReview"), ChatMessageType.System);
                return;
            }

            CancellationTokenSource reviewCts = null;
            ReviewSessionRecord record = null;
            ReviewSessionItem activeItem = null;
            bool cancelled = false;
            string error = null;
            try
            {
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;
                if (doc == null) throw new InvalidOperationException(UiText.T("I18n_UiOpenDocumentFirst"));
                if (resume)
                {
                    if (_issueRecord == null) throw new InvalidOperationException(UiText.T("I18n_ReviewSessionNotLoaded"));
                    ValidateReviewResume(_issueRecord);
                    record = _issueRecord;
                }

                SelectIssuesTab();
                SetWorkflowStatus(WorkflowState.Reviewing, UiText.T("I18n_UiReviewingIssues"));
                txtIssueSummary.Text = UiText.T("I18n_UiReviewing");
                statusBarCard.Visibility = Visibility.Visible;
                UpdateStatusIcon(true);
                lblStatus.Text = UiText.T("I18n_UiReviewingIssues");
                _isRunning = true;
                reviewCts = StartReviewGeneration();
                _issueReviewCoverage = null;
                _issueReviewIncomplete = true;
                _reviewPersistenceFailed = false;
                btnCancelIssueReview.Visibility = Visibility.Visible;
                var config = AppConfig.Instance;
                var budget = ReviewRequestBudget.Create(config, BuildIssueReviewPrompt(""), rewritesText: false);
                if (!resume)
                {
                    string plannedConfiguration = ReviewConfigurationFingerprint("Issues");
                    string sourceStamp = ReviewDocumentHash(doc);
                    List<ReviewParagraph> paragraphs;
                    Word.Range sourceRange = null;
                    int start, end;
                    try
                    {
                        sourceRange = selectionOnly ? app.Selection.Range.Duplicate : doc.Content.Duplicate;
                        start = sourceRange.Start; end = sourceRange.End;
                        if (selectionOnly) paragraphs = await SnapshotReviewParagraphsAsync(doc, sourceRange, reviewCts.Token);
                        else
                        {
                            paragraphs = new List<ReviewParagraph>();
                            var stories = ReviewDocumentScope.Capture(doc);
                            try
                            {
                                int nextIndex = 1;
                                foreach (var story in stories)
                                {
                                    paragraphs.AddRange(await SnapshotReviewParagraphsAsync(doc, story.Range, reviewCts.Token, nextIndex, story.Address));
                                    nextIndex += story.Range.Paragraphs.Count;
                                }
                            }
                            finally { foreach (var story in stories) story.Dispose(); }
                        }
                    }
                    finally { ComHelper.ReleaseCom(ref sourceRange); }
                    if (paragraphs.Count == 0)
                        throw new InvalidOperationException(UiText.T(selectionOnly ? "I18n_UiSelectTextFirst" : "I18n_UiNoReviewableText"));
                    if (sourceStamp != ReviewDocumentHash(doc)) throw new InvalidOperationException(UiText.T("I18n_ReviewSourceChanged"));
                    if (plannedConfiguration != ReviewConfigurationFingerprint("Issues"))
                        throw new InvalidOperationException(UiText.T("I18n_ReviewConfigurationChanged"));
                    var batches = ReviewChunkPlanner.CreateIssueBatches(paragraphs, budget, config.ReviewIssueBatchMaxParagraphs,
                        selectionOnly ? config.ReviewSelectionBatchMaxCharacters : config.ReviewIssueBatchMaxCharacters, BuildIssueReviewPrompt);
                    var paragraphHashes = paragraphs.ToDictionary(p => p.Index, p => p.FullTextHash);
                    var paragraphLocations = paragraphs.ToDictionary(p => p.Index, p => p.Location);
                    var planned = NewReviewRecord(doc, "Issues", selectionOnly ? "Selection" : "Document", start, end);
                    planned.Items = batches.Select((batch, index) => new ReviewSessionItem
                    {
                        Index = index + 1, FirstParagraph = batch.FirstParagraph, LastParagraph = batch.LastParagraph,
                        Context = batch.Context, Characters = batch.Characters,
                        ParagraphHashes = batch.Paragraphs.ToDictionary(i => i, i => paragraphHashes[i]),
                        ParagraphLocations = batch.Paragraphs.ToDictionary(i => i, i => paragraphLocations[i])
                    }).ToList();
                    await Task.Run(() => _reviewSessionStore.Create(planned));
                    record = planned;
                    _issueRecord = record;
                    _reviewIssues.Clear();
                    ShowReviewRecordSaved(record);
                }
                record.Metadata.StartedUtc = record.Metadata.StartedUtc ?? DateTime.UtcNow;
                record.Metadata.FinishedUtc = null;
                record.Metadata.State = ReviewRunState.Running;
                record.Metadata.Error = null;
                SaveReviewRecord(record);
                AddAgentBubble(UiText.F("I18n_ReviewPlan", record.Items.SelectMany(i => i.ParagraphHashes.Keys).Distinct().Count(),
                    record.Metadata.TotalCharacters, record.Items.Count, budget.ContextTokens, budget.OutputTokens), ChatMessageType.System);

                string source = selectionOnly ? UiText.T("I18n_UiManualSelectionReview") : UiText.T("I18n_UiManualDocumentReview");
                foreach (var item in record.Items.Where(i => i.NeedsProcessing).ToList())
                {
                    reviewCts.Token.ThrowIfCancellationRequested();
                    ValidateRunningReviewSource(record);
                    activeItem = item;
                    item.State = ReviewItemState.Running; item.Error = null; item.Attempts++;
                    item.StartedUtc = DateTime.UtcNow; item.FinishedUtc = null; item.BeginAttempt();
                    SaveReviewRecord(record, item);
                    string progress = UiText.F("I18n_ReviewBatchProgress", item.Index, record.Items.Count,
                        item.FirstParagraph, item.LastParagraph, record.Metadata.ReviewedCharacters, record.Metadata.TotalCharacters);
                    lblStatus.Text = progress;
                    lblIteration.Text = $"{record.Metadata.CompletedItems + record.Metadata.FailedItems}/{record.Items.Count}";
                    _issueReviewCoverage = progress;
                    UpdateIssueSummary();
                    try
                    {
                        string prompt = BuildIssueReviewPrompt(item.Context);
                        if (!budget.Fits(prompt, item.Context)) throw new InvalidOperationException(UiText.T("I18n_ReviewContextTooSmall"));
                        string result = await Globals.ThisAddIn.GenerateModelSeletor(prompt, requireJsonOutput: true,
                            cancellationToken: reviewCts.Token, ollamaOutputTokens: budget.OutputTokens, propagateErrors: true);
                        reviewCts.Token.ThrowIfCancellationRequested();
                        ValidateRunningReviewSource(record);
                        if (!ReviewChunkPlanner.TryReadIssues(result, new HashSet<int>(item.ParagraphHashes.Keys), out var parsed))
                            throw new InvalidOperationException(ReviewChunkPlanner.IsFailureResponse(result) && !string.IsNullOrWhiteSpace(result)
                                ? result : UiText.T("I18n_ReviewInvalidOutput"));
                        var issues = ParseIssuesFromJson(parsed.ToString(), source);
                        foreach (var issue in issues)
                        {
                            issue.DocumentIdentity = record.Metadata.DocumentIdentity;
                            issue.SourceParagraphHash = item.ParagraphHashes[issue.ParagraphIndex];
                            if (item.ParagraphLocations.TryGetValue(issue.ParagraphIndex, out var location)) issue.Location = location;
                            issue.ReviewSessionId = record.Metadata.Id; issue.ReviewItemId = item.Id;
                        }
                        item.Findings = issues;
                        item.State = ReviewItemState.Succeeded; item.FinishedUtc = DateTime.UtcNow;
                        item.FinishAttempt(ReviewItemState.Succeeded);
                        SaveReviewRecord(record, item);
                        AddReviewIssues(issues);
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (Exception ex)
                    {
                        if (_reviewPersistenceFailed) throw;
                        item.State = ReviewItemState.Failed; item.Error = ex.Message; item.FinishedUtc = DateTime.UtcNow;
                        item.FinishAttempt(ReviewItemState.Failed, ex.Message);
                        SaveReviewRecord(record, item);
                    }
                    activeItem = null;
                }
                ValidateRunningReviewSource(record);
            }
            catch (OperationCanceledException)
            {
                cancelled = true;
                AddAgentBubble(UiText.T("I18n_UiReviewQueueCancelled"), ChatMessageType.System);
            }
            catch (Exception ex)
            {
                error = ex.Message;
                AddAgentBubble(UiText.F("I18n_UiReviewFailed", error), ChatMessageType.System);
            }
            finally
            {
                if (activeItem != null && activeItem.State == ReviewItemState.Running)
                {
                    activeItem.State = cancelled ? ReviewItemState.Pending : ReviewItemState.Failed;
                    activeItem.Error = error;
                    activeItem.FinishAttempt(activeItem.State, error);
                    try { SaveReviewRecord(record, activeItem); }
                    catch (Exception ex) { AddAgentBubble(ex.Message, ChatMessageType.System); }
                }
                if (record != null)
                {
                    FinishReviewRecord(record, cancelled, error);
                    var m = record.Metadata;
                    _issueReviewCoverage = BuildReviewCoverage(m.CompletedItems, m.FailedItems, m.TotalItems,
                        m.ReviewedCharacters, m.FailedCharacters, m.TotalCharacters);
                    _issueReviewIncomplete = m.CompletedItems != m.TotalItems;
                    AddAgentBubble(_issueReviewCoverage, ChatMessageType.System);
                    var failures = record.Items.Where(i => i.State == ReviewItemState.Failed).ToList();
                    if (failures.Count > 0)
                        AddAgentBubble(string.Join("\n", failures.Take(20).Select(i => UiText.F("I18n_ReviewBatchFailed", i.FirstParagraph, i.LastParagraph, i.Error)))
                            + (failures.Count > 20 ? "\n" + UiText.F("I18n_ReviewMoreFailures", failures.Count - 20) : ""), ChatMessageType.System);
                    SetWorkflowStatus(cancelled ? WorkflowState.Cancelled : _issueReviewIncomplete || error != null ? WorkflowState.Failed : WorkflowState.Completed,
                        _issueReviewCoverage);
                }
                _isRunning = _agentReviewWorkflowActive;
                if (reviewCts != null) CompleteReviewGeneration(reviewCts);
                btnCancelIssueReview.Visibility = Visibility.Collapsed;
                statusBarCard.Visibility = Visibility.Collapsed;
                typingIndicator.Visibility = Visibility.Collapsed;
                SetReviewSessionActionsEnabled(!_isLoadingReviewScope && _reviewCts == null);
                UpdateIssueSummary();
            }
        }

        private static string BuildReviewCoverage(int completed, int failed, int total, long reviewedCharacters, long failedCharacters, long totalCharacters)
        {
            return UiText.F("I18n_ReviewCoverage", completed, failed, Math.Max(0, total - completed - failed),
                reviewedCharacters, totalCharacters, failedCharacters, Math.Max(0, totalCharacters - reviewedCharacters - failedCharacters));
        }

        private async Task<List<ReviewParagraph>> SnapshotReviewParagraphsAsync(Word.Document doc, Word.Range source, CancellationToken token, int? indexStart = null, ReviewStoryAddress storyAddress = null)
        {
            var result = new List<ReviewParagraph>();
            int firstParagraph = FindParagraphIndexByRange(doc, source);
            int index = indexStart ?? firstParagraph;
            var story = storyAddress ?? ReviewDocumentScope.AddressFor(doc, source);
            int start = source.Start, end = source.End;
            if (end <= start) return result;
            Word.Paragraphs paragraphs = null;
            try
            {
                paragraphs = source.Paragraphs;
                int count = paragraphs.Count;
                for (int i = 1; i <= count; i++, index++)
                {
                    token.ThrowIfCancellationRequested();
                    Word.Paragraph paragraph = null;
                    Word.Range range = null;
                    try
                    {
                        paragraph = paragraphs[i];
                        range = paragraph.Range.Duplicate;
                        string fullHash = EditSuggestionSession.Hash(range.Text);
                        range.SetRange(Math.Max(start, range.Start), Math.Min(end, range.End));
                        string text = EditSuggestionSession.NormalizeWordText(range.Text);
                        if (!string.IsNullOrWhiteSpace(text)) result.Add(new ReviewParagraph { Index = Math.Max(1, index), Text = text, FullTextHash = fullHash, Location = new ReviewParagraphLocation { Story = story, Paragraph = firstParagraph + i - 1, Start = range.Start, End = range.End } });
                    }
                    finally { ComHelper.ReleaseCom(ref range); ComHelper.ReleaseCom(ref paragraph); }
                    if (i % 50 == 0)
                    {
                        txtIssueSummary.Text = UiText.F("I18n_ReviewReadingProgress", i, count);
                        await Dispatcher.Yield(DispatcherPriority.Background);
                    }
                }
            }
            finally { ComHelper.ReleaseCom(ref paragraphs); }
            return result;
        }

        private static string GetReviewProviderKey()
        {
            var config = AppConfig.Instance;
            return string.Join("|", config.AIMode, config.OllamaEndPoint, config.OllamaModel, config.OllamaMaxToken,
                config.OnlineProviderType, config.ChatGPTEndPoint, config.ChatGPTModel, config.ChatGPTMaxToken,
                config.ClaudeEndPoint, config.ClaudeModel, config.GeminiEndPoint, config.GeminiModel, config.OnlineContextWindow,
                config.ReviewOllamaOutputTokens);
        }

        private static string TrimForReview(string text, int maxChars)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= maxChars) return text;
            return text.Substring(0, maxChars) + "\n[内容已截断]";
        }

        private string BuildIssueReviewPrompt(string context)
        {
            return PromptTemplateService.BuildIssueReviewPrompt(context, customInstruction: _issueReviewInstruction);
        }

        private List<ReviewIssue> ParseIssuesFromJson(string raw, string source)
        {
            var issues = new List<ReviewIssue>();
            string json = ExtractJsonArray(raw);
            if (string.IsNullOrWhiteSpace(json)) return ParseIssuesFromText(raw, source);

            try
            {
                var array = JArray.Parse(json);
                foreach (var item in array.OfType<JObject>())
                {
                    int paragraph = item["paragraph"]?.Value<int?>()
                        ?? item["paragraphIndex"]?.Value<int?>()
                        ?? item["段落"]?.Value<int?>()
                        ?? 0;
                    string title = item["title"]?.ToString() ?? item["标题"]?.ToString();
                    string description = item["description"]?.ToString() ?? item["问题说明"]?.ToString() ?? item["issue"]?.ToString();
                    string suggestion = item["suggestion"]?.ToString() ?? item["建议"]?.ToString();
                    string excerpt = item["excerpt"]?.ToString() ?? item["原文"]?.ToString();

                    if (paragraph <= 0 && string.IsNullOrWhiteSpace(description) && string.IsNullOrWhiteSpace(title))
                        continue;

                    issues.Add(new ReviewIssue
                    {
                        Source = source,
                        ParagraphIndex = paragraph,
                        Category = ReviewIssue.DisplayCategory(item["category"]?.ToString() ?? item["类型"]?.ToString()),
                        Severity = ReviewIssue.DisplaySeverity(item["severity"]?.ToString() ?? item["严重程度"]?.ToString()),
                        Title = string.IsNullOrWhiteSpace(title) ? UiText.T("I18n_UiReviewIssue") : title,
                        Description = description,
                        Suggestion = suggestion,
                        Excerpt = excerpt
                    });
                }
            }
            catch
            {
                return ParseIssuesFromText(raw, source);
            }

            return issues;
        }

        private static string ExtractJsonArray(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return "";
            int start = raw.IndexOf('[');
            int end = raw.LastIndexOf(']');
            if (start < 0 || end <= start) return "";
            return raw.Substring(start, end - start + 1);
        }

        private List<ReviewIssue> ParseIssuesFromText(string text, string source)
        {
            var issues = new List<ReviewIssue>();
            if (string.IsNullOrWhiteSpace(text)) return issues;

            var matches = Regex.Matches(text, @"(?:第\s*)?(?<p>\d+)\s*(?:段|段落)[：:\s]*(?<d>[^\r\n]+)");
            foreach (Match match in matches)
            {
                if (!int.TryParse(match.Groups["p"].Value, out int paragraph)) continue;
                string description = match.Groups["d"].Value.Trim();
                if (string.IsNullOrWhiteSpace(description)) continue;
                issues.Add(new ReviewIssue
                {
                    Source = source,
                    ParagraphIndex = paragraph,
                    Category = "Agent",
                    Severity = "中",
                    Title = Loc.T("I18n_AgentIssuesTitle"),
                    Description = description
                });
            }

            var altMatches = Regex.Matches(text, @"段落\s*(?<p>\d+)[：:\s]*(?<d>[^\r\n]+)");
            foreach (Match match in altMatches)
            {
                if (!int.TryParse(match.Groups["p"].Value, out int paragraph)) continue;
                string description = match.Groups["d"].Value.Trim();
                if (string.IsNullOrWhiteSpace(description)) continue;
                issues.Add(new ReviewIssue
                {
                    Source = source,
                    ParagraphIndex = paragraph,
                    Category = "Agent",
                    Severity = "中",
                    Title = Loc.T("I18n_AgentIssuesTitle"),
                    Description = description
                });
            }

            return issues;
        }

        private void AddDiffChangesAsIssues(IEnumerable<DiffChange> changes)
        {
            if (changes == null) return;
            var issues = changes.Select(change => new ReviewIssue
            {
                Source = UiText.T("I18n_UiAgentChangeSource"),
                ParagraphIndex = change.ParagraphIndex,
                Category = UiText.T("I18n_UiDocumentChangeCategory"),
                Severity = change.Type == "删除" ? UiText.T("I18n_UiSeverityHigh") : UiText.T("I18n_UiSeverityMedium"),
                Title = UiText.F("I18n_UiChangeTitle", change.Type, change.ParagraphIndex),
                Description = change.Type == "修改"
                    ? UiText.F("I18n_UiChangeBefore", TrimForReview(change.OldText, 120))
                    : TrimForReview(change.OldText ?? change.NewText, 160),
                Suggestion = change.Type == "修改"
                    ? UiText.F("I18n_UiChangeAfter", TrimForReview(change.NewText, 120))
                    : UiText.T("I18n_UiChangeReviewHint")
            });
            AddReviewIssues(issues, switchToIssuesTab: false);
        }

        private void JumpToIssue(ReviewIssue issue)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;
                if (doc == null) return;

                Word.Range range = GetIssueRange(doc, issue);
                if (range == null)
                {
                    MessageBox.Show(UiText.T("I18n_UiIssueParagraphNotFound"), UiText.IssuesTab, MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                DocumentAgent.FocusOnRange(range, app);
            }
            catch (Exception ex)
            {
                MessageBox.Show(UiText.F("I18n_UiFocusFailed", ex.Message), UiText.IssuesTab, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private Word.Range GetIssueRange(Word.Document doc, ReviewIssue issue)
        {
            if (!string.IsNullOrEmpty(issue.DocumentIdentity)
                && !string.Equals(issue.DocumentIdentity, doc.FullName, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(UiText.T("I18n_ReviewSourceChanged"));
            if (issue.Location != null)
            {
                Word.Range story = ReviewDocumentScope.Resolve(doc, issue.Location.Story);
                if (story == null) throw new InvalidOperationException(UiText.T("I18n_ReviewSourceChanged"));
                try
                {
                    if (issue.Location.Paragraph < 1 || issue.Location.Paragraph > story.Paragraphs.Count) throw new InvalidOperationException(UiText.T("I18n_ReviewSourceChanged"));
                    var located = story.Paragraphs[issue.Location.Paragraph].Range.Duplicate;
                    if (EditSuggestionSession.Hash(located.Text) != issue.SourceParagraphHash)
                    { ComHelper.ReleaseCom(ref located); throw new InvalidOperationException(UiText.T("I18n_ReviewSourceChanged")); }
                    return located;
                }
                finally { ComHelper.ReleaseCom(ref story); }
            }
            if (issue.ParagraphIndex > 0 && issue.ParagraphIndex <= doc.Paragraphs.Count)
            {
                var range = doc.Paragraphs[issue.ParagraphIndex].Range;
                if (!string.IsNullOrEmpty(issue.SourceParagraphHash) && issue.SourceParagraphHash != EditSuggestionSession.Hash(range.Text))
                {
                    ComHelper.ReleaseCom(ref range);
                    throw new InvalidOperationException(UiText.T("I18n_ReviewSourceChanged"));
                }
                return range;
            }

            if (!string.IsNullOrEmpty(issue.SourceParagraphHash))
                throw new InvalidOperationException(UiText.T("I18n_ReviewSourceChanged"));

            string keyword = !string.IsNullOrWhiteSpace(issue.Excerpt) ? issue.Excerpt : issue.TitleText;
            if (string.IsNullOrWhiteSpace(keyword)) return null;
            keyword = keyword.Trim();
            if (keyword.Length > 40) keyword = keyword.Substring(0, 40);

            Word.Range searchRange = doc.Range();
            searchRange.Find.ClearFormatting();
            searchRange.Find.Text = keyword;
            searchRange.Find.Forward = true;
            return searchRange.Find.Execute() ? searchRange : null;
        }

        private void AddIssueComment(ReviewIssue issue)
        {
            try
            {
                var doc = Globals.ThisAddIn.Application.ActiveDocument;
                if (doc == null) return;
                var range = GetIssueRange(doc, issue);
                if (range == null)
                {
                    MessageBox.Show(UiText.T("I18n_UiIssueParagraphNotFound"), UiText.IssuesTab, MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                object commentText = BuildIssueComment(issue);
                using (ReviewRevisionIdentity.Enter(Globals.ThisAddIn.Application, ReviewIdentity(_issueRecord)))
                    doc.Comments.Add(range, ref commentText);
                RecordIssueAction(issue, "commented");
                DocumentAgent.FocusOnRange(range, Globals.ThisAddIn.Application);
            }
            catch (Exception ex)
            {
                MessageBox.Show(UiText.F("I18n_UiAddCommentFailed", ex.Message), UiText.IssuesTab, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private string BuildIssueComment(ReviewIssue issue)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"[{issue.SeverityLabel}] {issue.TitleText}");
            if (!string.IsNullOrWhiteSpace(issue.Description))
                sb.AppendLine(issue.Description);
            if (!string.IsNullOrWhiteSpace(issue.Suggestion))
                sb.AppendLine(UiText.T("I18n_UiSuggestionLabel") + issue.Suggestion);
            sb.AppendLine(UiText.T("I18n_UiSourceLabel") + issue.SourceLabel);
            return sb.ToString().Trim();
        }

        private static int FindParagraphIndexByRange(Word.Document doc, Word.Range targetRange)
        {
            if (doc == null || targetRange == null) return 1;
            Word.Range contentRange = null;
            Word.Range prefixRange = null;
            try
            {
                contentRange = ReviewDocumentScope.Resolve(doc, ReviewDocumentScope.AddressFor(doc, targetRange));
                int start = contentRange.Start;
                int targetStart = Math.Max(start, Math.Min(targetRange.Start, contentRange.End));
                // A range ending at a paragraph's start counts only the preceding
                // paragraphs. Include its first character to keep the original index.
                prefixRange = contentRange.Duplicate;
                prefixRange.SetRange(start, Math.Min(targetStart + 1, contentRange.End));
                return Math.Max(1, prefixRange.Paragraphs.Count);
            }
            catch { return 1; }
            finally
            {
                ComHelper.ReleaseCom(ref prefixRange);
                ComHelper.ReleaseCom(ref contentRange);
            }
        }

        // ========== 图片粘贴处理 ==========

        /// <summary>显示粘贴图片预览。</summary>
        private void ShowPastePreview(System.Windows.Media.ImageSource image)
        {
            // 转换为 Base64 存储
            _pastedImageBase64 = ImageToBase64(image);

            // 显示预览
            pastePreviewImage.Source = image;
            pastePreview.Visibility = Visibility.Visible;
            txtInput.Tag = _pastedImageBase64; // 暂存用于后续发送
        }

        /// <summary>移除粘贴的图片。</summary>
        private void RemovePasteImage()
        {
            _pastedImageBase64 = null;
            pastePreviewImage.Source = null;
            pastePreview.Visibility = Visibility.Collapsed;
            txtInput.Tag = null;
        }

        /// <summary>将 WPF ImageSource 转换为 Base64 字符串（PNG 格式）。</summary>
        private static string ImageToBase64(System.Windows.Media.ImageSource image)
        {
            try
            {
                if (image is System.Windows.Media.Imaging.BitmapSource bitmap)
                {
                    var encoder = new System.Windows.Media.Imaging.PngBitmapEncoder();
                    encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(bitmap));
                    using (var ms = new System.IO.MemoryStream())
                    {
                        encoder.Save(ms);
                        return Convert.ToBase64String(ms.ToArray());
                    }
                }
            }
            catch { }
            return null;
        }
    }

    /// <summary>
    /// 聊天气泡数据模型 — 每个气泡只显示一条消息。
    /// </summary>
    public class ChatItem : INotifyPropertyChanged
    {
        public ChatMessageType MessageType { get; set; }
        public bool ShowAgentLabel { get; set; }
        private string _stepTitle;
        public string StepTitle
        {
            get => _stepTitle;
            set { _stepTitle = value; OnPropertyChanged(nameof(StepTitle)); }
        }

        /// <summary>兼容旧版属性：User 消息。</summary>
        public bool IsUser => MessageType == ChatMessageType.User;
        /// <summary>兼容旧版属性：Agent 消息（非 User 即 Agent）。</summary>
        public bool IsAgent => !IsUser;

        private string _message;
        public string Message
        {
            get => _message;
            set { _message = value; OnPropertyChanged(nameof(Message)); }
        }

        /// <summary>是否可折叠（长消息）。</summary>
        public bool IsCollapsible { get; set; }
        /// <summary>执行步骤是否展开。</summary>
        public bool IsExpanded { get; set; }

        // ========== 图片支持 ==========

        private ImageSource _imageSource;
        /// <summary>气泡中显示的图片（用户粘贴的图片或 AI 返回的图片）。</summary>
        public ImageSource ImageSource
        {
            get => _imageSource;
            set { _imageSource = value; OnPropertyChanged(nameof(ImageSource)); OnPropertyChanged(nameof(HasImage)); }
        }

        /// <summary>图片的 Base64 编码数据（用于发送给 AI）。</summary>
        public string ImageBase64 { get; set; }

        /// <summary>是否有图片显示。</summary>
        public bool HasImage => ImageSource != null;

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    /// <summary>
    /// 消息类型 — 决定气泡的显示样式。
    /// </summary>
    public enum ChatMessageType
    {
        System,     // 系统日志（🚀 开始、📦 压缩等）
        Thought,    // AI 思考过程（将被合并到 StepGroup）
        Action,     // 工具调用信息（将被合并到 StepGroup）
        Result,     // 最终结论
        Question,   // AI 提问
        User,       // 用户消息
        Error,      // 错误信息
        Plan,       // 执行计划
        StepGroup   // 合并后的步骤卡片（💭+🔧+👀）
    }

    /// <summary>
    /// 将 ChatMessageType 转换为对应的颜色笔刷或可见性。
    /// ConverterParameter:
    ///   "SystemVisible" → 仅 System 类型为 Visible，其余 Collapsed
    ///   "PlainVisible" → Result/Question 为 Visible
    ///   "CardVisible" → Error/Plan 为 Visible
    ///   其他值或不传 → 返回类型对应的颜色笔刷
    /// </summary>
    public class MessageTypeToBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (parameter is string param)
            {
                bool isSystem = value is ChatMessageType t && t == ChatMessageType.System;
                if (param == "SystemVisible") return isSystem ? Visibility.Visible : Visibility.Collapsed;
                if (param == "NonSystemVisible") return isSystem ? Visibility.Collapsed : Visibility.Visible;
                if (param == "PlanVisible") return (value is ChatMessageType pt && pt == ChatMessageType.Plan) ? Visibility.Visible : Visibility.Collapsed;
                if (param == "StepVisible") return (value is ChatMessageType st && st == ChatMessageType.StepGroup) ? Visibility.Visible : Visibility.Collapsed;
                if (param == "PlainVisible") return (value is ChatMessageType plain &&
                    (plain == ChatMessageType.Result || plain == ChatMessageType.Question)) ? Visibility.Visible : Visibility.Collapsed;
                if (param == "CardVisible") return (value is ChatMessageType card &&
                    (card == ChatMessageType.Error || card == ChatMessageType.Plan)) ? Visibility.Visible : Visibility.Collapsed;
            }

            // 返回颜色笔刷
            return (value as ChatMessageType?) switch
            {
                ChatMessageType.Thought => new SolidColorBrush(Color.FromRgb(0x9E, 0x9E, 0x9E)),
                ChatMessageType.Action => new SolidColorBrush(Color.FromRgb(0x21, 0x96, 0xF3)),
                ChatMessageType.Result => new SolidColorBrush(Color.FromRgb(0x4C, 0xAF, 0x50)),
                ChatMessageType.Question => new SolidColorBrush(Color.FromRgb(0xFF, 0x98, 0x00)),
                ChatMessageType.User => new SolidColorBrush(Color.FromRgb(0x3F, 0x51, 0xB5)),
                ChatMessageType.Error => new SolidColorBrush(Color.FromRgb(0xF4, 0x43, 0x36)),
                ChatMessageType.Plan => new SolidColorBrush(Color.FromRgb(0x21, 0x96, 0xF3)),
                ChatMessageType.StepGroup => new SolidColorBrush(Color.FromRgb(0x00, 0x96, 0x88)),
                _ => new SolidColorBrush(Color.FromRgb(0x9E, 0x9E, 0x9E)),
            };
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
            => Binding.DoNothing;
    }
}
