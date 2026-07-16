using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
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

        // ========== 图片粘贴 ==========
        /// <summary>用户粘贴的图片 Base64 数据（发送前暂存）。</summary>
        private string _pastedImageBase64;

        // ========== 步骤卡片缓存 ==========
        private int _stepCounter = 0;
        private string _bufferedThought = null;
        private string _bufferedAction = null;
        private readonly bool _reviewOnly;

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
            txtReviewInstruction.Text = UiText.DefaultReviewInstruction;
            _writingProfile = WritingProfileManager.LoadForDocument(Globals.ThisAddIn?.Application?.ActiveDocument);
            txtWritingProfile.Text = _writingProfile.Instructions ?? "";
            UpdateIssueSummary();
            UpdateSuggestionEmptyState();
            UpdateEmptyChatState();

            // 加载技能列表用于斜杠选择
            _skillManager.LoadSkills(SkillsRootPath);
            RefreshAgentApiDocumentation();

            ApplyTheme();
            ThemeManager.ThemeChanged += (_, _) => ApplyTheme();
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
                theme.SetBaseTheme(SharedSettings.themeType == "Dark" ? BaseTheme.Dark : BaseTheme.Light);
                Resources.SetTheme(theme);
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
                    _stepCounter = 0;
                    _bufferedThought = null;
                    _bufferedAction = null;
                    break;

                case "newchat":
                    // 清空窗口 + 重置 Agent
                    _chatItems.Clear();
                    _stepCounter = 0;
                    _bufferedThought = null;
                    _bufferedAction = null;
                    _agent = null;
                    _cts?.Cancel();
                    _pendingRequirement = null;
                    _pendingImageBase64 = null;
                    _isWaitingForResponse = false;
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

                // 添加气泡提示 Agent 正在继续
                AddAgentBubble(Properties.Resources.Agent_AnswerReceived);
                return;
            }

            // 如果 Agent 正在运行（包括等待用户回答），用户输入新需求 → 打断当前任务
            if (_isRunning)
            {
                // 取消当前任务并释放可能挂起的 TCS
                _agent?.SetUserResponse(UiText.T("I18n_UiUserInterrupted"));
                _isWaitingForResponse = false;
                _cts?.Cancel();

                // 标记被中断
                AddAgentBubble(Properties.Resources.Agent_Interrupted);

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
            if (_isRunning) return;
            _isRunning = true;
            _cts = new CancellationTokenSource();

            // 重置步骤卡片计数器
            _stepCounter = 0;
            _bufferedThought = null;
            _bufferedAction = null;
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
                _agent.MaxIterations = SharedSettings.AgentMaxIterations;
                if (mcpManager != null && mcpManager.ServerCount > 0)
                {
                    _agent.SetMcpManager(mcpManager);
                    RefreshAgentApiDocumentation();
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
                            AddAgentBubble(UiText.T("I18n_UiCheckpointResuming"));
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
                            _stepCounter++;
                            string stepTitle = UiText.F("I18n_UiAgentStep", _stepCounter);
                            string combined = $"{_bufferedThought}\n{_bufferedAction}\n{msg}";
                            AddAgentBubble(combined, ChatMessageType.StepGroup, stepTitle);
                            _bufferedThought = null;
                            _bufferedAction = null;
                            UpdateStatusBarFromLog(msg);
                            ScrollToEnd();
                            return;
                        }

                        // === 非步骤消息（🚀 ⚠️ ✅ 🤖 等）：直接显示 ===
                        FlushStepBuffer();
                        AddAgentBubble(msg);
                        UpdateStatusBarFromLog(msg);
                        ScrollToEnd();
                    });

                // 最终结果 → 新气泡
                _agent.ResultReady += summary =>
                    Dispatcher.Invoke(() =>
                    {
                        AddAgentBubble(summary);
                        AddReviewIssues(ParseIssuesFromText(summary, "Agent"));
                        UpdateStatusIcon(false);
                        lblStatus.Text = Properties.Resources.Agent_StatusDone;
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

                // [HITL] Agent 提问 → 新气泡
                _agent.QuestionAsked += question =>
                    Dispatcher.Invoke(() =>
                    {
                        _isWaitingForResponse = true;
                        AddAgentBubble(string.Format(Properties.Resources.Agent_AskLog, question));
                        AddAgentBubble(Properties.Resources.Agent_AskPrompt);
                        UpdateStatusIcon(null);
                        lblStatus.Text = Properties.Resources.Agent_StatusWaiting;
                        ScrollToEnd();
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

                await _agent.ExecuteAsync(requirement, _cts.Token, resumeCheckpoint);
            }
            catch (OperationCanceledException)
            {
                SetWorkflowStatus(WorkflowState.Cancelled, Properties.Resources.Agent_Cancelled);
                Dispatcher.Invoke(() =>
                {
                    AddAgentBubble(Properties.Resources.Agent_Cancelled);
                });
                _agent?.SetUserResponse(UiText.T("I18n_UiUserCancelled"));
            }
            catch (Exception ex)
            {
                SetWorkflowStatus(WorkflowState.Failed, string.Format(Properties.Resources.Agent_Error, ex.Message));
                Dispatcher.Invoke(() =>
                {
                    AddAgentBubble(string.Format(Properties.Resources.Agent_Error, ex.Message));
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
                AddAgentBubble(_bufferedThought, ChatMessageType.Thought);
                _bufferedThought = null;
            }
            if (_bufferedAction != null)
            {
                AddAgentBubble(_bufferedAction, ChatMessageType.Action);
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
            _agent?.SetUserResponse(UiText.T("I18n_UiUserCancelled"));
            _cts?.Cancel();
            btnCancelTask.IsEnabled = false;
        }

        // ========== 工具方法 ==========

        public void ClearChat()
        {
            Dispatcher.Invoke(() =>
            {
                _chatItems.Clear();
            });
        }

        /// <summary>供 SilentLinter 等外部组件调用 — 添加 Agent 气泡。</summary>
        public void StartAgentReply()
        {
            Dispatcher.Invoke(() =>
            {
                AddAgentBubble("");
                ScrollToEnd();
            });
        }

        /// <summary>供外部组件调用 — 添加 Agent 气泡。</summary>
        public void AppendLog(string message)
        {
            Dispatcher.Invoke(() =>
            {
                AddAgentBubble(message);
                ScrollToEnd();
            });
        }

        /// <summary>供外部组件调用 — 添加最终结论气泡。</summary>
        public void SetFinalMessage(string message)
        {
            Dispatcher.Invoke(() =>
            {
                AddAgentBubble(message);
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
            Dispatcher.Invoke(() =>
            {
                SelectReviewTab();
                if (!string.IsNullOrWhiteSpace(instruction))
                    txtReviewInstruction.Text = instruction.Trim();

                if (replaceExisting)
                    _editSuggestions.Clear();

                bool shouldGenerate = false;
                if (loadSelection && (_editSuggestions.Count == 0 || replaceExisting))
                {
                    LoadSelectionForReview(splitParagraphs: true);
                    shouldGenerate = _editSuggestions.Count > 0;
                }

                if (shouldGenerate)
                    _ = GenerateAllSuggestionsAsync();
            });
        }

        private void SelectReviewTab()
        {
            if (mainTabs != null)
                mainTabs.SelectedItem = ReviewTabItem;
        }

        private void BtnLoadReviewSelection_Click(object sender, RoutedEventArgs e)
        {
            LoadSelectionForReview(splitParagraphs: false);
        }

        private void BtnLoadReviewParagraphs_Click(object sender, RoutedEventArgs e)
        {
            LoadCurrentParagraphForReview();
        }

        private void BtnLoadReviewDocument_Click(object sender, RoutedEventArgs e)
        {
            LoadDocumentForReview();
        }

        private void BtnClearSuggestions_Click(object sender, RoutedEventArgs e)
        {
            _editSuggestions.Clear();
            UpdateSuggestionEmptyState();
        }

        private void BtnSaveWritingProfile_Click(object sender, RoutedEventArgs e)
        {
            _writingProfile = new WritingProfile { Instructions = txtWritingProfile.Text?.Trim() ?? "" };
            var document = Globals.ThisAddIn?.Application?.ActiveDocument;
            WritingProfileManager.SaveForDocument(document, _writingProfile);
            AddAgentBubble(UiText.T("I18n_UiWritingProfileSaved"), ChatMessageType.System);
        }

        private void LoadSelectionForReview(bool splitParagraphs)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;
                var selection = app.Selection;
                if (doc == null || selection == null)
                {
                    SetReviewScopeStatus(UiText.T("I18n_UiOpenDocumentFirst"));
                    return;
                }

                Word.Range selectedRange = selection.Range;
                LoadReviewRange(doc, selectedRange, splitParagraphs);
            }
            catch (Exception ex)
            {
                MessageBox.Show(UiText.F("I18n_UiLoadSelectionFailed", ex.Message),
                    UiText.ReviewTab, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadDocumentForReview()
        {
            try
            {
                var doc = Globals.ThisAddIn.Application.ActiveDocument;
                if (doc == null)
                {
                    SetReviewScopeStatus(UiText.T("I18n_UiOpenDocumentFirst"));
                    return;
                }

                Word.Range documentRange = doc.Content.Duplicate;
                if (documentRange.End > documentRange.Start)
                    documentRange.End -= 1;
                LoadReviewRange(doc, documentRange, splitParagraphs: true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(UiText.F("I18n_UiLoadSelectionFailed", ex.Message),
                    UiText.ReviewTab, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadCurrentParagraphForReview()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;
                var selection = app.Selection;
                if (doc == null || selection == null)
                {
                    SetReviewScopeStatus(UiText.T("I18n_UiOpenDocumentFirst"));
                    return;
                }

                Word.Range paragraphRange = selection.Paragraphs[1].Range.Duplicate;
                LoadReviewRange(doc, paragraphRange, splitParagraphs: false);
            }
            catch (Exception ex)
            {
                MessageBox.Show(UiText.F("I18n_UiLoadSelectionFailed", ex.Message),
                    UiText.ReviewTab, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadReviewRange(Word.Document doc, Word.Range reviewRange, bool splitParagraphs)
        {
            string selectedText = EditSuggestionSession.NormalizeWordText(reviewRange?.Text);
            if (string.IsNullOrWhiteSpace(selectedText))
            {
                SetReviewScopeStatus(UiText.ReviewScopeNone);
                return;
            }

            _editSuggestions.Clear();
            DeleteSuggestionBookmarks(doc);
            int index = 1;

            if (splitParagraphs && reviewRange.Paragraphs.Count > 1)
            {
                foreach (Word.Paragraph paragraph in reviewRange.Paragraphs)
                {
                    Word.Range range = paragraph.Range.Duplicate;
                    string text = EditSuggestionSession.NormalizeWordText(range.Text);
                    if (string.IsNullOrWhiteSpace(text)) continue;

                    _editSuggestions.Add(CreateSuggestionSession(doc, range, text, index++));
                }
            }
            else
            {
                _editSuggestions.Add(CreateSuggestionSession(doc, reviewRange.Duplicate, selectedText, index));
            }

            UpdateSuggestionEmptyState();
            SetReviewScopeStatus(UiText.F("I18n_UiReviewLoadedCount", _editSuggestions.Count));
            SelectReviewTab();
        }

        private void SetReviewScopeStatus(string message)
        {
            if (txtReviewScopeStatus != null)
                txtReviewScopeStatus.Text = message ?? "";
        }

        private EditSuggestionSession CreateSuggestionSession(Word.Document doc, Word.Range range, string text, int index)
        {
            var session = new EditSuggestionSession
            {
                Index = index,
                ParagraphIndex = FindParagraphIndexByRange(doc, range),
                RangeStart = range.Start,
                RangeEnd = range.End,
                OriginalText = text,
                OriginalHash = EditSuggestionSession.Hash(text),
                OriginalEndsWithParagraphMark = (range.Text ?? "").EndsWith("\r", StringComparison.Ordinal),
                Status = UiText.T("I18n_UiReady")
            };
            session.BookmarkName = "WO_SUG_" + session.Id.Substring(0, 32);
            try
            {
                if (doc.Bookmarks.Exists(session.BookmarkName))
                    doc.Bookmarks[session.BookmarkName].Delete();
                doc.Bookmarks.Add(session.BookmarkName, range.Duplicate);
            }
            catch
            {
                session.BookmarkName = null;
            }
            return session;
        }

        private async void BtnGenerateSuggestion_Click(object sender, RoutedEventArgs e)
        {
            if ((sender as Button)?.Tag is EditSuggestionSession session)
            {
                var generationCts = StartReviewGeneration();
                try { await GenerateSuggestionAsync(session, generationCts.Token); }
                catch (OperationCanceledException)
                {
                    // Cancelling a review generation is an expected user action.
                }
                finally { CompleteReviewGeneration(generationCts); }
            }
        }

        private async void BtnGenerateAllSuggestions_Click(object sender, RoutedEventArgs e)
        {
            await GenerateAllSuggestionsAsync();
        }

        private async Task GenerateAllSuggestionsAsync()
        {
            var generationCts = StartReviewGeneration();
            try
            {
                foreach (var session in _editSuggestions.ToList())
                {
                    generationCts.Token.ThrowIfCancellationRequested();
                    await GenerateSuggestionAsync(session, generationCts.Token);
                }
            }
            catch (OperationCanceledException)
            {
                AddAgentBubble(UiText.T("I18n_UiReviewQueueCancelled"), ChatMessageType.System);
            }
            finally
            {
                CompleteReviewGeneration(generationCts);
            }
        }

        private CancellationTokenSource StartReviewGeneration()
        {
            var previous = _reviewCts;
            var current = new CancellationTokenSource();
            _reviewCts = current;
            previous?.Cancel();
            return current;
        }

        private void CompleteReviewGeneration(CancellationTokenSource completed)
        {
            if (ReferenceEquals(_reviewCts, completed))
                _reviewCts = null;
            completed.Dispose();
        }

        private bool IsCurrentReviewGeneration(CancellationToken cancellationToken)
        {
            return _reviewCts != null && _reviewCts.Token == cancellationToken;
        }

        private void BtnCancelReviewGeneration_Click(object sender, RoutedEventArgs e)
        {
            _reviewCts?.Cancel();
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
            foreach (var session in _editSuggestions.ToList())
                DeleteSuggestionBookmark(session);
            _editSuggestions.Clear();
            UpdateSuggestionEmptyState();
        }

        private void ApplyAllSuggestions(string mode)
        {
            foreach (var session in _editSuggestions.ToList())
            {
                if (session.HasSuggestion)
                    ApplySuggestion(session, mode);
            }
        }

        private async Task GenerateSuggestionAsync(EditSuggestionSession session, CancellationToken cancellationToken)
        {
            if (session == null || session.IsBusy) return;
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                session.IsBusy = true;
                session.Status = UiText.T("I18n_UiGenerating");
                statusBarCard.Visibility = Visibility.Visible;
                UpdateStatusIcon(true);
                lblStatus.Text = UiText.T("I18n_UiGeneratingReviewSuggestion");

                string result = await Globals.ThisAddIn.GenerateModelSeletor(BuildSuggestionPrompt(session));
                cancellationToken.ThrowIfCancellationRequested();
                session.SuggestedText = CleanModelSuggestion(result);
                session.Status = string.IsNullOrWhiteSpace(session.SuggestedText)
                    ? UiText.T("I18n_UiNoOutput")
                    : UiText.T("I18n_UiGenerated");
            }
            catch (OperationCanceledException)
            {
                session.Status = UiText.T("I18n_UiReady");
                throw;
            }
            catch (Exception ex)
            {
                session.Status = UiText.T("I18n_UiGenerationFailed");
                MessageBox.Show(UiText.F("I18n_UiGenerateSuggestionFailed", ex.Message),
                    UiText.ReviewTab, MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                session.IsBusy = false;
                if (IsCurrentReviewGeneration(cancellationToken))
                {
                    statusBarCard.Visibility = Visibility.Collapsed;
                    typingIndicator.Visibility = Visibility.Collapsed;
                }
                UpdateSuggestionEmptyState();
            }
        }

        private string BuildSuggestionPrompt(EditSuggestionSession session)
        {
            string instruction = txtReviewInstruction.Text?.Trim();
            if (string.IsNullOrWhiteSpace(instruction))
                instruction = UiText.DefaultReviewInstruction;

            string profile = txtWritingProfile.Text?.Trim();
            return PromptTemplateService.BuildReviewSuggestionPrompt(session?.OriginalText, instruction, profile);
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
            if ((sender as Button)?.Tag is EditSuggestionSession session)
            {
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

        private void ApplySuggestion(EditSuggestionSession session, string mode)
        {
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
                    MessageBox.Show(UiText.T("I18n_UiOriginalNotLocated"),
                        UiText.ReviewTab, MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                if (!ConfirmUnchanged(session, range)) return;

                if (mode == "replace")
                {
                    ReplaceRangeWithSuggestion(range, session);
                    session.Status = UiText.T("I18n_UiReplaced");
                    DeleteSuggestionBookmark(session);
                    _editSuggestions.Remove(session);
                }
                else if (mode == "insert")
                {
                    Globals.ThisAddIn.InsertTextBelowSelection(range.Duplicate, session.SuggestedText);
                    session.Status = UiText.T("I18n_UiInserted");
                }
                else if (mode == "comment")
                {
                    Globals.ThisAddIn.AddCommentToSelection(range.Duplicate, session.SuggestedText);
                    session.Status = UiText.T("I18n_UiCommented");
                }

                UpdateSuggestionEmptyState();
            }
            catch (Exception ex)
            {
                MessageBox.Show(UiText.F("I18n_UiApplySuggestionFailed", ex.Message),
                    UiText.ReviewTab, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private Word.Range GetSuggestionRange(EditSuggestionSession session)
        {
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            if (doc == null || session == null) return null;

            try
            {
                if (!string.IsNullOrWhiteSpace(session.BookmarkName) && doc.Bookmarks.Exists(session.BookmarkName))
                    return doc.Bookmarks[session.BookmarkName].Range.Duplicate;

                int end = Math.Min(session.RangeEnd, doc.Content.End);
                if (session.RangeStart >= 0 && end > session.RangeStart)
                {
                    var offsetRange = doc.Range(session.RangeStart, end);
                    if (EditSuggestionSession.Hash(offsetRange.Text) == session.OriginalHash)
                        return offsetRange;
                }

                if (!string.IsNullOrWhiteSpace(session.OriginalText))
                {
                    var searchRange = doc.Content.Duplicate;
                    var find = searchRange.Find;
                    find.ClearFormatting();
                    find.Replacement.ClearFormatting();
                    find.Text = session.OriginalText;
                    find.Forward = true;
                    find.Wrap = Word.WdFindWrap.wdFindStop;
                    find.Format = false;
                    find.MatchWildcards = false;
                    if (find.Execute())
                        return searchRange.Duplicate;
                }
            }
            catch { }

            if (session.ParagraphIndex > 0 && session.ParagraphIndex <= doc.Paragraphs.Count)
            {
                var paragraphRange = doc.Paragraphs[session.ParagraphIndex].Range;
                if (EditSuggestionSession.Hash(paragraphRange.Text) == session.OriginalHash)
                    return paragraphRange;
            }

            return null;
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

        private void ReplaceRangeWithSuggestion(Word.Range range, EditSuggestionSession session)
        {
            var doc = range.Document;
            bool originalTrack = doc.TrackRevisions;
            try
            {
                doc.TrackRevisions = true;
                string text = session.SuggestedText ?? "";
                if (session.OriginalEndsWithParagraphMark && !text.EndsWith("\r", StringComparison.Ordinal))
                    text += "\r";
                range.Text = text;
                DocumentAgent.FocusOnRange(range, Globals.ThisAddIn.Application);
            }
            finally
            {
                doc.TrackRevisions = originalTrack;
            }
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
            txtIssueSummary.Text = count == 0
                ? UiText.NoIssues
                : UiText.F("I18n_UiIssueCount", count);
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
            if (sender is FrameworkElement fe && fe.Tag is ReviewIssue issue)
            {
                _reviewIssues.Remove(issue);
                UpdateIssueSummary();
            }
        }

        private async Task RunManualReviewAsync(bool selectionOnly)
        {
            if (_isRunning)
            {
                AddAgentBubble(UiText.T("I18n_UiAgentRunningReview"), ChatMessageType.System);
                return;
            }

            CancellationTokenSource reviewCts = null;
            try
            {
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show(UiText.T("I18n_UiOpenDocumentFirst"), UiText.IssuesTab, MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                string context = BuildReviewContext(doc, app, selectionOnly);
                if (string.IsNullOrWhiteSpace(context))
                {
                    MessageBox.Show(
                        selectionOnly
                            ? UiText.T("I18n_UiSelectTextFirst")
                            : UiText.T("I18n_UiNoReviewableText"),
                        UiText.IssuesTab,
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }

                SelectIssuesTab();
                SetWorkflowStatus(WorkflowState.Reviewing, UiText.T("I18n_UiReviewingIssues"));
                txtIssueSummary.Text = UiText.T("I18n_UiReviewing");
                statusBarCard.Visibility = Visibility.Visible;
                UpdateStatusIcon(true);
                lblStatus.Text = UiText.T("I18n_UiReviewingIssues");
                _isRunning = true;
                reviewCts = StartReviewGeneration();

                string result = await Globals.ThisAddIn.GenerateModelSeletor(BuildIssueReviewPrompt(context), requireJsonOutput: true);
                reviewCts.Token.ThrowIfCancellationRequested();
                var issues = ParseIssuesFromJson(result, selectionOnly ? UiText.T("I18n_UiManualSelectionReview") : UiText.T("I18n_UiManualDocumentReview"));
                if (issues.Count == 0)
                {
                    AddAgentBubble(UiText.T("I18n_UiReviewCompleteNoIssues"));
                    SetWorkflowStatus(WorkflowState.Completed, UiText.T("I18n_UiReviewComplete"));
                }
                else
                {
                    AddReviewIssues(issues);
                    AddAgentBubble(UiText.F("I18n_UiReviewAdded", issues.Count));
                    SetWorkflowStatus(WorkflowState.Completed, UiText.F("I18n_UiReviewAddedStatus", issues.Count));
                }
            }
            catch (OperationCanceledException)
            {
                AddAgentBubble(UiText.T("I18n_UiReviewQueueCancelled"), ChatMessageType.System);
                SetWorkflowStatus(WorkflowState.Cancelled, UiText.T("I18n_UiReviewQueueCancelled"));
            }
            catch (Exception ex)
            {
                SetWorkflowStatus(WorkflowState.Failed, UiText.F("I18n_UiReviewFailed", ex.Message));
                MessageBox.Show(UiText.F("I18n_UiReviewFailed", ex.Message), UiText.IssuesTab, MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (reviewCts != null)
                    CompleteReviewGeneration(reviewCts);
                _isRunning = false;
                statusBarCard.Visibility = Visibility.Collapsed;
                typingIndicator.Visibility = Visibility.Collapsed;
                UpdateIssueSummary();
            }
        }

        private string BuildReviewContext(Word.Document doc, Word.Application app, bool selectionOnly)
        {
            var sb = new StringBuilder();
            if (selectionOnly)
            {
                string selectedText = app.Selection?.Text?.TrimEnd('\r', '\a') ?? "";
                if (string.IsNullOrWhiteSpace(selectedText)) return "";
                int paragraphIndex = FindParagraphIndexByRange(doc, app.Selection.Range);
                sb.AppendLine($"[段落 {Math.Max(1, paragraphIndex)}]");
                sb.AppendLine(TrimForReview(selectedText, 6000));
                return sb.ToString();
            }

            int maxChars = 12000;
            int maxParagraphs = Math.Min(doc.Paragraphs.Count, 80);
            for (int i = 1; i <= maxParagraphs; i++)
            {
                string text = doc.Paragraphs[i].Range.Text?.TrimEnd('\r', '\a') ?? "";
                if (string.IsNullOrWhiteSpace(text)) continue;
                string block = $"[段落 {i}]\n{text}\n";
                if (sb.Length + block.Length > maxChars) break;
                sb.AppendLine(block);
            }
            return sb.ToString();
        }

        private static string TrimForReview(string text, int maxChars)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= maxChars) return text;
            return text.Substring(0, maxChars) + "\n[内容已截断]";
        }

        private string BuildIssueReviewPrompt(string context)
        {
            return PromptTemplateService.BuildIssueReviewPrompt(context);
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
            if (issue.ParagraphIndex > 0 && issue.ParagraphIndex <= doc.Paragraphs.Count)
                return doc.Paragraphs[issue.ParagraphIndex].Range;

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
                doc.Comments.Add(range, ref commentText);
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

        private int FindParagraphIndexByRange(Word.Document doc, Word.Range targetRange)
        {
            if (doc == null || targetRange == null) return 1;
            int start = targetRange.Start;
            for (int i = 1; i <= doc.Paragraphs.Count; i++)
            {
                if (doc.Paragraphs[i].Range.End >= start)
                    return i;
            }
            return 1;
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
        public string StepTitle { get; set; }

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
    ///   "NonSystemVisible" → 非 System 类型为 Visible，System 为 Collapsed
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
                if (param == "NonStepVisible") return (value is ChatMessageType nst && nst != ChatMessageType.StepGroup && nst != ChatMessageType.System) ? Visibility.Visible : Visibility.Collapsed;
                if (param == "ResultVisible") return (value is ChatMessageType rt && rt == ChatMessageType.Result) ? Visibility.Visible : Visibility.Collapsed;
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
