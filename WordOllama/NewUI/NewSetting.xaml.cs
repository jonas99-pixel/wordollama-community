using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using WordOllama.Forms;
using System.Diagnostics;
using MaterialDesignThemes.Wpf;
using Newtonsoft.Json;
using System.IO;
using MaterialDesignColors;
using System.Windows.Media.Media3D;
using WordOllama.lib;
using Word = Microsoft.Office.Interop.Word;
using WordOllama.lib.Providers;
using System.Runtime.InteropServices;
using System.Windows.Interop;

namespace WordOllama.NewUI
{
    /// <summary>
    /// NewSetting.xaml 的交互逻辑
    /// </summary>
    public partial class NewSetting : Window
    {
        private static readonly IntPtr HwndNotTopmost = new IntPtr(-2);
        private const uint SwpNoSize = 0x0001;
        private const uint SwpNoMove = 0x0002;
        private const uint SwpNoActivate = 0x0010;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(
            IntPtr hWnd,
            IntPtr hWndInsertAfter,
            int x,
            int y,
            int cx,
            int cy,
            uint flags);

        private HttpClient _httpClient;
        public ObservableCollection<OllamaModelInfo> OllamaModels { get; set; }
        public ObservableCollection<ChatGPTModelInfo> ChatGPTModels { get; set; }
        public ObservableCollection<UserMemoryItem> MemoryItems { get; set; } = new ObservableCollection<UserMemoryItem>();

        public bool KeyVisable;
        public JArray ChatGPTModelsJson;
        bool OllamaNetModeNow = SharedSettings.OllamaNetMode;

        // 弹窗拖拽
        private bool _isPopupDragging;
        private Point _popupDragStart;

        private void NonTopmostPopup_Opened(object sender, EventArgs e)
        {
            if (sender is Popup popup)
                Dispatcher.BeginInvoke(new Action(() => DemotePopup(popup)));
        }

        private void NewSetting_Deactivated(object sender, EventArgs e)
        {
            DemotePopup(AddModelPopup);
            DemotePopup(McpServerPopup);
            DemotePopup(McpToolPermissionsPopup);
            DemotePopup(McpImportPopup);
        }

        private static void DemotePopup(Popup popup)
        {
            if (popup?.IsOpen != true || popup.Child == null) return;
            var source = PresentationSource.FromVisual(popup.Child) as HwndSource;
            if (source?.Handle == IntPtr.Zero) return;

            SetWindowPos(
                source.Handle,
                HwndNotTopmost,
                0,
                0,
                0,
                0,
                SwpNoMove | SwpNoSize | SwpNoActivate);
        }

        // Provider 默认端点
        private Dictionary<string, string> providerDefaultEndpoints = new Dictionary<string, string>()
        {
            {"OpenAI", "https://api.openai.com/v1/chat/completions"},
            {"Claude", "https://api.anthropic.com/v1/messages"},
            {"Gemini", "https://generativelanguage.googleapis.com/v1beta"},
        };

        // 服务商端点映射 (新增/编辑弹窗中供快速选择)
        private Dictionary<string, string> serviceProviderEndpoints = new Dictionary<string, string>()
        {
            {"DeepSeek", "https://api.deepseek.com/chat/completions"},
            {"通义千问", "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"},
            {"豆包", "https://ark.cn-beijing.volces.com/api/v3/chat/completions"},
            {"智谱GLM","https://open.bigmodel.cn/api/paas/v4/chat/completions" },
            {"Kimi", "https://api.moonshot.cn/v1/chat/completions"},
            {"Siliconflow", "https://api.siliconflow.cn/v1/chat/completions"},
            {"Azure OpenAI", "https://<your-azure-openai-endpoint>.openai.azure.com/"},
            {"OpenAI", "https://api.openai.com/v1/chat/completions"},
            {"vllm", "http://localhost:8000/v1/chat/completions"},
            {"llama.cpp", "http://localhost:8080/v1/chat/completions"},
            {"LM Studio", "http://localhost:1234/v1/chat/completions"},
            {"MiniMax", "https://api.minimaxi.chat/v1/chat/completions"},
            {"Claude", "https://api.anthropic.com/v1/messages"},
            {"Gemini", "https://generativelanguage.googleapis.com/v1beta"},
        };

        // 服务商 → Provider 类型映射 (Claude/Gemini 使用原生 API, 其他用 OpenAI 兼容)
        private Dictionary<string, string> serviceProviderTypes = new Dictionary<string, string>()
        {
            {"DeepSeek", "OpenAI"},
            {"通义千问", "OpenAI"},
            {"豆包", "OpenAI"},
            {"智谱GLM", "OpenAI"},
            {"Kimi", "OpenAI"},
            {"Siliconflow", "OpenAI"},
            {"Azure OpenAI", "OpenAI"},
            {"OpenAI", "OpenAI"},
            {"vllm", "OpenAI"},
            {"llama.cpp", "OpenAI"},
            {"LM Studio", "OpenAI"},
            {"MiniMax", "OpenAI"},
            {"Claude", "Claude"},
            {"Gemini", "Gemini"},
        };

        private string _addModeProviderType = "OpenAI"; // kept for ServiceProviderComboBox_SelectionChanged
        private IDisposable _settingsDraft;
        private string _lastSavedSettingsSnapshot;
        private bool _isLoadingSettings;
        private bool _suppressSettingsChangeTracking;
        private bool _settingsDirty;


        public NewSetting()
        {
            InitializeComponent();
            EndPointTextBox.TextChanged += ProbeSettingsChanged;
            ModelNameTextBox.TextChanged += ProbeSettingsChanged;
            ApiKeyPasswordBox.PasswordChanged += ProbeSettingsChanged;
            PopupReasoningEffort.SelectionChanged += ProbeSettingsChanged;
            PopupOutputTokenParameter.SelectionChanged += ProbeSettingsChanged;
            PopupContextWindow.TextChanged += ProbeSettingsChanged;
            PopupProviderTypeCombo.SelectionChanged += ProbeSettingsChanged;
            ThinkingEnabledToggle.Checked += ProbeSettingsChanged;
            ThinkingEnabledToggle.Unchecked += ProbeSettingsChanged;

            RegisterSettingsChangeTracking();
            _lastSavedSettingsSnapshot = JsonConvert.SerializeObject(AppConfig.Instance, Formatting.None);
            _settingsDraft = AppConfig.BeginDraft();
            _isLoadingSettings = true;
            try
            {
                _httpClient = new HttpClient();
                OllamaModels = new ObservableCollection<OllamaModelInfo>();
                ChatGPTModels = new ObservableCollection<ChatGPTModelInfo>();
                InitSkillAndMcp();
                LoadSettings();
                // Keep change tracking suspended while bindings populate the controls.
                DataContext = this;
                KeyVisable = false;
            }
            finally
            {
                _isLoadingSettings = false;
                // Initial control population can raise SelectionChanged events. They are not user edits.
                _lastSavedSettingsSnapshot = JsonConvert.SerializeObject(AppConfig.Instance, Formatting.None);
                SetSettingsDirty(false);
            }
        }

        private bool HasUnsavedSettings()
        {
            return _settingsDirty;
        }

        private void RegisterSettingsChangeTracking()
        {
            AddHandler(TextBoxBase.TextChangedEvent,
                new TextChangedEventHandler(SettingsTextChanged), true);
            AddHandler(PasswordBox.PasswordChangedEvent,
                new RoutedEventHandler(SettingsControlChanged), true);
            AddHandler(Selector.SelectionChangedEvent,
                new SelectionChangedEventHandler(SettingsSelectionChanged), true);
            AddHandler(ToggleButton.CheckedEvent,
                new RoutedEventHandler(SettingsControlChanged), true);
            AddHandler(ToggleButton.UncheckedEvent,
                new RoutedEventHandler(SettingsControlChanged), true);
            AddHandler(RangeBase.ValueChangedEvent,
                new RoutedPropertyChangedEventHandler<double>(SettingsValueChanged), true);
        }

        private void SettingsTextChanged(object sender, TextChangedEventArgs e)
        {
            if (IsUserInitiatedSettingEvent(e))
                SetSettingsDirty(true);
        }

        private void SettingsSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // The outer TabControl is navigation, not a setting. Its selection changes
            // must not enable Save or trigger the unsaved-settings prompt.
            if (ReferenceEquals(e.OriginalSource, SettingsNavigation)
                || ReferenceEquals(e.Source, SettingsNavigation)
                || e.OriginalSource is TabControl
                || e.OriginalSource is TabItem
                || e.OriginalSource is DataGrid
                || e.OriginalSource is DataGridRow)
            {
                return;
            }

            if (IsUserInitiatedSettingEvent(e))
                SetSettingsDirty(true);
        }

        private void SettingsValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (IsUserInitiatedSettingEvent(e))
                SetSettingsDirty(true);
        }

        private void SettingsControlChanged(object sender, RoutedEventArgs e)
        {
            if (IsUserInitiatedSettingEvent(e))
                SetSettingsDirty(true);
        }

        private bool IsUserInitiatedSettingEvent(RoutedEventArgs e)
        {
            if (_isLoadingSettings || _suppressSettingsChangeTracking)
                return false;

            return HasKeyboardFocusWithin(e?.OriginalSource as DependencyObject)
                || HasKeyboardFocusWithin(e?.Source as DependencyObject);
        }

        private static bool HasKeyboardFocusWithin(DependencyObject source)
        {
            // Do not walk up to the navigation TabControl. Selecting a page can
            // give that parent control focus even though no setting was edited.
            return source is UIElement element && element.IsKeyboardFocusWithin;
        }

        private async void OptimizeTextBox_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            await Utils.OptimizeTextBoxAsync(button?.Tag as TextBox, button);
        }

        private void SetSettingsDirty(bool dirty)
        {
            // Collection refreshes can raise routed control events without a user edit.
            // Preserve an actual edit made before the refresh completes.
            if (_suppressSettingsChangeTracking)
            {
                if (!dirty)
                {
                    _settingsDirty = false;
                    if (SaveSettingsButton != null)
                        SaveSettingsButton.IsEnabled = false;
                }
                return;
            }

            _settingsDirty = !_isLoadingSettings && dirty;
            if (SaveSettingsButton != null)
                SaveSettingsButton.IsEnabled = _settingsDirty;
        }

        private void CancelSettingsDraft()
        {
            if (_settingsDraft == null) return;
            AppConfig.Restore(_lastSavedSettingsSnapshot);
            _settingsDraft.Dispose();
            _settingsDraft = null;
            SetSettingsDirty(false);
        }

        private void NewSetting_Closing(object sender, CancelEventArgs e)
        {
            if (_settingsDraft == null || !HasUnsavedSettings())
            {
                _settingsDraft?.Dispose();
                _settingsDraft = null;
                return;
            }

            MessageBoxResult result = MessageBox.Show(
                Loc.T("I18n_SettingsUnsavedPrompt"),
                Loc.T("I18n_SettingsUnsavedTitle"),
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Cancel)
            {
                e.Cancel = true;
                return;
            }

            if (result == MessageBoxResult.Yes)
            {
                if (!SaveSettings())
                {
                    e.Cancel = true;
                    return;
                }
            }
            else
            {
                CancelSettingsDraft();
            }

            _settingsDraft?.Dispose();
            _settingsDraft = null;
        }

        private void CloseSettings_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void LoadSettings()
        {
            // 通用
            UpdateLogTextBlock.Text = SharedSettings.UpdateLog;
            LoadOllamaPathEnvironmentVariable();
            ThemeManager.ApplyTheme(this.Resources);

            ThemeToggleBtn.IsChecked = SharedSettings.themeType == "Light"? true : false;


            // Ollama
            LoadOllamaLocalModels();
            OllamaModelDetails.ItemsSource = OllamaModels;
            OllamaEndPoint.Text = SharedSettings.OllamaEndPoint;
            OllamaMaxToken.Text = SharedSettings.OllamaMaxToken.ToString();
            OllamaTemperatureSettingSlider.Value = SharedSettings.OllamaTemperature;
            OllamaModelSeletor.SelectedValue = SharedSettings.OllamaModel;

            // ChatGPT / 在线模型
            EnsureFileExists();
            LoadChatGPTModels();
            SetChatGPTModelsInvisable();
            ChatGPTModelDetails.ItemsSource = ChatGPTModels;
            UserMemorySettings memory = UserMemoryStore.Get();
            MemoryItems.Clear();
            foreach (UserMemoryItem item in memory.Memories) MemoryItems.Add(item);
            OutputPreferenceBox.Text = memory.OutputPreference;
            AutoMemoryToggle.IsChecked = memory.AutoMemory;
            MemoryModelSelector.Text = memory.MemoryModel;

            // Load provider settings
            string currentProvider = SharedSettings.OnlineProviderType;
            LoadActiveProviderSettings(currentProvider);

            OutputTokenParameterSelector.SelectedValue = AppConfig.Instance.OutputTokenParameter;
            ChatGPTMaxToken.Text = SharedSettings.ChatGPTMaxToken.ToString();
            ChatGPTTemperatureSettingSlider.Value = SharedSettings.ChatGPTTemperature;
            OnlineContextWindow.Text = SharedSettings.OnlineContextWindow.ToString();

            // Agent
            int agentMax = SharedSettings.AgentMaxIterations;
            if (agentMax == 0)
            {
                AgentUnlimitedCheck.IsChecked = true;
                AgentMaxIterationsBox.IsEnabled = false;
                AgentMaxIterationsBox.Text = "20";
            }
            else
            {
                AgentUnlimitedCheck.IsChecked = false;
                AgentMaxIterationsBox.IsEnabled = true;
                AgentMaxIterationsBox.Text = agentMax.ToString();
            }

            string executionMode = AppConfig.Instance.AgentExecutionMode ?? "TrackedChanges";
            var modeItems = AgentExecutionModeSelector.Items.Cast<ComboBoxItem>().ToList();
            int modeIndex = modeItems.FindIndex(item => string.Equals(item.Tag?.ToString(), executionMode, StringComparison.OrdinalIgnoreCase));
            AgentExecutionModeSelector.SelectedIndex = modeIndex >= 0 ? modeIndex : 2;
            AgentExternalToolsToggle.IsChecked = SharedSettings.AgentExternalToolsEnabled;
            DiagnosticLoggingToggle.IsChecked = DebugLogger.Enabled;



            // 设置 AI 模式
            if (SharedSettings.AIMode == "Ollama")
            {
                OllamaMode.IsChecked = true;
            }
            else if (SharedSettings.AIMode == "ChatGPT")
            {
                ChatGPTMode.IsChecked = true;
            }

            SelectLanguageMode(SharedSettings.LanguageMode);

            SelectDefaultOutputMode(SharedSettings.DefaultPromptOutputMode);

            if (OllamaNetModeNow)
            {
                OllamaNetMode.IsChecked = true;
            }
            else
            {
                OllamaNetMode.IsChecked = false;
            }

            // 静默审查开关
            LinterEnabledToggle.IsChecked = SharedSettings.LinterEnabled;
            LinterModelBox.Text = AppConfig.Instance.LinterModel ?? "";
            LinterEndPointBox.Text = AppConfig.Instance.LinterEndPoint ?? "";
            LinterDebounceBox.Text = AppConfig.Instance.LinterDebounceSeconds.ToString();

            // 任务提醒开关
            SuppressAgentNotifToggle.IsChecked = SharedSettings.SuppressAgentNotifications;

            // 计划确认 / 变更审查 免打扰开关
            SuppressPlanConfirmationToggle.IsChecked = SharedSettings.SuppressPlanConfirmation;
            SuppressDiffReviewToggle.IsChecked = SharedSettings.SuppressDiffReviewPrompt;

            // 思维链设置
            ThinkingEnabledToggle.IsChecked = SharedSettings.ThinkingEnabled;
            string effort = SharedSettings.ReasoningEffort;
            var effortItems = ReasoningEffortSelector.Items.Cast<ComboBoxItem>().ToList();
            int effortIndex = effortItems.FindIndex(item => (item.Tag?.ToString() ?? "") == effort);
            ReasoningEffortSelector.SelectedIndex = effortIndex >= 0 ? effortIndex : 0; // 旧 none 或未知值显示服务默认，保存后更新

            // Markdown 转换设置
            LoadMdStyles();

        }

        // Ollama 部分功能
        private async void LoadOllamaLocalModels(bool userClick = false)
        {
            bool suppressRefreshChanges = _isLoadingSettings && !userClick;
            try
            {
                var requestUri = AppConfig.Instance.OllamaEndPoint + "/api/tags";
                var response = await _httpClient.GetAsync(requestUri);
                response.EnsureSuccessStatusCode();

                var responseBody = await response.Content.ReadAsStringAsync();
                var jsonObject = JObject.Parse(responseBody);

                var models = jsonObject["models"];
                if (models != null)
                {
                    await Dispatcher.InvokeAsync(() =>
                    {
                        bool previousSuppression = _suppressSettingsChangeTracking;
                        _suppressSettingsChangeTracking |= suppressRefreshChanges;
                        try
                        {
                            OllamaModels.Clear();
                            foreach (var model in models)
                            {
                                string name = model["name"]?.ToString();
                                long sizeBytes = model["size"]?.ToObject<long>() ?? 0;
                                string parameterSize = model["details"]?["parameter_size"]?.ToString();
                                string modifiedAt = model["modified_at"]?.ToString();

                                double sizeGB = sizeBytes / (1024.0 * 1024.0 * 1024.0);
                                DateTime modifiedDate = DateTime.Parse(modifiedAt);
                                string formattedDate = modifiedDate.ToString("yyyy-MM-dd HH:mm:ss");

                                OllamaModels.Add(new OllamaModelInfo
                                {
                                    Name = name,
                                    SizeGB = sizeGB.ToString("F2"),
                                    ParameterSize = parameterSize,
                                    ModifiedAt = formattedDate
                                });
                            }
                        }
                        finally
                        {
                            _suppressSettingsChangeTracking = previousSuppression;
                        }
                    });
                }
            }
            catch (HttpRequestException ex) when (ex.InnerException?.Message.Contains("由于目标计算机积极拒绝，无法连接") == true ||
                                                 ex.InnerException?.Message.Contains("Connection refused") == true)
            {
                // Ollama未启动，静默忽略
                Debug.WriteLine("Ollama服务未启动，已忽略");
                return;
            }
            catch (Exception ex)
            {
                // 其他异常仍需报告
                Debug.WriteLine(ex);
                if (userClick == false)
                {
                    return;
                }
                else
                {
                    Globals.Ribbons.Ribbon1.ShowNotification(WordOllama.Properties.Resources.获取模型信息错误请检查Ollama程序状态, 1000, System.Windows.Forms.ToolTipIcon.Error);
                }
            }
        }




        private void ToggleTheme_Click(object sender, RoutedEventArgs e)
        {
           if (SharedSettings.themeType == "Dark")
            {
                SharedSettings.themeType = "Light";
            }
           else
            {
                SharedSettings.themeType = "Dark";
            }

            ThemeManager.ApplyTheme(this.Resources);

        }


        private async void ModelDownloadBtn_Click(object sender, RoutedEventArgs e)
        {
            txtStatus.Text = "";
            progressBar.Value = 0;
            progressBar.Visibility = Visibility.Visible;
            txtStatus.Visibility = Visibility.Visible;
            lblProgress.Text = "";
            ModelDownloadBtn.IsEnabled = false;

            string modelName = ModelNameInput.Text;
            try
            {
                if (!string.IsNullOrEmpty(modelName))
                {
                    await DownloadModel(modelName);
                }
                else
                {
                    MessageBox.Show(WordOllama.Properties.Resources.请先填写模型名字);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(WordOllama.Properties.Resources.错误 + ex);
            }
            finally
            {

                ModelDownloadBtn.IsEnabled = true;
            }


        }

        private async Task DownloadModel(string modelName)
        {
            try
            {
                var requestUri = AppConfig.Instance.OllamaEndPoint + "/api/pull";
                var content = new StringContent($"{{ \"name\": \"{modelName}\" }}", System.Text.Encoding.UTF8, "application/json");

                var request = new HttpRequestMessage(HttpMethod.Post, requestUri)
                {
                    Content = content
                };

                using var response = await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead);
                response.EnsureSuccessStatusCode();

                using var stream = await response.Content.ReadAsStreamAsync();
                using var reader = new StreamReader(stream);

                string line;
                while ((line = await reader.ReadLineAsync()) != null)
                {
                    await Dispatcher.InvokeAsync(() => UpdateUI(line));
                }

                MessageBox.Show(modelName + WordOllama.Properties.Resources.下载完成, WordOllama.Properties.Resources.下载成功, MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(modelName+ WordOllama.Properties.Resources.下载错误 + ex.Message, WordOllama.Properties.Resources.下载错误, MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                await Dispatcher.InvokeAsync(() =>
                {
                    progressBar.Visibility = Visibility.Collapsed;
                    txtStatus.Visibility = Visibility.Collapsed;
                    LoadOllamaLocalModels(); // Refresh the model list
                });
            }
        }

        private void UpdateUI(string jsonLine)
        {
            try
            {
                var jsonObject = JObject.Parse(jsonLine);
                string status = jsonObject["status"]?.ToString();
                long totalBytes = jsonObject["total"]?.ToObject<long>() ?? 0L;
                long completedBytes = jsonObject["completed"]?.ToObject<long>() ?? 0L;

                double totalMB = totalBytes / (1024.0 * 1024.0);
                double completedMB = completedBytes / (1024.0 * 1024.0);

                txtStatus.Text = WordOllama.Properties.Resources.状态 + status;

                if (status != "success")
                {
                    if (totalBytes > 0)
                    {
                        progressBar.Maximum = totalMB;
                        progressBar.Value = completedMB;
                        lblProgress.Text = $"{completedMB:F2} MB / {totalMB:F2} MB";
                    }
                    else
                    {
                        progressBar.Value = 0;
                        lblProgress.Text = "";
                    }
                }
                else
                {
                    progressBar.Value = progressBar.Maximum;
                    lblProgress.Text = "";
                    txtStatus.Text = WordOllama.Properties.Resources.下载完成;
                }
            }
            catch (Exception ex)
            {
                txtStatus.Text = WordOllama.Properties.Resources.错误+ex.Message;
            }
        }

        private void RefreshBtn_Click(object sender, RoutedEventArgs e)
        {
            LoadOllamaLocalModels(true);
        }

        private async void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            if (OllamaModelDetails.SelectedItem is OllamaModelInfo selectedModel)
            {
                var confirmResult = MessageBox.Show($"{WordOllama.Properties.Resources.要删除} {selectedModel.Name} ？",
                                                     WordOllama.Properties.Resources.确认删除,
                                                     MessageBoxButton.YesNo,
                                                     MessageBoxImage.Warning);
                if (confirmResult == MessageBoxResult.Yes)
                {
                    try
                    {
                        var requestUri = AppConfig.Instance.OllamaEndPoint + "/api/delete";
                        var content = new StringContent($"{{ \"name\": \"{selectedModel.Name}\" }}", System.Text.Encoding.UTF8, "application/json");

                        var request = new HttpRequestMessage(HttpMethod.Delete, requestUri)
                        {
                            Content = content
                        };

                        var response = await _httpClient.SendAsync(request);

                        if (response.IsSuccessStatusCode)
                        {
                            MessageBox.Show(WordOllama.Properties.Resources.模型成功删除, WordOllama.Properties.Resources.成功, MessageBoxButton.OK, MessageBoxImage.Information);
                            OllamaModels.Remove(selectedModel);
                        }
                        else
                        {
                            MessageBox.Show(WordOllama.Properties.Resources.模型删除失败, WordOllama.Properties.Resources.失败, MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"{WordOllama.Properties.Resources.删除时遇到错误} {ex.Message}", WordOllama.Properties.Resources.错误, MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            else
            {
                MessageBox.Show(WordOllama.Properties.Resources.请先选择要删除的模型);
            }
        }

        public class OllamaModelInfo
        {
            public string Name { get; set; }
            public string SizeGB { get; set; }
            public string ParameterSize { get; set; }
            public string ModifiedAt { get; set; }
        }

        // ChatGPT

        public class ChatGPTModelInfo
        {
            public string ProviderType { get; set; }
            public string EndPoint { get; set; }
            public string ModelName { get; set; }
            public string ApiKey { get; set; }
            public string ToolCallingMode { get; set; } = "Auto";
            public string ReasoningEffort { get; set; } = "high";
            public string LastCapabilityTestStatus { get; set; }
            public DateTime? LastCapabilityTestAt { get; set; }
            public string LastCapabilityTestError { get; set; }

            public string RealApiKey { get; set; }
        }

        // ========== Provider 相关辅助方法 ==========

        private string GetCurrentProviderType()
        {
            return SharedSettings.OnlineProviderType;
        }

        private void LoadActiveProviderSettings(string providerType)
        {
            switch (providerType)
            {
                case "OpenAI":
                    ChatGPTEndPoint.Text = SharedSettings.ChatGPTEndPoint;
                    ChatGPTAPIKey.Password = SharedSettings.ChatGPTKey;
                    ChatGPTModel.Text = SharedSettings.ChatGPTModel;
                    break;
                case "Claude":
                    ChatGPTEndPoint.Text = SharedSettings.ClaudeEndPoint;
                    ChatGPTAPIKey.Password = SharedSettings.ClaudeKey;
                    ChatGPTModel.Text = SharedSettings.ClaudeModel;
                    break;
                case "Gemini":
                    ChatGPTEndPoint.Text = SharedSettings.GeminiEndPoint;
                    ChatGPTAPIKey.Password = SharedSettings.GeminiKey;
                    ChatGPTModel.Text = SharedSettings.GeminiModel;
                    break;
            }
        }

        private void SaveActiveProviderSettings(string providerType)
        {
            switch (providerType)
            {
                case "OpenAI":
                    SharedSettings.ChatGPTEndPoint = ChatGPTEndPoint.Text;
                    SharedSettings.ChatGPTKey = ChatGPTAPIKey.Password;
                    SharedSettings.ChatGPTModel = ChatGPTModel.Text;
                    break;
                case "Claude":
                    SharedSettings.ClaudeEndPoint = ChatGPTEndPoint.Text;
                    SharedSettings.ClaudeKey = ChatGPTAPIKey.Password;
                    SharedSettings.ClaudeModel = ChatGPTModel.Text;
                    break;
                case "Gemini":
                    SharedSettings.GeminiEndPoint = ChatGPTEndPoint.Text;
                    SharedSettings.GeminiKey = ChatGPTAPIKey.Password;
                    SharedSettings.GeminiModel = ChatGPTModel.Text;
                    break;
            }
        }

        private void ResetEndPointButton_Click(object sender, RoutedEventArgs e)
        {
            string provider = GetCurrentProviderType();
            if (providerDefaultEndpoints.ContainsKey(provider))
            {
                EndPointTextBox.Text = providerDefaultEndpoints[provider];
            }
        }

        private async void OAuth2LoginButton_Click(object sender, RoutedEventArgs e)
        {
            // Google OAuth2 requires a valid OAuth 2.0 Client ID from Google Cloud Console.
            // Ask user for their Client ID first.
            string clientId = Microsoft.VisualBasic.Interaction.InputBox(
                "请输入 Google OAuth 2.0 客户端 ID (Client ID):\n\n" +
                "你需要在 Google Cloud Console 中:\n" +
                "1. 创建项目并启用 Generative Language API\n" +
                "2. 创建 OAuth 2.0 客户端 ID (桌面应用类型)\n" +
                "3. 将客户端 ID 粘贴到此处",
                "Google OAuth2 配置",
                "");

            if (string.IsNullOrWhiteSpace(clientId))
            {
                Globals.Ribbons.Ribbon1.ShowNotification(Loc.T("I18n_OAuthCancelled"), 1000);
                return;
            }

            try
            {
                OAuth2LoginButton.IsEnabled = false;
                OAuth2LoginButton.Content = Loc.T("I18n_Authenticating");

                string token = await OAuth2Helper.AuthenticateAsync(
                    providerName: "Google",
                    clientId: clientId,
                    authUrl: "https://accounts.google.com/o/oauth2/v2/auth",
                    tokenUrl: "https://oauth2.googleapis.com/token",
                    scope: "https://www.googleapis.com/auth/generative-language.retriever",
                    redirectPort: "51889"
                );

                if (!string.IsNullOrEmpty(token))
                {
                    ApiKeyPasswordBox.Password = token;
                    Globals.Ribbons.Ribbon1.ShowNotification(Loc.T("I18n_OAuthSucceeded"), 2000);
                }
                else
                {
                    Globals.Ribbons.Ribbon1.ShowNotification(Loc.T("I18n_OAuthFailed"), 3000, System.Windows.Forms.ToolTipIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(Loc.F("I18n_OAuthError", ex.Message));
            }
            finally
            {
                OAuth2LoginButton.IsEnabled = true;
                OAuth2LoginButton.Content = "OAuth2";
            }
        }

        private async void FetchModelsButton_Click(object sender, RoutedEventArgs e)
        {
            // Determine provider type from popup
            string provider = (PopupProviderTypeCombo.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "OpenAI";

            string endpoint = EndPointTextBox.Text;
            string apiKey = ApiKeyPasswordBox.Password;

            if (string.IsNullOrWhiteSpace(apiKey) && provider != "Gemini")
            {
                MessageBox.Show(Loc.T("I18n_ApiKeyRequired"));
                return;
            }

            FetchModelsButton.IsEnabled = false;
            FetchModelsButton.Content = Loc.T("I18n_Fetching");

            try
            {
                var providerInst = ProviderFactory.Create(provider, endpoint, apiKey, "");
                var models = await providerInst.FetchModelsAsync();

                ModelSelector.Items.Clear();
                foreach (var m in models)
                {
                    ModelSelector.Items.Add(m);
                }

                if (models.Count > 0)
                {
                    Globals.Ribbons.Ribbon1.ShowNotification(Loc.F("I18n_FetchedModels", models.Count), 2000);
                }
                else
                {
                    Globals.Ribbons.Ribbon1.ShowNotification(Loc.T("I18n_NoModelsFetched"), 3000);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(Loc.F("I18n_FetchModelsFailed", ex.Message));
            }
            finally
            {
                FetchModelsButton.IsEnabled = true;
                FetchModelsButton.Content = Loc.T("获取模型列表");
            }
        }

        private int _probeRevision;
        private JObject _reasoningEvidence;
        private string ProbeScope => ((PopupProviderTypeCombo.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "OpenAI")
            + "|" + OpenAIProvider.NormalizeEndpoint(EndPointTextBox.Text?.Trim()) + "|" + ModelNameTextBox.Text?.Trim();

        private void ProbeSettingsChanged(object sender, RoutedEventArgs e)
        {
            _probeRevision++;
            _lastCapabilityTestStatus = null;
            _lastCapabilityTestAt = null;
            _lastCapabilityTestError = null;
            if (ModelConnectionStatus == null) return;
            ModelConnectionStatus.Text = "设置已修改，请重新测试。";
            AdoptDefaultButton.Visibility = Visibility.Collapsed;
            if (_reasoningEvidence?["Scope"]?.ToString() != ProbeScope) _reasoningEvidence = null;
            PopupOutputTokenParameter.IsEnabled = ((PopupProviderTypeCombo.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "OpenAI") == "OpenAI";
            ShowReasoningEvidence();
        }

        private void ShowReasoningEvidence()
        {
            ReasoningEvidenceText.Text = _reasoningEvidence == null ? "强度尚未验证；可使用服务默认值测试。"
                : "最近检测（" + _reasoningEvidence["CheckedAt"] + "）：" + _reasoningEvidence["Detail"] + "。可重新测试刷新。";
            foreach (ComboBoxItem item in PopupReasoningEffort.Items)
                item.ToolTip = ReasoningEvidenceText.Text;
        }

        private async void RetryDefaultButton_Click(object sender, RoutedEventArgs e) => await RunModelProbeAsync(true);
        private void AdoptDefaultButton_Click(object sender, RoutedEventArgs e)
        {
            SetReasoningEffortSelection("default");
            ModelConnectionStatus.Text = "已选择服务默认值，请保存模型设置。";
        }
        private async void TestModelConnectionButton_Click(object sender, RoutedEventArgs e) => await RunModelProbeAsync(false);

        private async Task RunModelProbeAsync(bool useDefault)
        {
            string provider = (PopupProviderTypeCombo.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "OpenAI";
            string endpoint = EndPointTextBox.Text?.Trim();
            string apiKey = ApiKeyPasswordBox.Password;
            string model = ModelNameTextBox.Text?.Trim();

            if (string.IsNullOrWhiteSpace(endpoint) || string.IsNullOrWhiteSpace(model))
            {
                ModelConnectionStatus.Text = Loc.T("I18n_ConnectionTestMissingSettings");
                return;
            }

            if (string.IsNullOrWhiteSpace(apiKey) && !string.Equals(provider, "Gemini", StringComparison.OrdinalIgnoreCase))
            {
                ModelConnectionStatus.Text = Loc.T("I18n_ApiKeyRequired");
                return;
            }

            int revision = _probeRevision;
            string effort = useDefault ? "default" : GetSelectedReasoningEffort();
            TestModelConnectionButton.IsEnabled = false;
            RetryDefaultButton.IsEnabled = false;
            AdoptDefaultButton.Visibility = Visibility.Collapsed;
            ModelConnectionStatus.Text = Loc.T("I18n_TestingConnection");

            try
            {
                var providerInst = ProviderFactory.Create(provider, endpoint, apiKey, model, reasoningEffort: effort, thinkingEnabled: ThinkingEnabledToggle.IsChecked == true, outputTokenParameter: PopupOutputTokenParameter.SelectedValue?.ToString() ?? "auto");
                CapabilityProbeResult probe = await ProviderCapabilityProbe.RunAsync(providerInst);
                if (revision != _probeRevision) return;
                if (!useDefault) SetCapabilityTestResult(probe.Success, probe.Success ? null : probe.Summary);
                string actualEndpoint = provider == "OpenAI" ? OpenAIProvider.NormalizeEndpoint(endpoint) : endpoint;
                ModelConnectionStatus.Text = (probe.Success ? "连接正常" : "聊天连接未通过")
                    + "；测试强度：" + effort + (ThinkingEnabledToggle.IsChecked == true ? "" : "（思考开关已关闭）")
                    + (provider == "OpenAI" ? "\n请求地址：" : "\n配置地址：") + actualEndpoint + "\n" + probe.Summary;
                AdoptDefaultButton.Visibility = useDefault && probe.Success ? Visibility.Visible : Visibility.Collapsed;
                string detail = probe.Success ? "普通聊天已验证：" + (ThinkingEnabledToggle.IsChecked == true ? effort : "服务默认（思考开关关闭）") : "本次测试未通过，尚不能确认强度是否支持：" + effort;
                var supported = System.Text.RegularExpressions.Regex.Match(probe.Summary, @"Supported (?:types|values)[^\r\n]*", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (supported.Success) detail = "服务声明：" + supported.Value;
                _reasoningEvidence = new JObject { ["Scope"] = ProbeScope, ["CheckedAt"] = DateTime.Now.ToString("s"), ["Detail"] = detail };
                ShowReasoningEvidence();
            }
            catch (Exception ex)
            {
                if (revision == _probeRevision) SetCapabilityTestResult(false, ProviderRequestPolicy.NormalizeError(provider, null, null, ex));
            }
            finally
            {
                TestModelConnectionButton.IsEnabled = true;
                RetryDefaultButton.IsEnabled = true;
            }
        }

        private void SetCapabilityTestResult(bool succeeded, string error)
        {
            _lastCapabilityTestStatus = succeeded ? "Passed" : "Failed";
            _lastCapabilityTestAt = DateTime.Now;
            _lastCapabilityTestError = succeeded ? null : (error ?? Loc.T("I18n_ConnectionTestUnknownError"));

            ModelConnectionStatus.Text = succeeded
                ? Loc.F("I18n_ConnectionTestSuccess", _lastCapabilityTestAt.Value.ToString("g"))
                : Loc.F("I18n_ConnectionTestFailed", _lastCapabilityTestError);
        }

        private string FormatCapabilityTestStatus(string status, DateTime? testedAt, string error)
        {
            if (string.Equals(status, "Passed", StringComparison.OrdinalIgnoreCase))
                return Loc.F("I18n_ConnectionTestSuccess", testedAt?.ToString("g") ?? "");
            if (string.Equals(status, "Failed", StringComparison.OrdinalIgnoreCase))
                return Loc.F("I18n_ConnectionTestFailed", error ?? Loc.T("I18n_ConnectionTestUnknownError"));
            return string.Empty;
        }

        // ========== Popup draggable ==========

        private void ModelPopupBorder_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            _isPopupDragging = true;
            _popupDragStart = e.GetPosition(this);
            ((UIElement)sender).CaptureMouse();
        }

        private void ModelPopupBorder_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (_isPopupDragging)
            {
                var pos = e.GetPosition(this);
                var offset = pos - _popupDragStart;
                AddModelPopup.HorizontalOffset += offset.X;
                AddModelPopup.VerticalOffset += offset.Y;
                _popupDragStart = pos;
            }
        }

        private void ModelPopupBorder_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            _isPopupDragging = false;
            ((UIElement)sender).ReleaseMouseCapture();
        }

        private void ModelSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ModelSelector.SelectedItem is string selectedModel && !string.IsNullOrEmpty(selectedModel))
            {
                ModelNameTextBox.Text = selectedModel;
            }
        }

        // ========== 模型列表编辑 ==========

        private string _editingModelName = null; // null = add mode, non-null = edit mode
        private string _lastCapabilityTestStatus;
        private DateTime? _lastCapabilityTestAt;
        private string _lastCapabilityTestError;

        private void AddModelButton_Click(object sender, RoutedEventArgs e)
        {
            _editingModelName = null;
            _addModeProviderType = "OpenAI";
            ModelPopupTitle.Text = WordOllama.Properties.Resources.新增模型;
            ServiceProviderComboBox.SelectedIndex = 0;
            PopupProviderTypeCombo.SelectedIndex = 0; // OpenAI
            PopupToolCallingMode.SelectedIndex = 0;
            PopupReasoningEffort.SelectedIndex = 0;
            SetCapabilitySelection(ProviderRequestPolicy.GetDefaultCapabilities("OpenAI"));
            ModelConnectionStatus.Text = "";
            _lastCapabilityTestStatus = null;
            _lastCapabilityTestAt = null;
            _lastCapabilityTestError = null;
            ModelNameTextBox.Text = string.Empty;
            EndPointTextBox.Text = string.Empty;
            ApiKeyPasswordBox.Password = string.Empty;
            OAuth2LoginButton.Visibility = Visibility.Collapsed;
            ShowReasoningEvidence();
            AddModelPopup.IsOpen = true;
        }

        private void EditModelButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedModel = ChatGPTModelDetails.SelectedItem as ChatGPTModelInfo;
            if (selectedModel == null)
            {
                MessageBox.Show(WordOllama.Properties.Resources.请先选择一个模型);
                return;
            }

            _editingModelName = selectedModel.ModelName;
            ModelPopupTitle.Text = Loc.T("I18n_EditModel");

            // Set provider type in popup
            string pt = selectedModel.ProviderType ?? "OpenAI";
            foreach (ComboBoxItem item in PopupProviderTypeCombo.Items)
            {
                if (item.Content?.ToString() == pt)
                {
                    PopupProviderTypeCombo.SelectedItem = item;
                    break;
                }
            }

            ModelNameTextBox.Text = selectedModel.ModelName;
            EndPointTextBox.Text = selectedModel.EndPoint;
            ApiKeyPasswordBox.Password = selectedModel.RealApiKey ?? selectedModel.ApiKey;
            SetToolCallingModeSelection(selectedModel.ToolCallingMode);
            SetReasoningEffortSelection(selectedModel.ReasoningEffort);
            var storedModel = ChatGPTModelsJson?.OfType<JObject>().FirstOrDefault(model => model["ModelName"]?.ToString() == selectedModel.ModelName);
            _reasoningEvidence = storedModel?["ReasoningEvidence"] as JObject;
            if (_reasoningEvidence?["Scope"]?.ToString() != ProbeScope) _reasoningEvidence = null;
            SetCapabilitySelection(ProviderRequestPolicy.ReadCapabilities(storedModel?["Capabilities"], pt));
            ModelConnectionStatus.Text = FormatCapabilityTestStatus(selectedModel.LastCapabilityTestStatus, selectedModel.LastCapabilityTestAt, selectedModel.LastCapabilityTestError);
            _lastCapabilityTestStatus = selectedModel.LastCapabilityTestStatus;
            _lastCapabilityTestAt = selectedModel.LastCapabilityTestAt;
            _lastCapabilityTestError = selectedModel.LastCapabilityTestError;
            OAuth2LoginButton.Visibility = pt == "Gemini" ? Visibility.Visible : Visibility.Collapsed;
            ShowReasoningEvidence();
            AddModelPopup.IsOpen = true;
        }

        public static void EnsureFileExists()
        {
            string APIFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WordOllama", "models.json");

            try
            {
                if (!File.Exists(APIFilePath))
                {
                    // 创建一个空的JSON文件
                    var emptyList = new List<ChatGPTModelInfo>();
                    var json = JsonConvert.SerializeObject(emptyList, Formatting.Indented);
                    File.WriteAllText(APIFilePath, json);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{WordOllama.Properties.Resources.创建ChatGPT模型文件错误}{ex}");
            }
        }

        private void LoadChatGPTModels()
        {
            string APIFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WordOllama", "models.json");
            try
            {
                var json = File.ReadAllText(APIFilePath);
                ChatGPTModelsJson = JArray.Parse(json);
                bool migratedLegacySecret = false;
                foreach (var model in ChatGPTModelsJson.OfType<JObject>())
                {
                    var apiKey = model["ApiKey"]?.ToString();
                    if (!string.IsNullOrEmpty(apiKey))
                    {
                        migratedLegacySecret |= !SecretProtector.IsProtected(apiKey);
                        model["ApiKey"] = SecretProtector.UnprotectOrLegacy(apiKey);
                    }
                }
                if (migratedLegacySecret) SaveChatGPTModelsToFile();
            }
            catch (Exception e)
            {
                MessageBox.Show($"{WordOllama.Properties.Resources.读取ChatGPT列表错误}{e.Message}");
            }

        }
        private void SetChatGPTModelsVisable()
        {
            try
            {
                if (ChatGPTModelsJson != null)
                {
                    ChatGPTModels.Clear();
                    foreach (var model in ChatGPTModelsJson)
                    {
                        string providerType = model["ProviderType"]?.ToString() ?? "OpenAI";
                        string modelName = model["ModelName"]?.ToString();
                        string modelEndPoint = model["EndPoint"]?.ToString();
                        string modelApiKey = model["ApiKey"]?.ToString();

                        ChatGPTModels.Add(new ChatGPTModelInfo
                        {
                            ProviderType = providerType,
                            ModelName = modelName,
                            EndPoint = modelEndPoint,
                            ApiKey = modelApiKey,
                            RealApiKey = modelApiKey,
                            ToolCallingMode = model["ToolCallingMode"]?.ToString() ?? "Auto"
                            , ReasoningEffort = model["ReasoningEffort"]?.ToString() ?? "high"
                            , LastCapabilityTestStatus = model["LastCapabilityTestStatus"]?.ToString()
                            , LastCapabilityTestAt = model["LastCapabilityTestAt"]?.ToObject<DateTime?>()
                            , LastCapabilityTestError = model["LastCapabilityTestError"]?.ToString()
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{WordOllama.Properties.Resources.加载ChatGPT模型时出错} {ex.Message}", WordOllama.Properties.Resources.错误);
            }
        }

        private void OpenConfigFolderButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string path = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "WordOllama");

                // 检查目录是否存在，如果不存在则创建
                if (!System.IO.Directory.Exists(path))
                {
                    System.IO.Directory.CreateDirectory(path);
                    MessageBox.Show(WordOllama.Properties.Resources.设置文件目录不存在已为您创建,
                                  WordOllama.Properties.Resources.提示,
                                  MessageBoxButton.OK,
                                  MessageBoxImage.Information);
                }

                // 使用 explorer.exe 打开文件夹
                Process.Start(new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = path,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{WordOllama.Properties.Resources.打开设置文件目录时出错} {ex.Message}",
                              WordOllama.Properties.Resources.错误,
                              MessageBoxButton.OK,
                              MessageBoxImage.Error);
            }
        }

        private void SetChatGPTModelsInvisable()
        {
            try
            {
                if (ChatGPTModelsJson != null)
                {
                    ChatGPTModels.Clear();
                    foreach (var model in ChatGPTModelsJson)
                    {
                        string providerType = model["ProviderType"]?.ToString() ?? "OpenAI";
                        string modelName = model["ModelName"]?.ToString();
                        string modelEndPoint = model["EndPoint"]?.ToString();
                        string modelApiKey = model["ApiKey"]?.ToString();

                        ChatGPTModels.Add(new ChatGPTModelInfo
                        {
                            ProviderType = providerType,
                            ModelName = modelName,
                            EndPoint = modelEndPoint,
                            ApiKey = "**********",
                            RealApiKey = modelApiKey,
                            ToolCallingMode = model["ToolCallingMode"]?.ToString() ?? "Auto"
                            , ReasoningEffort = model["ReasoningEffort"]?.ToString() ?? "high"
                            , LastCapabilityTestStatus = model["LastCapabilityTestStatus"]?.ToString()
                            , LastCapabilityTestAt = model["LastCapabilityTestAt"]?.ToObject<DateTime?>()
                            , LastCapabilityTestError = model["LastCapabilityTestError"]?.ToString()
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{WordOllama.Properties.Resources.加载ChatGPT模型时出错} {ex.Message}", WordOllama.Properties.Resources.错误);
            }
        }

        private void ResetChatGPTModelsDataFrame()
        {
            if (!KeyVisable)
            {
                SetChatGPTModelsInvisable();
            }
            else
            {
                SetChatGPTModelsVisable();
            }
        }


        private void ToggleButton_Click(object sender, RoutedEventArgs e)
        {
            if (KeyVisable)
            {
                SetChatGPTModelsInvisable();
                KeyVisable = false;
            }
            else
            {
                SetChatGPTModelsVisable();
                KeyVisable = true;
            }
        }

        private void CancelAddModel_Click(object sender, RoutedEventArgs e)
        {
            _editingModelName = null;
            _addModeProviderType = "OpenAI";
            ModelNameTextBox.Text = string.Empty;
            EndPointTextBox.Text = string.Empty;
            ApiKeyPasswordBox.Password = string.Empty;
            ServiceProviderComboBox.SelectedIndex = 0;
            PopupProviderTypeCombo.SelectedIndex = 0;
            AddModelPopup.IsOpen = false;
        }

        private void SaveChatGPTModelsToFile()
        {
            string APIFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WordOllama", "models.json");
            try
            {
                var persistedModels = (JArray)ChatGPTModelsJson.DeepClone();
                foreach (var model in persistedModels.OfType<JObject>())
                {
                    var apiKey = model["ApiKey"]?.ToString();
                    if (!string.IsNullOrEmpty(apiKey))
                        model["ApiKey"] = SecretProtector.Protect(apiKey);
                }
                string json = persistedModels.ToString(Formatting.Indented);

                // 写入文件
                File.WriteAllText(APIFilePath, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{WordOllama.Properties.Resources.保存模型信息时出错}{ex.Message}", WordOllama.Properties.Resources.错误);
            }
        }


        private void ConfirmAddModel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string providerType = (PopupProviderTypeCombo.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "OpenAI";
                if (!IsHttpEndpoint(EndPointTextBox.Text) || string.IsNullOrWhiteSpace(ModelNameTextBox.Text))
                { ModelConnectionStatus.Text = "请填写有效的 HTTP/HTTPS 地址和模型名称。"; return; }
                if (!ProviderRequestPolicy.TryTokenCount(PopupContextWindow.Text, out int popupContext) || popupContext > 2000000)
                { ModelConnectionStatus.Text = "上下文容量请输入 1 至 2000000 的整数或 K 值。"; return; }
                if (SharedSettings.ChatGPTMaxToken >= popupContext)
                { ModelConnectionStatus.Text = "上下文容量必须大于当前输出上限（" + SharedSettings.ChatGPTMaxToken + "）。请先调整输出上限。"; return; }
                ProviderCapabilityProfile capabilities = GetSelectedCapabilities(providerType);

                var modelObj = new JObject
                {
                    { "ModelName", ModelNameTextBox.Text },
                    { "EndPoint", EndPointTextBox.Text },
                    { "ApiKey", ApiKeyPasswordBox.Password },
                    { "ProviderType", providerType },
                    { "ToolCallingMode", GetSelectedToolCallingMode() },
                    { "ReasoningEffort", GetSelectedReasoningEffort() },
                    { "Capabilities", JObject.FromObject(capabilities) }
                };
                modelObj["ReasoningEvidence"] = _reasoningEvidence?.DeepClone();
                modelObj["LastCapabilityTestStatus"] = _lastCapabilityTestStatus;
                modelObj["LastCapabilityTestAt"] = _lastCapabilityTestAt;
                modelObj["LastCapabilityTestError"] = _lastCapabilityTestError;

                JObject modelToActivate = modelObj;
                if (_editingModelName != null)
                {
                    // Edit mode: replace existing entry
                    var existing = ChatGPTModelsJson.OfType<JObject>().FirstOrDefault(m =>
                        m["ModelName"]?.ToString() == _editingModelName);
                    if (existing != null)
                    {
                        existing["ModelName"] = ModelNameTextBox.Text;
                        existing["EndPoint"] = EndPointTextBox.Text;
                        existing["ApiKey"] = ApiKeyPasswordBox.Password;
                        existing["ProviderType"] = providerType;
                        existing["ToolCallingMode"] = GetSelectedToolCallingMode();
                        existing["ReasoningEffort"] = GetSelectedReasoningEffort();
                        existing["Capabilities"] = JObject.FromObject(capabilities);
                        existing["ReasoningEvidence"] = _reasoningEvidence?.DeepClone();
                        existing["LastCapabilityTestStatus"] = _lastCapabilityTestStatus;
                        existing["LastCapabilityTestAt"] = _lastCapabilityTestAt;
                        existing["LastCapabilityTestError"] = _lastCapabilityTestError;
                        modelToActivate = existing;
                    }
                    _editingModelName = null;
                }
                else
                {
                    // Add mode
                    ChatGPTModelsJson.Add(modelObj);
                }

                ApplyModelAsActive(modelToActivate);

                ModelNameTextBox.Text = string.Empty;
                EndPointTextBox.Text = string.Empty;
                ApiKeyPasswordBox.Password = string.Empty;
                ResetChatGPTModelsDataFrame();
                SaveChatGPTModelsToFile();
                ServiceProviderComboBox.SelectedIndex = 0;
                AddModelPopup.IsOpen = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{WordOllama.Properties.Resources.新增模型错误}{ex}");
            }
        }

        private void ServiceProviderComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ServiceProviderComboBox.SelectedItem is ComboBoxItem selectedItem)
            {
                string name = selectedItem.Content.ToString();
                if (string.IsNullOrEmpty(name)) return;

                // Fill endpoint
                if (serviceProviderEndpoints.ContainsKey(name))
                    EndPointTextBox.Text = serviceProviderEndpoints[name];

                // Sync provider type
                string pt = serviceProviderTypes.ContainsKey(name) ? serviceProviderTypes[name] : "OpenAI";
                _addModeProviderType = pt;
                foreach (ComboBoxItem item in PopupProviderTypeCombo.Items)
                {
                    if (item.Content?.ToString() == pt)
                    {
                        PopupProviderTypeCombo.SelectedItem = item;
                        break;
                    }
                }

                // Show/hide OAuth2 button (Gemini only)
                OAuth2LoginButton.Visibility = name == "Gemini" ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        private void DeleteChatGPTModel_Click(object sender, RoutedEventArgs e)
        {
            var selectedModel = ChatGPTModelDetails.SelectedItem as ChatGPTModelInfo;
            if (selectedModel != null)
            {
                var result = MessageBox.Show($"{WordOllama.Properties.Resources.确定要删除模型} '{selectedModel.ModelName}' ？", WordOllama.Properties.Resources.确认删除, MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
                {
                    ChatGPTModelsJson.Remove(ChatGPTModelsJson.FirstOrDefault(m => m["ModelName"].ToString() == selectedModel.ModelName));
                    SaveChatGPTModelsToFile();
                    ResetChatGPTModelsDataFrame();
                    MessageBox.Show(WordOllama.Properties.Resources.模型已删除);
                }
            }
            else
            {
                MessageBox.Show(WordOllama.Properties.Resources.请先选择一个模型);
            }
        }

        private void LoadThisModel_Click(object sender, RoutedEventArgs e)
        {
            var selectedModel = ChatGPTModelDetails.SelectedItem as ChatGPTModelInfo;
            if (selectedModel != null)
            {
                var matchedModel = ChatGPTModelsJson.OfType<JObject>()
                    .FirstOrDefault(model => model["ModelName"]?.ToString() == selectedModel.ModelName);

                if (matchedModel != null)
                {
                    ApplyModelAsActive(matchedModel);
                    Globals.Ribbons.Ribbon1.ShowNotification(Loc.F("I18n_ModelLoaded", matchedModel["ModelName"]), 2000);
                }
                else
                {
                    MessageBox.Show(WordOllama.Properties.Resources.未找到匹配的模型信息);
                }
            }
            else
            {
                MessageBox.Show(WordOllama.Properties.Resources.请先选择一个模型);
            }
        }

        private void SetToolCallingModeSelection(string mode)
        {
            string value = ProviderCompatibility.ParseToolCallingMode(mode).ToString();
            foreach (ComboBoxItem item in PopupToolCallingMode.Items)
            {
                if (string.Equals(item.Tag?.ToString(), value, StringComparison.OrdinalIgnoreCase))
                {
                    PopupToolCallingMode.SelectedItem = item;
                    return;
                }
            }
            PopupToolCallingMode.SelectedIndex = 0;
        }

        private string GetSelectedToolCallingMode()
        {
            return (PopupToolCallingMode.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "Auto";
        }

        private string GetSelectedReasoningEffort() =>
            (PopupReasoningEffort.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "default";

        private void SetReasoningEffortSelection(string effort)
        {
            if (string.Equals(effort, "none", StringComparison.OrdinalIgnoreCase))
                effort = "default";
            foreach (ComboBoxItem item in PopupReasoningEffort.Items)
            {
                if (string.Equals(item.Tag?.ToString(), effort, StringComparison.OrdinalIgnoreCase))
                {
                    PopupReasoningEffort.SelectedItem = item;
                    return;
                }
            }
            PopupReasoningEffort.SelectedIndex = 0;
        }

        private void ApplyModelAsActive(JObject model)
        {
            string providerType = model?["ProviderType"]?.ToString() ?? "OpenAI";
            string endpoint = model?["EndPoint"]?.ToString() ?? "";
            string apiKey = model?["ApiKey"]?.ToString() ?? "";
            string modelName = model?["ModelName"]?.ToString() ?? "";
            SaveActiveProviderSettings(SharedSettings.OnlineProviderType);
            SharedSettings.OnlineProviderType = providerType;
            switch (providerType)
            {
                case "Claude":
                    SharedSettings.ClaudeEndPoint = endpoint;
                    SharedSettings.ClaudeKey = apiKey;
                    SharedSettings.ClaudeModel = modelName;
                    break;
                case "Gemini":
                    SharedSettings.GeminiEndPoint = endpoint;
                    SharedSettings.GeminiKey = apiKey;
                    SharedSettings.GeminiModel = modelName;
                    break;
                default:
                    SharedSettings.ChatGPTEndPoint = endpoint;
                    SharedSettings.ChatGPTKey = apiKey;
                    SharedSettings.ChatGPTModel = modelName;
                    break;
            }

            ProviderCapabilityProfile capabilities = ProviderRequestPolicy.ReadCapabilities(model?["Capabilities"], providerType);
            capabilities.ToolCallingMode = model?["ToolCallingMode"]?.ToString() ?? capabilities.ToolCallingMode;
            ProviderRequestPolicy.ApplyActiveCapabilities(capabilities);
            OutputTokenParameterSelector.SelectedValue = capabilities.OutputTokenParameter ?? "auto";
            OnlineContextWindow.Text = capabilities.ContextWindow.ToString();
            SharedSettings.ReasoningEffort = model?["ReasoningEffort"]?.ToString() ?? "high";
            LoadActiveProviderSettings(providerType);
            ChatGPTModel.Text = modelName;
            AppConfig.Save();
        }

        private ProviderCapabilityProfile GetSelectedCapabilities(string providerType)
        {
            var defaults = ProviderRequestPolicy.GetDefaultCapabilities(providerType);
            return new ProviderCapabilityProfile
            {
                ToolCallingMode = GetSelectedToolCallingMode(),
                SupportsStreaming = PopupSupportsStreaming.IsChecked == true,
                SupportsVision = PopupSupportsVision.IsChecked == true,
                SupportsJsonOutput = PopupSupportsJsonOutput.IsChecked == true,
                OutputTokenParameter = PopupOutputTokenParameter.SelectedValue?.ToString() ?? "auto",
                ContextWindow = ProviderRequestPolicy.TryTokenCount(PopupContextWindow.Text, out int contextWindow) && contextWindow > 0
                    ? contextWindow
                    : defaults.ContextWindow
            };
        }

        private void SetCapabilitySelection(ProviderCapabilityProfile capabilities)
        {
            capabilities ??= ProviderRequestPolicy.GetDefaultCapabilities("OpenAI");
            PopupSupportsStreaming.IsChecked = capabilities.SupportsStreaming;
            PopupSupportsVision.IsChecked = capabilities.SupportsVision;
            PopupSupportsJsonOutput.IsChecked = capabilities.SupportsJsonOutput;
            PopupContextWindow.Text = capabilities.ContextWindow.ToString();
            PopupOutputTokenParameter.SelectedValue = capabilities.OutputTokenParameter ?? "auto";
        }



        // 让 DataGrid 在鼠标悬停时也能滚动父级 ScrollViewer
        private void ChatGPTModelDetails_PreviewMouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs e)
        {
            var scrollViewer = FindVisualChild<ScrollViewer>(this);
            if (scrollViewer != null)
            {
                scrollViewer.ScrollToVerticalOffset(scrollViewer.VerticalOffset - e.Delta);
                e.Handled = true;
            }
        }

        private static T FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T found)
                    return found;
                var result = FindVisualChild<T>(child);
                if (result != null)
                    return result;
            }
            return null;
        }

        // ========== Markdown 转换设置 ==========

        private void LoadMdStyles()
        {
            // Populate combo boxes with Word styles from active document
            var styleNames = new List<string>();
            try
            {
                var doc = Globals.ThisAddIn.Application.ActiveDocument;
                if (doc != null)
                {
                    foreach (Word.Style style in doc.Styles)
                    {
                        if (style.Type == Microsoft.Office.Interop.Word.WdStyleType.wdStyleTypeParagraph
                            || style.Type == Microsoft.Office.Interop.Word.WdStyleType.wdStyleTypeCharacter)
                        {
                            string sn = style.NameLocal;
                            if (!string.IsNullOrEmpty(sn) && !styleNames.Contains(sn))
                                styleNames.Add(sn);
                        }
                    }
                }
            }
            catch { }

            var allCombos = new[] { MdStyleH1, MdStyleH2, MdStyleH3, MdStyleP, MdStyleCode, MdStyleBlockquote, MdStyleUl, MdStyleOl };
            foreach (var combo in allCombos)
            {
                combo.ItemsSource = styleNames;
            }

            // Apply saved values
            string json = SharedSettings.MarkdownStyleMappingsJson;
            if (string.IsNullOrEmpty(json))
            {
                // Defaults
                MdStyleH1.Text = "Heading 1";
                MdStyleH2.Text = "Heading 2";
                MdStyleH3.Text = "Heading 3";
                MdStyleP.Text = "Normal";
                MdStyleCode.Text = "Normal";
                MdStyleBlockquote.Text = "Normal";
                MdStyleUl.Text = "List Bullet";
                MdStyleOl.Text = "List Number";
                return;
            }

            try
            {
                var mappings = Newtonsoft.Json.JsonConvert.DeserializeObject<List<MarkdownStyleMapping>>(json);
                if (mappings == null) return;

                foreach (var m in mappings)
                {
                    switch (m.MarkdownElement)
                    {
                        case "h1": MdStyleH1.Text = m.WordStyleName; break;
                        case "h2": MdStyleH2.Text = m.WordStyleName; break;
                        case "h3": MdStyleH3.Text = m.WordStyleName; break;
                        case "p": MdStyleP.Text = m.WordStyleName; break;
                        case "code": MdStyleCode.Text = m.WordStyleName; break;
                        case "blockquote": MdStyleBlockquote.Text = m.WordStyleName; break;
                        case "ul": MdStyleUl.Text = m.WordStyleName; break;
                        case "ol": MdStyleOl.Text = m.WordStyleName; break;
                    }
                }
            }
            catch { }
        }

        private void RefreshMdStyles_Click(object sender, RoutedEventArgs e)
        {
            LoadMdStyles();
            Globals.Ribbons.Ribbon1.ShowNotification(Loc.T("I18n_StyleListRefreshed"), 1000);
        }

        private void CreateNewWordStyle_Click(object sender, RoutedEventArgs e)
        {
            string styleName = NewStyleNameBox.Text?.Trim();
            if (string.IsNullOrEmpty(styleName))
            {
                MessageBox.Show(Loc.T("I18n_StyleNameRequired"));
                return;
            }

            try
            {
                var doc = Globals.ThisAddIn.Application.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show(Loc.T("I18n_OpenWordDocument"));
                    return;
                }

                // Check if style already exists
                bool exists = false;
                try { exists = doc.Styles[styleName] != null; } catch { }

                if (!exists)
                {
                    var newStyle = doc.Styles.Add(styleName, Microsoft.Office.Interop.Word.WdStyleType.wdStyleTypeParagraph);
                    newStyle.Font.Name = "等线";
                    newStyle.Font.Size = 11;
                    Globals.Ribbons.Ribbon1.ShowNotification(Loc.F("I18n_StyleCreated", styleName), 2000);
                }
                else
                {
                    Globals.Ribbons.Ribbon1.ShowNotification(Loc.F("I18n_StyleExists", styleName), 2000);
                }
                NewStyleNameBox.Text = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show(Loc.F("I18n_CreateStyleFailed", ex.Message));
            }
        }

        // 保存
        private void Save_btn_Click(object sender, RoutedEventArgs e)
        {
            if (SaveSettings())
            {
                Globals.Ribbons.Ribbon1.ShowNotification(WordOllama.Properties.Resources.保存成功, 1000);
            }
        }

        private bool SaveSettings()
        {
            ValidationMessage.Text = "";
            if (!ValidateSettings(out string validationError))
            {
                ValidationMessage.Text = validationError;
                return false;
            }

            // 通用
            if (OllamaMode.IsChecked == true)
            {
                SharedSettings.AIMode = "Ollama";
            }
            else if (ChatGPTMode.IsChecked == true)
            {
                SharedSettings.AIMode = "ChatGPT";
            }

            SharedSettings.LanguageMode = (LangMode.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "";



            // Ollama
            SharedSettings.OllamaEndPoint = OllamaEndPoint.Text;
            SharedSettings.OllamaMaxToken = int.Parse(OllamaMaxToken.Text.Trim());
            SharedSettings.OllamaTemperature = OllamaTemperatureSettingSlider.Value;
            SharedSettings.OllamaModel = OllamaModelSeletor.SelectedValue?.ToString() ?? "";

            // Online model — save current provider settings
            string currentProvider = GetCurrentProviderType();
            SharedSettings.OnlineProviderType = currentProvider;
            SaveActiveProviderSettings(currentProvider);

            AppConfig.Instance.OutputTokenParameter = OutputTokenParameterSelector.SelectedValue?.ToString() ?? "auto";
            SharedSettings.ChatGPTMaxToken = int.Parse(ChatGPTMaxToken.Text.Trim());
            SharedSettings.ChatGPTTemperature = ChatGPTTemperatureSettingSlider.Value;
            if (ProviderRequestPolicy.TryTokenCount(OnlineContextWindow.Text, out int ctxWindow) && ctxWindow > 0)
                SharedSettings.OnlineContextWindow = ctxWindow;

            UserMemoryStore.Save(new UserMemorySettings
            {
                Memories = MemoryItems.ToList(),
                OutputPreference = OutputPreferenceBox.Text,
                AutoMemory = AutoMemoryToggle.IsChecked == true,
                MemoryModel = MemoryModelSelector.Text
            });

            // 思维链
            SharedSettings.ThinkingEnabled = ThinkingEnabledToggle.IsChecked == true;
            if (ReasoningEffortSelector.SelectedItem is ComboBoxItem effortItem)
                SharedSettings.ReasoningEffort = effortItem.Tag?.ToString() ?? "default";

            // Agent
            if (AgentUnlimitedCheck.IsChecked == true)
            {
                SharedSettings.AgentMaxIterations = 0;
            }
            else
            {
                int val;
                if (int.TryParse(AgentMaxIterationsBox.Text, out val) && val > 0)
                    SharedSettings.AgentMaxIterations = val;
                else
                    SharedSettings.AgentMaxIterations = 20;
            }

            // Markdown 转换样式
            var mdMappings = new List<MarkdownStyleMapping>
            {
                new() { MarkdownElement = "h1", WordStyleName = MdStyleH1.Text },
                new() { MarkdownElement = "h2", WordStyleName = MdStyleH2.Text },
                new() { MarkdownElement = "h3", WordStyleName = MdStyleH3.Text },
                new() { MarkdownElement = "p", WordStyleName = MdStyleP.Text },
                new() { MarkdownElement = "code", WordStyleName = MdStyleCode.Text },
                new() { MarkdownElement = "blockquote", WordStyleName = MdStyleBlockquote.Text },
                new() { MarkdownElement = "ul", WordStyleName = MdStyleUl.Text },
                new() { MarkdownElement = "ol", WordStyleName = MdStyleOl.Text },
            };
            SharedSettings.MarkdownStyleMappingsJson =
                Newtonsoft.Json.JsonConvert.SerializeObject(mdMappings, Newtonsoft.Json.Formatting.Indented);

            // 在设置窗口草稿期间也立即落盘，但继续保留草稿状态，便于用户继续修改后再次保存或取消。
            string provider = SharedSettings.OnlineProviderType;
            string activeModel = provider == "Claude" ? SharedSettings.ClaudeModel : provider == "Gemini" ? SharedSettings.GeminiModel : SharedSettings.ChatGPTModel;
            string activeEndpoint = provider == "Claude" ? SharedSettings.ClaudeEndPoint : provider == "Gemini" ? SharedSettings.GeminiEndPoint : SharedSettings.ChatGPTEndPoint;
            var stored = ChatGPTModelsJson?.OfType<JObject>().FirstOrDefault(m =>
                (m["ProviderType"]?.ToString() ?? "OpenAI") == provider && m["ModelName"]?.ToString() == activeModel
                && OpenAIProvider.NormalizeEndpoint(m["EndPoint"]?.ToString()) == OpenAIProvider.NormalizeEndpoint(activeEndpoint));
            if (stored != null)
            {
                stored["Capabilities"] = JObject.FromObject(ProviderRequestPolicy.GetActiveCapabilities());
                stored["ReasoningEffort"] = SharedSettings.ReasoningEffort;
                stored["LastCapabilityTestStatus"] = null;
                stored["LastCapabilityTestAt"] = null;
                stored["LastCapabilityTestError"] = null;
                SaveChatGPTModelsToFile();
            }
            AppConfig.SaveNow();
            _lastSavedSettingsSnapshot = JsonConvert.SerializeObject(AppConfig.Instance, Formatting.None);
            SetSettingsDirty(false);
            return true;
        }

        private void SelectLanguageMode(string languageMode)
        {
            string target = languageMode?.Trim() ?? "";
            foreach (ComboBoxItem item in LangMode.Items)
            {
                if (string.Equals(item.Tag?.ToString() ?? "", target, StringComparison.OrdinalIgnoreCase))
                {
                    LangMode.SelectedItem = item;
                    return;
                }
            }

            LangMode.SelectedIndex = 0;
        }

        private bool ValidateSettings(out string message)
        {
            var errors = new List<string>();
            bool ollamaSelected = OllamaMode.IsChecked == true;
            if (ollamaSelected && !IsHttpEndpoint(OllamaEndPoint.Text))
                errors.Add(Loc.T("I18n_SettingsInvalidEndpoint").Replace("{0}", "Ollama"));

            if (!TryReadRange(OllamaMaxToken, "Ollama token limit", 1, 1000000, errors))
                errors.Add(Loc.F("I18n_SettingsInvalidNumber", "Ollama token limit", 1, 1000000));
            if (!TryReadRange(ChatGPTMaxToken, "Online model token limit", 1, 1000000, errors))
                errors.Add(Loc.F("I18n_SettingsInvalidNumber", "Online model token limit", 1, 1000000));
            if (!TryReadRange(OnlineContextWindow, "Online context window", 1, 2000000, errors))
                errors.Add(Loc.F("I18n_SettingsInvalidNumber", "Online context window", 1, 2000000));
            if (AgentUnlimitedCheck.IsChecked != true &&
                !TryReadRange(AgentMaxIterationsBox, "Agent iteration limit", 1, 1000, errors))
                errors.Add(Loc.F("I18n_SettingsInvalidNumber", "Agent iteration limit", 1, 1000));

            if (ProviderRequestPolicy.TryTokenCount(OnlineContextWindow.Text, out int context) && int.TryParse(ChatGPTMaxToken.Text, out int output) && output >= context)
                errors.Add("输出上限必须小于上下文容量，为输入留出空间。");
            message = errors.Count == 0
                ? ""
                : Loc.F("I18n_SettingsValidationFailed", string.Join(" ", errors.Distinct()));
            return errors.Count == 0;
        }

        private void ContextPreset_Click(object sender, RoutedEventArgs e)
        {
            string[] parts = ((Button)sender).Tag.ToString().Split('|');
            ((TextBox)FindName(parts[0])).Text = parts[1];
            SetSettingsDirty(true);
        }

        private static bool IsHttpEndpoint(string value)
        {
            return Uri.TryCreate(value?.Trim(), UriKind.Absolute, out Uri uri) &&
                   (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps);
        }

        private static bool TryReadRange(TextBox box, string _, int min, int max, List<string> errors)
        {
            return (box?.Name == "OnlineContextWindow" ? ProviderRequestPolicy.TryTokenCount(box.Text, out int value) : int.TryParse(box?.Text?.Trim(), out value)) && value >= min && value <= max;
        }

        // Agent 设置
        private void AgentUnlimitedCheck_Click(object sender, RoutedEventArgs e)
        {
            AgentMaxIterationsBox.IsEnabled = AgentUnlimitedCheck.IsChecked != true;
        }

        private void AddMemory_Click(object sender, RoutedEventArgs e)
        {
            string content = MemoryNewBox.Text?.Trim();
            if (string.IsNullOrWhiteSpace(content)) return;
            MemoryItems.Add(new UserMemoryItem { Content = content });
            MemoryNewBox.Text = "";
            SetSettingsDirty(true);
        }

        private void DeleteSelectedMemory_Click(object sender, RoutedEventArgs e)
        {
            foreach (UserMemoryItem item in MemoryItemsGrid.SelectedItems.Cast<UserMemoryItem>().ToList()) MemoryItems.Remove(item);
            SetSettingsDirty(true);
        }

        private void DeleteMemoryRow_Click(object sender, RoutedEventArgs e)
        {
            if ((sender as FrameworkElement)?.Tag is UserMemoryItem item && MemoryItems.Remove(item))
                SetSettingsDirty(true);
        }

        private void DeleteAllMemory_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show(Loc.T("I18n_DeleteAllMemoryConfirm"), Loc.T("I18n_MemorySettings"), MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes) return;
            MemoryItems.Clear();
            SetSettingsDirty(true);
        }

        private void AgentExecutionModeSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isLoadingSettings) return;

            if (AgentExecutionModeSelector.SelectedItem is ComboBoxItem item)
            {
                AppConfig.Instance.AgentExecutionMode = item.Tag?.ToString() ?? "TrackedChanges";
                AppConfig.Save();
            }
        }

        private void AgentExternalToolsToggle_Click(object sender, RoutedEventArgs e)
        {
            if (sender is ToggleButton toggle)
                SharedSettings.AgentExternalToolsEnabled = toggle.IsChecked == true;
        }

        private void DiagnosticLoggingToggle_Click(object sender, RoutedEventArgs e)
        {
            if (DiagnosticLoggingToggle.IsChecked == true)
                DebugLogger.EnableForCurrentSession();
            else
                DebugLogger.Disable();
        }

        private void ClearDiagnosticLog_Click(object sender, RoutedEventArgs e)
        {
            DebugLogger.Clear();
            Globals.Ribbons.Ribbon1.ShowNotification(Loc.T("I18n_DiagnosticLogCleared"), 1500);
        }

        private void CopyDiagnostics_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Clipboard.SetText(DiagnosticsService.BuildSupportBundle());
                Globals.Ribbons.Ribbon1.ShowNotification(Loc.T("I18n_DiagnosticsCopied"), 1800);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Loc.T("I18n_Diagnostics"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        // 进阶
        private void SetEnvironmentVariableButton_Click(object sender, RoutedEventArgs e)
        {
            string variableName = "OLLAMA_MODELS";
            string variableValue = OllamaModelPath.Text;
            string value = Environment.GetEnvironmentVariable("OLLAMA_MODELS", EnvironmentVariableTarget.User);
            try
            {
                if (IsValidPath(variableValue))
                {

                    Task.Factory.StartNew(() => Environment.SetEnvironmentVariable(variableName, variableValue, EnvironmentVariableTarget.User));
                    var confirmResult = MessageBox.Show(WordOllama.Properties.Resources.设置成功需要移动原有模型文件吗,
                                                     WordOllama.Properties.Resources.移动模型文件,
                                                     MessageBoxButton.YesNo,
                                                     MessageBoxImage.Question);
                    if (confirmResult == MessageBoxResult.Yes)
                    {
                        if (!string.IsNullOrEmpty(value))
                        {
                            MoveFilesAndFolders(value, variableValue);
                        }
                        else
                        {
                            string sourcePath = Path.Combine(
                                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                                ".ollama",
                                "models");
                            MoveFilesAndFolders(sourcePath, variableValue);
                        }
                    }
                    MessageBox.Show(WordOllama.Properties.Resources.设置成功请重启Ollama);
                }
                else
                {
                    MessageBox.Show(WordOllama.Properties.Resources.路径格式错误请检查);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{WordOllama.Properties.Resources.设置遇到错误}{ex}");
            }
            finally
            {
                LoadOllamaPathEnvironmentVariable();
            }

        }

        private void LoadOllamaPathEnvironmentVariable()
        {
            try
            {
                string value = Environment.GetEnvironmentVariable("OLLAMA_MODELS", EnvironmentVariableTarget.User);
                if (!string.IsNullOrEmpty(value))
                {
                    OllamaModelPath.Text = value;
                }
                else
                {
                    OllamaModelPath.Text = string.Empty;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{WordOllama.Properties.Resources.读取Ollama环境变量错误}{ex}");
            }
        }


        public static bool IsValidPath(string path)
        {
            try
            {
                Path.GetFullPath(path);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }


        private void ResetOllamaModelPath_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string value = Environment.GetEnvironmentVariable("OLLAMA_MODELS", EnvironmentVariableTarget.User);
                Task.Factory.StartNew(() => Environment.SetEnvironmentVariable("OLLAMA_MODELS", null, EnvironmentVariableTarget.User));
                if (!string.IsNullOrEmpty(value))
                {
                    var confirmResult = MessageBox.Show(WordOllama.Properties.Resources.恢复成功需要移动原有模型文件吗,
                                                     WordOllama.Properties.Resources.移动模型文件,
                                                     MessageBoxButton.YesNo,
                                                     MessageBoxImage.Question);
                    if (confirmResult == MessageBoxResult.Yes)
                    {
                        string destinationPath = Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                            ".ollama",
                            "models");
                        MoveFilesAndFolders(value, destinationPath);
                    }

                    MessageBox.Show(WordOllama.Properties.Resources.已恢复默认路径请重启Ollama);
                }
                else
                {
                    MessageBox.Show(WordOllama.Properties.Resources.路径为默认状态无需恢复);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{WordOllama.Properties.Resources.恢复遇到错误}{ex}");
            }
            finally
            {
                LoadOllamaPathEnvironmentVariable();
            }


        }

        private void MoveFilesAndFolders(string sourcePath, string destinationPath)
        {
            try
            {
                // 获取源路径中的所有文件和文件夹
                var directories = Directory.GetDirectories(sourcePath);
                var files = Directory.GetFiles(sourcePath);

                // 创建目标路径（如果不存在）
                if (!Directory.Exists(destinationPath))
                {
                    Directory.CreateDirectory(destinationPath);
                }

                // 移动文件夹
                foreach (var dir in directories)
                {
                    string dirName = Path.GetFileName(dir);
                    string destDir = Path.Combine(destinationPath, dirName);

                    // 如果源路径和目标路径在同一卷上，使用 Directory.Move
                    if (Path.GetPathRoot(dir) == Path.GetPathRoot(destDir))
                    {
                        Directory.Move(dir, destDir);
                    }
                    else
                    {
                        // 如果源路径和目标路径在不同卷上，使用 Directory.Move 和 File.Move
                        DirectoryCopy(dir, destDir, true);
                        Directory.Delete(dir, true);
                    }
                }

                // 移动文件
                foreach (var file in files)
                {
                    string fileName = Path.GetFileName(file);
                    string destFile = Path.Combine(destinationPath, fileName);

                    // 如果源路径和目标路径在同一卷上，使用 File.Move
                    if (Path.GetPathRoot(file) == Path.GetPathRoot(destFile))
                    {
                        File.Move(file, destFile);
                    }
                    else
                    {
                        // 如果源路径和目标路径在不同卷上，使用 File.Copy 和 File.Delete
                        File.Copy(file, destFile, true);
                        File.Delete(file);
                    }
                }

                MessageBox.Show(WordOllama.Properties.Resources.文件和文件夹已成功移动);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{WordOllama.Properties.Resources.移动文件和文件夹时出错} {ex.Message}");
            }
            finally
            {
                DialogHost.CloseDialogCommand.Execute(null, null);
            }
        }

        private void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
        {
            // 获取源目录信息
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);

            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException(
                    WordOllama.Properties.Resources.源目录不存在或无法找到
                    + sourceDirName);
            }

            // 如果目标目录不存在，则创建它
            if (!Directory.Exists(destDirName))
            {
                Directory.CreateDirectory(destDirName);
            }

            // 获取源目录中的所有文件并复制它们
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string temppath = Path.Combine(destDirName, file.Name);
                file.CopyTo(temppath, false);
            }

            // 如果需要复制子目录，则递归调用此方法
            if (copySubDirs)
            {
                DirectoryInfo[] dirs = dir.GetDirectories();
                foreach (DirectoryInfo subdir in dirs)
                {
                    string temppath = Path.Combine(destDirName, subdir.Name);
                    DirectoryCopy(subdir.FullName, temppath, copySubDirs);
                }
            }
        }

        private void OllamaNetMode_Click(object sender, RoutedEventArgs e)
        {

            if (OllamaNetModeNow)
            {
                if (ResetEnvironment("OLLAMA_HOST"))
                {
                    SharedSettings.OllamaNetMode = false;
                    OllamaNetModeNow = false;
                }
            }
            else
            {
                if (SetEnvironment("OLLAMA_HOST", "0,0,0,0"))
                {
                    SharedSettings.OllamaNetMode = true;
                    OllamaNetModeNow = true;
                }
            }
        }

        private bool SetEnvironment(string env, string value)
        {
            try
            {
                Task.Factory.StartNew(() => Environment.SetEnvironmentVariable(env, value, EnvironmentVariableTarget.User));
                MessageBox.Show(string.Format(WordOllama.Properties.Resources.已为0的值修改成1, env, value));
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format(WordOllama.Properties.Resources._0写入错误1, env, ex));
                return false;
            }
        }

        private bool ResetEnvironment(string env)
        {
            try
            {
                Task.Factory.StartNew(() => Environment.SetEnvironmentVariable(env, null, EnvironmentVariableTarget.User));
                MessageBox.Show(string.Format(WordOllama.Properties.Resources.已恢复0的默认值, env));
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format(WordOllama.Properties.Resources._0重置错误1,env,ex));
                return false;
            }
        }

        private void SetOllamaPort_Click(object sender, RoutedEventArgs e)
        {
            SetEnvironment("OLLAMA_PORT", OllamaPortText.Text);
        }

        private void ResetOllamaPort_Click(object sender, RoutedEventArgs e)
        {
            ResetEnvironment("OLLAMA_PORT");
        }

        private void SetOllamaAlive_Click(object sender, RoutedEventArgs e)
        {
            SetEnvironment("OLLAMA_KEEP_ALIVE", OllamaAliveText.Text);
        }

        private void ResetOllamaAlive_Click(object sender, RoutedEventArgs e)
        {
            ResetEnvironment("OLLAMA_KEEP_ALIVE");
        }

        private void SetOllamaMaxLoadModel_Click(object sender, RoutedEventArgs e)
        {
            SetEnvironment("OLLAMA_MAX_LOADED_MODELS", OllamaMaxLoadModelText.Text);
        }

        private void ResetOllamaMaxLoadModel_Click(object sender, RoutedEventArgs e)
        {
            ResetEnvironment("OLLAMA_MAX_LOADED_MODELS");
        }

        #region ========== SKILL & MCP 设置 ==========

        private string _skillsPath;
        private string _mcpConfigPath;
        private SkillManager _skillManager = new SkillManager();
        private ObservableCollection<McpServerConfig> _mcpServers = new ObservableCollection<McpServerConfig>();

        /// <summary>
        /// SKILL & MCP Tab 初始化 — 在 NewSetting() 构造函数中调用。
        /// </summary>
        private void InitSkillAndMcp()
        {
            _skillsPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "WordAgentSkills");
            _mcpConfigPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "WordOllama", "mcp-servers.json");

            LoadSkillsList();
            LoadMcpServers();
        }

        // ========== Skills ==========

        private void LoadSkillsList()
        {
            _skillManager.LoadSkills(_skillsPath);
            var items = new ObservableCollection<SkillDisplayItem>();
            foreach (var s in _skillManager.GetAllSkills())
            {
                items.Add(new SkillDisplayItem
                {
                    Name = s.Name,
                    Description = s.Description,
                    FolderPath = Path.GetDirectoryName(s.FilePath)
                });
            }
            SkillDataGrid.ItemsSource = items;
        }

        private void OpenSkillFolder_Click(object sender, RoutedEventArgs e)
        {
            if (!Directory.Exists(_skillsPath))
                Directory.CreateDirectory(_skillsPath);
            Process.Start("explorer.exe", _skillsPath);
        }

        private void AboutLink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
            e.Handled = true;
        }

        private void ImportSkillZip_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = Loc.T("I18n_SelectSkillZipTitle"),
                Filter = Loc.T("I18n_ZipFileFilter"),
                Multiselect = false
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                if (!Directory.Exists(_skillsPath))
                    Directory.CreateDirectory(_skillsPath);

                // 解压到以 ZIP 文件名为名的子目录
                string zipName = Path.GetFileNameWithoutExtension(dlg.FileName);
                string extractPath = Path.Combine(_skillsPath, zipName);

                // 如果目录已存在，加数字后缀
                int suffix = 1;
                while (Directory.Exists(extractPath))
                {
                    extractPath = Path.Combine(_skillsPath, $"{zipName}_{suffix}");
                    suffix++;
                }

                System.IO.Compression.ZipFile.ExtractToDirectory(dlg.FileName, extractPath);

                // 展平：如果 SKILL.md 在唯一的子目录中，上移一层去除多余嵌套
                if (!File.Exists(Path.Combine(extractPath, "SKILL.md")))
                {
                    var found = Directory.GetFiles(extractPath, "SKILL.md", SearchOption.AllDirectories);
                    if (found.Length > 0)
                    {
                        string innerDir = Path.GetDirectoryName(found[0]);
                        // 检查是否为单层子目录结构（ZIP 自带的顶层文件夹）
                        var topDirs = Directory.GetDirectories(extractPath);
                        if (topDirs.Length == 1 && string.Equals(topDirs[0], innerDir, StringComparison.OrdinalIgnoreCase))
                        {
                            // 将子目录中所有内容上移一层
                            foreach (var dir in Directory.GetDirectories(innerDir))
                            {
                                string targetDir = Path.Combine(extractPath, Path.GetFileName(dir));
                                if (!Directory.Exists(targetDir))
                                    Directory.Move(dir, targetDir);
                            }
                            foreach (var file in Directory.GetFiles(innerDir))
                            {
                                string targetFile = Path.Combine(extractPath, Path.GetFileName(file));
                                if (!File.Exists(targetFile))
                                    File.Move(file, targetFile);
                            }
                            // 删除空的子目录
                            try { Directory.Delete(innerDir, true); } catch { }
                        }
                    }
                    else
                    {
                        Directory.Delete(extractPath, true);
                        MessageBox.Show(Loc.T("I18n_SkillZipNoSkillMd"), Loc.T("I18n_ImportFailedTitle"),
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                }

                LoadSkillsList();
                Globals.Ribbons.Ribbon1.ShowNotification(Loc.F("I18n_SkillImportSucceeded", zipName), 2000);
            }
            catch (Exception ex)
            {
                MessageBox.Show(Loc.F("I18n_ImportFailed", ex.Message), Loc.T("I18n_ErrorTitle"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RefreshSkills_Click(object sender, RoutedEventArgs e)
        {
            LoadSkillsList();
        }

        private void DeleteSkill_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string folderPath)
            {
                var result = MessageBox.Show(
                    Loc.F("I18n_DeleteSkillConfirm", folderPath),
                    Loc.T("I18n_DeleteConfirmTitle"), MessageBoxButton.YesNo, MessageBoxImage.Warning);

                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        if (Directory.Exists(folderPath))
                            Directory.Delete(folderPath, true);
                        LoadSkillsList();
                        Globals.Ribbons.Ribbon1.ShowNotification(Loc.T("I18n_SkillDeleted"), 2000);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(Loc.F("I18n_DeleteFailed", ex.Message), Loc.T("I18n_ErrorTitle"), MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }

        // ========== MCP 服务器 ==========

        private void LoadMcpServers()
        {
            _mcpServers.Clear();
            bool migratedLegacySecrets = false;
            try
            {
                if (File.Exists(_mcpConfigPath))
                {
                    string json = File.ReadAllText(_mcpConfigPath);
                    foreach (var item in EnumerateMcpServerItems(JToken.Parse(json)))
                    {
                        migratedLegacySecrets |= item["envProtected"] == null || item["headersProtected"] == null || item["enabledTools"] == null;
                        _mcpServers.Add(CreateMcpServerConfig(item));
                    }
                }
            }
            catch { }
            if (migratedLegacySecrets) SaveMcpServers();
            McpServerDataGrid.ItemsSource = _mcpServers;
        }

        private void SaveMcpServers()
        {
            try
            {
                string dir = Path.GetDirectoryName(_mcpConfigPath);
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

                var arr = new JArray();
                foreach (var s in _mcpServers)
                {
                    arr.Add(new JObject
                    {
                        ["name"] = s.Name,
                        ["transport"] = string.IsNullOrWhiteSpace(s.Transport) ? "stdio" : s.Transport,
                        ["url"] = s.Url ?? "",
                        ["command"] = s.Command,
                        ["args"] = s.Args,
                        ["cwd"] = s.WorkingDirectory ?? "",
                        ["envProtected"] = SecretProtector.Protect(s.EnvironmentVariables ?? ""),
                        ["headersProtected"] = SecretProtector.Protect(s.Headers ?? ""),
                        ["enabled"] = s.Enabled
                        , ["trustLevel"] = s.TrustLevel ?? "Untrusted"
                        , ["availableTools"] = new JArray(s.AvailableToolNames ?? new List<string>())
                        , ["enabledTools"] = new JArray(s.EnabledToolNames ?? new List<string>())
                        , ["connectionState"] = s.ConnectionState ?? "Not tested"
                        , ["endpointDisplay"] = s.EndpointDisplay ?? ""
                        , ["lastHandshakeAt"] = s.LastHandshakeAt
                        , ["lastTestedAt"] = s.LastTestedAt
                        , ["lastToolCount"] = s.LastToolCount
                        , ["lastError"] = s.LastError
                    });
                }
                File.WriteAllText(_mcpConfigPath, arr.ToString(Newtonsoft.Json.Formatting.Indented));
            }
            catch { }
        }

        private static string ReadProtectedMcpValue(JToken item, string protectedProperty, JToken legacyValue)
        {
            string protectedValue = item[protectedProperty]?.ToString();
            if (!string.IsNullOrWhiteSpace(protectedValue))
                return SecretProtector.UnprotectOrLegacy(protectedValue);

            return McpToolManager.EnvironmentToText(legacyValue);
        }

        private static List<string> ReadMcpToolNames(JToken token)
        {
            return token is JArray values
                ? values.Values<string>().Where(name => !string.IsNullOrWhiteSpace(name)).Distinct(StringComparer.OrdinalIgnoreCase).ToList()
                : new List<string>();
        }

        private static McpServerConfig CreateMcpServerConfig(JObject item, bool forceUntrusted = false)
        {
            return new McpServerConfig
            {
                Name = item["name"]?.ToString()?.Trim() ?? "",
                Transport = McpToolManager.TransportToText(McpToolManager.ParseTransportKind(
                    item["transport"]?.ToString()
                    ?? item["type"]?.ToString()
                    ?? "stdio")),
                Url = item["url"]?.ToString()
                    ?? item["endpoint"]?.ToString()
                    ?? "",
                Command = item["command"]?.ToString() ?? "",
                Args = ReadMcpArgs(item["args"]),
                WorkingDirectory = item["cwd"]?.ToString()
                    ?? item["workingDirectory"]?.ToString()
                    ?? "",
                EnvironmentVariables = ReadProtectedMcpValue(item, "envProtected", item["env"] ?? item["environment"]),
                Headers = ReadProtectedMcpValue(item, "headersProtected", item["headers"]),
                Enabled = item["enabled"]?.ToObject<bool>() ?? true,
                TrustLevel = forceUntrusted ? "Untrusted" : item["trustLevel"]?.ToString() ?? "Untrusted",
                AvailableToolNames = ReadMcpToolNames(item["availableTools"]),
                EnabledToolNames = ReadMcpToolNames(item["enabledTools"]),
                ConnectionState = item["connectionState"]?.ToString() ?? "Not tested",
                EndpointDisplay = item["endpointDisplay"]?.ToString(),
                LastHandshakeAt = item["lastHandshakeAt"]?.ToObject<DateTime?>(),
                LastTestedAt = item["lastTestedAt"]?.ToObject<DateTime?>(),
                LastToolCount = item["lastToolCount"]?.ToObject<int>() ?? 0,
                LastError = item["lastError"]?.ToString()
            };
        }

        private string _editingMcpName = null; // null = 添加模式，非null = 编辑模式
        private McpServerConfig _toolPermissionsServer;
        private bool _updatingMcpToolSelection;

        private void ConfigureMcpTools_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is string name)) return;
            var server = _mcpServers.FirstOrDefault(item => item.Name == name);
            if (server == null) return;
            if (server.AvailableToolNames == null || server.AvailableToolNames.Count == 0)
            {
                MessageBox.Show(Loc.T("I18n_McpDiscoverToolsFirst"), Loc.T("I18n_McpTools"), MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            _toolPermissionsServer = server;
            McpToolPermissionsPanel.Children.Clear();
            var enabled = new HashSet<string>(server.EnabledToolNames ?? new List<string>(), StringComparer.OrdinalIgnoreCase);
            foreach (string toolName in server.AvailableToolNames.OrderBy(tool => tool, StringComparer.OrdinalIgnoreCase))
            {
                var toolCheckBox = new CheckBox
                {
                    Content = toolName,
                    Tag = toolName,
                    IsChecked = enabled.Contains(toolName),
                    Margin = new Thickness(0, 0, 0, 8)
                };
                toolCheckBox.Checked += McpToolCheckBox_Changed;
                toolCheckBox.Unchecked += McpToolCheckBox_Changed;
                McpToolPermissionsPanel.Children.Add(toolCheckBox);
            }
            UpdateMcpSelectAllState();
            McpToolPermissionsPopup.IsOpen = true;
        }

        private void McpSelectAllToolsCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            if (_updatingMcpToolSelection) return;
            SetAllMcpToolSelections(true);
        }

        private void McpSelectAllToolsCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            if (_updatingMcpToolSelection) return;
            SetAllMcpToolSelections(false);
        }

        private void McpToolCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            if (!_updatingMcpToolSelection)
                UpdateMcpSelectAllState();
        }

        private void SetAllMcpToolSelections(bool selected)
        {
            _updatingMcpToolSelection = true;
            try
            {
                foreach (var checkBox in McpToolPermissionsPanel.Children.OfType<CheckBox>())
                    checkBox.IsChecked = selected;
                McpSelectAllToolsCheckBox.IsChecked = selected;
            }
            finally
            {
                _updatingMcpToolSelection = false;
            }
        }

        private void UpdateMcpSelectAllState()
        {
            var checkBoxes = McpToolPermissionsPanel.Children.OfType<CheckBox>().ToList();
            int selectedCount = checkBoxes.Count(checkBox => checkBox.IsChecked == true);

            _updatingMcpToolSelection = true;
            try
            {
                McpSelectAllToolsCheckBox.IsChecked = selectedCount == 0
                    ? false
                    : selectedCount == checkBoxes.Count
                        ? true
                        : (bool?)null;
            }
            finally
            {
                _updatingMcpToolSelection = false;
            }
        }

        private void SaveMcpToolPermissions_Click(object sender, RoutedEventArgs e)
        {
            if (_toolPermissionsServer != null)
            {
                _toolPermissionsServer.EnabledToolNames = McpToolPermissionsPanel.Children
                    .OfType<CheckBox>()
                    .Where(checkBox => checkBox.IsChecked == true)
                    .Select(checkBox => checkBox.Tag?.ToString())
                    .Where(name => !string.IsNullOrWhiteSpace(name))
                    .ToList();
                SaveMcpServers();
            }
            _toolPermissionsServer = null;
            McpToolPermissionsPopup.IsOpen = false;
        }

        private void CancelMcpToolPermissions_Click(object sender, RoutedEventArgs e)
        {
            _toolPermissionsServer = null;
            McpToolPermissionsPopup.IsOpen = false;
        }

        private void AddMcpServer_Click(object sender, RoutedEventArgs e)
        {
            _editingMcpName = null;
            McpPopupTitle.Text = Loc.T("I18n_AddMcpServer");
            McpNameBox.Text = "";
            SetMcpTransportSelection("stdio");
            McpUrlBox.Text = "";
            McpCommandBox.Text = "";
            McpArgsBox.Text = "";
            McpWorkingDirectoryBox.Text = "";
            McpEnvBox.Text = "";
            McpHeadersBox.Text = "";
            McpTrustedCheckBox.IsChecked = false;
            McpNameBox.IsEnabled = true;
            McpServerPopup.IsOpen = true;
        }

        private void ImportMcpJson_Click(object sender, RoutedEventArgs e)
        {
            McpImportJsonBox.Text = "";
            McpImportPopup.IsOpen = true;
            McpImportJsonBox.Focus();
        }

        private void CancelMcpImport_Click(object sender, RoutedEventArgs e)
        {
            McpImportPopup.IsOpen = false;
        }

        private void ConfirmMcpImport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                JToken root = JToken.Parse(McpImportJsonBox.Text ?? "");
                var imported = EnumerateMcpServerItems(root)
                    .Select(item => CreateMcpServerConfig(item, forceUntrusted: true))
                    .Where(IsValidImportedMcpServer)
                    .GroupBy(item => item.Name, StringComparer.OrdinalIgnoreCase)
                    .Select(group => group.Last())
                    .ToList();

                if (imported.Count == 0)
                {
                    MessageBox.Show(Loc.T("I18n_McpImportInvalid"), Loc.T("I18n_ImportMcpJson"),
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                int updateCount = imported.Count(item =>
                    _mcpServers.Any(existing => string.Equals(existing.Name, item.Name, StringComparison.OrdinalIgnoreCase)));
                int addCount = imported.Count - updateCount;
                if (MessageBox.Show(
                        Loc.F("I18n_McpImportPreview", addCount, updateCount),
                        Loc.T("I18n_ImportMcpJson"),
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question) != MessageBoxResult.Yes)
                    return;

                foreach (var server in imported)
                {
                    var existing = _mcpServers.FirstOrDefault(item =>
                        string.Equals(item.Name, server.Name, StringComparison.OrdinalIgnoreCase));
                    if (existing == null)
                    {
                        _mcpServers.Add(server);
                        continue;
                    }

                    UpdateMcpServerFromImport(existing, server);
                }

                SaveMcpServers();
                McpServerDataGrid.Items.Refresh();
                McpImportPopup.IsOpen = false;
                MessageBox.Show(
                    Loc.F("I18n_McpImportComplete", imported.Count, addCount, updateCount),
                    Loc.T("I18n_ImportMcpJson"),
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    Loc.F("I18n_McpImportFailed", ex.Message),
                    Loc.T("I18n_ImportMcpJson"),
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private static bool IsValidImportedMcpServer(McpServerConfig server)
        {
            if (server == null || string.IsNullOrWhiteSpace(server.Name)) return false;
            if (McpToolManager.ParseTransportKind(server.Transport) == McpTransportKind.Stdio)
                return !string.IsNullOrWhiteSpace(server.Command);

            return Uri.TryCreate(server.Url, UriKind.Absolute, out var uri)
                && (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps);
        }

        private static void UpdateMcpServerFromImport(McpServerConfig target, McpServerConfig source)
        {
            target.Transport = source.Transport;
            target.Url = source.Url;
            target.Command = source.Command;
            target.Args = source.Args;
            target.WorkingDirectory = source.WorkingDirectory;
            target.EnvironmentVariables = source.EnvironmentVariables;
            target.Headers = source.Headers;
            target.Enabled = source.Enabled;
            target.TrustLevel = "Untrusted";
            target.AvailableToolNames = new List<string>();
            target.EnabledToolNames = new List<string>();
            target.ConnectionState = "Not tested";
            target.EndpointDisplay = null;
            target.LastHandshakeAt = null;
            target.LastTestedAt = null;
            target.LastToolCount = 0;
            target.LastError = null;
        }

        private void EditMcpServer_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string name)
            {
                var server = _mcpServers.FirstOrDefault(s => s.Name == name);
                if (server == null) return;

                _editingMcpName = name;
                McpPopupTitle.Text = Loc.T("I18n_EditMcpServer");
                McpNameBox.Text = server.Name;
                SetMcpTransportSelection(server.Transport);
                McpUrlBox.Text = server.Url;
                McpCommandBox.Text = server.Command;
                McpArgsBox.Text = server.Args;
                McpWorkingDirectoryBox.Text = server.WorkingDirectory;
                McpEnvBox.Text = server.EnvironmentVariables;
                McpHeadersBox.Text = server.Headers;
                McpTrustedCheckBox.IsChecked = string.Equals(server.TrustLevel, "Trusted", StringComparison.OrdinalIgnoreCase);
                McpNameBox.IsEnabled = false; // 编辑时不允许改名
                McpServerPopup.IsOpen = true;
            }
        }

        private void DeleteMcpServer_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string name)
            {
                var result = MessageBox.Show(Loc.F("I18n_DeleteMcpConfirm", name),
                    Loc.T("I18n_DeleteConfirmTitle"), MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (result == MessageBoxResult.Yes)
                {
                    var item = _mcpServers.FirstOrDefault(s => s.Name == name);
                    if (item != null) _mcpServers.Remove(item);
                    SaveMcpServers();
                }
            }
        }

        private async void TestMcpServer_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button btn) || !(btn.Tag is string name)) return;
            var server = _mcpServers.FirstOrDefault(s => s.Name == name);
            if (server == null) return;

            btn.IsEnabled = false;
            var manager = new McpToolManager();
            try
            {
                var kind = McpToolManager.ParseTransportKind(server.Transport);
                string endpointOrCommand = kind == McpTransportKind.Stdio ? server.Command : server.Url;
                await manager.AddServerAsync(
                    server.Name,
                    kind,
                    endpointOrCommand,
                    server.Args,
                    server.WorkingDirectory,
                    McpToolManager.ParseEnvironment(server.EnvironmentVariables == null ? null : JToken.FromObject(server.EnvironmentVariables)),
                    McpToolManager.ParseEnvironment(server.Headers == null ? null : JToken.FromObject(server.Headers)));
                if (manager.ServerCount > 0 && string.IsNullOrWhiteSpace(manager.LastError))
                {
                    server.AvailableToolNames = manager.GetDiscoveredTools()
                        .Where(tool => string.Equals(tool.ServerName, server.Name, StringComparison.OrdinalIgnoreCase))
                        .Select(tool => tool.Name)
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();
                    server.EnabledToolNames = (server.EnabledToolNames ?? new List<string>())
                        .Where(name => server.AvailableToolNames.Contains(name, StringComparer.OrdinalIgnoreCase))
                        .ToList();
                    server.ConnectionState = "Connected";
                    server.EndpointDisplay = RedactMcpEndpoint(server);
                    server.LastHandshakeAt = DateTime.UtcNow;
                    server.LastTestedAt = DateTime.UtcNow;
                    server.LastToolCount = server.AvailableToolNames.Count;
                    server.LastError = manager.LastError == null ? null : $"[{manager.LastErrorCategory}] {manager.LastError}";
                    SaveMcpServers();
                    MessageBox.Show(
                        Loc.F("I18n_McpTestSuccess", manager.ToolCount),
                        Loc.T("I18n_McpTestTitle"),
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
                else
                {
                    server.ConnectionState = "Failed";
                    server.EndpointDisplay = RedactMcpEndpoint(server);
                    server.LastHandshakeAt = null;
                    server.LastTestedAt = DateTime.UtcNow;
                    server.LastToolCount = 0;
                    server.LastError = manager.LastError == null
                        ? Loc.T("I18n_McpTestFailed")
                        : $"[{manager.LastErrorCategory}] {manager.LastError}";
                    SaveMcpServers();
                    MessageBox.Show(
                        Loc.F("I18n_McpTestError", server.LastError ?? Loc.T("I18n_McpTestFailed")),
                        Loc.T("I18n_McpTestTitle"),
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                server.ConnectionState = "Failed";
                server.EndpointDisplay = RedactMcpEndpoint(server);
                server.LastHandshakeAt = null;
                server.LastTestedAt = DateTime.UtcNow;
                server.LastToolCount = 0;
                server.LastError = ex.Message;
                SaveMcpServers();
                MessageBox.Show(
                    Loc.F("I18n_McpTestError", ex.Message),
                    Loc.T("I18n_McpTestTitle"),
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                manager.Dispose();
                btn.IsEnabled = true;
            }
        }

        private static string RedactMcpEndpoint(McpServerConfig server)
        {
            string value = McpToolManager.ParseTransportKind(server.Transport) == McpTransportKind.Stdio
                ? $"{server.Command} {server.Args}".Trim()
                : server.Url;
            if (string.IsNullOrWhiteSpace(value)) return "";

            if (Uri.TryCreate(value, UriKind.Absolute, out var uri))
                return $"{uri.Scheme}://{uri.Host}{uri.AbsolutePath}";

            return System.Text.RegularExpressions.Regex.Replace(
                value,
                @"(?i)(api[_-]?key|token|secret|password)\s*=\s*[^\s]+",
                "$1=[REDACTED]");
        }

        private void McpConfirm_Click(object sender, RoutedEventArgs e)
        {
            string name = McpNameBox.Text?.Trim();
            string transport = GetSelectedMcpTransport();
            string url = McpUrlBox.Text?.Trim();
            string command = McpCommandBox.Text?.Trim();
            string args = McpArgsBox.Text?.Trim();
            string workingDirectory = McpWorkingDirectoryBox.Text?.Trim();
            string env = McpEnvBox.Text?.Trim();
            string headers = McpHeadersBox.Text?.Trim();
            string trustLevel = McpTrustedCheckBox.IsChecked == true ? "Trusted" : "Untrusted";

            if (string.IsNullOrEmpty(name))
            {
                MessageBox.Show(Loc.T("I18n_NameRequired"), WordOllama.Properties.Resources.提示, MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            bool isStdio = McpToolManager.ParseTransportKind(transport) == McpTransportKind.Stdio;
            if (isStdio && string.IsNullOrEmpty(command))
            {
                MessageBox.Show(Loc.T("I18n_StdioCommandRequired"), WordOllama.Properties.Resources.提示, MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (!isStdio && string.IsNullOrEmpty(url))
            {
                MessageBox.Show(Loc.T("I18n_RemoteMcpUrlRequired"), WordOllama.Properties.Resources.提示, MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_editingMcpName == null)
            {
                // 添加模式 — 检查重名
                if (_mcpServers.Any(s => s.Name == name))
                {
                    MessageBox.Show(Loc.F("I18n_McpNameExists", name), WordOllama.Properties.Resources.提示, MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                _mcpServers.Add(new McpServerConfig
                {
                    Name = name,
                    Transport = transport,
                    Url = url,
                    Command = command,
                    Args = args,
                    WorkingDirectory = workingDirectory,
                    EnvironmentVariables = env,
                    Headers = headers,
                    TrustLevel = trustLevel
                });
            }
            else
            {
                // 编辑模式
                var server = _mcpServers.FirstOrDefault(s => s.Name == _editingMcpName);
                if (server != null)
                {
                    server.Transport = transport;
                    server.Url = url;
                    server.Command = command;
                    server.Args = args;
                    server.WorkingDirectory = workingDirectory;
                    server.EnvironmentVariables = env;
                    server.Headers = headers;
                    server.TrustLevel = trustLevel;
                }
            }

            SaveMcpServers();
            McpServerDataGrid.Items.Refresh();
            McpServerPopup.IsOpen = false;
        }

        private void McpCancel_Click(object sender, RoutedEventArgs e)
        {
            McpServerPopup.IsOpen = false;
        }

        private void ToggleMcpEnabled_Click(object sender, RoutedEventArgs e)
        {
            SaveMcpServers();
        }

        private static IEnumerable<JObject> EnumerateMcpServerItems(JToken root)
        {
            if (root is JArray arr)
            {
                foreach (var item in arr.OfType<JObject>())
                    yield return item;
                yield break;
            }

            var servers = root?["mcpServers"] as JObject;
            if (servers == null) yield break;

            foreach (var prop in servers.Properties())
            {
                if (prop.Value is JObject server)
                {
                    var copy = (JObject)server.DeepClone();
                    copy["name"] = copy["name"]?.ToString() ?? prop.Name;
                    yield return copy;
                }
            }
        }

        private static string ReadMcpArgs(JToken token)
        {
            if (token is JArray arr)
            {
                var parts = new List<string>();
                foreach (var item in arr)
                    parts.Add(QuoteMcpArg(item?.ToString() ?? ""));
                return string.Join(" ", parts);
            }
            return token?.ToString() ?? "";
        }

        private static string QuoteMcpArg(string value)
        {
            if (string.IsNullOrEmpty(value)) return "\"\"";
            if (value.IndexOfAny(new[] { ' ', '\t', '"' }) < 0) return value;
            return "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
        }

        private static JObject ParseMcpEnvText(string text)
        {
            var obj = new JObject();
            foreach (var item in McpToolManager.ParseEnvironment(text == null ? null : JToken.FromObject(text)))
                obj[item.Key] = item.Value;
            return obj;
        }

        private string GetSelectedMcpTransport()
        {
            if (McpTransportBox.SelectedItem is ComboBoxItem item)
                return item.Content?.ToString() ?? "stdio";
            return McpTransportBox.Text?.Trim();
        }

        private void SetMcpTransportSelection(string transport)
        {
            string value = string.IsNullOrWhiteSpace(transport) ? "stdio" : transport;
            foreach (var item in McpTransportBox.Items.OfType<ComboBoxItem>())
            {
                if (string.Equals(item.Content?.ToString(), value, StringComparison.OrdinalIgnoreCase))
                {
                    McpTransportBox.SelectedItem = item;
                    return;
                }
            }
            McpTransportBox.Text = value;
        }

        #endregion

        #region 静默审查模型设置

        private void LinterOllamaModelSelector_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var model = LinterOllamaModelSelector.SelectedItem;
            if (model != null)
            {
                var prop = model.GetType().GetProperty("Name");
                if (prop != null) LinterModelBox.Text = prop.GetValue(model, null)?.ToString() ?? "";
            }
        }

        private void LinterOnlineModelSelector_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var model = LinterOnlineModelSelector.SelectedItem;
            if (model != null)
            {
                var prop = model.GetType().GetProperty("ModelName");
                if (prop != null) LinterModelBox.Text = prop.GetValue(model, null)?.ToString() ?? "";
            }
        }

        private void LinterEnabledToggle_Click(object sender, RoutedEventArgs e)
        {
            if (sender is ToggleButton tb)
            {
                SharedSettings.LinterEnabled = tb.IsChecked == true;
                // 即时启停 SilentLinter 定时器
                if (tb.IsChecked == true)
                    Globals.ThisAddIn.StartSilentLinter();
                else
                    Globals.ThisAddIn.StopSilentLinter();
            }
        }

        private void SuppressAgentNotifToggle_Click(object sender, RoutedEventArgs e)
        {
            if (sender is ToggleButton tb)
                SharedSettings.SuppressAgentNotifications = tb.IsChecked == true;
        }

        private void SuppressPlanConfirmationToggle_Click(object sender, RoutedEventArgs e)
        {
            if (sender is ToggleButton tb)
                SharedSettings.SuppressPlanConfirmation = tb.IsChecked == true;
        }

        private void SuppressDiffReviewToggle_Click(object sender, RoutedEventArgs e)
        {
            if (sender is ToggleButton tb)
                SharedSettings.SuppressDiffReviewPrompt = tb.IsChecked == true;
        }

        private void DefaultOutputModeSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isLoadingSettings) return;

            if (DefaultOutputModeSelector?.SelectedItem is ComboBoxItem item)
                SharedSettings.DefaultPromptOutputMode = item.Tag?.ToString() ?? "Auto";
        }

        private void SelectDefaultOutputMode(string mode)
        {
            string targetMode = string.IsNullOrWhiteSpace(mode) ? "Auto" : mode.Trim();
            foreach (ComboBoxItem item in DefaultOutputModeSelector.Items)
            {
                if (string.Equals(item.Tag?.ToString(), targetMode, StringComparison.OrdinalIgnoreCase))
                {
                    DefaultOutputModeSelector.SelectedItem = item;
                    return;
                }
            }

            DefaultOutputModeSelector.SelectedIndex = 0;
        }

        private void SaveLinterSettings_Click(object sender, RoutedEventArgs e)
        {
            AppConfig.Instance.LinterModel = LinterModelBox.Text?.Trim() ?? "";
            AppConfig.Instance.LinterEndPoint = LinterEndPointBox.Text?.Trim() ?? "";

            if (int.TryParse(LinterDebounceBox.Text?.Trim(), out int debounce) && debounce > 0)
            {
                AppConfig.Instance.LinterDebounceSeconds = debounce;
            }

            AppConfig.Save();
            // 使间隔更改即时生效
            Globals.ThisAddIn.ApplyLinterInterval();
            MessageBox.Show(Loc.T("I18n_SaveReviewSettingsSaved"), Loc.T("I18n_WordOllamaTitle"), MessageBoxButton.OK, MessageBoxImage.Information);
        }

        #endregion
    }

    /// <summary>
    /// Skill 列表显示模型。
    /// </summary>
    public class SkillDisplayItem
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string FolderPath { get; set; }
    }

}
