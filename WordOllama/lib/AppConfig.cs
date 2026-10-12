using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace WordOllama.lib
{
    /// <summary>
    /// 统一配置中心 — 所有应用设置由 JSON 文件持久化。
    /// 替代 Properties.Settings + SettingsManager(.ini) 的混乱架构。
    /// 文件路径: %LocalAppData%\WordOllama\appconfig.json
    /// </summary>
    public class AppConfig
    {
        #region AI 模式

        public string AIMode { get; set; } = "Ollama";
        public string LanguageMode { get; set; } = "Auto";
        public string ThemeType { get; set; } = "Light";
        public string DefaultPromptOutputMode { get; set; } = "Auto";
        public bool SetupAssistantCompleted { get; set; } = false;

        #endregion

        #region Ollama

        public string OllamaEndPoint { get; set; } = "http://localhost:11434";
        public string OllamaModel { get; set; } = "";
        public int OllamaMaxToken { get; set; } = 4096;
        public int OllamaTimeoutSeconds { get; set; } = 600;
        public double OllamaTemperature { get; set; } = 0.5;
        public bool IsOllamaModelRefresh { get; set; } = false;
        public bool OllamaNetMode { get; set; } = false;

        #endregion

        #region ChatGPT / 在线模型

        public string OnlineProviderType { get; set; } = "OpenAI"; // "OpenAI", "Claude", "Gemini"

        public string ChatGPTEndPoint { get; set; } = "";
        [JsonIgnore]
        public string ChatGPTKey { get; set; } = "";
        [JsonProperty("ChatGPTKey")]
        private string StoredChatGPTKey
        {
            get { return SecretProtector.Protect(ChatGPTKey); }
            set { ChatGPTKey = SecretProtector.UnprotectOrLegacy(value); }
        }
        public string ChatGPTModel { get; set; } = "";
        public int ChatGPTMaxToken { get; set; } = 1000;
        public string OutputTokenParameter { get; set; } = "auto";
        public int OnlineContextWindow { get; set; } = 256000;
        public double ChatGPTTemperature { get; set; } = 0.5;
        public string ProviderToolCallingMode { get; set; } = "Auto";
        public bool ProviderSupportsStreaming { get; set; } = true;
        public bool ProviderSupportsVision { get; set; } = true;
        public bool ProviderSupportsJsonOutput { get; set; } = true;
        public int ProviderTimeoutSeconds { get; set; } = 120;
        public int ProviderConnectTimeoutSeconds { get; set; } = 15;
        public int ProviderMaxRetries { get; set; } = 1;

        public string ClaudeEndPoint { get; set; } = "https://api.anthropic.com/v1";
        [JsonIgnore]
        public string ClaudeKey { get; set; } = "";
        [JsonProperty("ClaudeKey")]
        private string StoredClaudeKey
        {
            get { return SecretProtector.Protect(ClaudeKey); }
            set { ClaudeKey = SecretProtector.UnprotectOrLegacy(value); }
        }
        public string ClaudeModel { get; set; } = "claude-sonnet-4-20250514";

        public string GeminiEndPoint { get; set; } = "https://generativelanguage.googleapis.com/v1beta";
        [JsonIgnore]
        public string GeminiKey { get; set; } = "";
        [JsonProperty("GeminiKey")]
        private string StoredGeminiKey
        {
            get { return SecretProtector.Protect(GeminiKey); }
            set { GeminiKey = SecretProtector.UnprotectOrLegacy(value); }
        }
        public string GeminiModel { get; set; } = "gemini-3-flash-preview";

        #endregion

        #region 思维链 (Thinking)

        public bool ThinkingEnabled { get; set; } = true;
        public string ReasoningEffort { get; set; } = "high";

        #endregion

        #region Markdown 转换样式映射

        /// <summary>
        /// JSON 序列化的 Markdown 样式映射列表。
        /// 格式: [{"MarkdownElement":"h1","WordStyleName":"Heading 1"},...]
        /// </summary>
        public string MarkdownStyleMappingsJson { get; set; } = "";

        #endregion

        #region 图片模型

        public string IMGModel { get; set; } = "";
        public string IMGEndPoint { get; set; } = "";
        [JsonIgnore]
        public string IMGAPIKey { get; set; } = "";
        [JsonProperty("IMGAPIKey")]
        private string StoredIMGAPIKey
        {
            get { return SecretProtector.Protect(IMGAPIKey); }
            set { IMGAPIKey = SecretProtector.UnprotectOrLegacy(value); }
        }
        public string IMGMode { get; set; } = "Global";

        #endregion

        #region HTML 模型

        public string HTMLModel { get; set; } = "";
        public string HTMLEndPoint { get; set; } = "";
        [JsonIgnore]
        public string HTMLAPIKey { get; set; } = "";
        [JsonProperty("HTMLAPIKey")]
        private string StoredHTMLAPIKey
        {
            get { return SecretProtector.Protect(HTMLAPIKey); }
            set { HTMLAPIKey = SecretProtector.UnprotectOrLegacy(value); }
        }
        public string HTMLMode { get; set; } = "Global";

        #endregion

        #region Agent

        public int AgentMaxIterations { get; set; } = 20;
        public int AgentTaskTimeoutMinutes { get; set; } = 30;
        public int AgentMaxResultSize { get; set; } = 0;
        /// <summary>Agent 侧边栏宽度，范围由宿主侧统一限制。</summary>
        public int AgentPaneWidth { get; set; } = 900;
        /// <summary>用户是否主动调整过 Agent 侧边栏宽度。</summary>
        public bool AgentPaneWidthUserSet { get; set; } = false;
        /// <summary>审阅工作台侧边栏宽度，范围由宿主侧统一限制。</summary>
        public int ReviewPaneWidth { get; set; } = 900;
        /// <summary>用户是否主动调整过审阅工作台宽度。</summary>
        public bool ReviewPaneWidthUserSet { get; set; } = false;
        /// <summary>ViewOnly / ProposeChanges / TrackedChanges.</summary>
        public string AgentExecutionMode { get; set; } = "TrackedChanges";
        /// <summary>Whether high-risk external Agent tools are included without a task-level prompt.</summary>
        public bool AgentExternalToolsEnabled { get; set; } = false;

        #endregion

        #region Agent 确认设置

        /// <summary>永久免计划确认。</summary>
        public bool SuppressPlanConfirmation { get; set; } = false;
        /// <summary>永久免变更审查弹窗。</summary>
        public bool SuppressDiffReviewPrompt { get; set; } = false;
        /// <summary>永久免错误恢复询问。</summary>
        public bool SuppressErrorRecoveryPrompt { get; set; } = false;
        /// <summary>免 LLM 任务通知（开始/完成的气泡提示）。</summary>
        public bool SuppressAgentNotifications { get; set; } = false;

        #endregion

        #region 长文档审校（每批上限，不限制整篇文档）

        public int ReviewIssueBatchMaxParagraphs { get; set; } = 80;
        public string ReviewReviewerName { get; set; } = "AI审批";
        public string ReviewReviewerIdentityMode { get; set; } = "Date";
        public int ReviewIssueBatchMaxCharacters { get; set; } = 12000;
        public int ReviewSelectionBatchMaxCharacters { get; set; } = 6000;
        public int ReviewSuggestionBatchMaxItems { get; set; } = 8;
        public int ReviewSuggestionBatchMaxCharacters { get; set; } = 6000;
        public int ReviewMaxConcurrentRequests { get; set; } = 3;
        public int ReviewOllamaOutputTokens { get; set; } = 2048;

        #endregion

        #region 静默审查 (SilentLinter)

        public bool LinterEnabled { get; set; } = false;
        public string LinterModel { get; set; } = "";
        public string LinterEndPoint { get; set; } = "";
        public int LinterDebounceSeconds { get; set; } = 30;
        public int LinterContextTokens { get; set; } = 4096;
        public int LinterTimeoutSeconds { get; set; } = 300;

        #endregion

        #region 快捷指令

        public string C1_label { get; set; } = "C11";
        public string C2_label { get; set; } = "C22";
        public string C3_label { get; set; } = "C33";
        public string C4_label { get; set; } = "C44";
        public string C1_Prompt { get; set; } = "";
        public string C2_Prompt { get; set; } = "";
        public string C3_Prompt { get; set; } = "";
        public string C4_Prompt { get; set; } = "";
        public string C1_Funt { get; set; } = "";
        public string C2_Funt { get; set; } = "";
        public string C3_Funt { get; set; } = "";
        public string C4_Funt { get; set; } = "";

        #endregion

        #region 动态属性访问（兼容 CustomPromptUI 等索引器用法）

        /// <summary>
        /// 动态获取设置值（兼容旧版 Properties.Settings.Default[key] 索引器模式）。
        /// </summary>
        public object GetSetting(string key)
        {
            switch (key)
            {
                case "C1_label": return C1_label;
                case "C2_label": return C2_label;
                case "C3_label": return C3_label;
                case "C4_label": return C4_label;
                case "C1_Prompt": return C1_Prompt;
                case "C2_Prompt": return C2_Prompt;
                case "C3_Prompt": return C3_Prompt;
                case "C4_Prompt": return C4_Prompt;
                case "C1_Funt": return C1_Funt;
                case "C2_Funt": return C2_Funt;
                case "C3_Funt": return C3_Funt;
                case "C4_Funt": return C4_Funt;
                default: return null;
            }
        }

        /// <summary>
        /// 动态设置设置值（兼容旧版 Properties.Settings.Default[key] = value 模式）。
        /// </summary>
        public void SetSetting(string key, object value)
        {
            string val = value?.ToString() ?? "";
            switch (key)
            {
                case "C1_label": C1_label = val; break;
                case "C2_label": C2_label = val; break;
                case "C3_label": C3_label = val; break;
                case "C4_label": C4_label = val; break;
                case "C1_Prompt": C1_Prompt = val; break;
                case "C2_Prompt": C2_Prompt = val; break;
                case "C3_Prompt": C3_Prompt = val; break;
                case "C4_Prompt": C4_Prompt = val; break;
                case "C1_Funt": C1_Funt = val; break;
                case "C2_Funt": C2_Funt = val; break;
                case "C3_Funt": C3_Funt = val; break;
                case "C4_Funt": C4_Funt = val; break;
            }
        }

        #endregion

        #region 单例 + 持久化

        [JsonIgnore]
        private static string ConfigDir => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "WordOllama");

        [JsonIgnore]
        private static string ConfigPath => Path.Combine(ConfigDir, "appconfig.json");

        [JsonIgnore]
        public static AppConfig Instance { get; private set; }

        private static int _saveSuspensionCount;

        /// <summary>
        /// 暂停自动落盘，用于设置窗口编辑草稿。内存中的配置仍会更新，关闭窗口时由调用方提交或恢复。
        /// </summary>
        public static IDisposable BeginDraft()
        {
            Interlocked.Increment(ref _saveSuspensionCount);
            return new DraftScope();
        }

        /// <summary>
        /// 加载配置。优先从 appconfig.json 读取；若不存在，从旧版 Properties.Settings 迁移。
        /// 在 ThisAddIn_Startup 中调用一次。
        /// </summary>
        public static void Load()
        {
            if (File.Exists(ConfigPath))
            {
                try
                {
                    string json = File.ReadAllText(ConfigPath);
                    Instance = Deserialize(json);
                    if (HasLegacySecrets(json)) Save();
                }
                catch
                {
                    Instance = new AppConfig();
                }
            }
            else
            {
                // 首次运行：从旧版 Properties.Settings 迁移
                Instance = new AppConfig();
                MigrateFromLegacy();
            }
        }

        /// <summary>Deserializes persisted settings while retaining defaults for fields added in newer versions.</summary>
        public static AppConfig Deserialize(string json)
        {
            return JsonConvert.DeserializeObject<AppConfig>(json ?? "") ?? new AppConfig();
        }

        /// <summary>恢复设置草稿，不写入磁盘。</summary>
        public static void Restore(string json)
        {
            Instance = Deserialize(json);
        }

        /// <summary>
        /// 保存配置到 appconfig.json。
        /// </summary>
        public static void Save()
        {
            if (Volatile.Read(ref _saveSuspensionCount) > 0)
                return;

            SaveNow();
        }

        /// <summary>无视草稿状态立即保存当前内存配置。</summary>
        public static void SaveNow()
        {
            try
            {
                Directory.CreateDirectory(ConfigDir);
                string json = JsonConvert.SerializeObject(Instance, Formatting.Indented);
                File.WriteAllText(ConfigPath, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AppConfig.Save error: {ex.Message}");
            }
        }

        private sealed class DraftScope : IDisposable
        {
            private bool _disposed;

            public void Dispose()
            {
                if (_disposed) return;
                _disposed = true;
                Interlocked.Decrement(ref _saveSuspensionCount);
            }
        }

        /// <summary>
        /// 从旧版 Properties.Settings 迁移（一次性）。
        /// </summary>
        private static void MigrateFromLegacy()
        {
            try
            {
                SettingsManager.LoadLegacyIniIntoProperties();
                var s = Properties.Settings.Default;
                Instance.AIMode = s.AIMode ?? "Ollama";
                Instance.LanguageMode = string.IsNullOrWhiteSpace(s.LanguageMode) ||
                                        string.Equals(s.LanguageMode, "null", StringComparison.OrdinalIgnoreCase)
                    ? "Auto"
                    : s.LanguageMode;
                Instance.ThemeType = s.themeType ?? "Light";
                Instance.DefaultPromptOutputMode = "Auto";
                Instance.OllamaEndPoint = s.OllamaEndPoint ?? "http://localhost:11434";
                Instance.OllamaModel = s.OllamaModel ?? "";
                Instance.OllamaMaxToken = s.OllamaMaxToken > 0 ? s.OllamaMaxToken : 4096;
                Instance.OllamaTemperature = s.OllamaTemperature;
                Instance.IsOllamaModelRefresh = s.isOllamaModelRefresh;
                Instance.OllamaNetMode = s.OllamaNetMode;
                Instance.ChatGPTEndPoint = s.ChatGPTEndPoint ?? "";
                Instance.ChatGPTKey = s.ChatGPTKey ?? "";
                Instance.ChatGPTModel = s.ChatGPTModel ?? "";
                Instance.ChatGPTMaxToken = s.ChatGPTMaxToken > 0 ? s.ChatGPTMaxToken : 1000;
                Instance.ChatGPTTemperature = s.ChatGPTTemperature;
                Instance.IMGModel = s.IMGModel ?? "";
                Instance.IMGEndPoint = s.IMGEndPint ?? "";
                Instance.IMGAPIKey = s.IMGAPIKEY ?? "";
                Instance.IMGMode = s.IMGMode ?? "Global";
                Instance.HTMLModel = s.HTMLModel ?? "";
                Instance.HTMLEndPoint = s.HTMLEndPint ?? "";
                Instance.HTMLAPIKey = s.HTMLAPIKEY ?? "";
                Instance.HTMLMode = s.HTMLMode ?? "Global";
                Instance.AgentMaxIterations = s.AgentMaxIterations > 0 ? s.AgentMaxIterations : 20;
                Instance.AgentMaxResultSize = s.AgentMaxResultSize;
                Instance.LinterEnabled = s.LinterEnabled;
                Instance.LinterModel = s.LinterModel ?? "";
                Instance.LinterEndPoint = s.LinterEndPoint ?? "";
                Instance.LinterDebounceSeconds = s.LinterDebounceSeconds > 0 ? s.LinterDebounceSeconds : 30;
                Instance.C1_label = s.C1_label ?? "C11";
                Instance.C2_label = s.C2_label ?? "C22";
                Instance.C3_label = s.C3_label ?? "C33";
                Instance.C4_label = s.C4_label ?? "C44";
                Instance.C1_Prompt = s.C1_Prompt ?? "";
                Instance.C2_Prompt = s.C2_Prompt ?? "";
                Instance.C3_Prompt = s.C3_Prompt ?? "";
                Instance.C4_Prompt = s.C4_Prompt ?? "";
                Instance.C1_Funt = s.C1_Funt ?? "";
                Instance.C2_Funt = s.C2_Funt ?? "";
                Instance.C3_Funt = s.C3_Funt ?? "";
                Instance.C4_Funt = s.C4_Funt ?? "";

                Save();
                ClearLegacySecrets(s);
                SettingsManager.RemoveLegacyIni();
                System.Diagnostics.Debug.WriteLine("AppConfig: migrated from Properties.Settings");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AppConfig migration error: {ex.Message}");
            }
        }

        private static bool HasLegacySecrets(string json)
        {
            try
            {
                var config = JObject.Parse(json);
                return IsLegacySecret(config["ChatGPTKey"])
                    || IsLegacySecret(config["ClaudeKey"])
                    || IsLegacySecret(config["GeminiKey"])
                    || IsLegacySecret(config["IMGAPIKey"])
                    || IsLegacySecret(config["HTMLAPIKey"]);
            }
            catch
            {
                return false;
            }
        }

        private static bool IsLegacySecret(JToken value)
        {
            return value != null
                && value.Type == JTokenType.String
                && !SecretProtector.IsProtected(value.ToString());
        }

        private static void ClearLegacySecrets(Properties.Settings s)
        {
            try
            {
                s.ChatGPTKey = string.Empty;
                s.IMGAPIKEY = string.Empty;
                s.HTMLAPIKEY = string.Empty;
                s.Save();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AppConfig legacy secret cleanup error: {ex.Message}");
            }
        }

        #endregion
    }
}
