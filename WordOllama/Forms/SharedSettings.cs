using System;
using System.IO;
using System.Reflection;
using System.Globalization;
using System.Windows;
using Microsoft.Office.Interop.Word;
using WordOllama.lib;

namespace WordOllama.Forms
{
    /// <summary>
    /// 设置门面 — 所有设置通过此类访问，底层统一由 AppConfig (JSON) 持久化。
    /// 保持与旧版 Properties.Settings 相同的静态属性 API，外部调用方无需改动。
    /// </summary>
    public static class SharedSettings
    {
        public static string UpdateLog
        {
            get
            {
                return ReadLocalizedUpdateLog(CultureInfo.CurrentUICulture);
            }
        }

        private static string ReadLocalizedUpdateLog(CultureInfo culture)
        {
            var assembly = Assembly.GetExecutingAssembly();
            string[] resourceNames = BuildUpdateLogResourceCandidates(culture);

            foreach (string resourceName in resourceNames)
            {
                using (var stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null) continue;
                    using (var reader = new StreamReader(stream))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }

            throw new FileNotFoundException("Update log resource not found", resourceNames[resourceNames.Length - 1]);
        }

        private static string[] BuildUpdateLogResourceCandidates(CultureInfo culture)
        {
            string name = culture?.Name ?? "";
            string neutral = culture?.TwoLetterISOLanguageName ?? "";

            if (string.Equals(neutral, "zh", StringComparison.OrdinalIgnoreCase))
            {
                return new[]
                {
                    $"WordOllama.Resources.updateLog.{name}.txt",
                    "WordOllama.Resources.updateLog.zh.txt",
                    "WordOllama.Resources.updateLog.txt",
                    "WordOllama.Resources.updateLog.en.txt"
                };
            }

            return new[]
            {
                $"WordOllama.Resources.updateLog.{name}.txt",
                $"WordOllama.Resources.updateLog.{neutral}.txt",
                "WordOllama.Resources.updateLog.en.txt",
                "WordOllama.Resources.updateLog.txt"
            };
        }

        #region AI 模式

        public static string AIMode
        {
            get { return AppConfig.Instance?.AIMode ?? "Ollama"; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.AIMode = value; AppConfig.Save(); ChangeLabel(); } }
        }

        public static string LanguageMode
        {
            get { return AppConfig.Instance?.LanguageMode ?? "Auto"; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.LanguageMode = value; AppConfig.Save(); } }
        }

        public static string themeType
        {
            get { return AppConfig.Instance?.ThemeType ?? "Light"; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.ThemeType = value; AppConfig.Save(); } }
        }

        public static string DefaultPromptOutputMode
        {
            get { return AppConfig.Instance?.DefaultPromptOutputMode ?? "Auto"; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.DefaultPromptOutputMode = value; AppConfig.Save(); } }
        }

        #endregion

        #region Ollama

        public static string OllamaEndPoint
        {
            get { return AppConfig.Instance?.OllamaEndPoint ?? "http://localhost:11434"; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.OllamaEndPoint = value; AppConfig.Save(); } }
        }

        public static string OllamaModel
        {
            get { return AppConfig.Instance?.OllamaModel ?? ""; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.OllamaModel = value; AppConfig.Save(); ChangeLabel(); } }
        }

        public static int OllamaMaxToken
        {
            get { return AppConfig.Instance?.OllamaMaxToken ?? 4096; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.OllamaMaxToken = value; AppConfig.Save(); } }
        }

        public static double OllamaTemperature
        {
            get { return AppConfig.Instance?.OllamaTemperature ?? 0.5; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.OllamaTemperature = value; AppConfig.Save(); } }
        }

        public static bool IsOllamaModelRefresh
        {
            get { return AppConfig.Instance?.IsOllamaModelRefresh ?? false; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.IsOllamaModelRefresh = value; AppConfig.Save(); } }
        }

        public static bool OllamaNetMode
        {
            get { return AppConfig.Instance?.OllamaNetMode ?? false; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.OllamaNetMode = value; AppConfig.Save(); } }
        }

        #endregion

        #region 在线模型 (通用)

        public static string OnlineProviderType
        {
            get { return AppConfig.Instance?.OnlineProviderType ?? "OpenAI"; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.OnlineProviderType = value; AppConfig.Save(); } }
        }

        public static string ChatGPTEndPoint
        {
            get { return AppConfig.Instance?.ChatGPTEndPoint ?? ""; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.ChatGPTEndPoint = value; AppConfig.Save(); } }
        }

        public static string ChatGPTModel
        {
            get { return AppConfig.Instance?.ChatGPTModel ?? ""; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.ChatGPTModel = value; AppConfig.Save(); ChangeLabel(); } }
        }

        public static string ChatGPTKey
        {
            get { return AppConfig.Instance?.ChatGPTKey ?? ""; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.ChatGPTKey = value; AppConfig.Save(); } }
        }

        public static int ChatGPTMaxToken
        {
            get { return AppConfig.Instance?.ChatGPTMaxToken ?? 1000; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.ChatGPTMaxToken = value; AppConfig.Save(); } }
        }
        public static int OnlineContextWindow
        {
            get { return AppConfig.Instance?.OnlineContextWindow ?? 256000; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.OnlineContextWindow = value; AppConfig.Save(); } }
        }

        public static double ChatGPTTemperature
        {
            get { return AppConfig.Instance?.ChatGPTTemperature ?? 0.5; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.ChatGPTTemperature = value; AppConfig.Save(); } }
        }

        // Claude
        public static string ClaudeEndPoint
        {
            get { return AppConfig.Instance?.ClaudeEndPoint ?? "https://api.anthropic.com/v1"; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.ClaudeEndPoint = value; AppConfig.Save(); } }
        }

        public static string ClaudeKey
        {
            get { return AppConfig.Instance?.ClaudeKey ?? ""; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.ClaudeKey = value; AppConfig.Save(); } }
        }

        public static string ClaudeModel
        {
            get { return AppConfig.Instance?.ClaudeModel ?? "claude-sonnet-4-20250514"; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.ClaudeModel = value; AppConfig.Save(); ChangeLabel(); } }
        }

        // Gemini
        public static string GeminiEndPoint
        {
            get { return AppConfig.Instance?.GeminiEndPoint ?? "https://generativelanguage.googleapis.com/v1beta"; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.GeminiEndPoint = value; AppConfig.Save(); } }
        }

        public static string GeminiKey
        {
            get { return AppConfig.Instance?.GeminiKey ?? ""; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.GeminiKey = value; AppConfig.Save(); } }
        }

        public static string GeminiModel
        {
            get { return AppConfig.Instance?.GeminiModel ?? "gemini-3-flash-preview"; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.GeminiModel = value; AppConfig.Save(); ChangeLabel(); } }
        }

        #endregion

        #region 图片模型

        public static string IMGModel
        {
            get { return AppConfig.Instance?.IMGModel ?? ""; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.IMGModel = value; AppConfig.Save(); } }
        }

        public static string IMGEndPint
        {
            get { return AppConfig.Instance?.IMGEndPoint ?? ""; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.IMGEndPoint = value; AppConfig.Save(); } }
        }

        public static string IMGAPIKEY
        {
            get { return AppConfig.Instance?.IMGAPIKey ?? ""; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.IMGAPIKey = value; AppConfig.Save(); } }
        }

        public static string IMGMode
        {
            get { return AppConfig.Instance?.IMGMode ?? "Global"; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.IMGMode = value; AppConfig.Save(); } }
        }

        #endregion

        #region HTML 模型

        public static string HTMLModel
        {
            get { return AppConfig.Instance?.HTMLModel ?? ""; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.HTMLModel = value; AppConfig.Save(); } }
        }

        public static string HTMLEndPint
        {
            get { return AppConfig.Instance?.HTMLEndPoint ?? ""; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.HTMLEndPoint = value; AppConfig.Save(); } }
        }

        public static string HTMLAPIKEY
        {
            get { return AppConfig.Instance?.HTMLAPIKey ?? ""; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.HTMLAPIKey = value; AppConfig.Save(); } }
        }

        public static string HTMLMode
        {
            get { return AppConfig.Instance?.HTMLMode ?? "Global"; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.HTMLMode = value; AppConfig.Save(); } }
        }

        #endregion

        #region Agent

        public static int AgentMaxIterations
        {
            get { return AppConfig.Instance?.AgentMaxIterations ?? 20; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.AgentMaxIterations = value; AppConfig.Save(); } }
        }

        public static int AgentMaxResultSize
        {
            get { return AppConfig.Instance?.AgentMaxResultSize ?? 0; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.AgentMaxResultSize = value; AppConfig.Save(); } }
        }

        public static bool AgentExternalToolsEnabled
        {
            get { return AppConfig.Instance?.AgentExternalToolsEnabled ?? false; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.AgentExternalToolsEnabled = value; AppConfig.Save(); } }
        }

        public static bool SuppressAgentNotifications
        {
            get { return AppConfig.Instance?.SuppressAgentNotifications ?? false; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.SuppressAgentNotifications = value; AppConfig.Save(); } }
        }

        #endregion

        #region 静默审查

        public static bool LinterEnabled
        {
            get { return AppConfig.Instance?.LinterEnabled ?? false; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.LinterEnabled = value; AppConfig.Save(); } }
        }

        public static string LinterModel
        {
            get { return AppConfig.Instance?.LinterModel ?? ""; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.LinterModel = value; AppConfig.Save(); } }
        }

        public static string LinterEndPoint
        {
            get { return AppConfig.Instance?.LinterEndPoint ?? ""; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.LinterEndPoint = value; AppConfig.Save(); } }
        }

        public static int LinterDebounceSeconds
        {
            get { return AppConfig.Instance?.LinterDebounceSeconds ?? 30; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.LinterDebounceSeconds = value; AppConfig.Save(); } }
        }

        #endregion

        #region 思维链 (Thinking)

        public static bool ThinkingEnabled
        {
            get { return AppConfig.Instance?.ThinkingEnabled ?? true; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.ThinkingEnabled = value; AppConfig.Save(); } }
        }

        public static string ReasoningEffort
        {
            get { return AppConfig.Instance?.ReasoningEffort ?? "high"; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.ReasoningEffort = value; AppConfig.Save(); } }
        }

        #endregion

        #region Markdown 转换

        public static string MarkdownStyleMappingsJson
        {
            get { return AppConfig.Instance?.MarkdownStyleMappingsJson ?? ""; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.MarkdownStyleMappingsJson = value; AppConfig.Save(); } }
        }

        #endregion

        #region Agent 确认设置

        public static bool SuppressPlanConfirmation
        {
            get { return AppConfig.Instance?.SuppressPlanConfirmation ?? false; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.SuppressPlanConfirmation = value; AppConfig.Save(); } }
        }

        public static bool SuppressDiffReviewPrompt
        {
            get { return AppConfig.Instance?.SuppressDiffReviewPrompt ?? false; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.SuppressDiffReviewPrompt = value; AppConfig.Save(); } }
        }

        public static bool SuppressErrorRecoveryPrompt
        {
            get { return AppConfig.Instance?.SuppressErrorRecoveryPrompt ?? false; }
            set { if (AppConfig.Instance != null) { AppConfig.Instance.SuppressErrorRecoveryPrompt = value; AppConfig.Save(); } }
        }

        #endregion

        // 辅助方法
        private static void ChangeLabel()
        {
            try { Globals.Ribbons.Ribbon1.ChangeRibbonLabels(); } catch { }
        }
    }

    public static class GetText
    {
        private static string GetDocumentText(string type)
        {
            try
            {
                Document activeDocument = Globals.ThisAddIn.Application.ActiveDocument;
                if (activeDocument != null)
                {
                    Selection selection = Globals.ThisAddIn.Application.Selection;
                    string result;
                    if (!string.IsNullOrEmpty(selection.Text.Trim()))
                    {
                        result = selection.Text;
                    }
                    else if (type == "FULL")
                    {
                        Range contentRange = activeDocument.Content;
                        result = contentRange.Text;
                        ComHelper.ReleaseCom(ref contentRange);
                    }
                    else
                    {
                        result = string.Empty;
                    }
                    // 释放临时 COM 引用（注意：ActiveDocument/Selection 由 VSTO 运行时管理，不应释放）
                    return result;
                }
                else
                {
                    MessageBox.Show(WordOllama.lib.Loc.T("I18n_NoActiveDocument"));
                    return string.Empty;
                }
            }
            catch { return string.Empty; }
        }

        public static string SelectedText => GetDocumentText("SELECT");
        public static string FullText => GetDocumentText("FULL");
    }
}
