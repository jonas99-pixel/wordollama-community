using System;
using System.Reflection;

namespace WordOllama.lib
{
    /// <summary>
    /// Builds a support bundle without API keys, document content, or raw log bodies.
    /// </summary>
    public static class DiagnosticsService
    {
        public static string BuildSupportBundle()
        {
            AppConfig config = AppConfig.Instance;
            string appVersion = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "unknown";
            string hostVersion = "unavailable";
            try { hostVersion = Globals.ThisAddIn?.Application?.Version ?? "unavailable"; } catch { }

            string provider = config.AIMode == "Ollama"
                ? $"Ollama / {config.OllamaModel} / {MaskEndpoint(config.OllamaEndPoint)}"
                : $"{config.OnlineProviderType} / {GetActiveModel(config)} / {MaskEndpoint(GetActiveEndpoint(config))}";

            return string.Join(Environment.NewLine, new[]
            {
                "WordOllama support diagnostics",
                $"App version: {appVersion}",
                $"Word version: {hostVersion}",
                $"OS: {Environment.OSVersion}",
                $".NET: {Environment.Version}",
                $"Language mode: {config.LanguageMode}",
                $"Theme: {config.ThemeType}",
                $"Provider: {provider}",
                $"Tool calling mode: {config.ProviderToolCallingMode}",
                $"Provider timeout: {config.ProviderTimeoutSeconds}s; connect timeout: {config.ProviderConnectTimeoutSeconds}s; retries: {config.ProviderMaxRetries}",
                $"Agent mode: {config.AgentExecutionMode}",
                $"Diagnostic logging enabled: {DebugLogger.Enabled}",
                "Recent diagnostic events (metadata only):",
                DebugLogger.GetMetadataSummary(),
            });
        }

        private static string GetActiveModel(AppConfig config)
        {
            switch (config.OnlineProviderType)
            {
                case "Claude": return config.ClaudeModel;
                case "Gemini": return config.GeminiModel;
                default: return config.ChatGPTModel;
            }
        }

        private static string GetActiveEndpoint(AppConfig config)
        {
            switch (config.OnlineProviderType)
            {
                case "Claude": return config.ClaudeEndPoint;
                case "Gemini": return config.GeminiEndPoint;
                default: return config.ChatGPTEndPoint;
            }
        }

        private static string MaskEndpoint(string endpoint)
        {
            if (string.IsNullOrWhiteSpace(endpoint)) return "(not configured)";
            if (!Uri.TryCreate(endpoint, UriKind.Absolute, out Uri uri)) return "(configured)";
            return uri.Scheme + "://" + uri.Host + (uri.IsDefaultPort ? "" : ":" + uri.Port);
        }
    }
}
