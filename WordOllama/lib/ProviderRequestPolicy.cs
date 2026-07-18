using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace WordOllama.lib
{
    public sealed class ProviderCapabilityProfile
    {
        public string ToolCallingMode { get; set; } = "Auto";
        public bool SupportsStreaming { get; set; }
        public bool SupportsVision { get; set; }
        public bool SupportsJsonOutput { get; set; }
        public int ContextWindow { get; set; }
    }

    public static class ProviderRequestPolicy
    {
        private static readonly HttpClientHandler SharedHandler = new HttpClientHandler();

        public static TimeSpan ChatTimeout => TimeSpan.FromSeconds(Math.Max(30, AppConfig.Instance?.ProviderTimeoutSeconds ?? 120));
        public static TimeSpan ConnectTimeout => TimeSpan.FromSeconds(Math.Max(5, AppConfig.Instance?.ProviderConnectTimeoutSeconds ?? 15));

        public static void ConfigureChatClient(HttpClient client)
        {
            client.Timeout = ChatTimeout;
        }

        public static void ConfigureConnectClient(HttpClient client)
        {
            client.Timeout = ConnectTimeout;
        }

        public static HttpClient CreateChatClient()
        {
            return CreateClient(ChatTimeout);
        }

        public static HttpClient CreateConnectClient()
        {
            return CreateClient(ConnectTimeout);
        }

        public static HttpClient CreateClient(TimeSpan timeout)
        {
            return new HttpClient(SharedHandler, disposeHandler: false)
            {
                Timeout = timeout
            };
        }

        public static ProviderCapabilityProfile GetDefaultCapabilities(string providerType)
        {
            switch (providerType)
            {
                case "Claude":
                    return new ProviderCapabilityProfile
                    {
                        SupportsStreaming = true,
                        SupportsVision = true,
                        SupportsJsonOutput = true,
                        ContextWindow = 200000
                    };
                case "Gemini":
                    return new ProviderCapabilityProfile
                    {
                        SupportsStreaming = true,
                        SupportsVision = true,
                        SupportsJsonOutput = true,
                        ContextWindow = 1000000
                    };
                default:
                    return new ProviderCapabilityProfile
                    {
                        SupportsStreaming = true,
                        SupportsVision = true,
                        SupportsJsonOutput = true,
                        ContextWindow = 128000
                    };
            }
        }

        public static ProviderCapabilityProfile ReadCapabilities(JToken token, string providerType)
        {
            var defaults = GetDefaultCapabilities(providerType);
            if (token == null || token.Type == JTokenType.Null) return defaults;

            try
            {
                var saved = token.ToObject<ProviderCapabilityProfile>();
                if (saved == null) return defaults;
                return new ProviderCapabilityProfile
                {
                    ToolCallingMode = string.IsNullOrWhiteSpace(saved.ToolCallingMode) ? defaults.ToolCallingMode : saved.ToolCallingMode,
                    SupportsStreaming = saved.SupportsStreaming,
                    SupportsVision = saved.SupportsVision,
                    SupportsJsonOutput = saved.SupportsJsonOutput,
                    ContextWindow = saved.ContextWindow > 0 ? saved.ContextWindow : defaults.ContextWindow
                };
            }
            catch
            {
                return defaults;
            }
        }

        public static ProviderCapabilityProfile GetActiveCapabilities()
        {
            var config = AppConfig.Instance;
            var defaults = GetDefaultCapabilities(config?.OnlineProviderType);
            if (config == null) return defaults;
            return new ProviderCapabilityProfile
            {
                ToolCallingMode = string.IsNullOrWhiteSpace(config.ProviderToolCallingMode) ? defaults.ToolCallingMode : config.ProviderToolCallingMode,
                SupportsStreaming = config.ProviderSupportsStreaming,
                SupportsVision = config.ProviderSupportsVision,
                SupportsJsonOutput = config.ProviderSupportsJsonOutput,
                ContextWindow = config.OnlineContextWindow > 0 ? config.OnlineContextWindow : defaults.ContextWindow
            };
        }

        public static void ApplyActiveCapabilities(ProviderCapabilityProfile profile)
        {
            if (AppConfig.Instance == null || profile == null) return;
            AppConfig.Instance.ProviderToolCallingMode = ProviderCompatibility.ParseToolCallingMode(profile.ToolCallingMode).ToString();
            AppConfig.Instance.ProviderSupportsStreaming = profile.SupportsStreaming;
            AppConfig.Instance.ProviderSupportsVision = profile.SupportsVision;
            AppConfig.Instance.ProviderSupportsJsonOutput = profile.SupportsJsonOutput;
            if (profile.ContextWindow > 0)
                AppConfig.Instance.OnlineContextWindow = profile.ContextWindow;
        }

        public static List<Dictionary<string, string>> TrimChatHistoryToContextWindow(List<Dictionary<string, string>> history)
        {
            if (history == null || history.Count == 0) return history ?? new List<Dictionary<string, string>>();

            var config = AppConfig.Instance;
            int outputTokens = Math.Max(256, config?.ChatGPTMaxToken ?? 1000);
            int contextTokens = Math.Max(outputTokens + 256, GetActiveCapabilities().ContextWindow);
            int characterBudget = Math.Max(1024, (contextTokens - outputTokens) * 4);
            int totalCharacters = history.Sum(message => message.TryGetValue("content", out var content) ? content?.Length ?? 0 : 0);
            if (totalCharacters <= characterBudget) return history;

            var systemMessages = history
                .Where(message => message.TryGetValue("role", out var role) && string.Equals(role, "system", StringComparison.OrdinalIgnoreCase))
                .ToList();
            int remaining = Math.Max(0, characterBudget - systemMessages.Sum(message => message.TryGetValue("content", out var content) ? content?.Length ?? 0 : 0));
            var recentMessages = new List<Dictionary<string, string>>();
            for (int index = history.Count - 1; index >= 0; index--)
            {
                var message = history[index];
                if (message.TryGetValue("role", out var role) && string.Equals(role, "system", StringComparison.OrdinalIgnoreCase))
                    continue;
                int length = message.TryGetValue("content", out var content) ? content?.Length ?? 0 : 0;
                if (recentMessages.Count > 0 && length > remaining) break;
                recentMessages.Add(message);
                remaining -= length;
            }
            recentMessages.Reverse();
            return systemMessages.Concat(recentMessages).ToList();
        }

        public static string NormalizeError(string provider, HttpStatusCode? statusCode, string detail, Exception exception = null)
        {
            string message = exception?.Message ?? detail ?? "Unknown provider error";
            if (exception is TimeoutException || message.IndexOf("timeout", StringComparison.OrdinalIgnoreCase) >= 0)
                return $"[API 错误] [{provider}:timeout] {message}";
            if (statusCode == HttpStatusCode.Unauthorized || statusCode == HttpStatusCode.Forbidden)
                return $"[API 错误] [{provider}:authentication] {message}";
            if (statusCode == HttpStatusCode.BadRequest)
                return $"[API 错误] [{provider}:request] {message}";
            if (statusCode.HasValue && (int)statusCode.Value == 429)
                return $"[API 错误] [{provider}:rate_limit] {message}";
            if (statusCode.HasValue && (int)statusCode.Value >= 500)
                return $"[API 错误] [{provider}:server] {message}";
            return $"[API 错误] [{provider}:transport] {message}";
        }

        public static bool IsTransientError(string response)
        {
            if (string.IsNullOrWhiteSpace(response)) return false;
            return response.IndexOf(":timeout]", StringComparison.OrdinalIgnoreCase) >= 0
                || response.IndexOf(":rate_limit]", StringComparison.OrdinalIgnoreCase) >= 0
                || response.IndexOf(":server]", StringComparison.OrdinalIgnoreCase) >= 0
                || response.IndexOf(":transport]", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static TimeSpan GetRetryDelay(int retryIndex)
        {
            int boundedIndex = Math.Max(0, Math.Min(retryIndex, 3));
            return TimeSpan.FromMilliseconds(500 * (boundedIndex + 1));
        }
    }
}
