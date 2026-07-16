using System;

namespace WordOllama.lib.Providers
{
    public static class ProviderFactory
    {
        public static IOnlineProvider Create(string providerType, string endpoint, string apiKey, string modelName, bool requestJsonOutput = false)
        {
            return providerType switch
            {
                "OpenAI" => new OpenAIProvider(endpoint, apiKey, modelName, requestJsonOutput),
                "Claude" => new ClaudeProvider(endpoint, apiKey, modelName),
                "Gemini" => new GeminiProvider(endpoint, apiKey, modelName),
                _ => new OpenAIProvider(endpoint, apiKey, modelName, requestJsonOutput)
            };
        }

        /// <summary>
        /// Create provider from current AppConfig settings.
        /// </summary>
        public static IOnlineProvider CreateFromConfig(bool requestJsonOutput = false)
        {
            string type = AppConfig.Instance?.OnlineProviderType ?? "OpenAI";

            return type switch
            {
                "OpenAI" => new OpenAIProvider(
                    AppConfig.Instance.ChatGPTEndPoint,
                    AppConfig.Instance.ChatGPTKey,
                    AppConfig.Instance.ChatGPTModel,
                    requestJsonOutput),

                "Claude" => new ClaudeProvider(
                    AppConfig.Instance.ClaudeEndPoint,
                    AppConfig.Instance.ClaudeKey,
                    AppConfig.Instance.ClaudeModel),

                "Gemini" => new GeminiProvider(
                    AppConfig.Instance.GeminiEndPoint,
                    AppConfig.Instance.GeminiKey,
                    AppConfig.Instance.GeminiModel),

                _ => new OpenAIProvider(
                    AppConfig.Instance.ChatGPTEndPoint,
                    AppConfig.Instance.ChatGPTKey,
                    AppConfig.Instance.ChatGPTModel,
                    requestJsonOutput)
            };
        }
    }
}
