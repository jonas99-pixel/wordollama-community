using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using WordOllama.lib.Providers;

namespace WordOllama.lib
{
    /// <summary>
    /// Backward-compatibility wrapper. All logic moved to OpenAIProvider.
    /// </summary>
    public static class OpenAIFunctions
    {
        public class ChatGPTClient
        {
            private readonly OpenAIProvider _inner;

            public ChatGPTClient(string endpoint, string apiKey, string modelName)
            {
                _inner = new OpenAIProvider(endpoint, apiKey, modelName);
            }

            public async Task<string> ChatGPTGenerateAsync(
                string userPrompt,
                string base64Image = null,
                double? temperature = null)
            {
                return await _inner.GenerateAsync(userPrompt, base64Image, temperature);
            }

            public async Task<string> ChatGPTChatAsync(
                List<Dictionary<string, string>> chatHistory,
                JArray tools = null)
            {
                return await _inner.ChatAsync(chatHistory, tools);
            }

            public async Task ChatGPTStreamAsync(string userPrompt, Action<string> processChunk)
            {
                await _inner.StreamAsync(userPrompt, processChunk);
            }
        }
    }
}
