using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Threading;
using WordOllama.NewUI;

namespace WordOllama.lib.Providers
{
    public class OpenAIProvider : IOnlineProvider, ICancellableOnlineProvider
    {
        public string ProviderType => "OpenAI";

        private readonly string _endpoint;
        private readonly string _apiKey;
        private readonly string _modelName;
        private readonly bool _requestJsonOutput;

        private static readonly string thinkPattern = @"<think>([\s\S]*?)</think>";

        public OpenAIProvider(string endpoint, string apiKey, string modelName, bool requestJsonOutput = false)
        {
            _endpoint = endpoint;
            _apiKey = apiKey;
            _modelName = modelName;
            _requestJsonOutput = requestJsonOutput;
        }

        public async Task<string> GenerateAsync(string prompt, string imageBase64 = null, double? temperature = null)
        {
            return await CommonChatRequestAsync(prompt, imageBase64, null, null, temperature);
        }

        public async Task<string> GenerateAsync(
            string prompt,
            string imageBase64,
            double? temperature,
            CancellationToken cancellationToken)
        {
            return await CommonChatRequestAsync(prompt, imageBase64, null, null, temperature, cancellationToken);
        }

        public async Task<string> ChatAsync(List<Dictionary<string, string>> chatHistory, JArray tools = null)
        {
            return await CommonChatRequestAsync(null, null, chatHistory, tools);
        }

        public async Task<string> ChatAsync(List<Dictionary<string, string>> chatHistory, JArray tools, CancellationToken cancellationToken)
        {
            return await CommonChatRequestAsync(null, null, chatHistory, tools, null, cancellationToken);
        }

        public async Task StreamAsync(string prompt, Action<string> processChunk)
        {
            await ChatGPTStreamAsync(prompt, processChunk);
        }

        public async Task<List<string>> FetchModelsAsync()
        {
            var models = new List<string>();
            try
            {
                using (var httpClient = ProviderRequestPolicy.CreateConnectClient())
                {
                    if (!string.IsNullOrEmpty(_apiKey))
                        httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + _apiKey);
                    string baseUrl = GetBaseUrl(_endpoint);
                    string modelsUrl = baseUrl.TrimEnd('/') + "/models";

                    var response = await httpClient.GetAsync(modelsUrl);
                    if (response.IsSuccessStatusCode)
                    {
                        string body = await response.Content.ReadAsStringAsync();
                        var json = JObject.Parse(body);
                        var data = json["data"] as JArray;
                        if (data != null)
                        {
                            foreach (var item in data)
                            {
                                string id = item["id"]?.ToString();
                                if (!string.IsNullOrEmpty(id))
                                    models.Add(id);
                            }
                        }
                    }
                }
            }
            catch { }
            return models;
        }

        private string GetBaseUrl(string endpoint)
        {
            if (string.IsNullOrEmpty(endpoint)) return "https://api.openai.com/v1";
            int idx = endpoint.IndexOf("/chat/completions", StringComparison.OrdinalIgnoreCase);
            if (idx > 0) return endpoint.Substring(0, idx);
            idx = endpoint.IndexOf("/v1/", StringComparison.OrdinalIgnoreCase);
            if (idx > 0) return endpoint.Substring(0, idx + 3);
            return endpoint.TrimEnd('/');
        }

        private string GetMaxTokenParamName()
        {
            var reasoningModels = new[] { "o1", "o3", "o4", "gpt-5" };
            string model = _modelName?.Trim() ?? "";
            foreach (var prefix in reasoningModels)
            {
                if (model.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                    || model.StartsWith("ft:" + prefix, StringComparison.OrdinalIgnoreCase))
                    return "max_completion_tokens";
            }
            return "max_tokens";
        }

        private void UpdateThinkingPane(string text)
        {
            if (string.IsNullOrEmpty(text)) return;
            try
            {
                var pane = Globals.ThisAddIn?.thinkingPane;
                if (pane != null && pane.IsHandleCreated)
                {
                    if (pane.InvokeRequired)
                        pane.Invoke(new Action(() => pane.UpdateThinkText(text)));
                    else
                        pane.UpdateThinkText(text);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"更新思考窗格失败: {ex.Message}");
            }
        }

        private async Task<string> CommonChatRequestAsync(
            string prompt,
            string base64Image,
            List<Dictionary<string, string>> history,
            JArray tools = null,
            double? temperature = null,
            CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(_endpoint))
                return "服务地址为空，请输入有效的服务地址";

            if (!Uri.TryCreate(_endpoint, UriKind.Absolute, out Uri uriResult)
                || (uriResult.Scheme != Uri.UriSchemeHttp && uriResult.Scheme != Uri.UriSchemeHttps))
            {
                return "服务地址格式错误，请输入有效的URL";
            }

            using (HttpClient httpClient = ProviderRequestPolicy.CreateChatClient())
            {
                httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + _apiKey);
                if (_endpoint.IndexOf(".openai.azure.com", StringComparison.OrdinalIgnoreCase) >= 0)
                    httpClient.DefaultRequestHeaders.Add("api-key", _apiKey);

                var requestData = new Dictionary<string, object>
                {
                    { "model", _modelName },
                    { "stream", false }
                };
                if (_requestJsonOutput)
                    requestData["response_format"] = new Dictionary<string, string> { { "type", "json_object" } };
                requestData.Add(GetMaxTokenParamName(), AppConfig.Instance.ChatGPTMaxToken);

                if (AppConfig.Instance.ThinkingEnabled)
                {
                    requestData["thinking"] = new Dictionary<string, object> { { "type", "enabled" } };
                    requestData["reasoning_effort"] = AppConfig.Instance.ReasoningEffort;
                }

                if (temperature.HasValue)
                    requestData["temperature"] = temperature.Value;

                if (history != null)
                {
                    string pendingImage = DocumentAgent.PendingImageBase64;
                    if (!string.IsNullOrEmpty(pendingImage))
                    {
                        var multiModalMessages = new List<object>();
                        bool firstUserFound = false;
                        foreach (var msg in history)
                        {
                            if (!firstUserFound && msg.TryGetValue("role", out var role) && role == "user")
                            {
                                firstUserFound = true;
                                multiModalMessages.Add(new Dictionary<string, object>
                                {
                                    ["role"] = "user",
                                    ["content"] = new List<object>
                                    {
                                        new Dictionary<string, string> { ["type"] = "text", ["text"] = msg["content"] },
                                        new Dictionary<string, object>
                                        {
                                            ["type"] = "image_url",
                                            ["image_url"] = new Dictionary<string, string> { ["url"] = $"data:image/jpeg;base64,{pendingImage}" }
                                        }
                                    }
                                });
                            }
                            else
                            {
                                multiModalMessages.Add(msg.ToDictionary(kvp => kvp.Key, kvp => (object)kvp.Value));
                            }
                        }
                        requestData["messages"] = multiModalMessages;
                        DocumentAgent.PendingImageBase64 = null;
                    }
                    else
                    {
                        requestData["messages"] = history;
                    }
                }
                else
                {
                    var messagesList = new List<Dictionary<string, object>>();
                    var message = new Dictionary<string, object> { ["role"] = "user" };
                    if (string.IsNullOrEmpty(base64Image))
                    {
                        message["content"] = prompt;
                    }
                    else
                    {
                        message["content"] = new List<object>
                        {
                            new Dictionary<string, string> { { "type", "text" }, { "text", prompt } },
                            new Dictionary<string, object>
                            {
                                { "type", "image_url" },
                                { "image_url", new Dictionary<string, string> { { "url", $"data:image/jpeg;base64,{base64Image}" } } }
                            }
                        };
                    }
                    messagesList.Add(message);
                    requestData["messages"] = messagesList;
                }

                if (tools != null && tools.Count > 0)
                    requestData["tools"] = tools;

                bool hadImage = !string.IsNullOrEmpty(DocumentAgent.PendingImageBase64) || !string.IsNullOrEmpty(base64Image);

                var requestBody = new StringContent(JsonConvert.SerializeObject(requestData), Encoding.UTF8, "application/json");

                try
                {
                    var startTime = DateTime.Now;
                    HttpResponseMessage response = await httpClient.PostAsync(_endpoint, requestBody, cancellationToken);
                    string responseBody = await response.Content.ReadAsStringAsync();

                    if (string.IsNullOrEmpty(responseBody))
                        return "服务器未响应";

                    if (response.IsSuccessStatusCode)
                    {
                        dynamic responseJson = JsonConvert.DeserializeObject(responseBody);
                        string content = responseJson.choices[0].message.content ?? "";

                        try
                        {
                            var toolCalls = responseJson.choices[0].message.tool_calls;
                            if (toolCalls != null && toolCalls.Count > 0)
                            {
                                var firstCall = toolCalls[0];
                                string funcName = firstCall.function.name;
                                string funcArgs = firstCall.function.arguments;
                                return $"[TOOL_CALL]\nname: {funcName}\narguments: {funcArgs}";
                            }
                        }
                        catch { }

                        string refusal = responseJson.choices[0].message.refusal;
                        if (!string.IsNullOrEmpty(refusal))
                            return $"[API 拒绝] {refusal}";

                        string reasoningField = "";
                        try { reasoningField = responseJson.choices[0].message.reasoning_content; } catch { }

                        string tagContent = "";
                        Match match = Regex.Match(content, thinkPattern, RegexOptions.Singleline);
                        if (match.Success)
                        {
                            tagContent = match.Groups[1].Value;
                            content = Regex.Replace(content, thinkPattern, string.Empty, RegexOptions.Singleline).Trim();
                            content = content.TrimStart('\n', '\r');
                        }

                        string combinedThinking = (reasoningField ?? "") + (tagContent ?? "");
                        if (!string.IsNullOrEmpty(combinedThinking))
                            UpdateThinkingPane(combinedThinking);

                        HandleUsageStats(startTime, responseJson);
                        return content;
                    }
                    else
                    {
                        if (hadImage)
                        {
                            DocumentAgent.PendingImageBase64 = null;
                            return await CommonChatRequestAsync(prompt, null, history, tools, null, cancellationToken);
                        }
                        return HandleErrorResponse(response, responseBody);
                    }
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex) when (hadImage)
                {
                    DocumentAgent.PendingImageBase64 = null;
                    Debug.WriteLine($"带图片请求失败，降级重试: {ex.Message}");
                    return await CommonChatRequestAsync(prompt, null, history, tools, null, cancellationToken);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"异常: {ex.Message}");
                    return ProviderRequestPolicy.NormalizeError("OpenAI", null, null, ex);
                }
            }
        }

        private async Task ChatGPTStreamAsync(string userPrompt, Action<string> processChunk)
        {
            using (HttpClient httpClient = ProviderRequestPolicy.CreateChatClient())
            {
                httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + _apiKey);
                if (_endpoint.IndexOf(".openai.azure.com", StringComparison.OrdinalIgnoreCase) >= 0)
                    httpClient.DefaultRequestHeaders.Add("api-key", _apiKey);

                var messagesList = new List<Dictionary<string, string>>
                {
                    new Dictionary<string, string> { { "role", "user" }, { "content", userPrompt } }
                };

                var requestData = new Dictionary<string, object>
                {
                    { "model", _modelName },
                    { "messages", messagesList },
                    { "stream", true }
                };
                requestData.Add(GetMaxTokenParamName(), Math.Max(1, Math.Min(AppConfig.Instance.ChatGPTMaxToken, 8192)));

                if (AppConfig.Instance.ThinkingEnabled)
                {
                    requestData["thinking"] = new Dictionary<string, object> { { "type", "enabled" } };
                    requestData["reasoning_effort"] = AppConfig.Instance.ReasoningEffort;
                }

                var requestBody = new StringContent(JsonConvert.SerializeObject(requestData), Encoding.UTF8, "application/json");

                try
                {
                    HttpResponseMessage response = await httpClient.PostAsync(_endpoint, requestBody);
                    response.EnsureSuccessStatusCode();

                    bool isThinkingInTag = false;

                    using (var responseStream = await response.Content.ReadAsStreamAsync())
                    using (var streamReader = new StreamReader(responseStream))
                    {
                        while (!streamReader.EndOfStream)
                        {
                            var line = await streamReader.ReadLineAsync();
                            if (string.IsNullOrWhiteSpace(line) || !line.StartsWith("data:"))
                                continue;

                            var json = line.Substring("data:".Length).Trim();
                            if (json == "[DONE]")
                                break;

                            try
                            {
                                dynamic responseJson = JsonConvert.DeserializeObject<dynamic>(json);
                                if (responseJson.choices == null || responseJson.choices.Count == 0)
                                {
                                    if (responseJson.usage != null)
                                        HandleUsageStats(DateTime.Now, responseJson);
                                    continue;
                                }

                                var delta = responseJson.choices[0].delta;

                                string explicitReasoning = delta?.reasoning_content?.ToString();
                                if (!string.IsNullOrEmpty(explicitReasoning))
                                    UpdateThinkingPane(explicitReasoning);

                                string contentChunk = delta?.content?.ToString();
                                if (!string.IsNullOrEmpty(contentChunk))
                                {
                                    string textToDoc = "";
                                    string textToThink = "";

                                    for (int i = 0; i < contentChunk.Length; i++)
                                    {
                                        if (!isThinkingInTag)
                                        {
                                            if (contentChunk.Substring(i).StartsWith("<think>"))
                                            {
                                                isThinkingInTag = true;
                                                i += 6;
                                                continue;
                                            }
                                            textToDoc += contentChunk[i];
                                        }
                                        else
                                        {
                                            if (contentChunk.Substring(i).StartsWith("</think>"))
                                            {
                                                isThinkingInTag = false;
                                                i += 7;
                                                continue;
                                            }
                                            textToThink += contentChunk[i];
                                        }
                                    }

                                    if (isThinkingInTag && string.IsNullOrEmpty(textToThink) && string.IsNullOrEmpty(textToDoc) && contentChunk.Length > 0)
                                        textToThink = contentChunk;

                                    if (!string.IsNullOrEmpty(textToThink))
                                        UpdateThinkingPane(textToThink);
                                    if (!string.IsNullOrEmpty(textToDoc))
                                        processChunk(textToDoc);
                                }
                            }
                            catch (JsonException)
                            {
                                // A streaming chunk may contain an incomplete JSON fragment.
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Stream Exception: {ex.Message}");
                    throw;
                }
            }
        }

        private string HandleErrorResponse(HttpResponseMessage response, string responseBody)
        {
            Debug.WriteLine($"Error {(int)response.StatusCode} - {responseBody}");
            string errorMsg = "";
            try
            {
                dynamic errorJson = JsonConvert.DeserializeObject(responseBody);
                errorMsg = errorJson?.error?.message ?? "";
            }
            catch { }

            if (string.IsNullOrWhiteSpace(errorMsg))
            {
                switch ((int)response.StatusCode)
                {
                    case 400: errorMsg = "请求格式错误，请检查输入"; break;
                    case 401: errorMsg = "API密钥无效"; break;
                    case 403: errorMsg = "没有权限访问"; break;
                    case 404: errorMsg = "地址错误"; break;
                    case 429: errorMsg = "请求过于频繁"; break;
                    case 500: errorMsg = "服务器内部错误"; break;
                    default: errorMsg = $"未知错误 ({(int)response.StatusCode})"; break;
                }
            }
            if (errorMsg.Length > 500)
                errorMsg = errorMsg.Substring(0, 500) + "...";
            return ProviderRequestPolicy.NormalizeError("OpenAI", response.StatusCode, errorMsg);
        }

        private void HandleUsageStats(DateTime startTime, dynamic responseJson)
        {
            double totalDurationSeconds = (DateTime.Now - startTime).TotalSeconds;
            string promptTokens = "?";
            string completionTokens = "?";
            string total_tokens = "?";
            string reasoningTokens = "";
            string cacheHit = "";
            string cacheMiss = "";

            try
            {
                if (responseJson.usage != null)
                {
                    promptTokens = responseJson.usage.prompt_tokens?.ToString() ?? "?";
                    completionTokens = responseJson.usage.completion_tokens?.ToString() ?? "?";
                    total_tokens = responseJson.usage.total_tokens?.ToString() ?? "?";
                    if (responseJson.usage.completion_tokens_details?.reasoning_tokens != null)
                        reasoningTokens = $" | R: {responseJson.usage.completion_tokens_details.reasoning_tokens}";
                    if (responseJson.usage.prompt_cache_hit_tokens != null)
                        cacheHit = $" | CH: {responseJson.usage.prompt_cache_hit_tokens}";
                    if (responseJson.usage.prompt_cache_miss_tokens != null)
                        cacheMiss = $" | CM: {responseJson.usage.prompt_cache_miss_tokens}";
                }
                if (responseJson.total_duration != null)
                {
                    long totalDurationNs = (long)responseJson.total_duration;
                    totalDurationSeconds = totalDurationNs / 1_000_000_000.0;
                }
            }
            catch { }

            if (!DocumentAgent.IsAgentRunning)
            {
                Globals.Ribbons.Ribbon1.ShowNotification(
                    $"耗时 {totalDurationSeconds:F2}s | P: {promptTokens} | C: {completionTokens} | T: {total_tokens}{reasoningTokens}{cacheHit}{cacheMiss}",
                    1000,
                    System.Windows.Forms.ToolTipIcon.Info);
            }
        }
    }
}
