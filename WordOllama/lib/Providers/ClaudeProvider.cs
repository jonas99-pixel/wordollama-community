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
using System.Threading.Tasks;
using System.Threading;
using WordOllama.NewUI;

namespace WordOllama.lib.Providers
{
    public class ClaudeProvider : IOnlineProvider, ICancellableOnlineProvider
    {
        public string ProviderType => "Claude";

        private readonly string _endpoint;
        private readonly string _apiKey;
        private readonly string _modelName;
        private readonly string _reasoningEffort;
        private readonly bool _thinkingEnabled;
        private readonly string _authScheme; // "x-api-key" | "Bearer"

        public ClaudeProvider(string endpoint, string apiKey, string modelName, string reasoningEffort = null, bool? thinkingEnabled = null)
        {
            _endpoint = endpoint;
            _apiKey = apiKey;
            _modelName = modelName;
            _reasoningEffort = reasoningEffort;
            _thinkingEnabled = thinkingEnabled ?? AppConfig.Instance?.ThinkingEnabled ?? false;
            // Detect auth scheme
            _authScheme = string.IsNullOrEmpty(apiKey) || apiKey.StartsWith("sk-") ? "x-api-key" : "Bearer";
        }

        public async Task<string> GenerateAsync(string prompt, string imageBase64 = null, double? temperature = null)
        {
            var messages = BuildMessages(prompt, imageBase64);
            return await SendRequestAsync(messages, null, temperature);
        }

        public async Task<string> GenerateAsync(
            string prompt,
            string imageBase64,
            double? temperature,
            CancellationToken cancellationToken)
        {
            var messages = BuildMessages(prompt, imageBase64);
            return await SendRequestAsync(messages, null, temperature, cancellationToken);
        }

        public async Task<string> ChatAsync(List<Dictionary<string, string>> chatHistory, JArray tools = null)
        {
            var messages = ConvertHistory(chatHistory);
            return await SendRequestAsync(messages, tools, null);
        }

        public async Task<string> ChatAsync(List<Dictionary<string, string>> chatHistory, JArray tools, CancellationToken cancellationToken)
        {
            var messages = ConvertHistory(chatHistory);
            return await SendRequestAsync(messages, tools, null, cancellationToken);
        }

        public async Task StreamAsync(string prompt, Action<string> processChunk)
        {
            await SendStreamRequestAsync(prompt, processChunk);
        }

        public Task<List<string>> FetchModelsAsync()
        {
            return Task.FromResult(new List<string>());
        }

        private string NormalizeEndpoint()
        {
            string ep = _endpoint;
            if (string.IsNullOrEmpty(ep))
                ep = "https://api.anthropic.com/v1/messages";
            if (!ep.EndsWith("/v1/messages", StringComparison.OrdinalIgnoreCase)
                && !ep.EndsWith("/v1/messages/", StringComparison.OrdinalIgnoreCase))
            {
                ep = ep.TrimEnd('/');
                if (!ep.EndsWith("/v1/messages"))
                    ep = ep + "/v1/messages";
            }
            return ep;
        }

        private JArray BuildMessages(string prompt, string imageBase64)
        {
            var content = new JArray();
            if (!string.IsNullOrEmpty(imageBase64))
            {
                content.Add(new JObject
                {
                    ["type"] = "image",
                    ["source"] = new JObject
                    {
                        ["type"] = "base64",
                        ["media_type"] = "image/jpeg",
                        ["data"] = imageBase64
                    }
                });
            }
            content.Add(new JObject { ["type"] = "text", ["text"] = prompt ?? "" });

            return new JArray
            {
                new JObject { ["role"] = "user", ["content"] = content }
            };
        }

        private JArray ConvertHistory(List<Dictionary<string, string>> chatHistory)
        {
            var messages = new JArray();
            string pendingImage = DocumentAgent.PendingImageBase64;

            foreach (var msg in chatHistory)
            {
                if (msg.TryGetValue("native", out var native))
                { messages.Add(new JObject { ["role"] = "assistant", ["content"] = JObject.Parse(native)["raw"].DeepClone() }); continue; }
                if (msg.TryGetValue("role", out var toolRole) && toolRole == "tool")
                { messages.Add(new JObject { ["role"] = "user", ["content"] = new JArray(new JObject { ["type"] = "tool_result", ["tool_use_id"] = msg["tool_call_id"], ["content"] = msg["content"] }) }); continue; }
                string role = msg.ContainsKey("role") ? msg["role"] : "user";
                string content = msg.ContainsKey("content") ? msg["content"] : "";

                // Inject pending image into the first user message
                bool isFirstUser = messages.Count == 0 && role == "user" && !string.IsNullOrEmpty(pendingImage);
                if (isFirstUser)
                {
                    var contentArr = new JArray();
                    contentArr.Add(new JObject { ["type"] = "text", ["text"] = content });
                    contentArr.Add(new JObject
                    {
                        ["type"] = "image",
                        ["source"] = new JObject
                        {
                            ["type"] = "base64",
                            ["media_type"] = "image/jpeg",
                            ["data"] = pendingImage
                        }
                    });
                    messages.Add(new JObject { ["role"] = role, ["content"] = contentArr });
                    DocumentAgent.PendingImageBase64 = null;
                }
                else
                {
                    messages.Add(new JObject { ["role"] = role, ["content"] = content });
                }
            }
            return messages;
        }

        private async Task<string> SendRequestAsync(JArray messages, JArray tools, double? temperature, CancellationToken cancellationToken = default)
        {
            string endpoint = NormalizeEndpoint();

            if (string.IsNullOrWhiteSpace(_apiKey))
                return ProviderRequestPolicy.NormalizeError("Claude", null, "API Key 为空，请输入有效的 API Key");

            using (var httpClient = ProviderRequestPolicy.CreateChatClient())
            {

                httpClient.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01");
                if (_authScheme == "x-api-key")
                    httpClient.DefaultRequestHeaders.Add("x-api-key", _apiKey);
                else
                    httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + _apiKey);

                var requestBody = new Dictionary<string, object>
                {
                    ["model"] = _modelName,
                    ["max_tokens"] = AppConfig.Instance.ChatGPTMaxToken,
                    ["messages"] = messages,
                    ["stream"] = false
                };

                if (temperature.HasValue)
                    requestBody["temperature"] = temperature.Value;

                ApplyReasoning(requestBody);

                if (tools != null && tools.Count > 0)
                    requestBody["tools"] = ConvertTools(tools);

                var jsonBody = new StringContent(
                    JsonConvert.SerializeObject(requestBody),
                    Encoding.UTF8,
                    "application/json");

                try
                {
                    var startTime = DateTime.Now;
                    var response = await httpClient.PostAsync(endpoint, jsonBody, cancellationToken);
                    string responseBody = await response.Content.ReadAsStringAsync();

                    if (!response.IsSuccessStatusCode)
                        return HandleError(response, responseBody);

                    var result = JObject.Parse(responseBody);
                    if (ProviderRequestPolicy.IsOutputTruncated(result))
                        return ProviderRequestPolicy.NormalizeError("Claude", null, Loc.T("I18n_ReviewOutputTruncated"));

                    // Extract thinking content
                    var contentArray = result["content"] as JArray;
                    if (contentArray != null)
                    {
                        foreach (var block in contentArray)
                        {
                            string type = block["type"]?.ToString();
                            if (type == "thinking")
                            {
                                string thinking = block["thinking"]?.ToString();
                                if (!string.IsNullOrEmpty(thinking))
                                    UpdateThinkingPane(thinking);
                            }
                        }
                    }

                    // Extract main text
                    string text = "";
                    JObject toolUse = null;
                    if (contentArray != null)
                    {
                        foreach (var block in contentArray)
                        {
                            if (block["type"]?.ToString() == "text")
                            {
                                text += block["text"]?.ToString() ?? "";
                            }
                            else if (block["type"]?.ToString() == "tool_use" && toolUse == null)
                            {
                                toolUse = block as JObject;
                            }
                        }
                    }

                    if (toolUse != null)
                    {
                        if (DocumentAgent.IsAgentRunning)
                        {
                            var calls = new JArray();
                            foreach (var block in contentArray)
                                if (block["type"]?.ToString() == "tool_use")
                                    calls.Add(AgentWireProtocol.Call(block["id"]?.ToString(), block["name"]?.ToString(), block["input"]));
                            return AgentWireProtocol.Encode("Claude", contentArray, calls);
                        }
                        string name = toolUse["name"]?.ToString() ?? "";
                        string arguments = (toolUse["input"] as JObject ?? new JObject())
                            .ToString(Formatting.None);
                        return $"[TOOL_CALL]\nname: {name}\narguments: {arguments}";
                    }

                    // Token usage stats
                    var usage = result["usage"];
                    if (usage != null)
                    {
                        double duration = (DateTime.Now - startTime).TotalSeconds;
                        string inTokens = usage["input_tokens"]?.ToString() ?? "?";
                        string outTokens = usage["output_tokens"]?.ToString() ?? "?";
                        string cacheCreate = usage["cache_creation_input_tokens"]?.ToString();
                        string cacheRead = usage["cache_read_input_tokens"]?.ToString();
                        string cacheStr = "";
                        if (!string.IsNullOrEmpty(cacheRead) && cacheRead != "0")
                            cacheStr += $" | CR: {cacheRead}";
                        if (!string.IsNullOrEmpty(cacheCreate) && cacheCreate != "0")
                            cacheStr += $" | CC: {cacheCreate}";

                        if (!DocumentAgent.IsAgentRunning)
                        {
                            Globals.Ribbons.Ribbon1.ShowNotification(
                                $"耗时 {duration:F2}s | P: {inTokens} | C: {outTokens}{cacheStr}",
                                1000,
                                System.Windows.Forms.ToolTipIcon.Info);
                        }
                    }

                    return text;
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Claude API 异常: {ex.Message}");
                    return ProviderRequestPolicy.NormalizeError("Claude", null, null, ex);
                }
            }
        }

        private static JArray ConvertTools(JArray tools)
        {
            var converted = new JArray();
            foreach (var tool in tools.OfType<JObject>())
            {
                var function = tool["function"] as JObject;
                if (function == null) continue;

                converted.Add(new JObject
                {
                    ["name"] = function["name"]?.ToString() ?? "",
                    ["description"] = function["description"]?.ToString() ?? "",
                    ["input_schema"] = function["parameters"]?.DeepClone() ?? new JObject
                    {
                        ["type"] = "object",
                        ["properties"] = new JObject()
                    }
                });
            }
            return converted;
        }

        private async Task SendStreamRequestAsync(string prompt, Action<string> processChunk)
        {
            string endpoint = NormalizeEndpoint();

            using (var httpClient = ProviderRequestPolicy.CreateChatClient())
            {

                httpClient.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01");
                if (_authScheme == "x-api-key")
                    httpClient.DefaultRequestHeaders.Add("x-api-key", _apiKey);
                else
                    httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + _apiKey);

                var requestBody = new Dictionary<string, object>
                {
                    ["model"] = _modelName,
                    ["max_tokens"] = AppConfig.Instance.ChatGPTMaxToken,
                    ["messages"] = BuildMessages(prompt, null),
                    ["stream"] = true
                };

                ApplyReasoning(requestBody);

                var jsonBody = new StringContent(
                    JsonConvert.SerializeObject(requestBody),
                    Encoding.UTF8,
                    "application/json");

                try
                {
                    var response = await httpClient.PostAsync(endpoint, jsonBody);
                    response.EnsureSuccessStatusCode();

                    using (var stream = await response.Content.ReadAsStreamAsync())
                    using (var reader = new StreamReader(stream))
                    {
                        string line;
                        while ((line = await reader.ReadLineAsync()) != null)
                        {
                            if (string.IsNullOrWhiteSpace(line)) continue;
                            if (!line.StartsWith("data: ")) continue;

                            var json = line.Substring(6).Trim();
                            try
                            {
                                var evt = JObject.Parse(json);
                                string type = evt["type"]?.ToString();

                                switch (type)
                                {
                                    case "content_block_delta":
                                        var delta = evt["delta"];
                                        if (delta?["type"]?.ToString() == "text_delta")
                                        {
                                            string text = delta["text"]?.ToString() ?? "";
                                            if (!string.IsNullOrEmpty(text))
                                                processChunk(text);
                                        }
                                        else if (delta?["type"]?.ToString() == "thinking_delta")
                                        {
                                            string thinking = delta["thinking"]?.ToString() ?? "";
                                            if (!string.IsNullOrEmpty(thinking))
                                                UpdateThinkingPane(thinking);
                                        }
                                        break;

                                    case "content_block_start":
                                        var block = evt["content_block"];
                                        if (block?["type"]?.ToString() == "thinking")
                                        {
                                            string thinking = block["thinking"]?.ToString() ?? "";
                                            if (!string.IsNullOrEmpty(thinking))
                                                UpdateThinkingPane(thinking);
                                        }
                                        break;
                                }
                            }
                            catch { }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Claude Stream Exception: {ex.Message}");
                    throw;
                }
            }
        }

        private void ApplyReasoning(Dictionary<string, object> body)
        {
            string effort = _reasoningEffort ?? AppConfig.Instance.ReasoningEffort;
            if (!_thinkingEnabled || effort == "default") return;
            if (effort == "none")
            {
                body["thinking"] = new { type = "disabled" };
                return;
            }
            if (effort != "low" && effort != "medium" && effort != "high" && effort != "max")
                throw new InvalidOperationException("Claude 不支持此强度，请选择 default、low、medium、high 或 max，并测试当前模型。");
            body.Remove("temperature");
            body["thinking"] = new { type = "adaptive" };
            body["output_config"] = new { effort };
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
            catch { }
        }

        private string HandleError(HttpResponseMessage response, string responseBody)
        {
            Debug.WriteLine($"Claude Error {(int)response.StatusCode} - {responseBody}");
            string errorMsg = "";
            try
            {
                var err = JObject.Parse(responseBody);
                errorMsg = err["error"]?["message"]?.ToString() ?? "";
            }
            catch { }

            if (string.IsNullOrWhiteSpace(errorMsg))
            {
                switch ((int)response.StatusCode)
                {
                    case 400: errorMsg = "请求格式错误"; break;
                    case 401: errorMsg = "API密钥无效"; break;
                    case 403: errorMsg = "没有权限访问"; break;
                    case 404: errorMsg = "地址错误，请检查 Endpoint"; break;
                    case 429: errorMsg = "请求过于频繁 (Rate Limited)"; break;
                    case 500: errorMsg = "服务器内部错误"; break;
                    default: errorMsg = $"未知错误 ({(int)response.StatusCode})"; break;
                }
            }
            return ProviderRequestPolicy.NormalizeError("Claude", response.StatusCode, errorMsg);
        }
    }
}
