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
    public class GeminiProvider : IOnlineProvider, ICancellableOnlineProvider
    {
        public string ProviderType => "Gemini";

        private readonly string _endpoint;
        private readonly string _apiKey;
        private readonly string _modelName;

        private const string DefaultBase = "https://generativelanguage.googleapis.com/v1beta";

        public GeminiProvider(string endpoint, string apiKey, string modelName)
        {
            _endpoint = endpoint;
            _apiKey = apiKey;
            _modelName = modelName;
        }

        public async Task<string> GenerateAsync(string prompt, string imageBase64 = null, double? temperature = null)
        {
            var parts = BuildParts(prompt, imageBase64);
            return await SendGenerateContentAsync(parts, null, temperature, null);
        }

        public async Task<string> ChatAsync(List<Dictionary<string, string>> chatHistory, JArray tools = null)
        {
            var contents = ConvertHistory(chatHistory);
            return await SendGenerateContentAsync(null, contents, null, tools);
        }

        public async Task<string> ChatAsync(List<Dictionary<string, string>> chatHistory, JArray tools, CancellationToken cancellationToken)
        {
            var contents = ConvertHistory(chatHistory);
            return await SendGenerateContentAsync(null, contents, null, tools, cancellationToken);
        }

        public async Task StreamAsync(string prompt, Action<string> processChunk)
        {
            await SendStreamRequestAsync(prompt, processChunk);
        }

        public async Task<List<string>> FetchModelsAsync()
        {
            var models = new List<string>();

            // Try API call if key is provided
            if (!string.IsNullOrWhiteSpace(_apiKey))
            {
                try
                {
                    string baseUrl = GetBaseUrl();
                    using (var httpClient = new HttpClient())
                    {
                        ProviderRequestPolicy.ConfigureConnectClient(httpClient);
                        string url = $"{baseUrl}/models?key={_apiKey}";
                        var response = await httpClient.GetAsync(url);
                        if (response.IsSuccessStatusCode)
                        {
                            string body = await response.Content.ReadAsStringAsync();
                            var json = JObject.Parse(body);
                            var data = json["models"] as JArray;
                            if (data != null)
                            {
                                foreach (var m in data)
                                {
                                    string name = m["name"]?.ToString();
                                    if (!string.IsNullOrEmpty(name))
                                    {
                                        if (name.StartsWith("models/"))
                                            name = name.Substring(7);
                                        var methods = m["supportedGenerationMethods"] as JArray;
                                        if (methods != null)
                                        {
                                            foreach (var method in methods)
                                            {
                                                if (method.ToString().Contains("generateContent"))
                                                {
                                                    models.Add(name);
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Gemini FetchModels error: {ex.Message}");
                }
            }

            return models;
        }

        private string GetBaseUrl()
        {
            if (string.IsNullOrEmpty(_endpoint))
                return DefaultBase;
            string url = _endpoint.TrimEnd('/');
            // If user provided a full path like /v1beta/models, strip to base
            if (url.Contains("/models"))
                url = url.Substring(0, url.IndexOf("/models", StringComparison.OrdinalIgnoreCase));
            return url;
        }

        private string GetGenerateUrl()
        {
            string baseUrl = GetBaseUrl();
            string model = string.IsNullOrEmpty(_modelName) ? "gemini-3-flash-preview" : _modelName;
            return $"{baseUrl}/models/{model}:generateContent";
        }

        private string GetStreamUrl()
        {
            string baseUrl = GetBaseUrl();
            string model = string.IsNullOrEmpty(_modelName) ? "gemini-3-flash-preview" : _modelName;
            return $"{baseUrl}/models/{model}:streamGenerateContent?alt=sse";
        }

        private JArray BuildParts(string prompt, string imageBase64)
        {
            var parts = new JArray();
            if (!string.IsNullOrEmpty(imageBase64))
            {
                parts.Add(new JObject
                {
                    ["inline_data"] = new JObject
                    {
                        ["mime_type"] = "image/jpeg",
                        ["data"] = imageBase64
                    }
                });
            }
            parts.Add(new JObject { ["text"] = prompt ?? "" });
            return parts;
        }

        private JArray ConvertHistory(List<Dictionary<string, string>> chatHistory)
        {
            var contents = new JArray();
            string pendingImage = DocumentAgent.PendingImageBase64;
            bool imageInjected = false;

            foreach (var msg in chatHistory)
            {
                string role = msg.ContainsKey("role") ? msg["role"] : "user";
                string content = msg.ContainsKey("content") ? msg["content"] : "";

                // Map "assistant" -> "model" for Gemini
                string geminiRole = role == "assistant" ? "model" : role;

                var parts = new JArray();
                bool isFirstUser = !imageInjected && geminiRole == "user" && !string.IsNullOrEmpty(pendingImage);
                if (isFirstUser)
                {
                    parts.Add(new JObject
                    {
                        ["inline_data"] = new JObject
                        {
                            ["mime_type"] = "image/jpeg",
                            ["data"] = pendingImage
                        }
                    });
                    imageInjected = true;
                    DocumentAgent.PendingImageBase64 = null;
                }
                parts.Add(new JObject { ["text"] = content });

                contents.Add(new JObject
                {
                    ["role"] = geminiRole,
                    ["parts"] = parts
                });
            }

            return contents;
        }

        private async Task<string> SendGenerateContentAsync(JArray parts, JArray contents, double? temperature, JArray tools, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(_apiKey))
                return "API Key 为空，请输入有效的 API Key";

            string url = GetGenerateUrl() + "?key=" + Uri.EscapeDataString(_apiKey);

            using (var httpClient = new HttpClient())
            {
                ProviderRequestPolicy.ConfigureChatClient(httpClient);
                ServicePointManager.SecurityProtocol =
                    SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

                var requestData = new Dictionary<string, object>();

                if (contents != null)
                {
                    requestData["contents"] = contents;
                }
                else
                {
                    requestData["contents"] = new JArray
                    {
                        new JObject
                        {
                            ["role"] = "user",
                            ["parts"] = parts
                        }
                    };
                }

                // System instruction from language mode
                string systemInstruction = GetSystemInstruction();
                if (!string.IsNullOrEmpty(systemInstruction))
                {
                    requestData["system_instruction"] = new JObject
                    {
                        ["parts"] = new JArray { new JObject { ["text"] = systemInstruction } }
                    };
                }

                // Generation config
                var genConfig = new Dictionary<string, object>();
                if (temperature.HasValue)
                    genConfig["temperature"] = temperature.Value;

                if (AppConfig.Instance.ThinkingEnabled)
                {
                    genConfig["thinkingConfig"] = new Dictionary<string, object>
                    {
                        ["thinkingLevel"] = MapReasoningEffort()
                    };
                }

                if (genConfig.Count > 0)
                    requestData["generationConfig"] = genConfig;

                if (tools != null && tools.Count > 0)
                    requestData["tools"] = ConvertTools(tools);

                var jsonBody = new StringContent(
                    JsonConvert.SerializeObject(requestData),
                    Encoding.UTF8,
                    "application/json");

                try
                {
                    var startTime = DateTime.Now;
                    var response = await httpClient.PostAsync(url, jsonBody, cancellationToken);
                    string responseBody = await response.Content.ReadAsStringAsync();

                    if (!response.IsSuccessStatusCode)
                        return HandleError(response, responseBody);

                    var result = JObject.Parse(responseBody);

                    // Extract thinking from candidates
                    var candidates = result["candidates"] as JArray;
                    if (candidates != null && candidates.Count > 0)
                    {
                        var first = candidates[0];
                        var finishReason = first["finishReason"]?.ToString();
                        if (finishReason == "SAFETY")
                            return "[API 拒绝] 内容被安全过滤";

                        // Check for thinking in avgLogprobs or groundingMetadata
                        // Gemini doesn't expose thinking text in the REST API easily
                    }

                    // Extract text or native function call
                    string text = "";
                    try
                    {
                        var responseParts = result["candidates"]?[0]?["content"]?["parts"] as JArray;
                        if (responseParts != null)
                        {
                            foreach (var part in responseParts.OfType<JObject>())
                            {
                                if (part["functionCall"] is JObject functionCall)
                                {
                                    string name = functionCall["name"]?.ToString() ?? "";
                                    string arguments = (functionCall["args"] as JObject ?? new JObject())
                                        .ToString(Newtonsoft.Json.Formatting.None);
                                    return $"[TOOL_CALL]\nname: {name}\narguments: {arguments}";
                                }
                                text += part["text"]?.ToString() ?? "";
                            }
                        }
                    }
                    catch { }

                    // Usage stats (Gemini v1beta returns promptFeedback)
                    var usageMetadata = result["usageMetadata"];
                    if (usageMetadata != null)
                    {
                        double duration = (DateTime.Now - startTime).TotalSeconds;
                        string promptTokens = usageMetadata["promptTokenCount"]?.ToString() ?? "?";
                        string candidatesTokens = usageMetadata["candidatesTokenCount"]?.ToString() ?? "?";
                        string totalTokens = usageMetadata["totalTokenCount"]?.ToString() ?? "?";

                        if (!DocumentAgent.IsAgentRunning)
                        {
                            Globals.Ribbons.Ribbon1.ShowNotification(
                                $"耗时 {duration:F2}s | P: {promptTokens} | C: {candidatesTokens} | T: {totalTokens}",
                                1000,
                                System.Windows.Forms.ToolTipIcon.Info);
                        }
                    }

                    return text;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Gemini API 异常: {ex.Message}");
                    return ProviderRequestPolicy.NormalizeError("Gemini", null, null, ex);
                }
            }
        }

        private static JArray ConvertTools(JArray tools)
        {
            var declarations = new JArray();
            foreach (var tool in tools.OfType<JObject>())
            {
                var function = tool["function"] as JObject;
                if (function == null) continue;

                declarations.Add(new JObject
                {
                    ["name"] = function["name"]?.ToString() ?? "",
                    ["description"] = function["description"]?.ToString() ?? "",
                    ["parameters"] = function["parameters"]?.DeepClone() ?? new JObject
                    {
                        ["type"] = "object",
                        ["properties"] = new JObject()
                    }
                });
            }

            return new JArray
            {
                new JObject { ["functionDeclarations"] = declarations }
            };
        }

        private async Task SendStreamRequestAsync(string prompt, Action<string> processChunk)
        {
            if (string.IsNullOrWhiteSpace(_apiKey)) return;

            string url = GetStreamUrl() + "&key=" + Uri.EscapeDataString(_apiKey);

            using (var httpClient = new HttpClient())
            {
                ServicePointManager.SecurityProtocol =
                    SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

                var requestData = new Dictionary<string, object>
                {
                    ["contents"] = new JArray
                    {
                        new JObject
                        {
                            ["role"] = "user",
                            ["parts"] = new JArray { new JObject { ["text"] = prompt } }
                        }
                    }
                };

                string systemInstruction = GetSystemInstruction();
                if (!string.IsNullOrEmpty(systemInstruction))
                {
                    requestData["system_instruction"] = new JObject
                    {
                        ["parts"] = new JArray { new JObject { ["text"] = systemInstruction } }
                    };
                }

                var jsonBody = new StringContent(
                    JsonConvert.SerializeObject(requestData),
                    Encoding.UTF8,
                    "application/json");

                try
                {
                    var response = await httpClient.PostAsync(url, jsonBody);
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
                            if (json == "[DONE]") break;

                            try
                            {
                                var result = JObject.Parse(json);
                                var candidates = result["candidates"] as JArray;
                                if (candidates == null || candidates.Count == 0) continue;

                                var parts = candidates[0]?["content"]?["parts"] as JArray;
                                if (parts == null || parts.Count == 0) continue;

                                string text = parts[0]?["text"]?.ToString() ?? "";
                                if (!string.IsNullOrEmpty(text))
                                    processChunk(text);
                            }
                            catch { }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Gemini Stream Exception: {ex.Message}");
                    throw;
                }
            }
        }

        private string GetSystemInstruction()
        {
            if (AppConfig.Instance == null) return "";
            string lang = AppConfig.Instance.LanguageMode;
            if (lang == "Chinese" || lang == "English") return PromptTemplateService.BuildReplyLanguagePrefix(lang).Trim();
            return "";
        }

        private string MapReasoningEffort()
        {
            // Map OpenAI reasoning_effort values to Gemini thinking levels
            return AppConfig.Instance.ReasoningEffort switch
            {
                "low" => "low",
                "medium" => "medium",
                "high" => "high",
                "xhigh" => "high",
                "max" => "high",
                _ => "low"
            };
        }

        private string HandleError(HttpResponseMessage response, string responseBody)
        {
            Debug.WriteLine($"Gemini Error {(int)response.StatusCode} - {responseBody}");
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
                    case 401:
                    case 403: errorMsg = "API密钥无效或没有权限"; break;
                    case 404: errorMsg = "地址错误，请检查 Endpoint"; break;
                    case 429: errorMsg = "请求过于频繁 (Quota Exceeded)"; break;
                    case 500: errorMsg = "服务器内部错误"; break;
                    default: errorMsg = $"未知错误 ({(int)response.StatusCode})"; break;
                }
            }
            return ProviderRequestPolicy.NormalizeError("Gemini", response.StatusCode, errorMsg);
        }
    }
}
