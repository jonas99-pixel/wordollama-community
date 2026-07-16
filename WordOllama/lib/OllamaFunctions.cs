using WordOllama.lib;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Channels;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace WordOllama.lib
{
    public static class OllamaFunctions
    {
        private static readonly int _spanMinutes = 10;
        private static readonly string thinkPattern = @"]*?)</think>\n\n";

        /// <summary>格式化 API 错误响应，统一返回 [API 错误] 前缀供 Agent 识别。</summary>
        private static string FormatApiError(System.Net.HttpStatusCode statusCode, string responseBody)
        {
            string errorMsg = "";
            try
            {
                var errorJson = JObject.Parse(responseBody);
                errorMsg = errorJson["error"]?.ToString() ?? "";
            }
            catch
            {
                errorMsg = responseBody;
            }
            if (string.IsNullOrWhiteSpace(errorMsg)) errorMsg = ((int)statusCode).ToString();
            if (errorMsg.Length > 500) errorMsg = errorMsg.Substring(0, 500) + "...";
            return ProviderRequestPolicy.NormalizeError("Ollama", statusCode, errorMsg);
        }
        public static async Task<string> OllamaGenerateAsync(string systemPrompt, string image = null)
        {
            string ollamaEndPoint = AppConfig.Instance.OllamaEndPoint + "/api/generate";

            using (HttpClient client = new HttpClient())
            {
                client.Timeout = TimeSpan.FromMinutes(_spanMinutes); // 超时设置
                var requestData = new Dictionary<string, object>
        {
            { "model", AppConfig.Instance.OllamaModel },
            { "prompt", systemPrompt },
            { "stream", false },
            { "options", new Dictionary<string, object>
                {
                    { "temperature", AppConfig.Instance.OllamaTemperature },
                    { "num_ctx", AppConfig.Instance.OllamaMaxToken }
                }
            }
        };

                // 图片支持：优先用 OpenAI 兼容格式（Ollama 0.3+），否则用旧版 images 字段
                if (!string.IsNullOrEmpty(image))
                {
                    try
                    {
                        // 检测 Ollama 版本是否支持 OpenAI 兼容格式
                        requestData["images"] = new[] { image };
                    }
                    catch { }
                }

                var postData = new StringContent(
                    JsonConvert.SerializeObject(requestData),
                    Encoding.UTF8,
                    "application/json"
                );


                try
                {
                    HttpResponseMessage response = await client.PostAsync(ollamaEndPoint, postData);

                    if (response.IsSuccessStatusCode)
                    {
                        string jsonResponse = await response.Content.ReadAsStringAsync();
                        JObject responseObject = JObject.Parse(jsonResponse);
                        string responseText = (string)responseObject["response"];



                        // 检查是否存在 reasoning_content 且不为空
                        try
                        {
                            MatchCollection matches = Regex.Matches(responseText, thinkPattern);
                            string thinkContent = string.Empty;
                            foreach (Match match in matches)
                            {
                                thinkContent = match.Groups[1].Value; 
                            }
                            responseText = Regex.Replace(responseText, thinkPattern, string.Empty, RegexOptions.Singleline);
                            string reasoningContent = (string)responseObject["reasoning_content"];
                            if (!string.IsNullOrEmpty(reasoningContent) || !string.IsNullOrEmpty(thinkContent))
                            {
                                try
                                {
                                    var pane = Globals.ThisAddIn?.thinkingPane;
                                    if (pane != null && pane.IsHandleCreated)
                                    {
                                        string combined = thinkContent + reasoningContent;
                                        if (pane.InvokeRequired)
                                            pane.Invoke(new Action(() => pane.UpdateThinkText(combined)));
                                        else
                                            pane.UpdateThinkText(combined);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine($"更新思考窗格失败: {ex.Message}");
                                }
                            }
                        }
                        catch (Exception)
                        {
                            // 属性不存在,继续使用原始内容
                        }

                        long total_duration = (long)responseObject["total_duration"];
                        // 将纳秒转换为秒（1秒 = 1,000,000,000纳秒）
                        double total_duration_s = total_duration / 1_000_000_000.0;
                        // 将秒数格式化为字符串
                        string total_duration_s_str = total_duration_s.ToString("F2");
                        string prompt_eval_count = (string)responseObject["prompt_eval_count"];
                        string eval_count = (string)responseObject["eval_count"];
                        if (!DocumentAgent.IsAgentRunning)
                        {
                            Globals.Ribbons.Ribbon1.ShowNotification($"生成完成！总耗时 {total_duration_s_str} 秒，提示词共 {prompt_eval_count} Tokens，结果共 {eval_count} Tokens", 3000, ToolTipIcon.Info);
                        }
                        return responseText;
                    }
                    else
                    {
                        string errorDetail = await response.Content.ReadAsStringAsync();
                        return FormatApiError(response.StatusCode, errorDetail);
                    }
                }
                catch (Exception e)
                {
                    return ProviderRequestPolicy.NormalizeError("Ollama", null, null, e);
                }}
            
        }

        public static async IAsyncEnumerable<string> OllamaStreamAsync(string systemPrompt, Action<string> processChunk, [EnumeratorCancellation] CancellationToken cancellationToken)
        {
            string ollamaEndPoint = AppConfig.Instance.OllamaEndPoint + "/api/generate";

            using (HttpClient client = new HttpClient())
            {
                var postData = new StringContent(
                    JsonConvert.SerializeObject(new
                    {
                        model = AppConfig.Instance.OllamaModel,
                        prompt = systemPrompt,
                        stream = true,
                        options = new Dictionary<string, object>
                        {
                    { "temperature", AppConfig.Instance.OllamaTemperature },
                    { "num_ctx", AppConfig.Instance.OllamaMaxToken }
                        }
                    }),
                    Encoding.UTF8,
                    "application/json"
                );

                var channel = Channel.CreateUnbounded<string>();

                // 异步线程运行POST请求
                _ = Task.Run(async () =>
                {
                    try
                    {
                        Debug.WriteLine("Posting to Ollama API!");
                        HttpResponseMessage response = await client.PostAsync(ollamaEndPoint, postData, cancellationToken);
                        response.EnsureSuccessStatusCode();

                        using (var responseStream = await response.Content.ReadAsStreamAsync())
                        using (var streamReader = new StreamReader(responseStream))
                        {
                            while (!streamReader.EndOfStream)
                            {
                                cancellationToken.ThrowIfCancellationRequested(); // 支持取消操作

                                var line = await streamReader.ReadLineAsync();
                                Debug.WriteLine(line);

                                var responseLine = JObject.Parse(line);
                                if (responseLine["done"]?.ToObject<bool>() != true)
                                {
                                    var chunk = responseLine["response"]?.ToString();
                                    if (!string.IsNullOrEmpty(chunk))
                                    {
                                        processChunk(chunk); // 处理数据
                                        await channel.Writer.WriteAsync(chunk, cancellationToken); // 写入通道
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error: {ex.Message}");
                        await channel.Writer.WriteAsync($"[API 错误] Ollama 流式请求失败: {ex.Message}", cancellationToken);
                    }
                    finally
                    {
                        channel.Writer.Complete();
                    }
                }, cancellationToken);

                // 通过通道返回数据
                await foreach (var chunk in channel.Reader.ReadAllAsync(cancellationToken))
                {
                    yield return chunk;
                }
            }
        }


        public static async Task<string> OllamaChatAsync(List<Dictionary<string, string>> ChatHistory, JArray tools = null, CancellationToken cancellationToken = default)
        {
            string ollamaEndPoint = AppConfig.Instance.OllamaEndPoint + "/api/chat";
            using (HttpClient client = new HttpClient())
            {
                ProviderRequestPolicy.ConfigureChatClient(client);

                // 构建请求体 —— 支持可选的 tools 参数
                var requestBodyObj = new Dictionary<string, object>
                {
                    ["model"] = AppConfig.Instance.OllamaModel,
                    ["messages"] = ChatHistory,
                    ["stream"] = false,
                    ["options"] = new Dictionary<string, double> { { "temperature", AppConfig.Instance.ChatGPTTemperature } }
                };

                // 图片支持：若有待发送图片，将首条 user 消息转为多模态格式
                // Ollama 格式：{"role":"user","content":"text","images":["base64"]}
                string pendingImage = DocumentAgent.PendingImageBase64;
                string requestPendingImage = pendingImage; // 为降级重试保留副本
                if (!string.IsNullOrEmpty(pendingImage))
                {
                    var multiModalMessages = new List<object>();
                    bool firstUserFound = false;
                    foreach (var msg in ChatHistory)
                    {
                        if (!firstUserFound && msg.TryGetValue("role", out var role) && role == "user")
                        {
                            firstUserFound = true;
                            var multiMsg = new Dictionary<string, object>
                            {
                                ["role"] = "user",
                                ["content"] = msg["content"],
                                ["images"] = new List<string> { pendingImage }
                            };
                            multiModalMessages.Add(multiMsg);
                        }
                        else
                        {
                            multiModalMessages.Add(msg.ToDictionary(kvp => kvp.Key, kvp => (object)kvp.Value));
                        }
                    }
                    requestBodyObj["messages"] = multiModalMessages;
                    DocumentAgent.PendingImageBase64 = null; // 首轮消费后清理
                }
                if (tools != null && tools.Count > 0)
                    requestBodyObj["tools"] = tools;

                var postData = new StringContent(
                    JsonConvert.SerializeObject(requestBodyObj),
                    Encoding.UTF8,
                    "application/json"
                );

                try
                {
                    HttpResponseMessage response = await client.PostAsync(ollamaEndPoint, postData, cancellationToken);

                    if (response.IsSuccessStatusCode)
                    {
                        string jsonResponse = await response.Content.ReadAsStringAsync();
                        JObject responseObject = JObject.Parse(jsonResponse);

                        // ===== 原生 Tool Calling 检测 =====
                        var toolCalls = responseObject["message"]?["tool_calls"];
                        if (toolCalls != null && toolCalls.HasValues)
                        {
                            // 只取第一个 tool_call（当前 ReAct 每轮只执行一个工具）
                            var firstCall = toolCalls[0];
                            string funcName = firstCall["function"]?["name"]?.ToString() ?? "";
                            string funcArgs = firstCall["function"]?["arguments"]?.ToString() ?? "{}";
                            return $"[TOOL_CALL]\nname: {funcName}\narguments: {funcArgs}";
                        }

                        string responseText = (string)responseObject["message"]["content"];

                        // 检查是否存在 reasoning_content 且不为空
                        try
                        {
                            MatchCollection matches = Regex.Matches(responseText, thinkPattern);
                            string thinkContent = string.Empty;
                            foreach (Match match in matches)
                            {
                                thinkContent = match.Groups[1].Value;
                            }
                            responseText = Regex.Replace(responseText, thinkPattern, string.Empty, RegexOptions.Singleline);
                            string reasoningContent = (string)responseObject["reasoning_content"];
                            if (!string.IsNullOrEmpty(reasoningContent) || !string.IsNullOrEmpty(thinkContent))
                            {
                                try
                                {
                                    var pane = Globals.ThisAddIn?.thinkingPane;
                                    if (pane != null && pane.IsHandleCreated)
                                    {
                                        string combined = thinkContent + reasoningContent;
                                        if (pane.InvokeRequired)
                                            pane.Invoke(new Action(() => pane.UpdateThinkText(combined)));
                                        else
                                            pane.UpdateThinkText(combined);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine($"更新思考窗格失败: {ex.Message}");
                                }
                            }
                        }
                        catch (Exception)
                        {
                            // 属性不存在,继续使用原始内容
                        }

                        return responseText;
                    }
                    else
                    {
                        string errorDetail = await response.Content.ReadAsStringAsync();
                        // 图片导致 API 拒绝（模型不支持视觉），降级重试
                        if (!string.IsNullOrEmpty(requestPendingImage))
                        {
                            DocumentAgent.PendingImageBase64 = null;
                            return await OllamaChatAsync(ChatHistory, tools, cancellationToken);
                        }
                        // 如果是因为 tools 参数不被支持，回退到无 tools 的调用
                        if (errorDetail.Contains("tools") || errorDetail.Contains("function"))
                            return "[TOOL_CALL_NOT_SUPPORTED]";
                        return FormatApiError(response.StatusCode, errorDetail);
                    }
                }
                catch (Exception) when (!string.IsNullOrEmpty(requestPendingImage))
                {
                    // 图片导致请求失败（模型不支持视觉），降级重试
                    DocumentAgent.PendingImageBase64 = null;
                    return await OllamaChatAsync(ChatHistory, tools, cancellationToken);
                }
                catch (Exception e)
                {
                    return AgentToolText.F("I18n_AgentResultOllamaConnectionFailed", e.Message);
                }

            }
        }
    }
}
