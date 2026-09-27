using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace WordOllama.lib
{
    public sealed class CapabilityProbeStep
    {
        public string Name { get; set; }
        public bool Passed { get; set; }
        public string Detail { get; set; }
    }

    public sealed class CapabilityProbeResult
    {
        public List<CapabilityProbeStep> Steps { get; } = new List<CapabilityProbeStep>();
        public long LatencyMs { get; set; }
        public int Passed => Steps.Count(step => step.Passed);
        public int Total => Steps.Count;
        public bool Success => Steps.Any(step => step.Name == "普通聊天" && step.Passed);
        public string Summary => string.Join(Environment.NewLine, Steps.Select(step => step.Name + ": " + (step.Passed ? "OK" : step.Detail)));
    }

    public static class ProviderCapabilityProbe
    {
        public static async Task<CapabilityProbeResult> RunAsync(IOnlineProvider provider)
        {
            if (provider == null) throw new ArgumentNullException(nameof(provider));
            var result = new CapabilityProbeResult();
            var watch = Stopwatch.StartNew();
            await RunStep(result, "模型列表", async () =>
            {
                List<string> models = await provider.FetchModelsAsync();
                return models != null ? "models=" + models.Count : "no response";
            });
            await RunStep(result, "普通聊天", async () =>
            {
                string text = await provider.GenerateAsync("Reply with exactly: OK", null, null);
                ThrowIfProviderError(text);
                if (string.IsNullOrWhiteSpace(text)) throw new InvalidOperationException("empty response");
                return text.Trim();
            });
            await RunStep(result, "流式输出", async () =>
            {
                int chunks = 0;
                await provider.StreamAsync("Reply with exactly: OK", chunk => { if (!string.IsNullOrEmpty(chunk)) chunks++; });
                if (chunks == 0) throw new InvalidOperationException("no stream chunks");
                return "chunks=" + chunks;
            });
            await RunStep(result, "工具调用", async () =>
            {
                var tools = new JArray(new JObject
                {
                    ["type"] = "function",
                    ["function"] = new JObject
                    {
                        ["name"] = "probe_echo",
                        ["description"] = "Call this tool with value OK.",
                        ["parameters"] = new JObject
                        {
                            ["type"] = "object",
                            ["properties"] = new JObject
                            {
                                ["value"] = new JObject
                                {
                                    ["type"] = "string",
                                    ["description"] = "The exact capability probe value."
                                }
                            },
                            ["required"] = new JArray("value"),
                            ["additionalProperties"] = false
                        }
                    }
                });
                string response = await provider.ChatAsync(new List<Dictionary<string, string>>
                {
                    new Dictionary<string, string>
                    {
                        ["role"] = "system",
                        ["content"] = "This is a native function-calling capability test. You must call probe_echo. Do not answer with text."
                    },
                    new Dictionary<string, string>
                    {
                        ["role"] = "user",
                        ["content"] = "Call probe_echo now with value exactly OK."
                    }
                }, tools);
                ThrowIfProviderError(response);
                if (string.IsNullOrWhiteSpace(response) || response.IndexOf("probe_echo", StringComparison.OrdinalIgnoreCase) < 0)
                    throw new InvalidOperationException("未收到原生工具调用；可尝试自动检测或 ReAct 模式。");
                return "native tool call observed";
            });
            watch.Stop();
            result.LatencyMs = watch.ElapsedMilliseconds;
            return result;
        }

        private static void ThrowIfProviderError(string response)
        {
            if (string.IsNullOrWhiteSpace(response)) return;

            string trimmed = response.TrimStart();
            if (trimmed.StartsWith("[API 错误]", StringComparison.OrdinalIgnoreCase)
                || trimmed.StartsWith("[API 拒绝]", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(trimmed);
            }
        }

        private static async Task RunStep(CapabilityProbeResult result, string name, Func<Task<string>> action)
        {
            try
            {
                string detail = await action();
                result.Steps.Add(new CapabilityProbeStep { Name = name, Passed = true, Detail = detail });
            }
            catch (Exception ex)
            {
                result.Steps.Add(new CapabilityProbeStep { Name = name, Passed = false, Detail = ex.Message });
            }
        }
    }
}
