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
        public bool Success => Total > 0 && Passed == Total;
        public string Summary => string.Join("; ", Steps.Select(step => step.Name + ": " + (step.Passed ? "OK" : step.Detail)));
    }

    public static class ProviderCapabilityProbe
    {
        public static async Task<CapabilityProbeResult> RunAsync(IOnlineProvider provider)
        {
            if (provider == null) throw new ArgumentNullException(nameof(provider));
            var result = new CapabilityProbeResult();
            var watch = Stopwatch.StartNew();
            await RunStep(result, "model-list", async () =>
            {
                List<string> models = await provider.FetchModelsAsync();
                return models != null ? "models=" + models.Count : "no response";
            });
            await RunStep(result, "chat", async () =>
            {
                string text = await provider.GenerateAsync("Reply with exactly: OK", null, 0.0);
                if (string.IsNullOrWhiteSpace(text)) throw new InvalidOperationException("empty response");
                return text.Trim();
            });
            await RunStep(result, "streaming", async () =>
            {
                int chunks = 0;
                await provider.StreamAsync("Reply with exactly: OK", chunk => { if (!string.IsNullOrEmpty(chunk)) chunks++; });
                if (chunks == 0) throw new InvalidOperationException("no stream chunks");
                return "chunks=" + chunks;
            });
            await RunStep(result, "tool-calling", async () =>
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
                            ["properties"] = new JObject { ["value"] = new JObject { ["type"] = "string" } },
                            ["required"] = new JArray("value")
                        }
                    }
                });
                string response = await provider.ChatAsync(new List<Dictionary<string, string>>
                {
                    new Dictionary<string, string> { ["role"] = "user", ["content"] = "Use probe_echo now with value OK." }
                }, tools);
                if (string.IsNullOrWhiteSpace(response) || response.IndexOf("probe_echo", StringComparison.OrdinalIgnoreCase) < 0)
                    throw new InvalidOperationException("native tool call was not observed");
                return "native tool call observed";
            });
            watch.Stop();
            result.LatencyMs = watch.ElapsedMilliseconds;
            return result;
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
