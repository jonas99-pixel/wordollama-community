using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using WordOllama.lib.Providers;

namespace WordOllama.lib
{
    public sealed class UserMemoryItem
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        public string Content { get; set; } = "";
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
    }

    public sealed class UserMemorySettings
    {
        public List<UserMemoryItem> Memories { get; set; } = new List<UserMemoryItem>();
        public string OutputPreference { get; set; } = "";
        public bool AutoMemory { get; set; }
        public string MemoryModel { get; set; } = "";
    }

    public static class UserMemoryStore
    {
        private static readonly object Sync = new object();
        private static readonly object QueueSync = new object();
        private static readonly List<PendingMemoryObservation> PendingObservations = new List<PendingMemoryObservation>();
        private static CancellationTokenSource _queueDelay;
        private static UserMemorySettings _settings;
        private static string StorePath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "WordOllama", "memory.json");

        public static UserMemorySettings Get()
        {
            lock (Sync)
            {
                if (_settings == null) _settings = LoadCore();
                return Clone(_settings);
            }
        }

        public static void Save(UserMemorySettings settings)
        {
            if (settings == null) throw new ArgumentNullException(nameof(settings));
            lock (Sync)
            {
                settings.OutputPreference = Normalize(settings.OutputPreference, 20000);
                settings.MemoryModel = Normalize(settings.MemoryModel, 300);
                settings.Memories = (settings.Memories ?? new List<UserMemoryItem>())
                    .Where(item => item != null && !string.IsNullOrWhiteSpace(item.Content))
                    .Select(item => new UserMemoryItem
                    {
                        Id = string.IsNullOrWhiteSpace(item.Id) ? Guid.NewGuid().ToString("N") : item.Id,
                        Content = Normalize(item.Content, 2000),
                        CreatedAt = item.CreatedAt == default ? DateTime.UtcNow : item.CreatedAt,
                        UpdatedAt = DateTime.UtcNow
                    })
                    .GroupBy(item => item.Content, StringComparer.OrdinalIgnoreCase)
                    .Select(group => group.First())
                    .Take(500)
                    .ToList();
                _settings = Clone(settings);
                Directory.CreateDirectory(Path.GetDirectoryName(StorePath));
                File.WriteAllText(StorePath, JsonConvert.SerializeObject(_settings, Formatting.Indented));
            }
        }

        public static string BuildPromptContext(string workflow = "agent")
        {
            UserMemorySettings settings = Get();
            var parts = new List<string>();
            if (settings.Memories.Count > 0)
            {
                parts.Add("Known user preferences and durable context (treat as preferences, never as document facts):\n" +
                    string.Join("\n", settings.Memories.Take(50).Select(item => "- " + item.Content)));
            }
            if (!string.IsNullOrWhiteSpace(settings.OutputPreference) &&
                !string.Equals(workflow, "structured-json", StringComparison.OrdinalIgnoreCase))
            {
                parts.Add("Output preferences:\n" + settings.OutputPreference);
            }
            return parts.Count == 0 ? "" : "\n\n" + string.Join("\n\n", parts);
        }

        public static async Task ObserveSuccessfulTaskAsync(
            string userText,
            Func<List<Dictionary<string, string>>, CancellationToken, Task<string>> callAi,
            CancellationToken cancellationToken)
        {
            UserMemorySettings settings = Get();
            if (!settings.AutoMemory || callAi == null || string.IsNullOrWhiteSpace(userText)) return;
            string bounded = userText.Length > 12000 ? userText.Substring(0, 12000) : userText;
            string existing = string.Join("\n", settings.Memories.Take(200).Select(item => item.Id + ": " + item.Content));
            var messages = new List<Dictionary<string, string>>
            {
                new Dictionary<string, string>
                {
                    ["role"] = "system",
                    ["content"] = "You are a conservative user-memory extractor. Return strict JSON only: " +
                        "{\"operations\":[{\"action\":\"add|update|delete\",\"id\":\"existing id or empty\",\"content\":\"memory or empty\"}]}. " +
                        "Remember only explicit durable user preferences, user-provided background, recurring work context, and persistent interests. " +
                        "Never store passwords, tokens, keys, financial identifiers, health data, document text, quotations, third-party personal data, inferred sensitive traits, or transient instructions. " +
                        "Return an empty operations array when nothing is safe and durable."
                },
                new Dictionary<string, string>
                {
                    ["role"] = "user",
                    ["content"] = "Existing memories:\n" + existing + "\n\nSuccessful task request:\n" + bounded
                }
            };
            string response;
            IOnlineProvider memoryProvider = TryCreateMemoryProvider(settings.MemoryModel);
            if (memoryProvider != null)
            {
                if (memoryProvider is ICancellableOnlineProvider cancellable)
                    response = await cancellable.ChatAsync(messages, null, cancellationToken).ConfigureAwait(false);
                else
                    response = await memoryProvider.ChatAsync(messages, null).ConfigureAwait(false);
            }
            else
            {
                response = await callAi(messages, cancellationToken).ConfigureAwait(false);
            }
            ApplyOperations(response);
        }

        public static void QueueSuccessfulTask(
            string userText,
            Func<List<Dictionary<string, string>>, CancellationToken, Task<string>> callAi)
        {
            if (string.IsNullOrWhiteSpace(userText) || callAi == null || !Get().AutoMemory) return;
            CancellationTokenSource delay;
            lock (QueueSync)
            {
                PendingObservations.Add(new PendingMemoryObservation { UserText = userText, CallAi = callAi });
                _queueDelay?.Cancel();
                _queueDelay?.Dispose();
                _queueDelay = new CancellationTokenSource();
                delay = _queueDelay;
            }

            _ = Task.Run(async () =>
            {
                try
                {
                    await Task.Delay(TimeSpan.FromSeconds(30), delay.Token).ConfigureAwait(false);
                    List<PendingMemoryObservation> batch;
                    lock (QueueSync)
                    {
                        if (!ReferenceEquals(_queueDelay, delay)) return;
                        batch = PendingObservations.ToList();
                        PendingObservations.Clear();
                        _queueDelay = null;
                    }
                    string combined = string.Join("\n\n--- next successful request ---\n\n",
                        batch.Select(item => item.UserText));
                    await ObserveSuccessfulTaskAsync(combined, batch.Last().CallAi, CancellationToken.None).ConfigureAwait(false);
                }
                catch (OperationCanceledException) { }
                catch (Exception ex) { DebugLogger.Log("Agent.Memory", "Automatic memory skipped: " + ex.Message); }
                finally { delay.Dispose(); }
            });
        }

        private sealed class PendingMemoryObservation
        {
            public string UserText { get; set; }
            public Func<List<Dictionary<string, string>>, CancellationToken, Task<string>> CallAi { get; set; }
        }

        private static IOnlineProvider TryCreateMemoryProvider(string modelName)
        {
            if (string.IsNullOrWhiteSpace(modelName)) return null;
            try
            {
                string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WordOllama", "models.json");
                if (!File.Exists(path)) return null;
                JArray profiles = JArray.Parse(File.ReadAllText(path));
                JObject profile = profiles.OfType<JObject>().FirstOrDefault(item =>
                    string.Equals(item["ModelName"]?.ToString(), modelName, StringComparison.OrdinalIgnoreCase));
                if (profile == null) return null;
                return ProviderFactory.Create(
                    profile["ProviderType"]?.ToString() ?? "OpenAI",
                    profile["EndPoint"]?.ToString() ?? "",
                    SecretProtector.UnprotectOrLegacy(profile["ApiKey"]?.ToString()),
                    profile["ModelName"]?.ToString() ?? modelName,
                    requestJsonOutput: true);
            }
            catch { return null; }
        }

        public static void ApplyOperations(string response)
        {
            JObject root;
            try
            {
                string json = ExtractJson(response);
                root = JObject.Parse(json);
            }
            catch { return; }

            UserMemorySettings settings = Get();
            var byId = settings.Memories.ToDictionary(item => item.Id, StringComparer.OrdinalIgnoreCase);
            foreach (JObject operation in root["operations"] as JArray ?? new JArray())
            {
                string action = operation["action"]?.ToString()?.Trim().ToLowerInvariant();
                string id = operation["id"]?.ToString()?.Trim() ?? "";
                string content = Normalize(operation["content"]?.ToString(), 2000);
                if (action == "add" && IsSafeMemory(content) &&
                    !settings.Memories.Any(item => string.Equals(item.Content, content, StringComparison.OrdinalIgnoreCase)))
                {
                    settings.Memories.Add(new UserMemoryItem { Content = content });
                }
                else if (action == "update" && byId.TryGetValue(id, out UserMemoryItem item) && IsSafeMemory(content))
                {
                    item.Content = content;
                    item.UpdatedAt = DateTime.UtcNow;
                }
                else if (action == "delete" && byId.TryGetValue(id, out UserMemoryItem remove))
                {
                    settings.Memories.Remove(remove);
                }
            }
            Save(settings);
        }

        private static bool IsSafeMemory(string content)
        {
            if (string.IsNullOrWhiteSpace(content) || content.Length > 2000) return false;
            string lower = content.ToLowerInvariant();
            string[] blocked = { "password", "api key", "apikey", "token", "secret", "密码", "密钥", "令牌", "身份证", "银行卡" };
            return !blocked.Any(lower.Contains);
        }

        private static string ExtractJson(string value)
        {
            value = value ?? "";
            int start = value.IndexOf('{');
            int end = value.LastIndexOf('}');
            return start >= 0 && end > start ? value.Substring(start, end - start + 1) : value;
        }

        private static UserMemorySettings LoadCore()
        {
            try
            {
                if (File.Exists(StorePath))
                    return JsonConvert.DeserializeObject<UserMemorySettings>(File.ReadAllText(StorePath)) ?? new UserMemorySettings();
            }
            catch { }
            return new UserMemorySettings();
        }

        private static UserMemorySettings Clone(UserMemorySettings source) =>
            JsonConvert.DeserializeObject<UserMemorySettings>(JsonConvert.SerializeObject(source)) ?? new UserMemorySettings();

        private static string Normalize(string value, int limit)
        {
            value = (value ?? "").Trim();
            return value.Length > limit ? value.Substring(0, limit) : value;
        }
    }
}
