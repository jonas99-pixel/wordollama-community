using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Threading;

namespace WordOllama.lib
{
    public interface IOnlineProvider
    {
        string ProviderType { get; }
        Task<string> GenerateAsync(string prompt, string imageBase64 = null, double? temperature = null);
        Task<string> ChatAsync(List<Dictionary<string, string>> chatHistory, JArray tools = null);
        Task StreamAsync(string prompt, Action<string> processChunk);
        Task<List<string>> FetchModelsAsync();
    }

    public interface ICancellableOnlineProvider
    {
        Task<string> GenerateAsync(string prompt, string imageBase64, double? temperature, CancellationToken cancellationToken);
        Task<string> ChatAsync(List<Dictionary<string, string>> chatHistory, JArray tools, CancellationToken cancellationToken);
    }
}

namespace WordOllama.lib
{
    // Internal envelope: provider-native assistant payload is retained verbatim.
    public static class AgentWireProtocol
    {
        public const string Prefix = "[AGENT_NATIVE]";
        public static string Encode(string provider, JToken raw, JArray calls) => Prefix + new JObject
        { ["provider"] = provider, ["raw"] = raw, ["calls"] = calls }.ToString(Newtonsoft.Json.Formatting.None);
        public static JObject Decode(string text) => text != null && text.StartsWith(Prefix, StringComparison.Ordinal)
            ? JObject.Parse(text.Substring(Prefix.Length)) : null;
        public static JObject Call(string id, string name, JToken args) => new JObject
        { ["id"] = id, ["name"] = name, ["args"] = args };
        public static JArray OpenAiMessages(List<Dictionary<string, string>> history)
        {
            var messages = new JArray();
            foreach (var item in history)
            {
                if (item.TryGetValue("native", out var native))
                    messages.Add(JObject.Parse(native)["raw"].DeepClone());
                else messages.Add(JObject.FromObject(item));
            }
            return messages;
        }
        public static JArray OllamaMessages(List<Dictionary<string, string>> history)
        {
            var result = OpenAiMessages(history);
            foreach (JObject message in result)
                if (message["role"]?.ToString() == "tool")
                {
                    message["tool_name"] = message["name"];
                    message.Remove("name");
                    message.Remove("tool_call_id");
                }
            return result;
        }

        public static string ValidateArguments(JObject schema, JObject args)
        {
            if (schema?["required"] is JArray required)
                foreach (var field in required)
                    if (args?[field.ToString()] == null || args[field.ToString()].Type == JTokenType.Null)
                        return "Missing required tool argument: " + field;
            if (schema?["properties"] is JObject properties && args != null)
                foreach (var property in properties.Properties())
                {
                    var value = args[property.Name];
                    if (value == null) continue;
                    string type = property.Value["type"]?.ToString();
                    bool valid = type == null || (type != "string" && type != "integer" && type != "number" && type != "boolean" && type != "object" && type != "array") || (type == "string" && value.Type == JTokenType.String)
                        || (type == "integer" && value.Type == JTokenType.Integer)
                        || (type == "number" && (value.Type == JTokenType.Integer || value.Type == JTokenType.Float))
                        || (type == "boolean" && value.Type == JTokenType.Boolean)
                        || (type == "object" && value.Type == JTokenType.Object)
                        || (type == "array" && value.Type == JTokenType.Array);
                    if (!valid) return "Invalid tool argument type: " + property.Name + "; expected " + type;
                }
            return null;
        }

        public static Dictionary<string, string> Result(string id, string name, string result, bool failed) =>
            new Dictionary<string, string> { ["role"] = "tool", ["tool_call_id"] = id, ["name"] = name,
                ["content"] = new JObject { ["ok"] = !failed, ["result"] = result }.ToString(Newtonsoft.Json.Formatting.None) };
    }
}
