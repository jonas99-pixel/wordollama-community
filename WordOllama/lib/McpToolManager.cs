using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.lib
{
    /// <summary>
    /// MCP 工具管理器 — 管理多个 MCP Bridge，统一对外暴露工具。
    /// </summary>
    public class McpToolManager : IDisposable
    {
        private readonly Dictionary<string, McpBridge> _bridges = new Dictionary<string, McpBridge>();
        private readonly List<McpToolDefinition> _discoveredTools = new List<McpToolDefinition>();
        private readonly Dictionary<string, HashSet<string>> _enabledToolsByServer =
            new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _trustedServers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, McpServerConnectionSpec> _serverSpecs =
            new Dictionary<string, McpServerConnectionSpec>(StringComparer.OrdinalIgnoreCase);
        public string LastError { get; private set; }
        public string LastErrorCategory { get; private set; }

        private sealed class McpServerConnectionSpec
        {
            public McpTransportKind Transport { get; set; }
            public string EndpointOrCommand { get; set; }
            public string Args { get; set; }
            public string WorkingDirectory { get; set; }
            public Dictionary<string, string> EnvironmentVariables { get; set; }
            public Dictionary<string, string> Headers { get; set; }
            public List<string> EnabledToolNames { get; set; }
            public bool Trusted { get; set; }
        }

        /// <summary>
        /// 连接 MCP Server（异步启动子进程 + 工具发现）。
        /// </summary>
        public async Task AddServerAsync(
            string serverName,
            string command,
            string args = "",
            string workingDirectory = "",
            Dictionary<string, string> environmentVariables = null)
        {
            await AddServerAsync(
                serverName,
                McpTransportKind.Stdio,
                command,
                args,
                workingDirectory,
                environmentVariables,
                null);
        }

        public async Task AddServerAsync(
            string serverName,
            McpTransportKind transportKind,
            string endpointOrCommand,
            string args = "",
            string workingDirectory = "",
            Dictionary<string, string> environmentVariables = null,
            Dictionary<string, string> httpHeaders = null,
            IEnumerable<string> enabledToolNames = null,
            CancellationToken cancellationToken = default,
            bool trusted = false)
        {
            LastError = null;
            LastErrorCategory = null;
            var spec = new McpServerConnectionSpec
            {
                Transport = transportKind,
                EndpointOrCommand = endpointOrCommand,
                Args = args,
                WorkingDirectory = workingDirectory,
                EnvironmentVariables = environmentVariables,
                Headers = httpHeaders,
                EnabledToolNames = enabledToolNames?.ToList(),
                Trusted = trusted
            };
            _serverSpecs[serverName] = spec;
            if (trusted) _trustedServers.Add(serverName); else _trustedServers.Remove(serverName);

            if (_bridges.TryGetValue(serverName, out var existing))
            {
                existing.Dispose();
                _bridges.Remove(serverName);
                _discoveredTools.RemoveAll(t => t.ServerName == serverName);
                _enabledToolsByServer.Remove(serverName);
            }

            var bridge = new McpBridge(serverName, transportKind, endpointOrCommand, args, workingDirectory, environmentVariables, httpHeaders);
            if (await bridge.ConnectAsync(cancellationToken))
            {
                _bridges[serverName] = bridge;
                var tools = await bridge.ListToolsAsync(cancellationToken);
                LastError = bridge.LastError;
                LastErrorCategory = bridge.LastErrorCategory;
                _discoveredTools.AddRange(tools);
                if (enabledToolNames != null)
                {
                    _enabledToolsByServer[serverName] = new HashSet<string>(
                        enabledToolNames.Where(name => !string.IsNullOrWhiteSpace(name)),
                        StringComparer.OrdinalIgnoreCase);
                }
                Debug.WriteLine($"MCP: 已连接 {serverName}，发现 {tools.Count} 个工具");
            }
            else
            {
                LastError = bridge.LastError ?? $"MCP server '{serverName}' failed to connect.";
                LastErrorCategory = bridge.LastErrorCategory ?? "transport";
                bridge.Dispose();
                Debug.WriteLine($"MCP: 连接 {serverName} 失败");
            }
        }

        /// <summary>
        /// 获取所有发现的外部工具定义。
        /// </summary>
        public List<McpToolDefinition> GetDiscoveredTools() => new List<McpToolDefinition>(_discoveredTools);

        /// <summary>Returns only the tools explicitly authorized for Agent use.</summary>
        public List<McpToolDefinition> GetEnabledTools()
        {
            return _discoveredTools.Where(tool =>
                !_enabledToolsByServer.TryGetValue(tool.ServerName, out var enabled)
                || enabled.Contains(tool.Name)).ToList();
        }

        public bool IsServerTrusted(string serverName) => _trustedServers.Contains(serverName ?? "");

        public async Task<bool> ReconnectServerAsync(string serverName, CancellationToken cancellationToken = default)
        {
            if (!_serverSpecs.TryGetValue(serverName, out var spec)) return false;

            await AddServerAsync(
                serverName,
                spec.Transport,
                spec.EndpointOrCommand,
                spec.Args,
                spec.WorkingDirectory,
                spec.EnvironmentVariables,
                spec.Headers,
                spec.EnabledToolNames,
                cancellationToken,
                spec.Trusted);
            return _bridges.ContainsKey(serverName);
        }

        /// <summary>
        /// 调用外部 MCP 工具。
        /// </summary>
        public async Task<string> CallToolAsync(string qualifiedName, JObject arguments, CancellationToken cancellationToken = default)
        {
            // 解析 mcp__serverName__toolName 格式
            string serverName = null;
            string toolName = null;

            if (qualifiedName.StartsWith("mcp__"))
            {
                var parts = qualifiedName.Split(new[] { "__" }, 3, StringSplitOptions.None);
                if (parts.Length == 3)
                {
                    serverName = parts[1];
                    toolName = parts[2];
                }
            }

            if (serverName == null || !_bridges.TryGetValue(serverName, out var bridge))
                return AgentToolText.F("I18n_AgentResultMcpToolServerDisconnected", qualifiedName);

            if (!bridge.IsConnected)
            {
                if (!await ReconnectServerAsync(serverName, cancellationToken)
                    || !_bridges.TryGetValue(serverName, out bridge))
                    return AgentToolText.F("I18n_AgentResultMcpReconnectFailed", serverName);
            }

            return await bridge.CallToolAsync(toolName, arguments, cancellationToken);
        }

        public async Task<string> CallToolAsync(string serverName, string toolName, JObject arguments, CancellationToken cancellationToken = default)
        {
            if (!_bridges.TryGetValue(serverName, out var bridge))
                return AgentToolText.F("I18n_AgentResultMcpServerDisconnected", serverName);
            if (!bridge.IsConnected)
            {
                if (!await ReconnectServerAsync(serverName, cancellationToken)
                    || !_bridges.TryGetValue(serverName, out bridge))
                    return AgentToolText.F("I18n_AgentResultMcpReconnectFailed", serverName);
            }
            return await bridge.CallToolAsync(toolName, arguments, cancellationToken);
        }

        public int ServerCount => _bridges.Count;
        public int ToolCount => _discoveredTools.Count;

        public void Dispose()
        {
            foreach (var b in _bridges.Values)
                b.Dispose();
            _bridges.Clear();
            _discoveredTools.Clear();
            _enabledToolsByServer.Clear();
            _serverSpecs.Clear();
            _trustedServers.Clear();
        }

        /// <summary>
        /// 从配置文件加载并连接所有启用的 MCP 服务器。
        /// 配置路径: %LOCALAPPDATA%/WordOllama/mcp-servers.json
        /// </summary>
        public static async Task<McpToolManager> LoadConfigAndConnectAsync()
        {
            var manager = new McpToolManager();
            try
            {
                string configPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "WordOllama", "mcp-servers.json");

                if (File.Exists(configPath))
                {
                    string json = File.ReadAllText(configPath);
                    foreach (var item in EnumerateConfiguredServers(JToken.Parse(json)))
                    {
                        bool enabled = item["enabled"]?.ToObject<bool>() ?? true;
                        if (!enabled) continue;

                        string name = item["name"]?.ToString() ?? "";
                        string transport = item["transport"]?.ToString() ?? "stdio";
                        var kind = ParseTransportKind(transport);
                        string command = item["command"]?.ToString() ?? "";
                        string url = item["url"]?.ToString()
                            ?? item["endpoint"]?.ToString()
                            ?? "";
                        string args = ReadArgs(item["args"]);
                        string workingDirectory = item["cwd"]?.ToString()
                            ?? item["workingDirectory"]?.ToString()
                            ?? "";
                        var env = ParseEnvironment(ReadProtectedConfigValue(item, "envProtected", item["env"] ?? item["environment"]));
                        var headers = ParseEnvironment(ReadProtectedConfigValue(item, "headersProtected", item["headers"]));
                        var enabledTools = ReadToolNames(item["enabledTools"]);
                        bool trusted = string.Equals(item["trustLevel"]?.ToString(), "Trusted", StringComparison.OrdinalIgnoreCase);
                        string endpointOrCommand = kind == McpTransportKind.Stdio ? command : url;

                        if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(endpointOrCommand))
                        {
                            await manager.AddServerAsync(name, kind, endpointOrCommand, args, workingDirectory, env, headers, enabledTools, default, trusted);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"MCP bootstrap failed: {ex.Message}");
            }
            return manager;
        }

        public static McpTransportKind ParseTransportKind(string transport)
        {
            string value = (transport ?? "stdio").Trim().ToLowerInvariant();
            if (value == "http" || value == "streamable" || value == "streamable-http" || value == "streamable_http")
                return McpTransportKind.StreamableHttp;
            if (value == "sse" || value == "legacy-sse" || value == "legacy_sse")
                return McpTransportKind.Sse;
            return McpTransportKind.Stdio;
        }

        public static string TransportToText(McpTransportKind kind)
        {
            switch (kind)
            {
                case McpTransportKind.StreamableHttp:
                    return "streamable-http";
                case McpTransportKind.Sse:
                    return "sse";
                default:
                    return "stdio";
            }
        }

        private static IEnumerable<JObject> EnumerateConfiguredServers(JToken root)
        {
            if (root is JArray arr)
            {
                foreach (var item in arr.OfType<JObject>())
                    yield return item;
                yield break;
            }

            var servers = root?["mcpServers"] as JObject;
            if (servers == null)
                yield break;

            foreach (var prop in servers.Properties())
            {
                if (prop.Value is JObject server)
                {
                    var copy = (JObject)server.DeepClone();
                    copy["name"] = copy["name"]?.ToString() ?? prop.Name;
                    yield return copy;
                }
            }
        }

        private static string ReadArgs(JToken token)
        {
            if (token is JArray arr)
            {
                var parts = new List<string>();
                foreach (var item in arr)
                    parts.Add(QuoteArgument(item?.ToString() ?? ""));
                return string.Join(" ", parts);
            }
            return token?.ToString() ?? "";
        }

        private static string ReadProtectedConfigValue(JObject item, string protectedProperty, JToken legacyValue)
        {
            string protectedValue = item[protectedProperty]?.ToString();
            return !string.IsNullOrWhiteSpace(protectedValue)
                ? SecretProtector.UnprotectOrLegacy(protectedValue)
                : EnvironmentToText(legacyValue);
        }

        private static IEnumerable<string> ReadToolNames(JToken token)
        {
            if (token is JArray arr)
                return arr.Values<string>().Where(name => !string.IsNullOrWhiteSpace(name)).ToList();
            return Array.Empty<string>();
        }

        private static string QuoteArgument(string value)
        {
            if (string.IsNullOrEmpty(value)) return "\"\"";
            if (value.IndexOfAny(new[] { ' ', '\t', '"' }) < 0) return value;
            return "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
        }

        public static Dictionary<string, string> ParseEnvironment(JToken token)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (token is JObject obj)
            {
                foreach (var prop in obj.Properties())
                    result[prop.Name] = prop.Value?.ToString() ?? "";
            }
            else
            {
                foreach (var line in (token?.ToString() ?? "").Split(new[] { "\r\n", "\n" }, StringSplitOptions.None))
                {
                    string trimmed = line.Trim();
                    if (string.IsNullOrWhiteSpace(trimmed) || trimmed.StartsWith("#")) continue;
                    int index = trimmed.IndexOf('=');
                    if (index <= 0) continue;
                    result[trimmed.Substring(0, index).Trim()] = trimmed.Substring(index + 1);
                }
            }
            return result;
        }

        public static string EnvironmentToText(JToken token)
        {
            var env = ParseEnvironment(token);
            var sb = new StringBuilder();
            foreach (var item in env)
                sb.AppendLine($"{item.Key}={item.Value}");
            return sb.ToString().TrimEnd();
        }
    }

    /// <summary>
    /// MCP 工具适配器 — 将外部 MCP 工具包装为本地 IAgentTool。
    /// 命名规范: mcp__&lt;serverName&gt;__&lt;toolName&gt;
    /// </summary>
    public class McpToolWrapper : IAgentTool
    {
        private readonly McpToolDefinition _def;
        private readonly McpToolManager _manager;

        public McpToolWrapper(McpToolDefinition def, McpToolManager manager)
        {
            _def = def;
            _manager = manager;
        }

        // Claude Code 风格命名: mcp__serverName__toolName
        public string Name => BuildSafeToolName(_def.ServerName, _def.Name);
        public string Description => AgentToolText.F("I18n_AgentToolDescMcp", _def.Description, _def.ServerName);
        // MCP annotations are informational only. Every MCP call receives a separate runtime permission check.
        public bool IsWriteOperation => false;
        public bool IsConcurrencySafe => false;
        public bool IsTrusted => _manager.IsServerTrusted(_def.ServerName);
        public string ServerName => _def.ServerName;
        public string ToolName => _def.Name;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => _def.InputSchema ?? new JObject();

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            // MCP 调用是异步的，这里阻塞等待
            return Task.Run(() => _manager.CallToolAsync(_def.ServerName, _def.Name, parameters)).Result;
        }

        public Task<string> ExecuteWithCancellationAsync(JObject parameters, CancellationToken cancellationToken)
        {
            return _manager.CallToolAsync(_def.ServerName, _def.Name, parameters, cancellationToken);
        }

        private static string BuildSafeToolName(string serverName, string toolName)
        {
            string raw = $"mcp__{serverName}__{toolName}";
            var sb = new StringBuilder(raw.Length);
            foreach (char c in raw)
            {
                bool asciiLetterOrDigit =
                    (c >= 'a' && c <= 'z') ||
                    (c >= 'A' && c <= 'Z') ||
                    (c >= '0' && c <= '9');
                sb.Append(asciiLetterOrDigit || c == '_' || c == '-' ? c : '_');
            }

            string safe = sb.ToString();
            if (safe == raw && safe.Length <= 64)
                return safe;

            uint hash = 2166136261;
            foreach (char c in raw)
            {
                hash ^= c;
                hash *= 16777619;
            }
            string suffix = "_" + hash.ToString("x8");
            int prefixLength = Math.Min(64 - suffix.Length, safe.Length);
            return safe.Substring(0, prefixLength) + suffix;
        }
    }
}
