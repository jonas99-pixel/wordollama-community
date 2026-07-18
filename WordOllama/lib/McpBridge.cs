using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace WordOllama.lib
{
    /// <summary>
    /// MCP 工具定义（从 MCP Server 的 tools/list 响应中获取）。
    /// </summary>
    public class McpToolDefinition
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public JObject InputSchema { get; set; }
        public JObject Annotations { get; set; }
        public string ServerName { get; set; }
    }

    public enum McpTransportKind
    {
        Stdio,
        StreamableHttp,
        Sse
    }

    /// <summary>
    /// MCP 桥接器 — 通过 stdio 与外部 MCP Server 通信。
    /// 实现 JSON-RPC 2.0 协议，支持 tools/list 和 tools/call。
    /// 
    /// 参考 Claude Code MCP 实现 (analysis/04d-mcp-implementation.md):
    /// - 传输: stdio (本地子进程)
    /// - 工具命名: mcp__&lt;serverName&gt;__&lt;toolName&gt;
    /// - 并发: 单连接串行调用
    /// </summary>
    public class McpBridge : IDisposable
    {
        private Process _process;
        private readonly string _serverName;
        private readonly string _command;
        private readonly string _args;
        private readonly string _workingDirectory;
        private readonly Dictionary<string, string> _environmentVariables;
        private readonly McpTransportKind _transportKind;
        private readonly Uri _httpEndpoint;
        private readonly Dictionary<string, string> _httpHeaders;
        private int _requestId;
        private readonly object _lock = new object();
        private readonly Dictionary<int, TaskCompletionSource<JObject>> _pendingRequests
            = new Dictionary<int, TaskCompletionSource<JObject>>();
        private readonly SemaphoreSlim _writeLock = new SemaphoreSlim(1, 1);
        private CancellationTokenSource _readLoopCts;
        private Task _readLoopTask;
        private Task _stderrLoopTask;
        private TaskCompletionSource<Uri> _legacySseEndpointSource;
        private HttpClient _httpClient;
        private bool _httpSessionHealthy;
        public string LastError { get; private set; }
        public string LastErrorCategory { get; private set; }
        public DateTime? LastHandshakeAt { get; private set; }
        public McpTransportKind TransportKind => _transportKind;
        private string _sessionId;
        private string _lastEventId;
        private bool _disposed;

        private const int HandshakeTimeoutMs = 15000;
        private const int ToolCallTimeoutMs = 120000;
        private const string ProtocolVersion = "2025-11-25";

        public string ServerName => _serverName;
        public bool IsConnected => _transportKind == McpTransportKind.Stdio
            ? _process != null && !_process.HasExited
            : !_disposed && _httpClient != null && _httpSessionHealthy;

        public McpBridge(
            string serverName,
            string command,
            string args = "",
            string workingDirectory = "",
            Dictionary<string, string> environmentVariables = null)
            : this(serverName, McpTransportKind.Stdio, command, args, workingDirectory, environmentVariables, null, null)
        {
        }

        public McpBridge(
            string serverName,
            McpTransportKind transportKind,
            string endpointOrCommand,
            string args = "",
            string workingDirectory = "",
            Dictionary<string, string> environmentVariables = null,
            Dictionary<string, string> httpHeaders = null)
            : this(serverName, transportKind, endpointOrCommand, args, workingDirectory, environmentVariables, httpHeaders, null)
        {
        }

        private McpBridge(
            string serverName,
            McpTransportKind transportKind,
            string endpointOrCommand,
            string args,
            string workingDirectory,
            Dictionary<string, string> environmentVariables,
            Dictionary<string, string> httpHeaders,
            Uri httpEndpoint)
        {
            _serverName = serverName;
            _transportKind = transportKind;
            _command = endpointOrCommand;
            _args = args;
            _workingDirectory = workingDirectory;
            _environmentVariables = environmentVariables ?? new Dictionary<string, string>();
            _httpHeaders = httpHeaders ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (transportKind != McpTransportKind.Stdio)
                _httpEndpoint = httpEndpoint ?? new Uri(endpointOrCommand);
        }

        /// <summary>
        /// 启动 MCP Server 子进程并完成初始化握手。
        /// </summary>
        public async Task<bool> ConnectAsync(CancellationToken cancellationToken = default)
        {
            try
            {
                if (_transportKind == McpTransportKind.StreamableHttp)
                    return await ConnectStreamableHttpAsync(cancellationToken);
                if (_transportKind == McpTransportKind.Sse)
                    return await ConnectLegacySseAsync(cancellationToken);

                cancellationToken.ThrowIfCancellationRequested();

                _process = new Process { StartInfo = BuildStartInfo() };
                _process.Start();
                _process.StandardInput.AutoFlush = true;

                // 启动读取循环（后台线程）
                _readLoopCts = new CancellationTokenSource();
                var token = _readLoopCts.Token;
                _readLoopTask = Task.Run(() => ReadLoopAsync(token), token);
                _stderrLoopTask = Task.Run(() => ReadStderrLoopAsync(token), token);

                // 发送初始化请求
                var initResult = await SendRequestAsync("initialize", new JObject
                {
                    ["protocolVersion"] = ProtocolVersion,
                    ["capabilities"] = new JObject(),
                    ["clientInfo"] = new JObject
                    {
                        ["name"] = "WordOllama",
                        ["version"] = "2.0"
                    }
                }, HandshakeTimeoutMs, cancellationToken);

                if (initResult?["result"] == null)
                    return false;

                // MCP 生命周期要求 initialize 成功后发送 initialized 通知。
                await SendNotificationAsync("notifications/initialized", new JObject(), cancellationToken);
                LastError = null;
                LastErrorCategory = null;
                LastHandshakeAt = DateTime.UtcNow;
                return true;
            }
            catch (OperationCanceledException)
            {
                Dispose();
                throw;
            }
            catch (Exception ex)
            {
                SetError(ex);
                Debug.WriteLine($"MCP Connect failed for {_serverName}: {ex.Message}");
                Dispose();
                return false;
            }
        }

        private async Task<bool> ConnectStreamableHttpAsync(CancellationToken cancellationToken)
        {
            _httpClient = CreateHttpClient();

            var initResult = await SendRequestAsync("initialize", BuildInitializeParams(), HandshakeTimeoutMs, cancellationToken);
            if (initResult?["result"] == null)
                return false;

            await SendNotificationAsync("notifications/initialized", new JObject(), cancellationToken);
            _httpSessionHealthy = true;
            LastError = null;
            LastErrorCategory = null;
            LastHandshakeAt = DateTime.UtcNow;
            return true;
        }

        private async Task<bool> ConnectLegacySseAsync(CancellationToken cancellationToken)
        {
            _httpClient = CreateHttpClient();
            _readLoopCts = new CancellationTokenSource();
            _legacySseEndpointSource = new TaskCompletionSource<Uri>(TaskCreationOptions.RunContinuationsAsynchronously);
            var token = _readLoopCts.Token;
            _readLoopTask = Task.Run(() => LegacySseReadLoopAsync(token), token);

            var endpointTask = _legacySseEndpointSource.Task;
            var timeoutTask = Task.Delay(HandshakeTimeoutMs, cancellationToken);
            if (await Task.WhenAny(endpointTask, timeoutTask) != endpointTask)
            {
                cancellationToken.ThrowIfCancellationRequested();
                throw new TimeoutException("Legacy SSE endpoint event timed out.");
            }

            var initResult = await SendRequestAsync("initialize", BuildInitializeParams(), HandshakeTimeoutMs, cancellationToken);
            if (initResult?["result"] == null)
                return false;

            await SendNotificationAsync("notifications/initialized", new JObject(), cancellationToken);
            _httpSessionHealthy = true;
            LastError = null;
            LastErrorCategory = null;
            LastHandshakeAt = DateTime.UtcNow;
            return true;
        }

        private JObject BuildInitializeParams()
        {
            return new JObject
            {
                ["protocolVersion"] = ProtocolVersion,
                ["capabilities"] = new JObject(),
                ["clientInfo"] = new JObject
                {
                    ["name"] = "WordOllama",
                    ["version"] = "2.0"
                }
            };
        }

        /// <summary>
        /// 获取 MCP Server 提供的所有工具定义。
        /// </summary>
        public async Task<List<McpToolDefinition>> ListToolsAsync(CancellationToken cancellationToken = default)
        {
            var tools = new List<McpToolDefinition>();
            try
            {
                string cursor = null;
                int pageCount = 0;
                do
                {
                    var parameters = new JObject();
                    if (!string.IsNullOrEmpty(cursor))
                        parameters["cursor"] = cursor;

                    var response = await SendRequestAsync("tools/list", parameters, HandshakeTimeoutMs, cancellationToken);
                    if (response?["error"] != null)
                        throw new InvalidOperationException(response["error"]["message"]?.ToString() ?? "tools/list failed");

                    if (response?["result"]?["tools"] is JArray toolArray)
                    {
                        foreach (var t in toolArray)
                        {
                            tools.Add(new McpToolDefinition
                            {
                                Name = t["name"]?.ToString() ?? "",
                                Description = TruncateDescription(t["description"]?.ToString() ?? ""),
                                InputSchema = t["inputSchema"] as JObject ?? new JObject(),
                                Annotations = t["annotations"] as JObject ?? new JObject(),
                                ServerName = _serverName
                            });
                        }
                    }

                    cursor = response?["result"]?["nextCursor"]?.ToString();
                    pageCount++;
                } while (!string.IsNullOrEmpty(cursor) && pageCount < 100);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                SetError(ex);
                Debug.WriteLine($"MCP ListTools failed: {ex.Message}");
            }
            return tools;
        }

        /// <summary>
        /// 调用 MCP Server 上的工具。
        /// </summary>
        public async Task<string> CallToolAsync(string toolName, JObject arguments, CancellationToken cancellationToken = default)
        {
            try
            {
                var response = await SendRequestAsync("tools/call", new JObject
                {
                    ["name"] = toolName,
                    ["arguments"] = arguments ?? new JObject()
                }, ToolCallTimeoutMs, cancellationToken);

                if (response?["error"] != null)
                {
                    return AgentToolText.F("I18n_AgentResultMcpError", response["error"]["message"]);
                }

                bool isError = response?["result"]?["isError"]?.ToObject<bool>() ?? false;

                // MCP 工具返回 content 数组，提取文本
                if (response?["result"]?["content"] is JArray contentArray)
                {
                    var texts = new List<string>();
                    foreach (var item in contentArray)
                    {
                        if (item["type"]?.ToString() == "text")
                            texts.Add(item["text"]?.ToString() ?? "");
                    }
                    string result = string.Join("\n", texts);
                    return isError ? AgentToolText.F("I18n_AgentResultMcpError", result) : result;
                }

                return response?["result"]?.ToString() ?? AgentToolText.T("I18n_AgentResultMcpEmptyResponse");
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                _httpSessionHealthy = false;
                SetError(ex);
                return AgentToolText.F("I18n_AgentResultMcpException", _serverName, toolName, ex.Message);
            }
        }

        /// <summary>
        /// 发送 JSON-RPC 2.0 请求并等待响应。
        /// </summary>
        private async Task<JObject> SendRequestAsync(string method, JObject parameters, int timeoutMs, CancellationToken cancellationToken = default)
        {
            if (_transportKind == McpTransportKind.StreamableHttp)
                return await SendStreamableHttpRequestAsync(method, parameters, timeoutMs, cancellationToken);

            ThrowIfDisconnected();

            int id;
            lock (_lock) { id = ++_requestId; }

            var tcs = new TaskCompletionSource<JObject>(TaskCreationOptions.RunContinuationsAsynchronously);
            lock (_lock) { _pendingRequests[id] = tcs; }

            var request = new JObject
            {
                ["jsonrpc"] = "2.0",
                ["id"] = id,
                ["method"] = method,
                ["params"] = parameters
            };

            try
            {
                await SendMessageAsync(request);
            }
            catch (Exception ex)
            {
                lock (_lock) { _pendingRequests.Remove(id); }
                throw new IOException($"Failed to send MCP request '{method}'.", ex);
            }

            using (var timeoutCts = new CancellationTokenSource(timeoutMs))
            using (var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(timeoutCts.Token, cancellationToken))
            {
                var timeoutTask = Task.Delay(Timeout.Infinite, linkedCts.Token);
                var completed = await Task.WhenAny(tcs.Task, timeoutTask);
                if (completed == tcs.Task)
                    return await tcs.Task;

                lock (_lock) { _pendingRequests.Remove(id); }
                cancellationToken.ThrowIfCancellationRequested();
                throw new TimeoutException($"MCP request '{method}' timed out after {timeoutMs} ms.");
            }
        }

        private async Task SendNotificationAsync(string method, JObject parameters, CancellationToken cancellationToken = default)
        {
            var notification = new JObject
            {
                ["jsonrpc"] = "2.0",
                ["method"] = method,
                ["params"] = parameters ?? new JObject()
            };
            await SendMessageAsync(notification, cancellationToken);
        }

        private async Task SendMessageAsync(JObject message, CancellationToken cancellationToken = default)
        {
            if (_transportKind == McpTransportKind.StreamableHttp)
            {
                await SendStreamableHttpMessageAsync(message, HandshakeTimeoutMs, cancellationToken);
                return;
            }

            if (_transportKind == McpTransportKind.Sse)
            {
                await SendLegacySseMessageAsync(message, HandshakeTimeoutMs, cancellationToken);
                return;
            }

            ThrowIfDisconnected();
            await _writeLock.WaitAsync(cancellationToken);
            try
            {
                string json = message.ToString(Formatting.None) + "\n";
                await _process.StandardInput.WriteAsync(json);
                await _process.StandardInput.FlushAsync();
            }
            finally
            {
                _writeLock.Release();
            }
        }

        private HttpClient CreateHttpClient()
        {
            var client = new HttpClient(new HttpClientHandler
            {
                AllowAutoRedirect = false
            });
            client.Timeout = Timeout.InfiniteTimeSpan;
            return client;
        }

        private HttpRequestMessage CreateHttpRequest(HttpMethod method, Uri uri)
        {
            var request = new HttpRequestMessage(method, uri);
            foreach (var header in _httpHeaders)
            {
                if (string.IsNullOrWhiteSpace(header.Key)) continue;
                request.Headers.TryAddWithoutValidation(header.Key.Trim(), header.Value ?? "");
            }

            if (!string.IsNullOrEmpty(_sessionId))
                request.Headers.TryAddWithoutValidation("MCP-Session-Id", _sessionId);

            request.Headers.TryAddWithoutValidation("MCP-Protocol-Version", ProtocolVersion);
            return request;
        }

        private async Task<JObject> SendStreamableHttpRequestAsync(string method, JObject parameters, int timeoutMs, CancellationToken cancellationToken)
        {
            int id;
            lock (_lock) { id = ++_requestId; }

            var requestObject = new JObject
            {
                ["jsonrpc"] = "2.0",
                ["id"] = id,
                ["method"] = method,
                ["params"] = parameters ?? new JObject()
            };

            using (var timeoutCts = new CancellationTokenSource(timeoutMs))
            using (var cts = CancellationTokenSource.CreateLinkedTokenSource(timeoutCts.Token, cancellationToken))
            using (var request = CreateHttpRequest(HttpMethod.Post, _httpEndpoint))
            {
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("text/event-stream"));
                request.Content = new StringContent(requestObject.ToString(Formatting.None), Encoding.UTF8, "application/json");

                using (var response = await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cts.Token))
                {
                    CaptureSessionId(response);
                    if (!response.IsSuccessStatusCode)
                        throw new IOException($"HTTP MCP request '{method}' failed: {(int)response.StatusCode} {response.ReasonPhrase}");

                    return await ReadHttpResponseForIdAsync(response, id, cts.Token);
                }
            }
        }

        private async Task SendStreamableHttpMessageAsync(JObject message, int timeoutMs, CancellationToken cancellationToken)
        {
            using (var timeoutCts = new CancellationTokenSource(timeoutMs))
            using (var cts = CancellationTokenSource.CreateLinkedTokenSource(timeoutCts.Token, cancellationToken))
            using (var request = CreateHttpRequest(HttpMethod.Post, _httpEndpoint))
            {
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("text/event-stream"));
                request.Content = new StringContent(message.ToString(Formatting.None), Encoding.UTF8, "application/json");

                using (var response = await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cts.Token))
                {
                    CaptureSessionId(response);
                    if (!response.IsSuccessStatusCode && response.StatusCode != HttpStatusCode.Accepted)
                        throw new IOException($"HTTP MCP message failed: {(int)response.StatusCode} {response.ReasonPhrase}");
                }
            }
        }

        private async Task SendLegacySseMessageAsync(JObject message, int timeoutMs, CancellationToken cancellationToken)
        {
            Uri endpoint = _legacySseEndpointSource?.Task.Status == TaskStatus.RanToCompletion
                ? _legacySseEndpointSource.Task.Result
                : null;
            if (endpoint == null)
                throw new IOException("Legacy SSE message endpoint is not available.");

            using (var timeoutCts = new CancellationTokenSource(timeoutMs))
            using (var cts = CancellationTokenSource.CreateLinkedTokenSource(timeoutCts.Token, cancellationToken))
            using (var request = CreateHttpRequest(HttpMethod.Post, endpoint))
            {
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("text/event-stream"));
                request.Content = new StringContent(message.ToString(Formatting.None), Encoding.UTF8, "application/json");

                using (var response = await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cts.Token))
                {
                    if (!response.IsSuccessStatusCode && response.StatusCode != HttpStatusCode.Accepted)
                        throw new IOException($"Legacy SSE MCP message failed: {(int)response.StatusCode} {response.ReasonPhrase}");
                }
            }
        }

        private void CaptureSessionId(HttpResponseMessage response)
        {
            if (response.Headers.TryGetValues("MCP-Session-Id", out var values))
            {
                foreach (var value in values)
                {
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        _sessionId = value.Trim();
                        break;
                    }
                }
            }
        }

        private async Task<JObject> ReadHttpResponseForIdAsync(HttpResponseMessage response, int id, CancellationToken token)
        {
            string contentType = response.Content.Headers.ContentType?.MediaType ?? "";
            if (contentType.IndexOf("text/event-stream", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                using (var stream = await response.Content.ReadAsStreamAsync())
                {
                    JObject matched = null;
                    await ReadSseEventsAsync(stream, async (eventName, data, eventId, retryMs) =>
                    {
                        if (!string.IsNullOrWhiteSpace(eventId))
                            _lastEventId = eventId;
                        if (string.IsNullOrWhiteSpace(data))
                            return false;

                        var message = JObject.Parse(data);
                        if (IsClientResponse(message, id))
                        {
                            matched = message;
                            return true;
                        }

                        await HandleServerMessageAsync(message);
                        return false;
                    }, token);

                    if (matched != null)
                        return matched;
                    throw new IOException($"HTTP MCP response stream ended before response id {id}.");
                }
            }

            string body = await response.Content.ReadAsStringAsync();
            if (string.IsNullOrWhiteSpace(body))
                return new JObject();
            var json = JObject.Parse(body);
            await HandleServerMessageAsync(json);
            return json;
        }

        private async Task LegacySseReadLoopAsync(CancellationToken token)
        {
            try
            {
                using (var request = CreateHttpRequest(HttpMethod.Get, _httpEndpoint))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("text/event-stream"));
                    using (var response = await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, token))
                    {
                        if (!response.IsSuccessStatusCode)
                            throw new IOException($"Legacy SSE GET failed: {(int)response.StatusCode} {response.ReasonPhrase}");

                        using (var stream = await response.Content.ReadAsStreamAsync())
                        {
                            await ReadSseEventsAsync(stream, async (eventName, data, eventId, retryMs) =>
                            {
                                if (!string.IsNullOrWhiteSpace(eventId))
                                    _lastEventId = eventId;

                                if (eventName == "endpoint")
                                {
                                    var endpoint = ResolveEndpoint(data);
                                    _legacySseEndpointSource?.TrySetResult(endpoint);
                                    return false;
                                }

                                if (eventName == "message" || string.IsNullOrEmpty(eventName))
                                {
                                    if (!string.IsNullOrWhiteSpace(data))
                                        await HandleServerMessageAsync(JObject.Parse(data));
                                }
                                return false;
                            }, token);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _legacySseEndpointSource?.TrySetException(ex);
                if (!token.IsCancellationRequested)
                    FailPendingRequests(ex);
            }
        }

        private Uri ResolveEndpoint(string endpoint)
        {
            if (string.IsNullOrWhiteSpace(endpoint))
                throw new IOException("Legacy SSE endpoint event did not include a message endpoint.");

            Uri resolved = Uri.TryCreate(endpoint.Trim(), UriKind.Absolute, out var absolute)
                ? absolute
                : new Uri(_httpEndpoint, endpoint.Trim());
            if (!IsSameOrigin(_httpEndpoint, resolved))
                throw new IOException("Legacy SSE message endpoint must use the same origin as the configured MCP server.");
            return resolved;
        }

        private static bool IsSameOrigin(Uri configured, Uri candidate)
        {
            if (configured == null || candidate == null) return false;
            return string.Equals(configured.Scheme, candidate.Scheme, StringComparison.OrdinalIgnoreCase)
                && string.Equals(configured.Host, candidate.Host, StringComparison.OrdinalIgnoreCase)
                && configured.Port == candidate.Port;
        }

        private async Task ReadSseEventsAsync(
            Stream stream,
            Func<string, string, string, int?, Task<bool>> onEvent,
            CancellationToken token)
        {
            using (var reader = new StreamReader(stream, Encoding.UTF8))
            {
                string eventName = "";
                string eventId = null;
                int? retryMs = null;
                var data = new StringBuilder();

                while (!token.IsCancellationRequested)
                {
                    string line = await reader.ReadLineAsync();
                    if (line == null) break;

                    if (line.Length == 0)
                    {
                        if (data.Length > 0 || !string.IsNullOrEmpty(eventName))
                        {
                            string payload = data.ToString();
                            if (payload.EndsWith("\n", StringComparison.Ordinal))
                                payload = payload.Substring(0, payload.Length - 1);
                            bool stop = await onEvent(eventName, payload, eventId, retryMs);
                            if (stop) break;
                        }
                        eventName = "";
                        eventId = null;
                        retryMs = null;
                        data.Clear();
                        continue;
                    }

                    if (line.StartsWith(":", StringComparison.Ordinal))
                        continue;

                    int colonIndex = line.IndexOf(':');
                    string field = colonIndex >= 0 ? line.Substring(0, colonIndex) : line;
                    string value = colonIndex >= 0 ? line.Substring(colonIndex + 1) : "";
                    if (value.StartsWith(" ", StringComparison.Ordinal))
                        value = value.Substring(1);

                    switch (field)
                    {
                        case "event":
                            eventName = value;
                            break;
                        case "data":
                            data.AppendLine(value);
                            break;
                        case "id":
                            eventId = value;
                            break;
                        case "retry":
                            if (int.TryParse(value, out var parsedRetry))
                                retryMs = parsedRetry;
                            break;
                    }
                }
            }
        }

        private bool IsClientResponse(JObject message, int id)
        {
            return message["id"]?.Type == JTokenType.Integer &&
                   message["id"]?.ToObject<int>() == id &&
                   (message["result"] != null || message["error"] != null);
        }

        private async Task HandleServerMessageAsync(JObject message)
        {
            JToken idToken = message["id"];
            bool isResponse = message["result"] != null || message["error"] != null;
            if (idToken != null && isResponse && idToken.Type == JTokenType.Integer)
            {
                int id = idToken.ToObject<int>();
                TaskCompletionSource<JObject> tcs;
                lock (_lock)
                {
                    if (_pendingRequests.TryGetValue(id, out tcs))
                        _pendingRequests.Remove(id);
                }
                tcs?.TrySetResult(message);
            }
            else if (idToken != null && message["method"] != null)
            {
                await RespondToServerRequestAsync(idToken.DeepClone(), message["method"]?.ToString());
            }
        }

        /// <summary>
        /// 后台读取循环 — 解析 MCP Server 的 JSON-RPC 响应流。
        /// </summary>
        private async Task ReadLoopAsync(CancellationToken token)
        {
            try
            {
                while (!_process.HasExited && !token.IsCancellationRequested)
                {
                    string line = await _process.StandardOutput.ReadLineAsync();
                    if (line == null) break;
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    try
                    {
                        var response = JObject.Parse(line);
                        JToken idToken = response["id"];
                        bool isResponse = response["result"] != null || response["error"] != null;
                        if (idToken != null && isResponse && idToken.Type == JTokenType.Integer)
                        {
                            int id = idToken.ToObject<int>();
                            TaskCompletionSource<JObject> tcs;
                            lock (_lock)
                            {
                                if (_pendingRequests.TryGetValue(id, out tcs))
                                    _pendingRequests.Remove(id);
                            }
                            tcs?.SetResult(response);
                        }
                        else if (idToken != null && response["method"] != null)
                        {
                            await RespondToServerRequestAsync(idToken.DeepClone(), response["method"]?.ToString());
                        }
                    }
                    catch { }
                }
            }
            catch { }
            finally
            {
                FailPendingRequests(new IOException($"MCP server '{_serverName}' disconnected."));
            }
        }

        private async Task RespondToServerRequestAsync(JToken id, string method)
        {
            JObject response;
            if (method == "ping")
            {
                response = new JObject
                {
                    ["jsonrpc"] = "2.0",
                    ["id"] = id,
                    ["result"] = new JObject()
                };
            }
            else
            {
                response = new JObject
                {
                    ["jsonrpc"] = "2.0",
                    ["id"] = id,
                    ["error"] = new JObject
                    {
                        ["code"] = -32601,
                        ["message"] = $"Unsupported server request: {method}"
                    }
                };
            }
            await SendMessageAsync(response);
        }

        private async Task ReadStderrLoopAsync(CancellationToken token)
        {
            try
            {
                while (!_process.HasExited && !token.IsCancellationRequested)
                {
                    string line = await _process.StandardError.ReadLineAsync();
                    if (line == null) break;
                    if (!string.IsNullOrWhiteSpace(line))
                        Debug.WriteLine($"MCP {_serverName} stderr: {line}");
                }
            }
            catch { }
        }

        private ProcessStartInfo BuildStartInfo()
        {
            string command = Environment.ExpandEnvironmentVariables(_command);
            string args = Environment.ExpandEnvironmentVariables(_args ?? "");
            command = ResolveWindowsCommand(command);

            if (Environment.OSVersion.Platform == PlatformID.Win32NT)
            {
                string extension = Path.GetExtension(command);
                if (extension.Equals(".cmd", StringComparison.OrdinalIgnoreCase) ||
                    extension.Equals(".bat", StringComparison.OrdinalIgnoreCase))
                {
                    string comSpec = Environment.GetEnvironmentVariable("ComSpec") ?? "cmd.exe";
                    args = $"/d /s /c \"\"{command}\"{(string.IsNullOrWhiteSpace(args) ? "" : " " + args)}\"";
                    command = comSpec;
                }
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = command,
                Arguments = args,
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };

            string workingDirectory = Environment.ExpandEnvironmentVariables(_workingDirectory ?? "").Trim();
            if (!string.IsNullOrWhiteSpace(workingDirectory) && Directory.Exists(workingDirectory))
                startInfo.WorkingDirectory = workingDirectory;

            foreach (var item in _environmentVariables)
            {
                if (string.IsNullOrWhiteSpace(item.Key)) continue;
                startInfo.EnvironmentVariables[item.Key.Trim()] = Environment.ExpandEnvironmentVariables(item.Value ?? "");
            }

            return startInfo;
        }

        private static string ResolveWindowsCommand(string command)
        {
            if (Environment.OSVersion.Platform != PlatformID.Win32NT ||
                string.IsNullOrWhiteSpace(command) ||
                Path.IsPathRooted(command) ||
                command.IndexOf(Path.DirectorySeparatorChar) >= 0 ||
                command.IndexOf(Path.AltDirectorySeparatorChar) >= 0)
            {
                return command;
            }

            string path = Environment.GetEnvironmentVariable("PATH") ?? "";
            string pathExt = Environment.GetEnvironmentVariable("PATHEXT") ?? ".COM;.EXE;.BAT;.CMD";
            var extensions = string.IsNullOrEmpty(Path.GetExtension(command))
                ? pathExt.Split(';')
                : new[] { "" };

            foreach (string directory in path.Split(Path.PathSeparator))
            {
                if (string.IsNullOrWhiteSpace(directory)) continue;
                foreach (string extension in extensions)
                {
                    try
                    {
                        string candidate = Path.Combine(directory.Trim(), command + extension);
                        if (File.Exists(candidate))
                            return candidate;
                    }
                    catch { }
                }
            }
            return command;
        }

        private void ThrowIfDisconnected()
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(McpBridge));
            if (_transportKind == McpTransportKind.Stdio && (_process == null || _process.HasExited))
                throw new IOException($"MCP server '{_serverName}' is not connected.");
            if (_transportKind != McpTransportKind.Stdio && _httpClient == null)
                throw new IOException($"MCP server '{_serverName}' is not connected.");
        }

        private void SetError(Exception exception)
        {
            LastError = exception?.Message ?? "Unknown MCP error";
            string text = LastError.ToLowerInvariant();
            if (exception is TimeoutException || text.Contains("timeout") || text.Contains("timed out"))
                LastErrorCategory = "timeout";
            else if (text.Contains("401") || text.Contains("403") || text.Contains("unauthorized") || text.Contains("forbidden"))
                LastErrorCategory = "authentication";
            else if (text.Contains("json") || text.Contains("protocol") || text.Contains("rpc"))
                LastErrorCategory = "protocol";
            else if (text.Contains("http") || text.Contains("endpoint") || text.Contains("connect"))
                LastErrorCategory = "endpoint";
            else
                LastErrorCategory = "transport";
        }

        private void FailPendingRequests(Exception exception)
        {
            List<TaskCompletionSource<JObject>> pending;
            lock (_lock)
            {
                pending = new List<TaskCompletionSource<JObject>>(_pendingRequests.Values);
                _pendingRequests.Clear();
            }
            foreach (var request in pending)
                request.TrySetException(exception);
        }

        /// <summary>
        /// 截断过长的工具描述，防止上下文窗口爆炸。
        /// 参考 Claude Code 实现：最大 2048 字符。
        /// </summary>
        private static string TruncateDescription(string desc)
        {
            if (string.IsNullOrEmpty(desc)) return "";
            return desc.Length > 2000 ? desc.Substring(0, 2000) + "..." : desc;
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            // 停止读取循环
            if (_readLoopCts != null)
            {
                _readLoopCts.Cancel();
            }

            try
            {
                if (_process != null && !_process.HasExited)
                {
                    _process.StandardInput.Close();
                    if (!_process.WaitForExit(2000))
                        _process.Kill();
                }
            }
            catch { }
            TryTerminateHttpSession();
            FailPendingRequests(new ObjectDisposedException(nameof(McpBridge)));
            try { _process?.Dispose(); } catch { }
            try { _httpClient?.Dispose(); } catch { }
            _process = null;
            _httpClient = null;
            _httpSessionHealthy = false;
            _readLoopCts?.Dispose();
            _readLoopCts = null;
            _writeLock.Dispose();
        }

        private void TryTerminateHttpSession()
        {
            if (_transportKind != McpTransportKind.StreamableHttp ||
                _httpClient == null ||
                _httpEndpoint == null ||
                string.IsNullOrEmpty(_sessionId))
            {
                return;
            }

            try
            {
                using (var request = CreateHttpRequest(HttpMethod.Delete, _httpEndpoint))
                {
                    var task = _httpClient.SendAsync(request);
                    task.Wait(500);
                }
            }
            catch { }
        }
    }
}
