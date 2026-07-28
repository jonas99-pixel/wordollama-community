using System;
using System.Collections.Generic;

namespace WordOllama.lib
{
    /// <summary>
    /// MCP 服务器配置 — 持久化到 JSON 文件。
    /// </summary>
    public class McpServerConfig
    {
        /// <summary>服务器名称（唯一标识，用于工具命名 mcp__name__tool）</summary>
        public string Name { get; set; }

        /// <summary>传输类型：stdio / streamable-http / sse。</summary>
        public string Transport { get; set; } = "stdio";

        /// <summary>远程 MCP endpoint（Streamable HTTP 或 legacy SSE）。</summary>
        public string Url { get; set; }

        /// <summary>启动命令（如 "npx", "node", "python"）</summary>
        public string Command { get; set; }

        /// <summary>命令参数（如 "mcp-server-企查查"）</summary>
        public string Args { get; set; }

        /// <summary>工作目录（可选）。部分 MCP Server 依赖相对路径或本地配置文件。</summary>
        public string WorkingDirectory { get; set; }

        /// <summary>环境变量（可选），每行 KEY=VALUE。常用于 API Key 或私有服务地址。</summary>
        public string EnvironmentVariables { get; set; }

        /// <summary>HTTP 请求头（可选），每行 Header=Value。常用于 Authorization。</summary>
        public string Headers { get; set; }

        /// <summary>是否启用</summary>
        public bool Enabled { get; set; } = true;
        public string TrustLevel { get; set; } = "Untrusted";

        /// <summary>测试后发现的工具名称，用于在设置界面中配置白名单。</summary>
        public List<string> AvailableToolNames { get; set; } = new List<string>();

        /// <summary>明确授权给 Agent 的工具名称。空列表表示不授权任何工具。</summary>
        public List<string> EnabledToolNames { get; set; } = new List<string>();

        public string ConnectionState { get; set; } = "Not tested";
        public string EndpointDisplay { get; set; }
        public DateTime? LastHandshakeAt { get; set; }
        public DateTime? LastTestedAt { get; set; }
        public int LastToolCount { get; set; }
        public string LastError { get; set; }

        public string DisplayTarget =>
            string.Equals(Transport, "stdio", StringComparison.OrdinalIgnoreCase)
                ? $"{Command} {Args}".Trim()
                : Url ?? "";
    }
}
