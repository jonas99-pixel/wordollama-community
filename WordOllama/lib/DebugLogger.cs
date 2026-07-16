using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace WordOllama.lib
{
    /// <summary>
    /// 调试用日志记录器 — 记录 Agent/API 的输入输出到本地文件。
    /// 通过 <see cref="Enabled"/> 开关控制，false 时无任何开销。
    /// 日志文件：%LOCALAPPDATA%/WordOllama/debug.log
    /// </summary>
    public static class DebugLogger
    {
        /// <summary>总开关。设为 false 则所有写入变为 no-op。</summary>
        public static bool Enabled { get; set; }
        public static bool IncludeContent { get; private set; }

        private static readonly object _lock = new object();
        private const long MaxLogBytes = 5 * 1024 * 1024;
        private static string _logPath;

        private static string LogPath
        {
            get
            {
                if (_logPath == null)
                {
                    var dir = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "WordOllama");
                    if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                    _logPath = Path.Combine(dir, "debug.log");
                }
                return _logPath;
            }
        }

        /// <summary>
        /// 记录一条日志。category 用于标记来源（如 "Ollama.Generate", "Agent.Tool"）。
        /// </summary>
        public static void Log(string category, string content)
        {
            if (!Enabled) return;
            try
            {
                lock (_lock)
                {
                    File.AppendAllText(LogPath,
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{category}] {Redact(content)}{Environment.NewLine}",
                        Encoding.UTF8);
                    TrimIfNeeded();
                }
            }
            catch { /* 日志写入失败不应影响主流程 */ }
        }

        /// <summary>
        /// 记录带分隔线的长内容（prompt/response），便于阅读。
        /// </summary>
        public static void LogBlock(string category, string label, string content)
        {
            if (!Enabled) return;
            try
            {
                lock (_lock)
                {
                    string value = IncludeContent
                        ? Redact(content)
                        : $"[content omitted; {content?.Length ?? 0} characters]";
                    File.AppendAllText(LogPath,
                        $"\n[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{category}] ===== {label} =====\n{value}\n===== END {label} =====\n\n",
                        Encoding.UTF8);
                    TrimIfNeeded();
                }
            }
            catch { }
        }

        /// <summary>Enables diagnostic logging for the current Word session only.</summary>
        public static void EnableForCurrentSession(bool includeContent = false)
        {
            Enabled = true;
            IncludeContent = includeContent;
        }

        public static void Disable()
        {
            Enabled = false;
            IncludeContent = false;
        }

        /// <summary>清空日志文件。</summary>
        public static void Clear()
        {
            try
            {
                lock (_lock)
                {
                    if (File.Exists(LogPath))
                        File.WriteAllText(LogPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] === 日志已清空 ===\n", Encoding.UTF8);
                }
            }
            catch { }
        }

        /// <summary>获取日志文件路径。</summary>
        public static string GetLogPath() => LogPath;

        /// <summary>Returns timestamps and categories only for support bundles.</summary>
        public static string GetMetadataSummary(int maxLines = 20)
        {
            try
            {
                if (!File.Exists(LogPath)) return "(no diagnostic events)";
                string[] lines = File.ReadAllLines(LogPath, Encoding.UTF8);
                var summary = new System.Collections.Generic.List<string>();
                for (int i = Math.Max(0, lines.Length - Math.Max(1, maxLines)); i < lines.Length; i++)
                {
                    Match match = Regex.Match(lines[i], @"^\[(?<time>[^\]]+)\]\s+\[(?<category>[^\]]+)\]");
                    if (!match.Success) continue;
                    bool error = lines[i].IndexOf("error", StringComparison.OrdinalIgnoreCase) >= 0
                        || lines[i].IndexOf("exception", StringComparison.OrdinalIgnoreCase) >= 0
                        || lines[i].IndexOf("失败", StringComparison.OrdinalIgnoreCase) >= 0
                        || lines[i].IndexOf("错误", StringComparison.OrdinalIgnoreCase) >= 0;
                    summary.Add($"{match.Groups["time"].Value} [{match.Groups["category"].Value}] {(error ? "error" : "event")}");
                }
                return summary.Count == 0 ? "(no diagnostic events)" : string.Join(Environment.NewLine, summary);
            }
            catch
            {
                return "(diagnostic metadata unavailable)";
            }
        }

        private static string Redact(string content)
        {
            if (string.IsNullOrEmpty(content)) return string.Empty;

            string result = Regex.Replace(content,
                @"(?i)(authorization\s*[:=]\s*(?:bearer\s+)?)[^\s,;\""']+",
                "$1[REDACTED]");
            return Regex.Replace(result,
                @"(?i)(api[_-]?key\s*[:=]\s*)[^\s,;\""']+",
                "$1[REDACTED]");
        }

        private static void TrimIfNeeded()
        {
            try
            {
                var file = new FileInfo(LogPath);
                if (file.Length <= MaxLogBytes) return;

                string text = File.ReadAllText(LogPath, Encoding.UTF8);
                int start = Math.Max(0, text.Length - (int)(MaxLogBytes * 0.75));
                int lineStart = text.IndexOf('\n', start);
                if (lineStart >= 0) start = lineStart + 1;
                File.WriteAllText(LogPath, "[diagnostic log truncated]\n" + text.Substring(start), Encoding.UTF8);
            }
            catch { }
        }
    }
}
