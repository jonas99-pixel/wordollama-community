using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.lib.Tools
{
    public sealed class FetchUrlTool : IAgentTool, ICancellableAgentTool
    {
        private const int MaxBytes = 1024 * 1024;
        public string Name => "fetch_url";
        public string Description => "Safely fetch text from a public HTTPS URL. Blocks local/private addresses, credentials, unsafe MIME types, excessive redirects and responses over 1 MiB.";
        public bool IsWriteOperation => false;
        public bool IsConcurrencySafe => true;
        public int MaxResultSizeChars => 120000;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["url"] = new JObject { ["type"] = "string", ["description"] = "A public credential-free HTTPS URL." },
                ["timeout_seconds"] = new JObject { ["type"] = "integer", ["minimum"] = 1, ["maximum"] = 60 }
            },
            ["required"] = new JArray("url")
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app) =>
            ExecuteAsync(parameters, doc, app, CancellationToken.None).GetAwaiter().GetResult();

        public async Task<string> ExecuteAsync(JObject parameters, Word.Document doc, Word.Application app, CancellationToken cancellationToken)
        {
            string value = parameters?["url"]?.ToString() ?? "";
            int timeout = Math.Max(1, Math.Min(60, parameters?["timeout_seconds"]?.Value<int?>() ?? 30));
            if (!Uri.TryCreate(value, UriKind.Absolute, out Uri current)) throw new InvalidOperationException("fetch_url requires an absolute HTTPS URL.");
            for (int redirects = 0; redirects <= 5; redirects++)
            {
                ValidatePublicHttps(current);
                var request = (HttpWebRequest)WebRequest.Create(current);
                request.Method = "GET";
                request.AllowAutoRedirect = false;
                request.Timeout = timeout * 1000;
                request.ReadWriteTimeout = timeout * 1000;
                request.UserAgent = "WordOllama/2 safe-fetch";
                request.Accept = "text/html,text/plain,text/markdown,application/json,application/xml;q=0.9";
                using (var requestCancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken))
                using (requestCancellation.Token.Register(() => { try { request.Abort(); } catch { } }))
                {
                    requestCancellation.CancelAfter(TimeSpan.FromSeconds(timeout));
                using (var response = (HttpWebResponse)await request.GetResponseAsync().ConfigureAwait(false))
                {
                    int status = (int)response.StatusCode;
                    if (status == 301 || status == 302 || status == 303 || status == 307 || status == 308)
                    {
                        string location = response.Headers[HttpResponseHeader.Location];
                        if (redirects == 5 || string.IsNullOrWhiteSpace(location)) throw new InvalidOperationException("fetch_url exceeded the redirect limit.");
                        current = Uri.TryCreate(location, UriKind.Absolute, out Uri absolute) ? absolute : new Uri(current, location);
                        continue;
                    }
                    string mediaType = (response.ContentType ?? "").Split(';')[0].Trim().ToLowerInvariant();
                    string[] allowed = { "text/html", "text/plain", "text/markdown", "text/xml", "application/json", "application/xml", "application/xhtml+xml" };
                    if (!allowed.Contains(mediaType)) throw new InvalidOperationException("fetch_url rejected MIME type '" + mediaType + "'.");
                    if (response.ContentLength > MaxBytes) throw new InvalidOperationException("fetch_url response exceeds 1 MiB.");
                    byte[] bytes = await ReadLimitedAsync(response.GetResponseStream(), requestCancellation.Token).ConfigureAwait(false);
                    string text = Encoding.UTF8.GetString(bytes);
                    string title = "";
                    if (mediaType.Contains("html"))
                    {
                        Match titleMatch = Regex.Match(text, "<title[^>]*>(?<v>[\\s\\S]*?)</title>", RegexOptions.IgnoreCase, TimeSpan.FromSeconds(1));
                        title = titleMatch.Success ? WebUtility.HtmlDecode(titleMatch.Groups["v"].Value).Trim() : "";
                        text = Regex.Replace(text, "<(script|style|noscript)[^>]*>[\\s\\S]*?</\\1>", " ", RegexOptions.IgnoreCase, TimeSpan.FromSeconds(1));
                        text = WebUtility.HtmlDecode(Regex.Replace(text, "<[^>]+>", "\n", RegexOptions.None, TimeSpan.FromSeconds(1)));
                    }
                    if (text.Length > 120000) text = text.Substring(0, 120000);
                    return JsonConvert.SerializeObject(new { url = current.AbsoluteUri, status, mediaType, title, text, bytes = bytes.Length });
                }
                }
            }
            throw new InvalidOperationException("fetch_url failed.");
        }

        private static void ValidatePublicHttps(Uri uri)
        {
            if (uri.Scheme != Uri.UriSchemeHttps || string.IsNullOrWhiteSpace(uri.Host) || !string.IsNullOrWhiteSpace(uri.UserInfo))
                throw new InvalidOperationException("fetch_url only permits credential-free HTTPS URLs.");
            IPAddress[] addresses = Dns.GetHostAddresses(uri.DnsSafeHost);
            if (addresses.Length == 0 || addresses.Any(IsPrivateOrSpecial))
                throw new InvalidOperationException("fetch_url blocks loopback, private, link-local and special network addresses.");
        }

        private static bool IsPrivateOrSpecial(IPAddress address)
        {
            if (address.IsIPv4MappedToIPv6) address = address.MapToIPv4();
            if (IPAddress.IsLoopback(address) || address.Equals(IPAddress.Any) || address.Equals(IPAddress.IPv6Any) ||
                address.IsIPv6LinkLocal || address.IsIPv6SiteLocal || address.IsIPv6Multicast) return true;
            byte[] bytes = address.GetAddressBytes();
            if (address.AddressFamily == AddressFamily.InterNetwork)
                return bytes[0] == 0 || bytes[0] == 10 || bytes[0] == 127 || bytes[0] >= 224 ||
                    (bytes[0] == 169 && bytes[1] == 254) || (bytes[0] == 172 && bytes[1] >= 16 && bytes[1] <= 31) ||
                    (bytes[0] == 192 && bytes[1] == 168);
            return (bytes[0] & 0xfe) == 0xfc || bytes.All(b => b == 0);
        }

        private static async Task<byte[]> ReadLimitedAsync(Stream stream, CancellationToken cancellationToken)
        {
            using (var output = new MemoryStream())
            {
                byte[] buffer = new byte[16384];
                while (true)
                {
                    int read = await stream.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false);
                    if (read == 0) return output.ToArray();
                    if (output.Length + read > MaxBytes) throw new InvalidOperationException("fetch_url response exceeds 1 MiB.");
                    output.Write(buffer, 0, read);
                }
            }
        }
    }

    public static class AgentWorkspaceManager
    {
        private static readonly object Sync = new object();
        private static string _sessionId;
        private static string _root;
        public static string BeginSession()
        {
            lock (Sync)
            {
                _sessionId = Guid.NewGuid().ToString("N");
                _root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WordOllama", "AgentWorkspaces", _sessionId);
                Directory.CreateDirectory(_root);
                return _sessionId;
            }
        }
        public static string CurrentRoot
        {
            get { lock (Sync) { if (string.IsNullOrWhiteSpace(_root)) BeginSession(); return _root; } }
        }
        public static void EndSession()
        {
            string root;
            lock (Sync) { root = _root; _root = null; _sessionId = null; }
            if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root)) return;
            try
            {
                foreach (string entry in Directory.EnumerateFileSystemEntries(root, "*", SearchOption.AllDirectories))
                    if ((File.GetAttributes(entry) & FileAttributes.ReparsePoint) != 0) return;
                Directory.Delete(root, true);
            }
            catch { }
        }
        public static string Resolve(string relative, bool allowRoot = false)
        {
            string root = Path.GetFullPath(CurrentRoot);
            string path = Path.GetFullPath(Path.Combine(root, relative ?? ""));
            string prefix = root.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if ((!allowRoot && string.Equals(path, root, StringComparison.OrdinalIgnoreCase)) ||
                (!string.Equals(path, root, StringComparison.OrdinalIgnoreCase) && !path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("Workspace path escapes the isolated task directory.");
            string cursor = path;
            while (!string.Equals(cursor, root, StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(cursor))
            {
                if ((File.Exists(cursor) || Directory.Exists(cursor)) && (File.GetAttributes(cursor) & FileAttributes.ReparsePoint) != 0)
                    throw new InvalidOperationException("Workspace reparse points are not allowed.");
                cursor = Path.GetDirectoryName(cursor);
            }
            return path;
        }
        public static void EnforceTotalLimit(long additional)
        {
            long total = Directory.Exists(CurrentRoot) ? Directory.EnumerateFiles(CurrentRoot, "*", SearchOption.AllDirectories).Sum(file => new FileInfo(file).Length) : 0;
            if (total + additional > 10 * 1024 * 1024) throw new InvalidOperationException("Workspace exceeds the 10 MiB limit.");
        }
    }

    public abstract class WorkspaceToolBase : IAgentTool
    {
        public abstract string Name { get; }
        public abstract string Description { get; }
        public abstract JObject ParameterSchema { get; }
        public virtual bool IsWriteOperation => false;
        public bool IsConcurrencySafe => true;
        public int MaxResultSizeChars => 120000;
        public abstract string Execute(JObject parameters, Word.Document doc, Word.Application app);
    }

    public sealed class ListWorkspaceFilesTool : WorkspaceToolBase
    {
        public override string Name => "list_workspace_files";
        public override string Description => "List files in this Agent task's isolated workspace.";
        public override JObject ParameterSchema => Schema.OptionalPath();
        public override string Execute(JObject p, Word.Document doc, Word.Application app)
        {
            string dir = AgentWorkspaceManager.Resolve(p?["path"]?.ToString() ?? "", true);
            if (!Directory.Exists(dir)) throw new DirectoryNotFoundException();
            string root = AgentWorkspaceManager.CurrentRoot.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            return JsonConvert.SerializeObject(Directory.EnumerateFiles(dir, "*", SearchOption.AllDirectories).Take(500)
                .Select(file => file.StartsWith(root, StringComparison.OrdinalIgnoreCase) ? file.Substring(root.Length).Replace('\\', '/') : file));
        }
    }

    public sealed class ReadWorkspaceFileTool : WorkspaceToolBase
    {
        public override string Name => "read_workspace_file";
        public override string Description => "Read a UTF-8 text file from this Agent task's isolated workspace.";
        public override JObject ParameterSchema => Schema.RequiredPath();
        public override string Execute(JObject p, Word.Document doc, Word.Application app)
        {
            string path = AgentWorkspaceManager.Resolve(p?["path"]?.ToString());
            if (new FileInfo(path).Length > 1024 * 1024) throw new InvalidOperationException("Workspace file exceeds 1 MiB.");
            return File.ReadAllText(path, Encoding.UTF8);
        }
    }

    public sealed class WriteWorkspaceFileTool : WorkspaceToolBase
    {
        public override string Name => "write_workspace_file";
        public override string Description => "Create or replace a UTF-8 text file inside this Agent task's isolated workspace.";
        public override JObject ParameterSchema => new JObject { ["type"] = "object", ["properties"] = new JObject { ["path"] = new JObject { ["type"] = "string" }, ["content"] = new JObject { ["type"] = "string" } }, ["required"] = new JArray("path", "content") };
        public override string Execute(JObject p, Word.Document doc, Word.Application app)
        {
            string content = p?["content"]?.ToString() ?? "";
            byte[] bytes = Encoding.UTF8.GetBytes(content);
            if (bytes.Length > 1024 * 1024) throw new InvalidOperationException("Workspace file exceeds 1 MiB.");
            AgentWorkspaceManager.EnforceTotalLimit(bytes.Length);
            string path = AgentWorkspaceManager.Resolve(p?["path"]?.ToString());
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.WriteAllText(path, content, new UTF8Encoding(false));
            return JsonConvert.SerializeObject(new { path = p?["path"]?.ToString(), bytes = bytes.Length });
        }
    }

    public abstract class RunWorkspaceCodeTool : WorkspaceToolBase, ICancellableAgentTool
    {
        protected abstract string Runtime { get; }
        public override JObject ParameterSchema => new JObject { ["type"] = "object", ["properties"] = new JObject { ["code"] = new JObject { ["type"] = "string" }, ["timeout_seconds"] = new JObject { ["type"] = "integer", ["minimum"] = 1, ["maximum"] = 120 } }, ["required"] = new JArray("code") };
        public override string Execute(JObject p, Word.Document doc, Word.Application app) => ExecuteAsync(p, doc, app, CancellationToken.None).GetAwaiter().GetResult();
        public async Task<string> ExecuteAsync(JObject p, Word.Document doc, Word.Application app, CancellationToken ct)
        {
            string code = p?["code"]?.ToString() ?? "";
            if (code.Length > 100000) throw new InvalidOperationException("Code exceeds 100,000 characters.");
            int timeout = Math.Max(1, Math.Min(120, p?["timeout_seconds"]?.Value<int?>() ?? 30));
            string extension = Runtime == "python" ? ".py" : ".js";
            string codePath = AgentWorkspaceManager.Resolve(".wordollama-task" + extension);
            File.WriteAllText(codePath, code, new UTF8Encoding(false));
            string exe = Runtime == "python" ? "python.exe" : "node.exe";
            var start = new ProcessStartInfo
            {
                FileName = exe,
                Arguments = "\"" + codePath + "\"",
                WorkingDirectory = AgentWorkspaceManager.CurrentRoot,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            start.EnvironmentVariables["PYTHONNOUSERSITE"] = "1";
            using (var process = Process.Start(start))
            using (var job = WindowsProcessJob.Attach(process))
            using (ct.Register(() => { try { if (!process.HasExited) process.Kill(); } catch { } }))
            {
                Task<string> stdout = process.StandardOutput.ReadToEndAsync();
                Task<string> stderr = process.StandardError.ReadToEndAsync();
                bool exited = await Task.Run(() => process.WaitForExit(timeout * 1000), ct).ConfigureAwait(false);
                if (!exited) { try { process.Kill(); } catch { } }
                string output = await stdout.ConfigureAwait(false);
                string error = await stderr.ConfigureAwait(false);
                if (output.Length > 120000) output = output.Substring(0, 120000);
                if (error.Length > 20000) error = error.Substring(0, 20000);
                return JsonConvert.SerializeObject(new { exitCode = exited ? process.ExitCode : -1, stdout = output, stderr = error, timedOut = !exited });
            }
        }
    }

    public sealed class RunPythonWorkspaceTool : RunWorkspaceCodeTool
    {
        protected override string Runtime => "python";
        public override string Name => "run_python";
        public override string Description => "Run Python code in the current task workspace with file, time and output limits. Requires explicit approval.";
    }
    public sealed class RunNodeWorkspaceTool : RunWorkspaceCodeTool
    {
        protected override string Runtime => "node";
        public override string Name => "run_node";
        public override string Description => "Run Node.js code in the current task workspace with file, time and output limits. Requires explicit approval.";
    }

    internal static class Schema
    {
        public static JObject OptionalPath() => new JObject { ["type"] = "object", ["properties"] = new JObject { ["path"] = new JObject { ["type"] = "string" } } };
        public static JObject RequiredPath() { JObject value = OptionalPath(); value["required"] = new JArray("path"); return value; }
    }

    internal sealed class WindowsProcessJob : IDisposable
    {
        private IntPtr _handle;
        private WindowsProcessJob(IntPtr handle) { _handle = handle; }
        public static WindowsProcessJob Attach(Process process)
        {
            if (process == null) throw new ArgumentNullException(nameof(process));
            IntPtr handle = CreateJobObject(IntPtr.Zero, null);
            if (handle == IntPtr.Zero) throw new InvalidOperationException("Unable to create a Windows job object.");
            var info = new JOBOBJECT_EXTENDED_LIMIT_INFORMATION();
            info.BasicLimitInformation.LimitFlags = 0x00002000 | 0x00000008 | 0x00000100;
            info.BasicLimitInformation.ActiveProcessLimit = 8;
            info.ProcessMemoryLimit = new UIntPtr(512UL * 1024UL * 1024UL);
            int length = Marshal.SizeOf(info);
            IntPtr buffer = Marshal.AllocHGlobal(length);
            try
            {
                Marshal.StructureToPtr(info, buffer, false);
                if (!SetInformationJobObject(handle, 9, buffer, (uint)length) || !AssignProcessToJobObject(handle, process.Handle))
                    throw new InvalidOperationException("Unable to apply the Windows process limits.");
                return new WindowsProcessJob(handle);
            }
            catch { CloseHandle(handle); throw; }
            finally { Marshal.FreeHGlobal(buffer); }
        }
        public void Dispose() { if (_handle != IntPtr.Zero) { CloseHandle(_handle); _handle = IntPtr.Zero; } }

        [StructLayout(LayoutKind.Sequential)] private struct IO_COUNTERS { public ulong ReadOperationCount, WriteOperationCount, OtherOperationCount, ReadTransferCount, WriteTransferCount, OtherTransferCount; }
        [StructLayout(LayoutKind.Sequential)] private struct JOBOBJECT_BASIC_LIMIT_INFORMATION { public long PerProcessUserTimeLimit, PerJobUserTimeLimit; public uint LimitFlags; public UIntPtr MinimumWorkingSetSize, MaximumWorkingSetSize; public uint ActiveProcessLimit; public UIntPtr Affinity; public uint PriorityClass, SchedulingClass; }
        [StructLayout(LayoutKind.Sequential)] private struct JOBOBJECT_EXTENDED_LIMIT_INFORMATION { public JOBOBJECT_BASIC_LIMIT_INFORMATION BasicLimitInformation; public IO_COUNTERS IoInfo; public UIntPtr ProcessMemoryLimit, JobMemoryLimit, PeakProcessMemoryUsed, PeakJobMemoryUsed; }
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr CreateJobObject(IntPtr attributes, string name);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool SetInformationJobObject(IntPtr job, int infoClass, IntPtr info, uint length);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);
        [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr handle);
    }
}
