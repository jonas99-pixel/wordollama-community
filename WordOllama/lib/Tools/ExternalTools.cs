using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.lib.Tools
{
    /// <summary>
    /// 执行外部命令/Python 脚本，返回输出结果。
    /// 用于 Skill 中需要调用 Python 等外部程序的场景。
    /// </summary>
    public class ExecuteCommandTool : IAgentTool, ICancellableAgentTool
    {
        public string Name => "execute_command";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => false;
        public bool IsConcurrencySafe => false;
        public int MaxResultSizeChars => 10000;
        public JObject ParameterSchema =>
            new JObject
            {
                ["type"] = "object",
                ["properties"] = new JObject
                {
                    ["command"] = new JObject
                    {
                        ["type"] = "string"
                    },
                    ["args"] = new JObject { ["type"] = "string" },
                    ["timeout_seconds"] = new JObject
                    {
                        ["type"] = "integer"
                    }
                },
                ["required"] = new JArray("command")
            };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
            => ExecuteCore(parameters, doc, app, CancellationToken.None);

        public Task<string> ExecuteAsync(JObject parameters, Word.Document doc, Word.Application app, CancellationToken cancellationToken)
            => Task.Run(() => ExecuteCore(parameters, doc, app, cancellationToken), cancellationToken);

        private string ExecuteCore(JObject parameters, Word.Document doc, Word.Application app, CancellationToken cancellationToken)
        {
            string command = parameters["command"]?.ToString() ?? "";
            string args = parameters["args"]?.ToString() ?? "";
            int timeout = parameters["timeout_seconds"]?.ToObject<int>() ?? 30;

            if (string.IsNullOrEmpty(command))
                return AgentToolText.T("I18n_AgentResultCommandRequired");

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = command,
                    Arguments = args,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8
                };
                if (command.IndexOf("python", StringComparison.OrdinalIgnoreCase) >= 0)
                    psi.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";

                using (var process = new Process { StartInfo = psi })
                {
                    var stdoutBuilder = new StringBuilder();
                    var stderrBuilder = new StringBuilder();

                    process.OutputDataReceived += (_, e) =>
                    {
                        if (e.Data != null)
                            stdoutBuilder.AppendLine(e.Data);
                    };
                    process.ErrorDataReceived += (_, e) =>
                    {
                        if (e.Data != null)
                            stderrBuilder.AppendLine(e.Data);
                    };

                    process.Start();
                    cancellationToken.Register(() =>
                    {
                        try { if (!process.HasExited) process.Kill(); } catch { }
                    });
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();

                    if (process.WaitForExit(timeout * 1000))
                    {
                        if (cancellationToken.IsCancellationRequested)
                            return AgentToolText.T("I18n_AgentResultCommandCancelled");
                        string stdout = stdoutBuilder.ToString().Trim();
                        string stderr = stderrBuilder.ToString().Trim();

                        var result = new StringBuilder();
                        result.AppendLine(AgentToolText.F("I18n_AgentResultCommandLabel", command, args));
                        result.AppendLine(AgentToolText.F("I18n_AgentResultExitCodeLabel", process.ExitCode));
                        if (!string.IsNullOrEmpty(stdout))
                            result.AppendLine(AgentToolText.F("I18n_AgentResultOutputLabel", stdout));
                        if (!string.IsNullOrEmpty(stderr))
                            result.AppendLine(AgentToolText.F("I18n_AgentResultErrorLabel", stderr));

                        return result.ToString().Trim();
                    }
                    else
                    {
                        process.Kill();
                        string partial = stdoutBuilder.ToString().Trim();
                        string partialErr = stderrBuilder.ToString().Trim();
                        return AgentToolText.F("I18n_AgentResultCommandTimedOut", command, timeout, partial, partialErr).Trim();
                    }
                }
            }
            catch (Exception ex)
            {
                return AgentToolText.F("I18n_AgentResultCommandFailed", ex.Message);
            }
        }
    }

    /// <summary>
    /// HTTP 请求工具 — 让 AI 可以调用外部 REST API。
    /// 支持 GET/POST/PUT/DELETE/PATCH，自定义 Header 和 Body。
    /// </summary>
    public class HttpRequestTool : IAgentTool, ICancellableAgentTool
    {
        public string Name => "http_request";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => false;
        public bool IsConcurrencySafe => false;
        public int MaxResultSizeChars => 10000;
        public JObject ParameterSchema =>
            new JObject
            {
                ["type"] = "object",
                ["properties"] = new JObject
                {
                    ["url"] = new JObject { ["type"] = "string" },
                    ["method"] = new JObject
                    {
                        ["type"] = "string"
                    },
                    ["headers"] = new JObject
                    {
                        ["type"] = "string"
                    },
                    ["body"] = new JObject
                    {
                        ["type"] = "string"
                    }
                },
                ["required"] = new JArray("url")
            };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
            => ExecuteCore(parameters, doc, app, CancellationToken.None);

        public Task<string> ExecuteAsync(JObject parameters, Word.Document doc, Word.Application app, CancellationToken cancellationToken)
            => Task.Run(() => ExecuteCore(parameters, doc, app, cancellationToken), cancellationToken);

        private string ExecuteCore(JObject parameters, Word.Document doc, Word.Application app, CancellationToken cancellationToken)
        {
            string url = parameters["url"]?.ToString() ?? "";
            string method = parameters["method"]?.ToString()?.ToUpper() ?? "GET";
            string headersJson = parameters["headers"]?.ToString() ?? "";
            string body = parameters["body"]?.ToString() ?? "";

            if (string.IsNullOrEmpty(url))
                return AgentToolText.T("I18n_AgentResultUrlRequired");

            try
            {
                var request = (HttpWebRequest)WebRequest.Create(url);
                cancellationToken.Register(() =>
                {
                    try { request.Abort(); } catch { }
                });
                request.Method = method;
                request.Timeout = 60000;
                request.ContentType = DetectContentType(headersJson);

                // 解析自定义请求头
                if (!string.IsNullOrEmpty(headersJson))
                {
                    try
                    {
                        var headers = JObject.Parse(headersJson);
                        foreach (var prop in headers.Properties())
                        {
                            string name = prop.Name;
                            string value = prop.Value?.ToString() ?? "";

                            // Content-Type 已通过属性设置，跳过
                            if (
                                string.Equals(
                                    name,
                                    "Content-Type",
                                    StringComparison.OrdinalIgnoreCase
                                )
                            )
                                continue;

                            try
                            {
                                request.Headers.Add(name, value);
                            }
                            catch
                            { /* 忽略无法添加的请求头 */
                            }
                        }
                    }
                    catch
                    {
                        return AgentToolText.F("I18n_AgentResultInvalidHeaders", headersJson);
                    }
                }

                // 写入请求体（POST/PUT/PATCH）
                if (
                    !string.IsNullOrEmpty(body)
                    && (method == "POST" || method == "PUT" || method == "PATCH")
                )
                {
                    byte[] data = Encoding.UTF8.GetBytes(body);
                    request.ContentLength = data.Length;
                    using (var stream = request.GetRequestStream())
                    {
                        stream.Write(data, 0, data.Length);
                    }
                }
                else if (method != "GET" && method != "DELETE")
                {
                    return AgentToolText.F("I18n_AgentResultUnsupportedHttpMethod", method);
                }

                // 发送请求并读取响应
                try
                {
                    using (var response = (HttpWebResponse)request.GetResponse())
                    {
                        using (
                            var reader = new StreamReader(
                                response.GetResponseStream(),
                                Encoding.UTF8
                            )
                        )
                        {
                            string responseBody = reader.ReadToEnd();
                            int maxBodyLen = 8000;
                            if (responseBody.Length > maxBodyLen)
                                responseBody =
                                    responseBody.Substring(0, maxBodyLen)
                                    + AgentToolText.F("I18n_AgentResultResponseTruncated", responseBody.Length);

                            return AgentToolText.F("I18n_AgentResultHttpResponse", (int)response.StatusCode, response.StatusDescription, url, responseBody).Trim();
                        }
                    }
                }
                catch (WebException ex)
                {
                    if (cancellationToken.IsCancellationRequested)
                        return AgentToolText.T("I18n_AgentResultHttpCancelled");
                    // 提取错误响应体
                    string errorBody = "";
                    if (ex.Response != null)
                    {
                        using (
                            var reader = new StreamReader(
                                ex.Response.GetResponseStream(),
                                Encoding.UTF8
                            )
                        )
                        {
                            errorBody = reader.ReadToEnd();
                            if (errorBody.Length > 2000)
                                errorBody = errorBody.Substring(0, 2000) + "...";
                        }

                        var httpResp = ex.Response as HttpWebResponse;
                        if (httpResp != null)
                        {
                            return AgentToolText.F("I18n_AgentResultHttpErrorResponse", (int)httpResp.StatusCode, httpResp.StatusDescription, url, errorBody).Trim();
                        }
                    }
                    return AgentToolText.F("I18n_AgentResultHttpFailed", ex.Message);
                }
            }
            catch (Exception ex)
            {
                return AgentToolText.F("I18n_AgentResultHttpFailed", ex.Message);
            }
        }

        private static string DetectContentType(string headersJson)
        {
            if (string.IsNullOrEmpty(headersJson))
                return "application/json";
            try
            {
                var headers = JObject.Parse(headersJson);
                return headers["Content-Type"]?.ToString() ?? "application/json";
            }
            catch
            {
                return "application/json";
            }
        }
    }

    /// <summary>
    /// 运行技能目录中的 Python 脚本。
    /// 专门用于执行 WordAgentSkills 下的 Python 脚本（如企业信息查询、法律检索等）。
    /// </summary>
    public class RunPythonScriptTool : IAgentTool, ICancellableAgentTool
    {
        public string Name => "run_python_script";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => false;
        public bool IsConcurrencySafe => false;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema =>
            new JObject
            {
                ["type"] = "object",
                ["properties"] = new JObject
                {
                    ["skill_name"] = new JObject
                    {
                        ["type"] = "string"
                    },
                    ["script_args"] = new JObject
                    {
                        ["type"] = "string"
                    }
                },
                ["required"] = new JArray("skill_name")
            };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
            => ExecuteCore(parameters, doc, app, CancellationToken.None);

        public Task<string> ExecuteAsync(JObject parameters, Word.Document doc, Word.Application app, CancellationToken cancellationToken)
            => Task.Run(() => ExecuteCore(parameters, doc, app, cancellationToken), cancellationToken);

        private string ExecuteCore(JObject parameters, Word.Document doc, Word.Application app, CancellationToken cancellationToken)
        {
            string skillName = parameters["skill_name"]?.ToString()?.Trim() ?? "";
            string scriptArgs = parameters["script_args"]?.ToString()?.Trim() ?? "";

            if (string.IsNullOrEmpty(skillName))
                return AgentToolText.T("I18n_AgentResultSkillNameRequired");

            // 优先 MyDocuments，其次 OneDrive 文档目录
            string skillsDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "WordAgentSkills"
            );
            string skillDir = Path.GetFullPath(Path.Combine(skillsDir, skillName));
            if (!IsPathUnderRoot(skillDir, skillsDir))
            {
                return AgentToolText.F("I18n_AgentResultSkillDirectoryNotFound", skillName);
            }

            if (!Directory.Exists(skillDir))
            {
                // 尝试 OneDrive 路径
                string onedriveDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                    "OneDrive",
                    "文档",
                    "WordAgentSkills"
                );
                string onedriveSkillDir = Path.GetFullPath(Path.Combine(onedriveDir, skillName));
                if (!IsPathUnderRoot(onedriveSkillDir, onedriveDir))
                {
                    return AgentToolText.F("I18n_AgentResultSkillDirectoryNotFound", skillName);
                }
                if (Directory.Exists(onedriveSkillDir))
                    skillDir = onedriveSkillDir;
                else
                    return AgentToolText.F("I18n_AgentResultSkillDirectoryNotFound", skillName);
            }

            // 查找 skill 目录下的 Python 脚本
            string[] pyFiles = Directory.GetFiles(skillDir, "*.py", SearchOption.AllDirectories);
            if (pyFiles.Length == 0)
                return AgentToolText.F("I18n_AgentResultPythonScriptNotFound", skillName);

            // 优先使用 scripts/ 子目录下的脚本，否则用第一个找到的
            string scriptPath =
                pyFiles.FirstOrDefault(f => f.Contains("\\scripts\\")) ?? pyFiles[0];

            // 检查 .env 是否存在（仅反馈，不阻塞执行）
            string envPath = Path.Combine(skillDir, ".env");
            bool hasEnv = File.Exists(envPath);

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "python",
                    Arguments = $"\"{scriptPath}\" {scriptArgs}",
                    WorkingDirectory = skillDir,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8
                };
                psi.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";

                using var process = new Process { StartInfo = psi };
                process.Start();
                using var cancellationRegistration = cancellationToken.Register(() =>
                {
                    try { if (!process.HasExited) process.Kill(); } catch { }
                });

                // Start both readers before waiting so a full stderr buffer cannot block stdout.
                Task<string> outputTask = process.StandardOutput.ReadToEndAsync();
                Task<string> errorTask = process.StandardError.ReadToEndAsync();
                Task exitTask = Task.Run(() => process.WaitForExit());
                bool completed = Task.WaitAny(new[] { exitTask }, 60000) == 0;
                if (cancellationToken.IsCancellationRequested)
                {
                    exitTask.GetAwaiter().GetResult();
                    return AgentToolText.T("I18n_AgentResultPythonCancelled");
                }
                if (!completed)
                {
                    try
                    {
                        if (!process.HasExited) process.Kill();
                    }
                    catch { }
                    exitTask.GetAwaiter().GetResult();
                }

                string output = outputTask.GetAwaiter().GetResult();
                string error = errorTask.GetAwaiter().GetResult();
                if (!completed)
                    return AgentToolText.F("I18n_AgentResultPythonTimedOut", output, error).Trim();

                if (process.ExitCode != 0)
                {
                    string errMsg;
                    if (!string.IsNullOrEmpty(error))
                        errMsg = error;
                    else if (!string.IsNullOrEmpty(output))
                        errMsg = $"(stdout)\n{output}";
                    else
                        errMsg = AgentToolText.T("I18n_AgentResultNoErrorOutput");
                    // API Key 缺失时给出明确的 .env 指引
                    if (errMsg.Contains("YUANDIAN_API_KEY") || errMsg.Contains("API_KEY"))
                    {
                        string hint = hasEnv
                            ? AgentToolText.F("I18n_AgentResultEnvFilePresent", envPath)
                            : AgentToolText.F("I18n_AgentResultEnvFileMissing", envPath);
                        return AgentToolText.F("I18n_AgentResultScriptFailedWithHint", process.ExitCode, errMsg, hint);
                    }
                    return AgentToolText.F("I18n_AgentResultScriptFailed", process.ExitCode, errMsg);
                }

                if (!string.IsNullOrEmpty(error))
                    output += AgentToolText.F("I18n_AgentResultStdErrLabel", error);

                return string.IsNullOrEmpty(output) ? AgentToolText.T("I18n_AgentResultScriptSucceededNoOutput") : output;
            }
            catch (Exception ex)
            {
                return AgentToolText.F("I18n_AgentResultExecutionException", ex.Message);
            }
        }

        private static bool IsPathUnderRoot(string path, string root)
        {
            string fullPath = Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string fullRoot = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            return fullPath.StartsWith(fullRoot + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
        }
    }

    /// <summary>
    /// grep 工具 — 在文件/目录中递归搜索文本模式，返回匹配的文件路径、行号和内容。
    /// 纯 C# 实现，不依赖外部 grep 命令。
    /// </summary>
    public class GrepTool : IAgentTool
    {
        public string Name => "grep";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => false;
        public bool IsConcurrencySafe => true;
        public int MaxResultSizeChars => 15000;
        public JObject ParameterSchema =>
            new JObject
            {
                ["type"] = "object",
                ["properties"] = new JObject
                {
                    ["pattern"] = new JObject
                    {
                        ["type"] = "string"
                    },
                    ["path"] = new JObject
                    {
                        ["type"] = "string"
                    },
                    ["glob"] = new JObject
                    {
                        ["type"] = "string"
                    },
                    ["ignore_case"] = new JObject
                    {
                        ["type"] = "boolean"
                    },
                    ["max_matches"] = new JObject
                    {
                        ["type"] = "integer"
                    },
                    ["max_file_size_kb"] = new JObject
                    {
                        ["type"] = "integer"
                    }
                },
                ["required"] = new JArray("pattern")
            };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            string pattern = parameters["pattern"]?.ToString() ?? "";
            string requestedPath = parameters["path"]?.ToString();
            string glob = parameters["glob"]?.ToString() ?? "*";
            bool ignoreCase = parameters["ignore_case"]?.Value<bool>() ?? false;
            int maxMatches = Math.Min(parameters["max_matches"]?.Value<int>() ?? 50, 200);
            int maxFileSizeKb = parameters["max_file_size_kb"]?.Value<int>() ?? 1024;

            if (string.IsNullOrEmpty(pattern))
                return AgentToolText.T("I18n_AgentResultPatternRequired");

            try
            {
                string defaultRoot = GetDefaultSearchRoot(doc);
                string searchPath = string.IsNullOrWhiteSpace(requestedPath)
                    ? defaultRoot
                    : Path.GetFullPath(Path.IsPathRooted(requestedPath)
                        ? requestedPath
                        : Path.Combine(defaultRoot, requestedPath));

                if (!IsAllowedSearchPath(searchPath, doc) || HasReparsePoint(searchPath))
                    return AgentToolText.F("I18n_AgentResultGrepPathDenied", searchPath);
                if (glob.IndexOf("..", StringComparison.Ordinal) >= 0
                    || glob.IndexOf(Path.DirectorySeparatorChar) >= 0
                    || glob.IndexOf(Path.AltDirectorySeparatorChar) >= 0)
                    return AgentToolText.F("I18n_AgentResultGrepGlobDenied", glob);
                if (File.Exists(searchPath) && IsSensitiveFile(searchPath))
                    return AgentToolText.F("I18n_AgentResultGrepSensitivePathDenied", searchPath);

                var regex = new Regex(
                    pattern,
                    ignoreCase
                        ? RegexOptions.IgnoreCase | RegexOptions.Compiled
                        : RegexOptions.Compiled,
                    TimeSpan.FromSeconds(5)
                );

                // 确定搜索路径
                if (!Directory.Exists(searchPath) && !File.Exists(searchPath))
                    return AgentToolText.F("I18n_AgentResultPathNotFound", searchPath);

                IEnumerable<string> files = File.Exists(searchPath)
                    ? new[] { searchPath }
                    : EnumerateSafeFiles(searchPath, glob);

                var results = new List<string>();
                int totalMatches = 0;
                long totalBytesScanned = 0;
                int skippedLarge = 0;
                int skippedBinary = 0;
                int skippedSensitive = 0;
                int scannedFiles = 0;

                foreach (string file in files)
                {
                    if (results.Count >= maxMatches)
                        break;

                    try
                    {
                        if (!IsAllowedSearchPath(file, doc) || HasReparsePoint(file))
                            continue;
                        if (IsSensitiveFile(file))
                        {
                            skippedSensitive++;
                            continue;
                        }

                        scannedFiles++;
                        var fi = new FileInfo(file);
                        // 跳过超大文件
                        if (fi.Length > maxFileSizeKb * 1024L)
                        {
                            skippedLarge++;
                            continue;
                        }
                        // 跳过明显是二进制的文件
                        if (IsBinaryFile(file))
                        {
                            skippedBinary++;
                            continue;
                        }

                        totalBytesScanned += fi.Length;
                        string[] lines = File.ReadAllLines(file, Encoding.UTF8);

                        for (int i = 0; i < lines.Length; i++)
                        {
                            if (results.Count >= maxMatches)
                                break;
                            if (regex.IsMatch(lines[i]))
                            {
                                string lineNum = (i + 1).ToString();
                                string displayPath = file;
                                // 如果路径在搜索目录下，用相对路径简洁些
                                if (Directory.Exists(searchPath) && IsPathUnderRootOrEqual(file, searchPath))
                                    displayPath = file.Substring(searchPath.Length)
                                        .TrimStart('\\', '/');
                                string preview = lines[i].Trim();
                                if (preview.Length > 200)
                                    preview = preview.Substring(0, 200) + "...";
                                results.Add($"{displayPath}:{lineNum}: {preview}");
                                totalMatches++;
                            }
                        }
                    }
                    catch (UnauthorizedAccessException)
                    {
                        continue;
                    }
                    catch (IOException)
                    {
                        continue;
                    }
                }

                if (scannedFiles == 0 && skippedSensitive == 0)
                    return AgentToolText.T("I18n_AgentResultNoMatchingFiles");

                if (results.Count == 0)
                    return AgentToolText.F("I18n_AgentResultNoGrepMatches", pattern)
                        + (skippedLarge > 0 ? AgentToolText.F("I18n_AgentResultSkippedLargeFiles", skippedLarge) : "")
                        + (skippedBinary > 0 ? AgentToolText.F("I18n_AgentResultSkippedBinaryFiles", skippedBinary) : "")
                        + (skippedSensitive > 0 ? AgentToolText.F("I18n_AgentResultSkippedSensitiveFiles", skippedSensitive) : "");

                var sb = new StringBuilder();
                sb.AppendLine(AgentToolText.F("I18n_AgentResultGrepPattern", pattern));
                sb.AppendLine(AgentToolText.F("I18n_AgentResultGrepPath", searchPath, glob));
                sb.AppendLine(AgentToolText.F("I18n_AgentResultGrepMatches", totalMatches, results.Count));
                if (skippedLarge > 0)
                    sb.AppendLine(AgentToolText.F("I18n_AgentResultGrepSkippedLarge", skippedLarge));
                if (skippedBinary > 0)
                    sb.AppendLine(AgentToolText.F("I18n_AgentResultGrepSkippedBinary", skippedBinary));
                if (skippedSensitive > 0)
                    sb.AppendLine(AgentToolText.F("I18n_AgentResultGrepSkippedSensitive", skippedSensitive));
                sb.AppendLine();
                sb.AppendLine(string.Join("\n", results));

                return sb.ToString();
            }
            catch (Exception ex)
            {
                return AgentToolText.F("I18n_AgentResultGrepFailed", ex.Message);
            }
        }

        private static string GetDefaultSearchRoot(Word.Document doc)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(doc?.Path))
                    return Path.GetFullPath(doc.Path);
            }
            catch { }

            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "WordAgentSkills");
        }

        private static bool IsAllowedSearchPath(string path, Word.Document doc)
        {
            if (string.IsNullOrWhiteSpace(path)) return false;

            var roots = new List<string>();
            try
            {
                if (!string.IsNullOrWhiteSpace(doc?.Path))
                    roots.Add(doc.Path);
            }
            catch { }

            roots.Add(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "WordAgentSkills"));

            string oneDrive = Environment.GetEnvironmentVariable("OneDrive");
            if (!string.IsNullOrWhiteSpace(oneDrive))
            {
                roots.Add(Path.Combine(oneDrive, "Documents", "WordAgentSkills"));
                roots.Add(Path.Combine(oneDrive, "文档", "WordAgentSkills"));
            }

            return roots.Any(root => IsPathUnderRootOrEqual(path, root));
        }

        private static bool IsPathUnderRootOrEqual(string path, string root)
        {
            if (string.IsNullOrWhiteSpace(path) || string.IsNullOrWhiteSpace(root)) return false;
            string fullPath = Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string fullRoot = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            return string.Equals(fullPath, fullRoot, StringComparison.OrdinalIgnoreCase)
                || fullPath.StartsWith(fullRoot + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
        }

        private static IEnumerable<string> EnumerateSafeFiles(string root, string glob)
        {
            var pending = new Stack<string>();
            pending.Push(root);
            while (pending.Count > 0)
            {
                string current = pending.Pop();
                if (HasReparsePoint(current)) continue;

                string[] files;
                try { files = Directory.GetFiles(current, glob, SearchOption.TopDirectoryOnly); }
                catch { files = Array.Empty<string>(); }
                foreach (string file in files)
                    yield return file;

                string[] directories;
                try { directories = Directory.GetDirectories(current); }
                catch { directories = Array.Empty<string>(); }
                foreach (string directory in directories)
                {
                    if (!HasReparsePoint(directory))
                        pending.Push(directory);
                }
            }
        }

        private static bool HasReparsePoint(string path)
        {
            try
            {
                return (File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0;
            }
            catch
            {
                return true;
            }
        }

        private static bool IsSensitiveFile(string path)
        {
            string name = Path.GetFileName(path)?.ToLowerInvariant() ?? "";
            string ext = Path.GetExtension(path)?.ToLowerInvariant() ?? "";
            string[] exactNames =
            {
                ".env", ".npmrc", ".yarnrc", ".netrc", "credentials", "credentials.json",
                "secrets.json", "id_rsa", "id_ed25519"
            };
            string[] sensitiveExtensions = { ".pfx", ".p12", ".key", ".pem", ".snk", ".kdbx" };
            return exactNames.Contains(name)
                || sensitiveExtensions.Contains(ext)
                || name.IndexOf("credential", StringComparison.OrdinalIgnoreCase) >= 0
                || name.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0
                || name.IndexOf("secret", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>快速判断是否为二进制文件（基于扩展名和前 512 字节）。</summary>
        private static bool IsBinaryFile(string path)
        {
            string ext = Path.GetExtension(path)?.ToLowerInvariant() ?? "";
            string[] textExts =
            {
                ".cs",
                ".py",
                ".js",
                ".ts",
                ".jsx",
                ".tsx",
                ".md",
                ".txt",
                ".json",
                ".xml",
                ".yaml",
                ".yml",
                ".toml",
                ".ini",
                ".cfg",
                ".conf",
                ".env",
                ".csv",
                ".html",
                ".css",
                ".scss",
                ".less",
                ".vue",
                ".svelte",
                ".java",
                ".cpp",
                ".c",
                ".h",
                ".hpp",
                ".rs",
                ".go",
                ".rb",
                ".php",
                ".swift",
                ".kt",
                ".scala",
                ".clj",
                ".sh",
                ".bat",
                ".ps1",
                ".sql",
                ".r",
                ".m",
                ".mm",
                ".pl",
                ".pm",
                ".lua",
                ".hs",
                ".erl",
                ".ex",
                ".exs",
                ".sln",
                ".csproj",
                ".props",
                ".targets",
                ".resx",
                ".xaml",
                ".razor",
                ".cshtml",
                ".psm1",
                ".psd1",
                ".gradle",
                ".cmake",
                ".makefile",
                ".dockerfile",
                ".gitignore",
                ".gitattributes",
                ".editorconfig",
                ".npmrc",
                ".yarnrc",
                ".prettierrc",
                ".eslintrc",
                ".terraform",
                ".tf",
                ".tfvars"
            };
            // 已知二进制扩展名直接跳过
            if (!textExts.Contains(ext) && ext != "" && ext.IndexOfAny(new[] { '.', '~' }) >= 0)
                return true;
            if (ext == "")
                return false; // 无扩展名视为文本

            // 额外检测前 512 字节是否含 null 字符（二进制特征）
            try
            {
                using (
                    var fs = new FileStream(
                        path,
                        FileMode.Open,
                        FileAccess.Read,
                        FileShare.ReadWrite
                    )
                )
                {
                    if (fs.Length > 1024 * 1024)
                        return false; // >1MB 跳过检测避免性能问题
                    byte[] buffer = new byte[Math.Min(512, (int)fs.Length)];
                    fs.Read(buffer, 0, buffer.Length);
                    for (int i = 0; i < buffer.Length; i++)
                    {
                        if (buffer[i] == 0)
                            return true; // null 字符 → 二进制
                    }
                }
            }
            catch { }
            return false;
        }
    }
}
