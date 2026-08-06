using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Security.Cryptography;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WordOllama.lib.Tools;
using WordOllama.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.lib
{
    public enum AgentPermissionDecision
    {
        Deny,
        AllowOnce,
        AllowSession
    }

    public sealed class AgentPermissionRequest
    {
        public string PermissionKey { get; set; }
        public string ToolName { get; set; }
        public string Operation { get; set; }
        public string EffectDescription { get; set; }
        public bool AllowSession { get; set; } = true;
    }

    public class DocumentAgent
    {
        private readonly Func<List<Dictionary<string, string>>, CancellationToken, Task<string>> _callAIChat;
        private readonly Func<List<Dictionary<string, string>>, JArray, CancellationToken, Task<string>> _callAIChatWithTools;
        private CancellationToken _ct;
        private Word.Document _doc;
        private Word.Application _app;

        // 工具调用能力探测：true = API 支持原生 tool calling，false = 降级到 JSON ReAct
        private bool _supportsToolCalling = true;
        private static readonly HashSet<string> UnsupportedToolCallingProfiles = new HashSet<string>(StringComparer.Ordinal);

        // Tool registry for built-in and explicitly enabled external tools.
        private readonly ToolRegistry _toolRegistry = new ToolRegistry();

        // Skill manager
        private readonly SkillManager _skillManager = new SkillManager();
        private readonly WorkflowStatus _workflowStatus = new WorkflowStatus();

        // MCP tool manager
        private McpToolManager _mcpManager;

        // [多轮对话] 聊天历史 — API 原生管理上下文记忆
        private readonly List<Dictionary<string, string>> _chatHistory = new List<Dictionary<string, string>>();

        // 里程碑追踪（防止重复修改同一段落）
        private readonly HashSet<string> _completedMilestones = new HashSet<string>();
        private readonly bool _externalToolsEnabled;

        // 文档快照（启动时自动生成）
        private string _documentSnapshot = null;

        // [HITL] Human-in-the-Loop 支持
        private TaskCompletionSource<string> _pendingResponse;
        private string _pendingQuestion;
        private TaskCompletionSource<AgentPermissionDecision> _pendingPermission;

        // [计划确认] Agent 执行前先展示计划
        private TaskCompletionSource<bool> _planConfirmation;
        private string _currentPlan;

        // [快照对比] 执行前文档快照
        private List<string> _preSnapshot;
        private bool _trackedChangesBatchStarted;
        private bool _trackRevisionsEnabledByAgent;
        private bool _taskUndoRecordStarted;
        private bool _taskUndoRecordFinalized;
        private string _taskUndoAfterXml;
        private string _taskDocumentIdentity;
        private Dictionary<string, int> _preRevisionFingerprints = new Dictionary<string, int>(StringComparer.Ordinal);
        private List<RevisionLocator> _taskRevisionLocators = new List<RevisionLocator>();
        private bool _taskRevisionIsolationSafe;
        private bool _revisionBaselineCaptured;
        private readonly Dictionary<string, AgentSource> _sources = new Dictionary<string, AgentSource>(StringComparer.OrdinalIgnoreCase);
        private DateTime _startedAt;
        private int _toolCallCount;
        private int _toolFailureCount;
        private string _goal = "";
        /// <summary>执行完成后可通过此属性获取变更列表。</summary>
        public List<DiffChange> DiffChanges { get; private set; }

        public event Action<string> StatusChanged;
        public event Action<WorkflowStatusSnapshot> StateChanged;
        public event Action<int, int> ProgressChanged;
        public event Action<string> LogEntry;
        public event Action<string> ResultReady;
        public event Action<AgentSource> SourceDiscovered;
        public event Action<AgentRunMetrics> MetricsChanged;

        public WorkflowStatus WorkflowStatus => _workflowStatus;

        /// <summary>[HITL] Agent 向用户提问时触发。UI 应显示问题并收集用户回答。</summary>
        public event Action<string> QuestionAsked;

        /// <summary>外部操作需要本地授权时触发。授权由 UI 完成，不进入模型对话。</summary>
        public event Action<AgentPermissionRequest> ExternalPermissionRequested;

        /// <summary>[计划确认] 计划生成后触发，UI 应展示计划并等待用户确认。</summary>
        public event Action<string> PlanReady;

        private int _maxIterations = 200;
        private bool _useEnglish = true;

        private bool UseEnglish => _useEnglish;

        /// <summary>最大任务轮数。0 表示无上限。</summary>
        public int MaxIterations { get => _maxIterations; set => _maxIterations = value; }

        /// <summary>
        /// Agent 模式开关 — 为 true 时，底层 AI 函数应抑制逐次通知，避免右下角弹框轰炸。
        /// </summary>
        public static bool IsAgentRunning { get; private set; }

        /// <summary>
        /// 待发送的图片 Base64 数据 — 由 UI 在启动 Agent 前设置，后端 API 调用时注入到首条用户消息中。
        /// </summary>
        public static string PendingImageBase64 { get; set; }

        public DocumentAgent(
            Func<List<Dictionary<string, string>>, CancellationToken, Task<string>> callAIChat,
            Func<List<Dictionary<string, string>>, JArray, CancellationToken, Task<string>> callAIChatWithTools = null,
            bool externalToolsEnabled = false)
        {
            _callAIChat = callAIChat;
            _callAIChatWithTools = callAIChatWithTools;
            _externalToolsEnabled = externalToolsEnabled;
            _toolRegistry.DiscoverAndRegister(_externalToolsEnabled);

            // Load skills and register read_skill.
            string skillsPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "WordAgentSkills");
            _skillManager.LoadSkills(skillsPath);
            _toolRegistry.Register(new Tools.ReadSkillTool(_skillManager));
            _toolRegistry.Register(new Tools.CreateSkillTool(_skillManager));

            // [HITL] 注册 ask_human 工具
            _toolRegistry.Register(new Tools.AskHumanTool());
        }

        /// <summary>
        /// 重新加载技能库（供 UI 在运行时刷新后调用）。
        /// </summary>
        public void RefreshSkills()
        {
            string skillsPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "WordAgentSkills");
            _skillManager.LoadSkills(skillsPath);
        }

        /// <summary>
        /// 配置 MCP 工具管理器（可选，在 ExecuteAsync 之前调用）。
        /// </summary>
        public void SetMcpManager(McpToolManager manager)
        {
            _mcpManager = manager;
            if (!_externalToolsEnabled || manager == null) return;
            // 将所有已发现的 MCP 工具注册为本地工具
            foreach (var toolDef in manager.GetEnabledTools())
            {
                _toolRegistry.Register(new McpToolWrapper(toolDef, manager));
            }
        }

        /// <summary>
        /// 获取已注册的工具注册表（可用于外部检查或添加自定义工具）。
        /// </summary>
        public ToolRegistry ToolRegistry => _toolRegistry;

        /// <summary>
        /// 生成 OpenAI / Ollama 兼容的 tools JSON Schema（为未来原生 Tool Calling 预备）。
        /// </summary>
        public JArray GetToolSchemas() => _toolRegistry.BuildToolSchemas();

        public async Task ExecuteAsync(string userRequirement, CancellationToken ct = default, bool resumeFromCheckpoint = false, string goal = null)
        {
            _ct = ct;
            _goal = (goal ?? "").Trim();
            _startedAt = DateTime.UtcNow;
            _toolCallCount = 0;
            _toolFailureCount = 0;
            _sources.Clear();
            AgentWorkspaceManager.BeginSession();
            string resolvedLanguage = PromptTemplateService.ResolveLanguage(SharedSettings.LanguageMode, userRequirement);
            _useEnglish = string.Equals(resolvedLanguage, "English", StringComparison.OrdinalIgnoreCase);
            AgentToolText.SetCurrentAgentLanguage(resolvedLanguage);
            IsAgentRunning = true;
            _trackedChangesBatchStarted = false;
            _trackRevisionsEnabledByAgent = false;
            _supportsToolCalling = !UnsupportedToolCallingProfiles.Contains(GetToolCallingProfileKey());
            try
            {
            UiThread.RunOfficeSync(() =>
            {
                _app = Globals.ThisAddIn.Application;
                _doc = _app.ActiveDocument;
            });
            _taskUndoRecordStarted = false;
            _taskUndoRecordFinalized = false;
            _taskUndoAfterXml = null;
            _taskDocumentIdentity = UiThread.RunOfficeSync(() => GetDocumentIdentity(_doc));
            _preRevisionFingerprints.Clear();
            _taskRevisionLocators.Clear();
            _taskRevisionIsolationSafe = true;
            _revisionBaselineCaptured = false;

            if (!resumeFromCheckpoint)
            {
                _completedMilestones.Clear();

                // [上下文注入] 自动生成文档快照 — 小文档喂全文，大文档喂骨架
                _documentSnapshot = UiThread.RunOfficeSync(() => BuildDocumentSnapshot(_doc));
                Log(AgentToolText.F("I18n_AgentLogSnapshotLoaded", _documentSnapshot.Length));

                // [快照对比] 执行前保存段落快照
                _preSnapshot = UiThread.RunOfficeSync(CaptureDocumentSnapshot);
                DiffChanges = null;
                Log(AgentToolText.F("I18n_AgentLogDiffSnapshotSaved", _preSnapshot.Count));

                // [选中区域注入] 如果用户在文档中选中了文本，优先在系统提示中强调
                string selectionHint = "";
                try
                {
                    selectionHint = UiThread.RunOfficeSync(() =>
                    {
                        string selText = _app.Selection.Text?.Trim();
                        if (string.IsNullOrEmpty(selText) || selText.Length <= 2 || selText.Length >= 5000)
                            return "";
                        int selStart = _app.Selection.Start;
                        int selParaIndex = GetParagraphIndexAtPosition(_doc, selStart);
                        string selPreview = selText.Length > 200 ? selText.Substring(0, 200) + "..." : selText;
                        string hint = PromptTemplateService.BuildSelectionHint(
                            selParaIndex, selPreview, UseEnglish);
                        Log(AgentToolText.F("I18n_AgentLogSelectionDetected", selParaIndex));
                        return hint;
                    });
                }
                catch { /* Selection 操作失败不影响主流程 */ }

                // ===== [多轮对话架构] 初始化聊天历史 =====
                _chatHistory.Clear();

                // 1. 构建系统提示词（只组装一次，包含工具列表、文档快照、规则）
                string systemPrompt = BuildSystemPrompt(userRequirement);
                // 附加选中区域提示
                if (!string.IsNullOrEmpty(selectionHint))
                    systemPrompt += selectionHint;
                _chatHistory.Add(new Dictionary<string, string> {
                    { "role", "system" },
                    { "content", systemPrompt }
                });

                // 2. 注入用户初始需求
                _chatHistory.Add(new Dictionary<string, string> {
                    { "role", "user" },
                    { "content", PromptTemplateService.BuildAgentUserRequirement(userRequirement, UseEnglish) }
                });

                ReportStatus(Properties.Resources.Agent_StatusStarting, WorkflowState.Starting);
                Log($"🚀 {Properties.Resources.Agent_LogStarting}: {userRequirement}");

                // ===== [计划确认] 先生成计划，等待用户确认 =====
                await RequestAndConfirmPlan(userRequirement);
            }
            else
            {
                // [Checkpoint 恢复] 重新生成文档快照（文档可能已变）
                UiThread.RunOfficeSync(() =>
                {
                    _documentSnapshot = BuildDocumentSnapshot(_doc);
                    _preSnapshot = CaptureDocumentSnapshot();
                });
                DiffChanges = null;

                // 重建 system prompt（第 0 条始终是 system）
                string newSystemPrompt = BuildSystemPrompt(userRequirement);
                if (_chatHistory.Count > 0 && _chatHistory[0].TryGetValue("role", out var role) && role == "system")
                    _chatHistory[0]["content"] = newSystemPrompt;
                else
                    _chatHistory.Insert(0, new Dictionary<string, string> { { "role", "system" }, { "content", newSystemPrompt } });

                Log(AgentToolText.F("I18n_AgentLogResumed", _completedMilestones.Count));
            }

            // Re-anchor immediately before execution. Planning can take time and the
            // user may have edited the document while reviewing the plan.
            UiThread.RunOfficeSync(() =>
            {
                _preSnapshot = CaptureDocumentSnapshot();
                _preRevisionFingerprints = CaptureRevisionFingerprintCounts();
                _taskDocumentIdentity = GetDocumentIdentity(_doc);
                BeginTaskUndoRecord();
            });

            string taskCompleteMsg = null;
            int currentIteration = 0;
            var recentActions = new List<(string Action, string Params)>(); // 循环检测
            int consecutiveFailures = 0;
            const int maxConsecutiveFailures = 3;

            while ((_maxIterations == 0 || currentIteration < _maxIterations) && DateTime.UtcNow - _startedAt < TimeSpan.FromMinutes(30))
            {
                _ct.ThrowIfCancellationRequested();
                currentIteration++;
                ProgressChanged?.Invoke(currentIteration, _maxIterations == 0 ? 0 : _maxIterations);
                PublishMetrics(currentIteration);
                ReportStatus(string.Format(Properties.Resources.Agent_StatusRound, currentIteration, _maxIterations == 0 ? "" : $"/{_maxIterations}"), WorkflowState.Thinking);

                // [错误升级] 连续失败超过阈值时主动询问用户
                if (consecutiveFailures >= maxConsecutiveFailures)
                {
                    Log(AgentToolText.F("I18n_AgentLogConsecutiveFailures", consecutiveFailures));
                    AskHumanQuestion(AgentToolText.F("I18n_AgentQuestionRetryAfterFailures", consecutiveFailures));
                    string userDecision = await _pendingResponse.Task;
                    _ct.ThrowIfCancellationRequested();
                    if (userDecision.Trim().ToLower() == "stop")
                    {
                        taskCompleteMsg = AgentToolText.T("I18n_AgentResultUserStoppedTask");
                        Log(AgentToolText.F("I18n_AgentLogUserStopped", taskCompleteMsg));
                        break;
                    }
                    consecutiveFailures = 0;
                    // 将用户建议作为 user 消息追加到聊天历史
                    _chatHistory.Add(new Dictionary<string, string> {
                        { "role", "user" },
                            { "content", PromptTemplateService.BuildUserAdjustmentMessage(userDecision, UseEnglish) }
                    });
                    continue;
                }

                // ===== 核心：将完整聊天历史发送给 API =====
                DebugLogger.LogBlock("Agent.Step", $"Step {currentIteration} ChatHistory", $"消息数: {_chatHistory.Count}");

                // 原生 Tool Calling 优先（若 API 支持）
                string aiResponse;
                bool isToolCall = false;
                string toolCallName = null;
                JObject toolCallArgs = null;

                if (_callAIChatWithTools != null && _supportsToolCalling)
                {
                    var tools = _toolRegistry.BuildToolSchemas();
                    aiResponse = await _callAIChatWithTools(_chatHistory, tools, _ct);

                    // 检测 [TOOL_CALL_NOT_SUPPORTED] → 降级到 JSON ReAct
                    if (aiResponse == "[TOOL_CALL_NOT_SUPPORTED]")
                    {
                        _supportsToolCalling = false;
                        UnsupportedToolCallingProfiles.Add(GetToolCallingProfileKey());
                        Log(AgentToolText.T("I18n_AgentLogToolCallingFallback"));
                        aiResponse = await _callAIChat(_chatHistory, _ct);
                    }
                    else
                    {
                        // 检测 [TOOL_CALL] 标记
                        var tc = ParseToolCall(aiResponse);
                        if (tc != null)
                        {
                            isToolCall = true;
                            toolCallName = tc.Value.Name;
                            toolCallArgs = tc.Value.Args;
                            // 将 AI 的 tool_call 响应存入历史（保留 thought）
                            _chatHistory.Add(new Dictionary<string, string> {
                                { "role", "assistant" },
                                { "content", aiResponse }
                            });
                        }
                    }
                }
                else
                {
                    aiResponse = await _callAIChat(_chatHistory, _ct);
                }

                // 若非 Tool Call 路径，走原有 JSON ReAct 解析
                AgentDecision decision = null;
                if (!isToolCall)
                {
                    DebugLogger.LogBlock("Agent.Step", $"Step {currentIteration} Response", aiResponse);

                    // 将 AI 回复存入历史
                    _chatHistory.Add(new Dictionary<string, string> {
                        { "role", "assistant" },
                        { "content", aiResponse }
                    });

                    decision = ParseDecision(aiResponse);

                    if (decision == null)
                    {
                        // 格式错误 → 用 user 角色纠正
                        if (aiResponse.StartsWith("[API ", StringComparison.OrdinalIgnoreCase))
                        {
                            string errorMsg = PromptTemplateService.BuildAgentApiErrorRepairPrompt(
                                aiResponse, UseEnglish);
                            Log($"⚠️ {aiResponse.Substring(0, Math.Min(120, aiResponse.Length))}");
                            _chatHistory.Add(new Dictionary<string, string> {
                                { "role", "user" },
                                { "content", errorMsg }
                            });
                        }
                        else
                        {
                            Log(AgentToolText.T("I18n_AgentLogInvalidModelOutput"));
                            _chatHistory.Add(new Dictionary<string, string> {
                                { "role", "user" },
                                { "content", PromptTemplateService.BuildAgentJsonRepairPrompt(aiResponse, UseEnglish) }
                            });
                        }
                        continue;
                    }
                }

                // [原生 Tool Calling] 将 tool call 转换为 AgentDecision
                if (isToolCall)
                {
                    string thought = ExtractThoughtFromToolResponse(aiResponse);
                    decision = new AgentDecision
                    {
                        Thought = thought,
                        Action = toolCallName,
                        Params = toolCallArgs ?? new JObject(),
                        Done = toolCallName == "finish"
                    };
                    Log(AgentToolText.F("I18n_AgentLogThought", thought));
                }

                // [循环检测] 窗口检测 — 比对最近 5 步内是否有重复的 Action+Params
                if (decision.Action != "finish" && decision.Action != "ask_human")
                {
                    string decisionParams = decision.Params?.ToString(Newtonsoft.Json.Formatting.None) ?? "{}";
                    bool isRedundant = recentActions.Any(a =>
                        a.Action == decision.Action &&
                        a.Params == decisionParams);

                    if (isRedundant)
                    {
                        string blockMsg = PromptTemplateService.BuildRepeatedActionWarning(decision.Action, UseEnglish);
                        Log(AgentToolText.F("I18n_AgentLogRepeatedActionBlocked", decision.Action));
                        _chatHistory.Add(new Dictionary<string, string> {
                            { "role", "user" },
                            { "content", blockMsg }
                        });
                        continue;
                    }
                }

                Log(AgentToolText.F("I18n_AgentLogThought", decision.Thought));

                if (decision.Action == "finish" || decision.Done)
                {
                    taskCompleteMsg = decision.Params?["message"]?.ToString() ?? AgentToolText.T("I18n_AgentResultTaskCompleted");
                    Log(AgentToolText.F("I18n_AgentLogTaskFinished", taskCompleteMsg));
                    break;
                }

                Log(AgentToolText.F("I18n_AgentLogAction", decision.Action));

                string toolResult = null;
                bool toolFailed = false;

                // [HITL] ask_human 特殊处理
                if (decision.Action == "ask_human")
                {
                    string question = decision.Params?["question"]?.ToString() ?? AgentToolText.T("I18n_AgentResultUnspecifiedQuestion");
                    AskHumanQuestion(question);
                    Log(AgentToolText.F("I18n_AgentLogQuestion", question));
                    string userAnswer = await _pendingResponse.Task;
                    _ct.ThrowIfCancellationRequested();
                    toolResult = AgentToolText.F("I18n_AgentResultUserAnswer", userAnswer);
                    Log(AgentToolText.F("I18n_AgentLogAnswer", userAnswer));
                }
                else
                {
                    var tool = _toolRegistry.GetTool(decision.Action);
                    if (tool == null)
                    {
                        toolResult = AgentToolText.F("I18n_AgentResultUnknownTool", decision.Action);
                        toolFailed = true;
                    }
                    else if (!IsToolAllowedByMode(tool))
                    {
                        toolResult = BuildModeBlockedMessage(tool);
                        toolFailed = false;
                    }
                    else if (!await ConfirmExternalEffectsAsync(tool, decision.Params))
                    {
                        toolResult = AgentToolText.T("I18n_AgentResultExternalPermissionDenied");
                        toolFailed = false;
                    }

                    // 危险命令二次确认（execute_command 删除/破坏类操作）
                    if (toolResult == null && decision.Action == "execute_command")
                    {
                        string cmd = decision.Params?["command"]?.ToString() ?? "";
                        string cmdArgs = decision.Params?["args"]?.ToString() ?? "";
                        if (IsDangerousCommand(cmd, cmdArgs))
                        {
                            string fullCmd = string.IsNullOrEmpty(cmdArgs) ? cmd : $"{cmd} {cmdArgs}";
                            Log(AgentToolText.F("I18n_AgentLogDangerousCommandPending", fullCmd));
                            var dangerousDecision = await RequestExternalPermissionAsync(new AgentPermissionRequest
                            {
                                PermissionKey = "execute_command|dangerous",
                                ToolName = "execute_command",
                                Operation = fullCmd,
                                EffectDescription = AgentToolText.T("I18n_AgentEffectLocalProcess"),
                                AllowSession = false
                            });
                            if (dangerousDecision == AgentPermissionDecision.Deny)
                            {
                                toolResult = AgentToolText.F("I18n_AgentResultDangerousCommandDenied", fullCmd);
                                toolFailed = false;
                                Log(AgentToolText.F("I18n_AgentLogDangerousCommandDenied", fullCmd));
                            }
                        }
                    }

                    if (toolResult == null) // 用户未拒绝，正常执行工具
                    {
                        // 执行工具 — 强制在主线程执行以保证 COM 线程安全
                        toolResult = await ExecuteToolSafeAsync(decision.Action, decision.Params, tool);
                        _toolCallCount++;
                        toolFailed =
                            toolResult.StartsWith("[错误]") ||
                            toolResult.StartsWith("[执行异常]") ||
                            toolResult.StartsWith("[MCP 错误]") ||
                            toolResult.StartsWith("[MCP 异常]") ||
                            toolResult.StartsWith("[Error]") ||
                            toolResult.StartsWith("[Execution error]") ||
                            toolResult.StartsWith("[MCP error]") ||
                            toolResult.StartsWith("[MCP exception]");
                    }
                }

                // 连续失败计数
                if (toolFailed)
                {
                    consecutiveFailures++;
                    _toolFailureCount++;
                }
                else
                {
                    consecutiveFailures = 0;
                    // 里程碑追踪
                    if (decision.Action == "replace_paragraph")
                    {
                        int paraIdx = decision.Params?["paragraph_index"]?.Value<int>() ?? 0;
                        if (paraIdx > 0)
                            _completedMilestones.Add($"replace_paragraph:{paraIdx}");
                    }
                    else if (decision.Action == "apply_style")
                    {
                        int paraIdx = decision.Params?["paragraph_index"]?.Value<int>() ?? 0;
                        if (paraIdx > 0)
                            _completedMilestones.Add($"apply_style:{paraIdx}");
                    }
                    else if (decision.Action == "add_comment")
                    {
                        int paraIdx = decision.Params?["paragraph_index"]?.Value<int>() ?? 0;
                        if (paraIdx > 0)
                            _completedMilestones.Add($"add_comment:{paraIdx}");
                    }
                    else if (decision.Action == "insert_text_at_end")
                    {
                        _completedMilestones.Add("insert_text_at_end");
                    }
                }

                // [关键] 将工具执行结果作为 user 消息追加到聊天历史
                // 这样 AI 的记忆里完整保存了 read_large_chunk 返回的真实段落，再也不会失忆！
                string shortParams = decision.Params?.ToString(Newtonsoft.Json.Formatting.None) ?? "{}";
                if (shortParams.Length > 100) shortParams = shortParams.Substring(0, 100) + "...";

                string observationContent = PromptTemplateService.BuildToolObservation(decision.Action, toolResult, UseEnglish);
                PublishSources(decision.Action, toolResult);
                _chatHistory.Add(new Dictionary<string, string> {
                    { "role", "user" },
                    { "content", observationContent }
                });

                // [上下文压缩] 防止历史撑爆上下文窗口
                CompressChatHistoryIfNeeded();

                recentActions.Add((decision.Action, shortParams));
                if (recentActions.Count > 5) recentActions.RemoveAt(0);

                // 日志只显示预览
                string displayPreview = toolResult.Length > 100 ? toolResult.Substring(0, 100) + "..." : toolResult;
                Log(AgentToolText.F("I18n_AgentLogObservation", displayPreview));

                // [Checkpoint] 每轮完成后保存断点（允许后续恢复）
                SaveCheckpoint(currentIteration, userRequirement);
            }

            if (taskCompleteMsg == null)
            {
                taskCompleteMsg = DateTime.UtcNow - _startedAt >= TimeSpan.FromMinutes(30)
                    ? (UseEnglish ? "The 30-minute task budget was reached. Work completed so far has been preserved." : "已达到 30 分钟任务预算，当前完成结果已保留。")
                    : Properties.Resources.Agent_MaxIterReached;
            }

            ReportStatus(Properties.Resources.Agent_StatusDone, WorkflowState.Completed);

            // [快照对比] 执行完毕后计算变更
            UiThread.RunOfficeSync(() =>
            {
                RestoreTrackedChangesMode();
                FinalizeTaskUndoRecord();
                CaptureTaskRevisionLocators();
                try { ComputeDiff(); } catch { /* diff 不应影响主流程 */ }
            });

            // [Checkpoint] 任务正常完成，清除断点
            ClearCheckpoint();

            ResultReady?.Invoke(taskCompleteMsg);
            PublishMetrics(currentIteration);
            UserMemoryStore.QueueSuccessfulTask(userRequirement, _callAIChat);
            }
            finally
            {
                // ExecuteAsync resumes on whichever synchronization context the model/provider
                // completes on. Never let final cleanup touch Word RCWs from that thread, and do
                // not allow a cleanup failure to hide the original Agent exception.
                try
                {
                    UiThread.RunOfficeSync(() =>
                    {
                        RestoreTrackedChangesMode();
                        FinalizeTaskUndoRecord();
                    });
                }
                catch (Exception cleanupException)
                {
                    DebugLogger.Log("Agent.Cleanup", cleanupException.ToString());
                    System.Diagnostics.Debug.WriteLine(cleanupException);
                }
                // [Checkpoint] 清理断点（取消/异常情况下也清除，避免残留）
                ClearCheckpoint();
                IsAgentRunning = false;
                AgentWorkspaceManager.EndSession();
            }
        }

        /// <summary>
        /// 构建系统提示词 — 包含工具列表、文档快照、技能目录和规则。
        /// 只在对话初始化时组装一次，后续由 API 多轮对话管理上下文。
        /// </summary>
        private string BuildSystemPrompt(string requirement)
        {
            return PromptTemplateService.BuildAgentSystemPrompt(
                _documentSnapshot,
                _toolRegistry.BuildToolDescriptions(),
                _toolRegistry.GetAllTools().Count(),
                _skillManager.GetSkillCatalog(),
                PromptTemplateService.ResolveLanguage(SharedSettings.LanguageMode, requirement),
                _goal,
                UserMemoryStore.BuildPromptContext("agent"));
        }

        private void PublishSources(string tool, string content)
        {
            foreach (AgentSource source in AgentSourceExtractor.Extract(tool, content))
            {
                if (_sources.ContainsKey(source.Url)) continue;
                _sources[source.Url] = source;
                SourceDiscovered?.Invoke(source);
            }
        }

        private void PublishMetrics(int iteration)
        {
            MetricsChanged?.Invoke(new AgentRunMetrics
            {
                Iteration = iteration,
                ToolCalls = _toolCallCount,
                ToolFailures = _toolFailureCount,
                Sources = _sources.Count,
                Elapsed = DateTime.UtcNow - _startedAt
            });
        }
        #region COM thread safety — marshal operations to the main thread

        /// <summary>
        /// 将 COM 操作强制调度到主线程执行。
        /// Word COM 对象是 STA（单线程单元），在后台线程操作会导致 0x8001010E (RPC_E_WRONG_THREAD) 崩溃。
        /// </summary>
        private async Task<string> ExecuteToolSafeAsync(string action, JObject parameters, IAgentTool tool)
        {
            ToolEffect effects = ToolExecutionPolicy.GetEffects(tool);
            if (tool is McpToolWrapper mcpTool)
                return await mcpTool.ExecuteWithCancellationAsync(parameters, _ct);
            if (tool is ICancellableAgentTool cancellableTool)
                return await cancellableTool.ExecuteAsync(parameters, _doc, _app, _ct);
            if (!ToolExecutionPolicy.RequiresWordCom(effects))
                return await Task.Run(() => ExecuteToolCore(action, parameters, tool), _ct);

            return UiThread.RunOfficeSync(() => ExecuteToolCore(action, parameters, tool));
        }

        /// <summary>实际执行工具 — 在主线程上执行 COM 操作，处理修订模式。</summary>
        private string ExecuteToolCore(string action, JObject parameters, IAgentTool registeredTool = null)
        {
            try
            {
                // 查找注册的工具
                var tool = registeredTool ?? _toolRegistry.GetTool(action);
                if (tool != null)
                {
                    bool isWriteTool = tool.IsWriteOperation;
                    bool writesWordDocument = isWriteTool &&
                        ToolExecutionPolicy.RequiresWordCom(ToolExecutionPolicy.GetEffects(tool));
                    if (writesWordDocument)
                        EnsureTrackedChangesForTask();

                    string result = tool.Execute(parameters, _doc, _app);
                    if (writesWordDocument)
                    {
                        result += AgentToolText.T("I18n_AgentResultTrackedChanges");
                    }

                    return result;
                }

                // 兼容：处理 finish 动作（不是工具注册表中的工具）
                if (action == "finish")
                {
                    return AgentToolText.T("I18n_AgentResultFinishHandled");
                }

                return AgentToolText.F("I18n_AgentResultUnknownTool", action);
            }
            catch (Exception ex)
            {
                return AgentToolText.F("I18n_AgentResultToolExecutionFailed", action, ex.Message);
            }
        }

        private void EnsureTrackedChangesForTask()
        {
            if (_trackedChangesBatchStarted || _doc == null) return;

            bool wasEnabled = _doc.TrackRevisions;
            _trackedChangesBatchStarted = true;
            _trackRevisionsEnabledByAgent = !wasEnabled;
            if (_trackRevisionsEnabledByAgent)
                _doc.TrackRevisions = true;
        }

        private void RestoreTrackedChangesMode()
        {
            if (!_trackedChangesBatchStarted) return;

            Action restore = () =>
            {
                if (_trackRevisionsEnabledByAgent && _doc != null)
                    _doc.TrackRevisions = false;
                _trackedChangesBatchStarted = false;
                _trackRevisionsEnabledByAgent = false;
            };

            try
            {
                UiThread.RunOfficeSync(restore);
            }
            catch
            {
                // Word may already be closing; shutdown should not be blocked.
            }
            finally
            {
                _trackedChangesBatchStarted = false;
                _trackRevisionsEnabledByAgent = false;
            }
        }

        private bool IsToolAllowedByMode(IAgentTool tool)
        {
            AgentExecutionMode mode = ToolExecutionPolicy.ParseMode(AppConfig.Instance?.AgentExecutionMode);
            ToolEffect effects = ToolExecutionPolicy.GetEffects(tool);
            return (effects & ToolEffect.DocumentWrite) == 0 || mode == AgentExecutionMode.TrackedChanges;
        }

        private string BuildModeBlockedMessage(IAgentTool tool)
        {
            AgentExecutionMode mode = ToolExecutionPolicy.ParseMode(AppConfig.Instance?.AgentExecutionMode);
            string modeName = mode == AgentExecutionMode.ViewOnly
                ? AgentToolText.T("I18n_AgentResultModeViewOnly")
                : AgentToolText.T("I18n_AgentResultModeProposeChanges");
            return AgentToolText.F("I18n_AgentResultModeBlocked", modeName, tool.Name);
        }

        private async Task<bool> ConfirmExternalEffectsAsync(IAgentTool tool, JObject parameters)
        {
            if (!_externalToolsEnabled && ToolExecutionPolicy.RequiresExternalToolPermission(tool))
                return false;

            ToolEffect required = ToolExecutionPolicy.GetEffects(tool)
                & (ToolEffect.ExternalNetwork | ToolEffect.LocalProcess | ToolEffect.ExternalMcp | ToolEffect.LocalFileRead | ToolEffect.LocalFileWrite);
            if (tool is McpToolWrapper trustedMcp && trustedMcp.IsTrusted)
                required &= ~ToolEffect.ExternalMcp;
            if (required == ToolEffect.None) return true;

            string permissionKey = BuildPermissionKey(tool.Name, required, parameters);
            if (ConfirmationManager.HasExternalPermissionForSession(permissionKey))
                return true;

            var decision = await RequestExternalPermissionAsync(new AgentPermissionRequest
            {
                PermissionKey = permissionKey,
                ToolName = tool.Name,
                Operation = BuildExternalOperation(tool, parameters),
                EffectDescription = ToolExecutionPolicy.DescribeEffects(required, UseEnglish),
                AllowSession = true
            });

            if (decision == AgentPermissionDecision.AllowSession)
                ConfirmationManager.GrantExternalPermissionForSession(permissionKey);
            return decision != AgentPermissionDecision.Deny;
        }

        private static string BuildPermissionKey(string toolName, ToolEffect effects, JObject parameters)
        {
            string canonical = parameters == null ? "{}" : parameters.ToString(Formatting.None);
            using (var sha = SHA256.Create())
            {
                string hash = Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(canonical)));
                return toolName + "|" + (int)effects + "|" + hash;
            }
        }

        private string BuildExternalOperation(IAgentTool tool, JObject parameters)
        {
            parameters = RedactSensitiveArguments(parameters);
            if (tool is ExecuteCommandTool)
            {
                return AgentToolText.F("I18n_AgentExternalOperationCommand",
                    parameters?["command"]?.ToString() ?? "",
                    parameters?["args"]?.ToString() ?? "");
            }

            if (tool is HttpRequestTool)
            {
                return AgentToolText.F("I18n_AgentExternalOperationHttp",
                    parameters?["method"]?.ToString()?.ToUpperInvariant() ?? "GET",
                    parameters?["url"]?.ToString() ?? "");
            }

            if (tool is RunPythonScriptTool)
            {
                return AgentToolText.F("I18n_AgentExternalOperationPython",
                    parameters?["skill_name"]?.ToString() ?? "",
                    parameters?["script_args"]?.ToString() ?? "");
            }

            if (tool is GrepTool)
            {
                return AgentToolText.F("I18n_AgentExternalOperationGrep",
                    parameters?["path"]?.ToString() ?? AgentToolText.T("I18n_AgentExternalOperationCurrentDocumentDirectory"),
                    parameters?["glob"]?.ToString() ?? "*");
            }

            if (tool is McpToolWrapper mcp)
            {
                string args = parameters?.ToString(Formatting.None) ?? "{}";
                return AgentToolText.F("I18n_AgentExternalOperationMcp", mcp.ServerName, mcp.ToolName, TrimForApproval(args));
            }

            return AgentToolText.F("I18n_AgentExternalOperationGeneric", tool?.Name ?? "", TrimForApproval(parameters?.ToString(Formatting.None) ?? "{}"));
        }

        private static JObject RedactSensitiveArguments(JObject parameters)
        {
            var clone = parameters == null ? new JObject() : (JObject)parameters.DeepClone();
            RedactSensitiveToken(clone);
            return clone;
        }

        private static void RedactSensitiveToken(JToken token)
        {
            if (token is JObject obj)
            {
                foreach (var property in obj.Properties().ToList())
                {
                    if (IsSensitiveArgumentName(property.Name))
                        property.Value = "[redacted]";
                    else
                        RedactSensitiveToken(property.Value);
                }
            }
            else if (token is JArray array)
            {
                foreach (var item in array)
                    RedactSensitiveToken(item);
            }
        }

        private static bool IsSensitiveArgumentName(string name)
        {
            string lower = (name ?? "").ToLowerInvariant();
            return lower.Contains("key") || lower.Contains("token") || lower.Contains("secret") ||
                   lower.Contains("password") || lower.Contains("authorization") || lower.Contains("cookie");
        }

        private static string TrimForApproval(string value)
        {
            if (string.IsNullOrEmpty(value) || value.Length <= 800) return value;
            return value.Substring(0, 800) + "...";
        }

        private static string GetToolCallingProfileKey()
        {
            var config = AppConfig.Instance;
            if (config == null) return "unknown";

            if (string.Equals(config.AIMode, "Ollama", StringComparison.OrdinalIgnoreCase))
                return $"ollama|{config.OllamaEndPoint}|{config.OllamaModel}";

            switch (config.OnlineProviderType)
            {
                case "Claude":
                    return $"claude|{config.ClaudeEndPoint}|{config.ClaudeModel}";
                case "Gemini":
                    return $"gemini|{config.GeminiEndPoint}|{config.GeminiModel}";
                default:
                    return $"openai|{config.ChatGPTEndPoint}|{config.ChatGPTModel}";
            }
        }

        #endregion

        #region 辅助方法

        /// <summary>
        /// 构建文档快照用于上下文注入。
        /// 文档 ≤ 8000 字符：直接返回全文；
        /// 文档 > 8000 字符：每 10 段抽取首句生成骨架导航。
        /// </summary>
        private string BuildDocumentSnapshot(Word.Document doc)
        {
            bool english = UseEnglish;
            int totalParas = doc.Paragraphs.Count;
            if (totalParas == 0) return PromptTemplateService.BuildDocumentSnapshotEmpty(english);
            string fullText = ReadDocumentText(doc);
            string[] paragraphs = fullText.Split(new[] { '\r' }, StringSplitOptions.None);
            int totalChars = fullText.Length;

            // 小文档：直接注入全文
            if (totalChars <= 8000)
            {
                var sb = new StringBuilder();
                sb.AppendLine(PromptTemplateService.BuildDocumentSnapshotFullHeader(totalParas, totalChars, english));
                sb.AppendLine("<document>");
                for (int i = 1; i <= totalParas && i <= paragraphs.Length; i++)
                {
                    string text = paragraphs[i - 1].Trim('\a');
                    if (!string.IsNullOrWhiteSpace(text))
                        sb.AppendLine(PromptTemplateService.FormatSnapshotParagraph(i, text, english));
                }
                sb.AppendLine("</document>");
                sb.AppendLine();
                sb.AppendLine(PromptTemplateService.BuildDocumentSnapshotFullTextNotice(english));
                return sb.ToString();
            }

            // 大文档：生成骨架导航（每 10 段抽首句）
            var map = new StringBuilder();
            map.AppendLine(PromptTemplateService.BuildDocumentSnapshotOutlineHeader(totalParas, totalChars, english));
            map.AppendLine("<outline>");
            for (int i = 1; i <= totalParas; i += 10)
            {
                int chunkEnd = Math.Min(i + 9, totalParas);
                string firstLine = paragraphs
                    .Skip(i - 1)
                    .Take(chunkEnd - i + 1)
                    .Select(text => (text ?? "").Trim('\a'))
                    .FirstOrDefault(text => !string.IsNullOrWhiteSpace(text)) ?? "";
                string summary = firstLine.Length > 60 ? firstLine.Substring(0, 60) + "..." : firstLine;
                if (string.IsNullOrWhiteSpace(summary)) summary = PromptTemplateService.BuildDocumentSnapshotEmptyParagraph(english);
                map.AppendLine(PromptTemplateService.FormatSnapshotOutlineEntry(i, chunkEnd, summary, english));
            }
            map.AppendLine("</outline>");
            map.AppendLine(PromptTemplateService.BuildDocumentSnapshotOutlineNotice(english));
            return map.ToString();
        }

        private static string ReadDocumentText(Word.Document doc)
        {
            Word.Range contentRange = null;
            try
            {
                contentRange = doc.Content;
                return contentRange.Text ?? "";
            }
            finally
            {
                ComHelper.ReleaseCom(ref contentRange);
            }
        }

        private static int GetParagraphIndexAtPosition(Word.Document doc, int position)
        {
            Word.Range contentRange = null;
            Word.Range prefixRange = null;
            try
            {
                contentRange = doc.Content;
                int start = contentRange.Start;
                int end = Math.Max(start, Math.Min(position, contentRange.End));
                prefixRange = doc.Range(start, end);
                return Math.Max(1, prefixRange.Paragraphs.Count);
            }
            catch
            {
                return 1;
            }
            finally
            {
                ComHelper.ReleaseCom(ref prefixRange);
                ComHelper.ReleaseCom(ref contentRange);
            }
        }

        /// <summary>
        /// 解析 [TOOL_CALL] 标记，返回 (Name, Args) 或 null。
        /// 格式: [TOOL_CALL]\nname: toolName\narguments: {"param":"value"}
        /// </summary>
        private ToolCallInfo? ParseToolCall(string response)
        {
            if (!response.StartsWith("[TOOL_CALL]")) return null;
            try
            {
                var lines = response.Split('\n');
                string name = null;
                string args = null;
                foreach (var l in lines)
                {
                    if (l.StartsWith("name:")) name = l.Substring(5).Trim();
                    if (l.StartsWith("arguments:")) args = l.Substring(10).Trim();
                }
                if (string.IsNullOrEmpty(name)) return null;
                return new ToolCallInfo(name, !string.IsNullOrEmpty(args) ? JObject.Parse(args) : new JObject());
            }
            catch { return null; }
        }

        /// <summary>
        /// 从 tool_call 响应中提取 thought（跳过 [TOOL_CALL] 标记行）。
        /// </summary>
        private string ExtractThoughtFromToolResponse(string response)
        {
            if (string.IsNullOrEmpty(response)) return "";
            // 如果 AI 在 tool_call 之前输出了思考文本（格式如下）：
            // 我想需要先搜索关键词...\n[TOOL_CALL]\nname: ...
            int markerIndex = response.IndexOf("[TOOL_CALL]");
            if (markerIndex > 0)
                return response.Substring(0, markerIndex).Trim();
            return "";
        }

        /// <summary>
        /// 解析 AI 输出为 AgentDecision。
        /// </summary>
        private AgentDecision ParseDecision(string json)
        {
            try
            {
                // 优先查找 ```json ... ``` 之间的内容（标准 Markdown 格式）
                var match = Regex.Match(json, @"```(?:json)?\s*(\{.*?\})\s*```", RegexOptions.Singleline);
                string cleanJson;

                if (match.Success)
                {
                    cleanJson = match.Groups[1].Value;
                }
                else
                {
                    // 退化为寻找最外层的 {}
                    int start = json.IndexOf('{');
                    int end = json.LastIndexOf('}');
                    if (start >= 0 && end > start)
                    {
                        cleanJson = json.Substring(start, end - start + 1);
                    }
                    else
                    {
                        return null;
                    }
                }

                var obj = JObject.Parse(cleanJson);
                return new AgentDecision
                {
                    Thought = obj["thought"]?.ToString() ?? "",
                    Action = obj["action"]?.ToString() ?? "",
                    Params = obj["params"] as JObject ?? new JObject(),
                    Done = obj["done"]?.ToObject<bool>() ?? false
                };
            }
            catch (Exception ex)
            {
                // 解析失败时返回 null，外层会将错误信息喂给 AI 让它修正
                Log(AgentToolText.F("I18n_AgentLogJsonParseFailed", ex.Message));
                return null;
            }
        }

        private void ReportStatus(string msg, WorkflowState state = WorkflowState.Thinking)
        {
            StatusChanged?.Invoke(msg);
            StateChanged?.Invoke(_workflowStatus.Set(state, msg));
        }

        private void Log(string msg) => LogEntry?.Invoke(msg);

        /// <summary>
        /// 检测是否为危险命令（删除/格式化/关机等），需要用户二次确认。
        /// </summary>
        private static bool IsDangerousCommand(string command, string args)
        {
            if (string.IsNullOrEmpty(command)) return false;

            string cmd = command.Trim().ToLowerInvariant();
            string cmdArgs = (args ?? "").Trim().ToLowerInvariant();

            // 危险命令列表（Windows/PowerShell/Git Bash）
            var dangerousPatterns = new (string Cmd, string ArgsHint)[]
            {
                // rm 递归/强制删除
                ("rm",      "-rf"),
                ("rm",      "-r"),
                ("rm",      "-f"),
                ("rm",      "/s"),
                // del 强制删除
                ("del",     "/f"),
                ("del",     "/s"),
                // rd / rmdir 递归删除
                ("rd",      "/s"),
                ("rmdir",   "/s"),
                // PowerShell Remove-Item
                ("remove-item", "-recurse"),
                ("remove-item", "-force"),
                ("rm",      "-recurse"),
                ("rm",      "-force"),
                // 格式化/磁盘/系统
                ("format",  ""),
                ("diskpart",""),
                ("shutdown","/s"),
                ("shutdown","/r"),
                ("shutdown","-s"),
                ("shutdown","-r"),
            };

            foreach (var (dangerCmd, argsHint) in dangerousPatterns)
            {
                if (cmd == dangerCmd)
                {
                    // ArgsHint 为空表示只要命令匹配就拦截（如 format、diskpart）
                    if (string.IsNullOrEmpty(argsHint)) return true;
                    if (cmdArgs.Contains(argsHint)) return true;
                }
            }

            // PowerShell 动词+名词模式：Remove-*, Clear-*, Format-*
            if (cmd.StartsWith("remove-") || cmd.StartsWith("clear-"))
                return true;
            if (cmd == "powershell" || cmd == "pwsh")
            {
                if (cmdArgs.Contains("remove-") || cmdArgs.Contains("rm ") ||
                    cmdArgs.Contains("del ") || cmdArgs.Contains("format"))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// 获取工具的 MaxResultSizeChars，0 表示使用全局动态默认值。
        /// </summary>
        private int GetToolMaxResultSize(string action)
        {
            var tool = _toolRegistry.GetTool(action);
            if (tool != null && tool.MaxResultSizeChars != 0)
                return tool.MaxResultSizeChars < 0 ? int.MaxValue : tool.MaxResultSizeChars;
            return GetEffectiveMaxResultSize();
        }

        /// <summary>
        /// 计算全局有效的工具结果截断上限。
        /// 0=自动模式：根据当前模型上下文窗口减去 2000（预留给 prompt + 生成）。
        /// 正数=用户手动指定。
        /// </summary>
        private int GetEffectiveMaxResultSize()
        {
            int userSetting = SharedSettings.AgentMaxResultSize;
            if (userSetting > 0) return userSetting;

            // 自动模式：根据模型上下文窗口计算可用余量
            int contextWindow = GetModelContextWindow();
            // 上下文窗口（token）× 3 ≈ 字符数，扣掉 prompt + 生成 headroom（~2000 tok ≈ 8000 chars）
            int dynamicLimit = contextWindow * 3 - 8000;
            // 上限 50000 字符（约 16667 token），避免 256K 上下文算出天文数字
            return Math.Max(2000, Math.Min(50000, dynamicLimit));
        }

        /// <summary>
        /// 智能上下文压缩 — 基于 token 水位线触发。
        /// 保留 system prompt + 用户需求 + 最近 2 轮完整对话，
        /// 中间轮次压缩为 ActionLog（工具名+参数+结果预览）。
        /// </summary>
        /// <summary>
        /// 获取当前模型的上下文窗口大小（token）。
        /// Ollama: num_ctx（用户配置的 OllamaMaxToken）。
        /// ChatGPT/在线: 默认 256K（API 服务商管理上下文，256K 覆盖绝大多数模型）。
        /// </summary>
        private static int GetModelContextWindow()
        {
            string aiMode = SharedSettings.AIMode ?? "";
            if (aiMode == "Ollama")
            {
                int ctx = SharedSettings.OllamaMaxToken; // num_ctx
                return ctx > 0 ? ctx : 4096;
            }
            // 在线模型：默认 256K，API 服务商管理上下文
            return 256000;
        }

        private void CompressChatHistoryIfNeeded()
        {
            // 1. 估算当前上下文总 token（中英混合按 chars/3 粗估）
            int totalChars = 0;
            foreach (var msg in _chatHistory)
            {
                if (msg.TryGetValue("content", out var content) && content != null)
                    totalChars += content.Length;
            }
            int estimatedTokens = totalChars / 3;

            // 2. 获取模型实际上下文窗口
            int contextWindow = GetModelContextWindow();

            // 3. 70% 水位线 — 超出则触发压缩
            int watermark = (int)(contextWindow * 0.7);
            if (estimatedTokens <= watermark) return;

            // 4. 保留结构: [sys] [需求] | [ActionLog] | [round N-1 AI+Obs] [round N AI+Obs]
            const int keepCompleteRounds = 2;
            int keepMessages = keepCompleteRounds * 2; // 每轮: assistant + observation
            int protectedCount = 2; // system(0) + user requirement(1)

            int removableStart = protectedCount;
            int removableEnd = _chatHistory.Count - keepMessages;
            if (removableEnd <= removableStart) return;

            // 5. 为被压缩的中间轮次生成 ActionLog
            bool isEnglish = UseEnglish;
            var actionLog = new List<string>();

            for (int i = removableStart; i < removableEnd; i++)
            {
                string content = _chatHistory[i]["content"];

                // Match the localized observation envelope emitted by BuildToolObservation.
                bool chineseObservation = content.StartsWith("[系统] 工具 ") && content.Contains(" 执行结果（观察）：");
                bool englishObservation = content.StartsWith("[System] Tool ") && content.Contains(" result (Observation):");
                if (chineseObservation || englishObservation)
                {
                    string prefix = chineseObservation ? "[系统] 工具 " : "[System] Tool ";
                    string separator = chineseObservation ? " 执行结果（观察）：" : " result (Observation):";
                    int toolEnd = content.IndexOf(separator, StringComparison.Ordinal);
                    string actionName = content.Substring(prefix.Length, toolEnd - prefix.Length).Trim();

                    // 提取结果第一行作为预览
                    string preview = "";
                    int obsStart = content.IndexOf('\n', toolEnd);
                    if (obsStart >= 0 && obsStart + 1 < content.Length)
                    {
                        string resultPart = content.Substring(obsStart + 1).TrimStart();
                        int nl = resultPart.IndexOf('\n');
                        preview = nl > 0
                            ? resultPart.Substring(0, nl).Trim()
                            : resultPart.Substring(0, Math.Min(80, resultPart.Length)).Trim();
                        if (preview.Length > 80) preview = preview.Substring(0, 80) + "...";
                    }

                    actionLog.Add($"  {actionName} → {preview}");
                }
            }

            string summary = PromptTemplateService.BuildConversationCompressionSummary(actionLog, keepCompleteRounds, isEnglish);

            // 6. 执行压缩：移除中间消息，插入 ActionLog
            int removeCount = removableEnd - removableStart;
            _chatHistory.RemoveRange(removableStart, removeCount);
            _chatHistory.Insert(removableStart, new Dictionary<string, string> {
                { "role", "user" },
                { "content", summary }
            });

            // 压缩后重新估算
            int newChars = 0;
            foreach (var msg in _chatHistory)
            {
                if (msg.TryGetValue("content", out var content) && content != null)
                    newChars += content.Length;
            }
            Log(AgentToolText.F("I18n_AgentLogContextCompressed", removeCount, estimatedTokens, newChars / 3, keepCompleteRounds));
        }

        #endregion

        #region [HITL] Human-in-the-Loop

        /// <summary>
        /// Agent 发起提问 — 创建 TCS 挂起主循环，触发事件通知 UI。
        /// </summary>
        internal void AskHumanQuestion(string question)
        {
            _pendingQuestion = question;
            _pendingResponse = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
            ReportStatus(AgentToolText.T("I18n_AgentStatusAwaitingInput"), WorkflowState.WaitingForUser);
            QuestionAsked?.Invoke(question);
        }

        private async Task<AgentPermissionDecision> RequestExternalPermissionAsync(AgentPermissionRequest request)
        {
            var tcs = new TaskCompletionSource<AgentPermissionDecision>(TaskCreationOptions.RunContinuationsAsynchronously);
            _pendingPermission = tcs;
            ReportStatus(AgentToolText.T("I18n_AgentStatusAwaitingInput"), WorkflowState.WaitingForUser);

            var handler = ExternalPermissionRequested;
            if (handler == null)
            {
                Interlocked.CompareExchange(ref _pendingPermission, null, tcs);
                return AgentPermissionDecision.Deny;
            }

            handler.Invoke(request);
            try
            {
                AgentPermissionDecision decision = await tcs.Task;
                _ct.ThrowIfCancellationRequested();
                return decision;
            }
            finally
            {
                Interlocked.CompareExchange(ref _pendingPermission, null, tcs);
            }
        }

        public void SetExternalPermissionDecision(AgentPermissionDecision decision)
        {
            Interlocked.Exchange(ref _pendingPermission, null)?.TrySetResult(decision);
        }

        /// <summary>
        /// UI 调用此方法提交用户回答，唤醒挂起的主循环。
        /// </summary>
        public void SetUserResponse(string response)
        {
            var tcs = Interlocked.Exchange(ref _pendingResponse, null);
            _pendingQuestion = null;
            tcs?.TrySetResult(response ?? AgentToolText.T("I18n_AgentResultNoUserAnswer"));
        }

        public void CancelPendingInteractions()
        {
            _pendingQuestion = null;
            Interlocked.Exchange(ref _pendingResponse, null)?.TrySetCanceled();
            Interlocked.Exchange(ref _pendingPermission, null)?.TrySetCanceled();
            Interlocked.Exchange(ref _planConfirmation, null)?.TrySetCanceled();
        }

        /// <summary>
        /// 获取当前待回答的问题（供 UI 检查是否在等待回答）。
        /// </summary>
        public string PendingQuestion => _pendingQuestion;

        /// <summary>
        /// 是否正在等待用户回答。
        /// </summary>
        public bool IsWaitingForUser => _pendingResponse != null;

        #endregion

        #region [视觉同步] Visual Focus Sync

        /// <summary>
        /// 将 Word 屏幕滚动并高亮到指定 Range，让用户看到 Agent 正在操作的位置。
        /// 在工具 Execute 方法中调用此静态方法。
        /// </summary>
        public static void FocusOnRange(Word.Range range, Word.Application app)
        {
            try
            {
                range.Select();
                app.ActiveWindow.ScrollIntoView(range);
            }
            catch { /* 滚动失败不应影响工具执行结果 */ }
        }

        #endregion

        #region [计划确认] Plan-Then-Execute

        /// <summary>
        /// 在正式执行前，先让 AI 输出执行计划供用户确认。
        /// </summary>
        private async Task RequestAndConfirmPlan(string requirement)
        {
            // 让 AI 自行判断：简单需求直接处理，复杂需求输出计划
            bool english = UseEnglish;
            string planPrompt = PromptTemplateService.BuildPlanDecisionUserPrompt(requirement, english);

            var planHistory = new List<Dictionary<string, string>>
            {
                new Dictionary<string, string> { { "role", "system" }, { "content", PromptTemplateService.BuildPlanDecisionSystemPrompt(english) } },
                new Dictionary<string, string> { { "role", "user" }, { "content", planPrompt } }
            };

            Log(AgentToolText.T("I18n_AgentLogEvaluatingPlan"));
            ReportStatus(AgentToolText.T("I18n_AgentStatusEvaluatingRequirement"), WorkflowState.Thinking);

            string planResponse;
            if (_callAIChatWithTools != null)
                planResponse = await _callAIChatWithTools(planHistory, new JArray(), _ct);
            else
                planResponse = await _callAIChat(planHistory, _ct);

            // AI 判定为简单需求 → 跳过计划
            if (planResponse?.TrimStart().StartsWith("[SIMPLE]") == true)
            {
                Log(AgentToolText.T("I18n_AgentLogPlanSkipped"));
                return;
            }

            // 提取计划文本（移除 [PLAN] 标记）
            _currentPlan = planResponse?.Replace("[PLAN]", "").Trim() ?? AgentToolText.T("I18n_AgentResultPlanGenerationFailed");
            string effectSummary = BuildPlanEffectSummary(_currentPlan);
            if (!string.IsNullOrWhiteSpace(effectSummary))
                _currentPlan = effectSummary + "\n\n" + _currentPlan;
            Log(AgentToolText.F("I18n_AgentLogPlan", _currentPlan));

            // 触发计划展示事件，等待用户确认
            var confirmation = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            _planConfirmation = confirmation;
            PlanReady?.Invoke(_currentPlan);

            // 挂起等待用户确认或拒绝
            bool confirmed;
            try
            {
                confirmed = await confirmation.Task;
            }
            finally
            {
                Interlocked.CompareExchange(ref _planConfirmation, null, confirmation);
            }
            _ct.ThrowIfCancellationRequested();

            if (confirmed)
            {
                Log(AgentToolText.T("I18n_AgentLogPlanConfirmed"));
            }
            else
            {
                Log(AgentToolText.T("I18n_AgentLogPlanNotConfirmed"));
                // 用户拒绝后，等待新需求（通过 HITL 机制）
                AskHumanQuestion(AgentToolText.T("I18n_AgentQuestionReviseRequirement"));
                string newReq = await _pendingResponse.Task;
                _ct.ThrowIfCancellationRequested();
                if (newReq.Trim().ToLower() != "go")
                {
                    // 用新需求重新生成计划
                    await RequestAndConfirmPlan(newReq);
                }
                else
                {
                    Log(AgentToolText.T("I18n_AgentLogDirectExecutionConfirmed"));
                }
            }
        }

        /// <summary>
        /// UI 调用此方法确认计划。
        /// </summary>
        public void ConfirmPlan()
        {
            _planConfirmation?.TrySetResult(true);
        }

        /// <summary>
        /// UI 调用此方法拒绝计划，可附带修改建议。
        /// </summary>
        public void RejectPlan()
        {
            _planConfirmation?.TrySetResult(false);
        }

        /// <summary>
        /// 是否正在等待计划确认。
        /// </summary>
        public bool IsWaitingForPlanConfirmation => _planConfirmation != null && !_planConfirmation.Task.IsCompleted;

        private string BuildPlanEffectSummary(string plan)
        {
            bool english = UseEnglish;
            var effects = new Dictionary<ToolEffect, List<string>>();
            foreach (var tool in _toolRegistry.GetAllTools().Where(tool => PlanMentionsTool(plan, tool.Name)))
            {
                ToolEffect effect = ToolExecutionPolicy.GetEffects(tool);
                if (effect == ToolEffect.None) continue;
                if (!effects.TryGetValue(effect, out var names))
                {
                    names = new List<string>();
                    effects[effect] = names;
                }
                names.Add(tool.Name);
            }

            if (effects.Count == 0) return "";
            var lines = new List<string>
            {
                AgentToolText.T("I18n_AgentPlanEffectsTitle")
            };
            foreach (var item in effects.OrderBy(pair => (int)pair.Key))
            {
                string label = ToolExecutionPolicy.DescribeEffects(item.Key, english);
                string names = string.Join(", ", item.Value.OrderBy(name => name, StringComparer.OrdinalIgnoreCase));
                lines.Add(AgentToolText.F("I18n_AgentPlanEffectsLine", label, names));
            }
            return string.Join("\n", lines);
        }

        private static bool PlanMentionsTool(string plan, string toolName)
        {
            if (string.IsNullOrWhiteSpace(plan) || string.IsNullOrWhiteSpace(toolName)) return false;
            return Regex.IsMatch(plan, @"(?<![A-Za-z0-9_-])" + Regex.Escape(toolName) + @"(?![A-Za-z0-9_-])", RegexOptions.IgnoreCase);
        }

        #endregion

        #region Task-scoped undo and revisions

        private const int MaxUndoSnapshotCharacters = 8_000_000;

        private void BeginTaskUndoRecord()
        {
            try
            {
                Word.UndoRecord undo = _app?.UndoRecord;
                if (undo == null || undo.IsRecordingCustomRecord || undo.CustomRecordLevel > 0) return;
                undo.StartCustomRecord("WordOllama Agent");
                _taskUndoRecordStarted = true;
            }
            catch (Exception ex)
            {
                DebugLogger.Log("Agent.Undo", "Custom undo record unavailable: " + ex.Message);
            }
        }

        private void FinalizeTaskUndoRecord()
        {
            if (!_taskUndoRecordStarted || _taskUndoRecordFinalized) return;
            try
            {
                _app?.UndoRecord?.EndCustomRecord();
                _taskUndoRecordFinalized = true;
                string xml = _doc?.WordOpenXML;
                if (!string.IsNullOrEmpty(xml) && xml.Length <= MaxUndoSnapshotCharacters)
                    _taskUndoAfterXml = xml;
            }
            catch (Exception ex)
            {
                DebugLogger.Log("Agent.Undo", "Custom undo record finalization failed: " + ex.Message);
                _taskUndoAfterXml = null;
            }
            finally
            {
                _taskUndoRecordStarted = false;
            }
        }

        private static string GetDocumentIdentity(Word.Document document)
        {
            if (document == null) return "";
            try { return document.FullName ?? document.Name ?? ""; }
            catch { try { return document.Name ?? ""; } catch { return ""; } }
        }

        private Dictionary<string, int> CaptureRevisionFingerprintCounts()
        {
            var counts = new Dictionary<string, int>(StringComparer.Ordinal);
            try
            {
                for (int i = 1; i <= _doc.Revisions.Count; i++)
                {
                    string fingerprint = RevisionFingerprint(_doc.Revisions[i]);
                    counts[fingerprint] = counts.TryGetValue(fingerprint, out int value) ? value + 1 : 1;
                }
                _revisionBaselineCaptured = true;
            }
            catch (Exception ex)
            {
                _revisionBaselineCaptured = false;
                DebugLogger.Log("Agent.Revisions", "Baseline capture failed: " + ex.Message);
            }
            return counts;
        }

        private void CaptureTaskRevisionLocators()
        {
            _taskRevisionLocators.Clear();
            var remaining = new Dictionary<string, int>(_preRevisionFingerprints, StringComparer.Ordinal);
            try
            {
                for (int i = 1; i <= _doc.Revisions.Count; i++)
                {
                    Word.Revision revision = _doc.Revisions[i];
                    string fingerprint = RevisionFingerprint(revision);
                    if (remaining.TryGetValue(fingerprint, out int count) && count > 0)
                    {
                        remaining[fingerprint] = count - 1;
                        continue;
                    }
                    _taskRevisionLocators.Add(new RevisionLocator
                    {
                        Fingerprint = fingerprint,
                        Start = revision.Range.Start,
                        End = revision.Range.End
                    });
                }
                _taskRevisionIsolationSafe = _revisionBaselineCaptured && !remaining.Values.Any(value => value > 0);
            }
            catch (Exception ex)
            {
                _taskRevisionLocators.Clear();
                _taskRevisionIsolationSafe = false;
                DebugLogger.Log("Agent.Revisions", "Task revision capture failed: " + ex.Message);
            }
        }

        private static string RevisionFingerprint(Word.Revision revision)
        {
            string text = "";
            string format = "";
            long ticks = 0;
            try { text = NormalizeSnapshotText(revision.Range?.Text); } catch { }
            try { format = revision.FormatDescription ?? ""; } catch { }
            try { ticks = revision.Date.ToUniversalTime().Ticks; } catch { }
            string value = (int)revision.Type + "\u001f" + (revision.Author ?? "") + "\u001f" +
                ticks + "\u001f" + format + "\u001f" + text;
            using (var sha = SHA256.Create())
                return Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(value)));
        }

        private List<RevisionCandidate> ResolveTaskRevisions()
        {
            var candidates = new List<RevisionCandidate>();
            for (int i = 1; i <= _doc.Revisions.Count; i++)
            {
                Word.Revision revision = _doc.Revisions[i];
                candidates.Add(new RevisionCandidate
                {
                    Revision = revision,
                    Fingerprint = RevisionFingerprint(revision),
                    Start = revision.Range.Start
                });
            }

            var resolved = new List<RevisionCandidate>();
            foreach (RevisionLocator locator in _taskRevisionLocators)
            {
                RevisionCandidate match = candidates
                    .Where(item => !item.Used && string.Equals(item.Fingerprint, locator.Fingerprint, StringComparison.Ordinal))
                    .OrderBy(item => Math.Abs(item.Start - locator.Start))
                    .FirstOrDefault();
                if (match == null) continue;
                match.Used = true;
                match.Locator = locator;
                resolved.Add(match);
            }
            return resolved;
        }

        private void EnsureTaskSnapshotCurrent()
        {
            if (!_taskUndoRecordFinalized || string.IsNullOrEmpty(_taskUndoAfterXml) || _doc == null ||
                !string.Equals(GetDocumentIdentity(_doc), _taskDocumentIdentity, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(AgentToolText.T("I18n_DiffTaskUndoUnavailable"));
            string current = _doc.WordOpenXML;
            if (!string.Equals(current, _taskUndoAfterXml, StringComparison.Ordinal))
                throw new InvalidOperationException(AgentToolText.T("I18n_DiffTaskDocumentChanged"));
        }

        private sealed class RevisionLocator
        {
            public string Fingerprint { get; set; }
            public int Start { get; set; }
            public int End { get; set; }
        }

        private sealed class RevisionCandidate
        {
            public Word.Revision Revision { get; set; }
            public RevisionLocator Locator { get; set; }
            public string Fingerprint { get; set; }
            public int Start { get; set; }
            public bool Used { get; set; }
        }

        #endregion

        #region [快照对比] Document Snapshot & Diff

        /// <summary>
        /// 捕获当前文档所有段落的文本快照。
        /// </summary>
        private List<string> CaptureDocumentSnapshot()
        {
            var snapshot = new List<string>();
            try
            {
                string[] paragraphs = ReadDocumentText(_doc)
                    .Split(new[] { '\r' }, StringSplitOptions.None);
                int paragraphCount = Math.Min(_doc.Paragraphs.Count, paragraphs.Length);
                for (int i = 1; i <= paragraphCount; i++)
                    snapshot.Add(NormalizeSnapshotText(paragraphs[i - 1]));
            }
            catch { /* 快照失败不影响主流程 */ }
            return snapshot;
        }

        private static string NormalizeSnapshotText(string text) =>
            (text ?? "").TrimEnd('\r', '\a');

        /// <summary>
        /// 对比执行前后的文档快照，生成变更列表。
        /// </summary>
        private void ComputeDiff()
        {
            try
            {
                var postSnapshot = CaptureDocumentSnapshot();
                var changes = DocumentDiffPlanner.CreateChanges(_preSnapshot, postSnapshot);
                foreach (DiffChange change in changes)
                    AttachDiffAnchor(change);

                DiffChanges = changes;

                if (changes.Count > 0)
                    Log(AgentToolText.F("I18n_AgentLogDocumentChanges", changes.Count(c => c.Type == "新增"), changes.Count(c => c.Type == "修改"), changes.Count(c => c.Type == "删除")));
            }
            catch { /* diff 不应影响主流程 */ }
        }

        private void AttachDiffAnchor(DiffChange change)
        {
            if (change == null || _doc == null) return;
            try
            {
                int postIndex = Math.Max(1, change.PostParagraphIndex);
                if (change.Type == "删除")
                {
                    int position = postIndex <= _doc.Paragraphs.Count
                        ? _doc.Paragraphs[postIndex].Range.Start
                        : Math.Max(0, _doc.Content.End - 1);
                    change.AnchorRange = _doc.Range(position, position);
                }
                else if (postIndex <= _doc.Paragraphs.Count)
                {
                    change.AnchorRange = _doc.Paragraphs[postIndex].Range.Duplicate;
                }
                change.AnchorStart = change.AnchorRange?.Start ?? 0;
            }
            catch { change.AnchorRange = null; }
        }

        private Word.Range ResolveDiffRange(DiffChange change, bool requireCurrentText)
        {
            if (change == null || _doc == null) return null;
            if (change.Type == "删除") return change.AnchorRange;
            string expected = NormalizeSnapshotText(change.NewText);
            try
            {
                if (change.AnchorRange != null &&
                    string.Equals(NormalizeSnapshotText(change.AnchorRange.Text), expected, StringComparison.Ordinal))
                    return change.AnchorRange;
            }
            catch { }

            Word.Range closest = null;
            int distance = int.MaxValue;
            for (int i = 1; i <= _doc.Paragraphs.Count; i++)
            {
                Word.Range candidate = null;
                try
                {
                    candidate = _doc.Paragraphs[i].Range;
                    if (!string.Equals(NormalizeSnapshotText(candidate.Text), expected, StringComparison.Ordinal))
                        continue;
                    int candidateDistance = Math.Abs(candidate.Start - change.AnchorStart);
                    if (candidateDistance >= distance) continue;
                    Word.Range previous = closest;
                    closest = candidate.Duplicate;
                    ComHelper.ReleaseCom(ref previous);
                    distance = candidateDistance;
                }
                finally { ComHelper.ReleaseCom(ref candidate); }
            }
            if (closest != null)
            {
                ReleaseDiffAnchor(change);
                change.AnchorRange = closest;
                change.AnchorStart = closest.Start;
                return closest;
            }
            if (requireCurrentText)
                throw new InvalidOperationException(AgentToolText.T("I18n_DiffTargetChanged"));
            return change.AnchorRange;
        }

        private static string PreserveParagraphSuffix(string original, string replacement) =>
            TextDiffService.PreserveStructuralSuffix(original ?? "", replacement ?? "");

        private static void ReleaseDiffAnchor(DiffChange change)
        {
            if (change?.AnchorRange == null) return;
            Word.Range range = change.AnchorRange;
            change.AnchorRange = null;
            ComHelper.ReleaseCom(ref range);
        }

        private void ReleaseAllDiffAnchors()
        {
            foreach (DiffChange change in DiffChanges ?? new List<DiffChange>())
                ReleaseDiffAnchor(change);
        }

        private void CompleteDiffReview()
        {
            ReleaseAllDiffAnchors();
            _taskRevisionLocators.Clear();
            _taskUndoAfterXml = null;
            if (DiffChanges != null) DiffChanges.Clear();
        }

        public void DiscardDiffReview(IEnumerable<DiffChange> changes)
        {
            foreach (DiffChange change in changes ?? Enumerable.Empty<DiffChange>())
                ReleaseDiffAnchor(change);
        }

        /// <summary>
        /// 接受或拒绝所有修订。
        /// UI 在 DiffReviewWindow 中调用。
        /// </summary>
        public void AcceptAllRevisions()
        {
            EnsureTaskSnapshotCurrent();
            if (!_taskRevisionIsolationSafe)
                throw new InvalidOperationException(AgentToolText.T("I18n_DiffTaskRevisionUnavailable"));
            List<RevisionCandidate> revisions = ResolveTaskRevisions();
            if (revisions.Count != _taskRevisionLocators.Count)
                throw new InvalidOperationException(AgentToolText.T("I18n_DiffTaskRevisionUnavailable"));
            foreach (RevisionCandidate item in revisions.OrderByDescending(item => item.Start))
                item.Revision.Accept();
            CompleteDiffReview();
        }

        /// <summary>
        /// 拒绝所有修订，还原文档到执行前状态。
        /// </summary>
        public void RejectAllRevisions()
        {
            EnsureTaskSnapshotCurrent();
            if (!_doc.Undo(1))
                throw new InvalidOperationException(AgentToolText.T("I18n_DiffTaskUndoUnavailable"));
            CompleteDiffReview();
        }

        /// <summary>接受单处变更（修改→应用新文本，新增/删除→保留现状）。</summary>
        public void AcceptChange(DiffChange change)
        {
            try
            {
                Word.Range range = ResolveDiffRange(change, requireCurrentText: change.Type != "删除");
                List<RevisionCandidate> revisions = ResolveTaskRevisionsForRange(range);
                foreach (RevisionCandidate item in revisions.OrderByDescending(item => item.Start))
                    item.Revision.Accept();
                RemoveResolvedRevisionLocators(revisions);
                ReleaseDiffAnchor(change);
            }
            catch { throw; }
        }

        /// <summary>拒绝单处变更（修改→恢复旧文本，新增→删除段落，删除→重新插入）。</summary>
        public void RejectChange(DiffChange change)
        {
            try
            {
                Word.Range range = ResolveDiffRange(change, requireCurrentText: change.Type != "删除");
                List<RevisionCandidate> revisions = ResolveTaskRevisionsForRange(range);
                if (revisions.Count > 0)
                {
                    foreach (RevisionCandidate item in revisions.OrderByDescending(item => item.Start))
                        item.Revision.Reject();
                    RemoveResolvedRevisionLocators(revisions);
                    ReleaseDiffAnchor(change);
                    return;
                }
                if (change.Type == "修改" && range != null)
                {
                    range.Text = PreserveParagraphSuffix(range.Text, change.OldText);
                }
                else if (change.Type == "新增" && range != null)
                {
                    range.Text = "";
                }
                else if (change.Type == "删除")
                {
                    string text = (change.OldText ?? "").TrimEnd('\r', '\a');
                    range?.InsertBefore(text + "\r");
                }
                ReleaseDiffAnchor(change);
            }
            catch { throw; }
        }

        private List<RevisionCandidate> ResolveTaskRevisionsForRange(Word.Range range)
        {
            if (range == null) return new List<RevisionCandidate>();
            int start = range.Start;
            int end = Math.Max(start, range.End);
            return ResolveTaskRevisions().Where(item =>
            {
                int revisionEnd;
                try { revisionEnd = item.Revision.Range.End; }
                catch { revisionEnd = item.Start; }
                return item.Start <= end + 1 && revisionEnd >= start - 1;
            }).ToList();
        }

        private void RemoveResolvedRevisionLocators(IEnumerable<RevisionCandidate> revisions)
        {
            foreach (RevisionCandidate item in revisions ?? Enumerable.Empty<RevisionCandidate>())
                if (item.Locator != null) _taskRevisionLocators.Remove(item.Locator);
            // An individual review action creates a newer Word undo entry, so the
            // original one-step task undo is intentionally no longer offered.
            _taskUndoAfterXml = null;
        }

        /// <summary>应用用户修改后的自定义文本（仅修改类型）。</summary>
        public void ModifyChange(DiffChange change, string customText)
        {
            try
            {
                if (change.Type == "修改")
                {
                    Word.Range range = ResolveDiffRange(change, requireCurrentText: true);
                    List<RevisionCandidate> revisions = ResolveTaskRevisionsForRange(range);
                    foreach (RevisionCandidate item in revisions.OrderByDescending(item => item.Start))
                        item.Revision.Reject();
                    RemoveResolvedRevisionLocators(revisions);
                    if (range != null)
                        range.Text = PreserveParagraphSuffix(range.Text, customText);
                    ReleaseDiffAnchor(change);
                }
            }
            catch { throw; }
        }

        #endregion

        #region [Checkpoint] 断点保存与恢复

        private static string GetCheckpointPath()
        {
            string dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "WordOllama");
            try { Directory.CreateDirectory(dir); } catch { }
            return Path.Combine(dir, "agent_checkpoint.dat");
        }

        private static string GetLegacyCheckpointPath()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "WordOllama",
                "agent_checkpoint.json");
        }

        /// <summary>检查是否存在可恢复的断点。</summary>
        public static bool HasCheckpoint()
        {
            return File.Exists(GetCheckpointPath()) || File.Exists(GetLegacyCheckpointPath());
        }

        /// <summary>加载断点数据到当前 Agent 实例，返回原始用户需求。</summary>
        public bool TryLoadCheckpoint(out string requirement)
        {
            requirement = null;
            try
            {
                string protectedPath = GetCheckpointPath();
                string legacyPath = GetLegacyCheckpointPath();
                string path = File.Exists(protectedPath) ? protectedPath : legacyPath;
                if (!File.Exists(path)) return false;

                string stored = File.ReadAllText(path);
                string json = SecretProtector.UnprotectOrLegacy(stored);
                if (string.IsNullOrWhiteSpace(json)) return false;
                var cp = JsonConvert.DeserializeObject<AgentCheckpoint>(json);
                if (cp?.ChatHistory == null || cp.ChatHistory.Count == 0) return false;

                _chatHistory.Clear();
                _chatHistory.AddRange(cp.ChatHistory);

                _completedMilestones.Clear();
                if (cp.CompletedMilestones != null)
                    foreach (var m in cp.CompletedMilestones)
                        _completedMilestones.Add(m);

                requirement = cp.UserRequirement ?? "";
                _goal = cp.Goal ?? "";
                if (string.Equals(path, legacyPath, StringComparison.OrdinalIgnoreCase))
                {
                    WriteProtectedCheckpoint(cp);
                    try { File.Delete(legacyPath); } catch { }
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void SaveCheckpoint(int iteration, string requirement)
        {
            try
            {
                var cp = new AgentCheckpoint
                {
                    ChatHistory = new List<Dictionary<string, string>>(_chatHistory),
                    CompletedMilestones = new List<string>(_completedMilestones),
                    CurrentIteration = iteration,
                    UserRequirement = requirement,
                    Goal = _goal,
                    SavedAt = DateTime.Now
                };
                WriteProtectedCheckpoint(cp);
            }
            catch { /* 断点保存失败不应影响主流程 */ }
        }

        private static void WriteProtectedCheckpoint(AgentCheckpoint checkpoint)
        {
            string json = JsonConvert.SerializeObject(checkpoint, Formatting.None);
            File.WriteAllText(GetCheckpointPath(), SecretProtector.Protect(json));
        }

        private void ClearCheckpoint()
        {
            try
            {
                string path = GetCheckpointPath();
                if (File.Exists(path))
                    File.Delete(path);
                string legacyPath = GetLegacyCheckpointPath();
                if (File.Exists(legacyPath))
                    File.Delete(legacyPath);
            }
            catch { }
        }

        #endregion
    }

    /// <summary>
    /// 文档变更条目 — 记录一处新增/修改/删除。
    /// </summary>
    public class DiffChange
    {
        public string Type { get; set; }
        public int ParagraphIndex { get; set; }
        public string OldText { get; set; }
        public string NewText { get; set; }
        internal int PostParagraphIndex { get; set; }
        internal int AnchorStart { get; set; }
        [JsonIgnore]
        internal Word.Range AnchorRange { get; set; }

        /// <summary>仅修改类型的 NewText 可编辑。</summary>
        public bool IsReadOnly => Type != "修改";
        /// <summary>是否为修改类型（决定是否显示"应用修改"按钮）。</summary>
        public bool IsModification => Type == "修改";

        /// <summary>变更类型的简标（用于 UI 显示）。</summary>
        public string TypeBadge => Type switch
        {
            "新增" => "➕",
            "修改" => "✏️",
            "删除" => "🗑️",
            _ => "❓"
        };
    }

    /// <summary>
    /// 断点数据 — 保存 Agent 执行状态以支持恢复。
    /// </summary>
    public class AgentCheckpoint
    {
        public List<Dictionary<string, string>> ChatHistory { get; set; }
        public List<string> CompletedMilestones { get; set; }
        public int CurrentIteration { get; set; }
        public string UserRequirement { get; set; }
        public DateTime SavedAt { get; set; }
        public string Goal { get; set; }
    }

    public sealed class AgentRunMetrics
    {
        public int Iteration { get; set; }
        public int ToolCalls { get; set; }
        public int ToolFailures { get; set; }
        public int Sources { get; set; }
        public TimeSpan Elapsed { get; set; }
    }

    public class AgentDecision
    {
        public string Thought { get; set; }
        public string Action { get; set; }
        public JObject Params { get; set; }
        public bool Done { get; set; }
    }

    /// <summary>
    /// 原生 Tool Calling 的解析结果。
    /// </summary>
    public struct ToolCallInfo
    {
        public string Name { get; }
        public JObject Args { get; }
        public ToolCallInfo(string name, JObject args) { Name = name; Args = args; }
    }
}
