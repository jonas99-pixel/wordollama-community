# WordOllama Agent 运行机制详解

> 最后更新：2026-07-16
> 本文描述当前代码结构；文件行数以源码为准，不在此处维护估算值。

---

## 目录

1. [整体架构概览](#1-整体架构概览)
2. [核心循环：ReAct 模式](#2-核心循环react-模式)
3. [提示词构建](#3-提示词构建)
4. [AI 响应解析](#4-ai-响应解析)
5. [工具系统](#5-工具系统)
6. [工具执行引擎](#6-工具执行引擎)
7. [安全机制](#7-安全机制)
8. [上下文管理](#8-上下文管理)
9. [HITL 人机交互](#9-hitl-人机交互)
10. [UI 层](#10-ui-层)
11. [技能系统](#11-技能系统)
12. [MCP 集成](#12-mcp-集成)
13. [配置与设置](#13-配置与设置)
14. [完整工具列表](#14-完整工具列表)

---

## 1. 整体架构概览

```
┌─────────────────────────────────────────────────────────┐
│                    AgentTaskPaneUI (WPF)                 │
│  用户输入 → 气泡渲染 → 状态栏 → 快捷指令 → HITL 回答     │
└───────────────────────┬─────────────────────────────────┘
                        │ 事件驱动（LogEntry / ResultReady / QuestionAsked）
                        ▼
┌─────────────────────────────────────────────────────────┐
│                   DocumentAgent (核心引擎)               │
│                                                         │
│  ExecuteAsync ─→ ReAct 循环 ─→ 提示词构建 ─→ AI 调用    │
│       │              │              │           │        │
│       │              │              │      ParseDecision │
│       │              │              │           │        │
│       │              │         BuildReActPrompt  │       │
│       │              │              │           │        │
│       │              └──── ExecuteToolSafe ─────┘        │
│       │                     (主线程 COM 安全)             │
│       │                                                  │
│       ├── 循环检测 (recentActions)                        │
│       ├── 错误升级 (consecutiveFailures)                  │
│       ├── 上下文压缩 (CompressActionHistoryIfNeeded)      │
│       └── HITL 挂起 (TaskCompletionSource)                │
└───────────────────────┬─────────────────────────────────┘
                        │
                        ▼
┌─────────────────────────────────────────────────────────┐
│                   ToolRegistry (工具注册表)               │
│                                                         │
│  自动发现 ─→ 注册 IAgentTool 实现 ─→ BuildToolDescriptions │
│       │                                                  │
│       ├── DocumentTools (8 个)                           │
│       ├── LegalTools (7 个)                              │
│       ├── ReadSkillTool                                  │
│       ├── AskHumanTool                                   │
│       └── McpToolWrapper (MCP 外部工具)                   │
└─────────────────────────────────────────────────────────┘
```

### 核心文件清单

| 文件 | 行数 | 职责 |
|------|------|------|
| `lib/DocumentAgent.cs` | 1609 | 核心引擎：ReAct 循环、提示词、上下文管理、HITL |
| `lib/IAgentTool.cs` | 176 | 工具接口 + ToolRegistry + DocumentAgentTool\<T\> 基类 |
| `lib/Tools/DocumentTools.cs` | 656 | 文档操作工具 |
| `lib/Tools/LegalTools.cs` | 905 | 法律审查工具 |
| `lib/Tools/AskHumanTool.cs` | ~40 | HITL 提问工具 |
| `lib/Tools/ReadSkillTool.cs` | ~50 | 技能文件加载工具 |
| `lib/SkillManager.cs` | ~96 | 技能管理 |
| `lib/McpToolManager.cs` | 445 | MCP 工具管理 + McpToolWrapper |
| `lib/McpBridge.cs` | 994 | MCP 协议通信（stdio、Streamable HTTP、SSE） |
| `lib/SilentLinter.cs` | ~183 | 静默审查（保存时自动审查） |
| `NewUI/AgentTaskPaneUI.xaml` | ~144 | 聊天式 UI 布局 |
| `NewUI/AgentTaskPaneUI.xaml.cs` | 2214 | UI 逻辑：气泡、输入、状态栏 |
| `Forms/SharedSettings.cs` | 428 | 全局设置 |

---

## 2. 核心循环：ReAct 模式

Agent 采用 **ReAct（Reasoning + Acting）** 模式，每一轮包含 Thought → Action → Observation 三步。

### 入口：`ExecuteAsync`

```csharp
public async Task ExecuteAsync(string userRequirement, CancellationToken ct)
```

**初始化阶段**（每次任务开始时执行一次）：
1. 设置 `_ct = ct`（取消令牌）
2. 设置 `IsAgentRunning = true`（全局标志，抑制底层 AI 弹框）
3. 获取 `_app = Globals.ThisAddIn.Application` 和 `_doc = _app.ActiveDocument`
4. 捕获主线程 `_mainDispatcher`（后续 COM 操作必须在主线程）
5. 清空 `_actionHistory`，重置 `_lastObservation`
6. 检查 `_doc == null`，若为空则触发 `ResultReady` 并返回

**循环阶段**：

```
while (未达到最大轮数)
{
    1. 检查取消令牌
    2. currentIteration++
    3. 触发 ProgressChanged 事件（更新进度条）
    4. 构建提示词 BuildReActPrompt()
    5. 循环检测（最近 2 步相同 → 注入警告）
    6. 错误升级（连续失败 ≥ 3 → ask_human 询问用户）
    7. 调用 AI 获取响应 _callAI(prompt, ct)
    8. 解析 JSON → AgentDecision { Thought, Action, Params, Done }
    9. 若 Action == "finish" 或 Done == true → 任务完成，break
    10. 执行工具 ExecuteToolSafe(action, params)
    11. 连续失败计数
    12. 结果截断（per-tool MaxResultSizeChars）
    13. 更新 _lastObservation
    14. 记录到 _actionHistory
    15. 上下文压缩（超 15 条时触发）
    16. 记录到 recentActions（循环检测用）
}
```

**收尾阶段**（finally 块）：
- 设置 `IsAgentRunning = false`
- 如果 `taskCompleteMsg == null`（达到最大轮数），设置默认消息
- 触发 `ResultReady` 事件（UI 显示最终结果）

### 关键状态变量

| 变量 | 类型 | 作用 |
|------|------|------|
| `_actionHistory` | `List<string>` | 所有步骤的摘要，注入到 prompt 中 |
| `_lastObservation` | `string` | 上一步工具返回的结果，注入到 prompt 中 |
| `recentActions` | `List<(Action, Params)>` | 最近 5 步，用于循环检测 |
| `consecutiveFailures` | `int` | 连续失败计数，≥3 时触发错误升级 |
| `currentIteration` | `int` | 当前轮数 |
| `_maxIterations` | `int` | 最大轮数（0=无限制，默认 20） |

---

## 3. 提示词构建

### `BuildReActPrompt(requirement)`

根据 `SharedSettings.LanguageMode` 选择中文或英文提示词。

#### 中文提示词结构（`BuildChinesePrompt`）

```
你是 WordOllama，一款集成在 Microsoft Word 中的独立 AI 助手...
请严格遵循 ReAct 模式：先输出 thought，再输出 action...

## 用户的最终需求
{requirement}

## 你当前的状态
- 已执行步骤数：{count}
- 执行历史：
  [1] 调用 get_doc_overview 参数: {} -> 文档共 3 段...
  [2] 调用 insert_text_at_end 参数: {"text":"..."} -> 成功...
  ...
- 上一步的执行结果（Observation）：{_lastObservation}

## 可用工具 (每次只能调用一个)
1. search_text(keyword) — 搜索文本...
2. get_doc_overview() — 获取文档总字数...
...
N. finish(message) — 认为用户需求已经完全满足时调用...

## 扩展技能库 (Skills)
{skillCatalog}

## 重要规则
- 写操作自动开启修订模式
- 不要重复调用同一工具
- 失败重试策略
- 主动提问（ask_human）
- [语言要求] 请始终用中文回复用户。  ← 仅在 LanguageMode=="Chinese" 时注入

## 输出格式要求
必须输出合法 JSON：{ thought, action, params, done }
```

#### 英文提示词结构（`BuildEnglishPrompt`）

与中文版对称，所有指令翻译为英文，并强制 `Always respond in English`。

#### 动态注入内容

每轮循环中，以下内容会动态变化：
- **已执行步骤数** — 实时更新
- **执行历史** — 只保留最近 5 条（含压缩摘要）
- **上一步 Observation** — 每轮更新为最新工具结果
- **工具列表** — 启动时固定，但如果 MCP 工具变化则会不同
- **技能目录** — 启动时固定

#### 循环检测追加内容

当 `recentActions` 最近 2 步完全相同时，追加警告：
```
[系统警告] 你已经连续两次调用 "toolName" 且参数完全相同（params），结果没有变化。
请认真阅读错误信息，修改参数后再试，不要盲目重复相同的尝试。
如果任务已完成，请直接调用 finish。
```

---

## 4. AI 响应解析

### `ParseDecision(json)`

AI 必须返回如下格式的 JSON：

```json
{
  "thought": "分析目前的观察结果，决定下一步该用什么工具",
  "action": "工具名称",
  "params": { "参数名": "参数值" },
  "done": false
}
```

**解析容错机制**（两层 fallback）：

1. **优先**：用正则匹配 `` ```json {...} ``` `` 之间的内容
2. **退化**：找最外层的 `{...}`（`IndexOf('{')` + `LastIndexOf('}')`）
3. **失败**：返回 `null`，外层会将原始输出作为错误信息喂给 AI 让它修正

解析成功后返回 `AgentDecision` 对象：
```csharp
public class AgentDecision
{
    public string Thought { get; set; }
    public string Action { get; set; }
    public JObject Params { get; set; }
    public bool Done { get; set; }
}
```

---

## 5. 工具系统

### IAgentTool 接口

```csharp
public interface IAgentTool
{
    string Name { get; }              // 工具唯一名称（对应 AI 的 action 字段）
    string Description { get; }       // 给大模型的描述文本
    JObject ParameterSchema { get; }  // JSON Schema（为未来原生 Tool Calling 预备）
    bool IsWriteOperation { get; }    // 写操作 → 自动开启修订模式
    bool IsConcurrencySafe { get; }   // 并发安全 → 为批量执行预备
    int MaxResultSizeChars { get; }   // 结果大小上限（0=默认2000, -1=无限）
    string Execute(JObject parameters, Word.Document doc, Word.Application app);
}
```

### ToolRegistry 自动发现

`DocumentAgent` 构造函数中调用 `_toolRegistry.DiscoverAndRegister(includeExternalTools)`：

```csharp
public void DiscoverAndRegister(bool includeExternalTools = false)
{
    var asm = typeof(IAgentTool).Assembly;
    foreach (var type in asm.GetTypes())
    {
        if (type.IsAbstract || type.IsInterface) continue;
        if (!typeof(IAgentTool).IsAssignableFrom(type)) continue;
        
        // 实例化；高风险外部工具仅在显式开启时注册
        IAgentTool tool = (IAgentTool)Activator.CreateInstance(type);
        _tools[tool.Name] = tool;
    }
}
```

扫描整个程序集，找到所有 `IAgentTool` 的非抽象实现类。命令、Python、HTTP 等高风险工具默认不注册，只有任务授权或设置明确开启后才会进入模型工具列表。

手动注册的工具（不在自动发现范围内）：
- `ReadSkillTool` — 构造函数中手动注册
- `AskHumanTool` — 构造函数中手动注册
- `McpToolWrapper` — `SetMcpManager()` 中批量注册

### DocumentAgentTool\<T\> 泛型基类

简化工具创建，子类只需实现 `ExecuteLogic`：

```csharp
public abstract class DocumentAgentTool<T> : IAgentTool where T : class
{
    public abstract string Name { get; }
    public abstract string Description { get; }
    public virtual bool IsWriteOperation => false;
    public virtual bool IsConcurrencySafe => !IsWriteOperation;
    public virtual int MaxResultSizeChars => 0;
    
    public string Execute(JObject parameters, Word.Document doc, Word.Application app)
    {
        return ExecuteLogic(parameters, doc, app);
    }
    
    protected abstract string ExecuteLogic(JObject parameters, Word.Document doc, Word.Application app);
}
```

### 工具描述生成

`BuildToolDescriptions()` 遍历所有已注册工具，拼接描述文本：

```
1. search_text(keyword) — 搜索文本，返回所在段落编号及上下文预览。
2. get_doc_overview() — 获取文档总字数、段落数及前2段预览。
3. get_selection() — 获取用户当前在文档中选中的文本。
...
```

`BuildToolSchemas()` 生成 OpenAI/Ollama 兼容的 JSON Schema（为未来原生 Tool Calling 预备）。

---

## 6. 工具执行引擎

### `ExecuteToolSafe(action, parameters)`

COM 线程安全入口：

```csharp
private string ExecuteToolSafe(string action, JObject parameters)
{
    if (_mainDispatcher != null && !_mainDispatcher.CheckAccess())
    {
        return _mainDispatcher.Invoke(() => ExecuteToolCore(action, parameters));
    }
    return ExecuteToolCore(action, parameters);
}
```

Word COM 对象是 STA（单线程单元），在非主线程操作会崩溃（`0x8001010E RPC_E_WRONG_THREAD`）。

### `ExecuteToolCore(action, parameters)`

实际执行逻辑：

```
1. 从 _toolRegistry.GetTool(action) 查找工具
2. 若找到：
   a. 检查 tool.IsWriteOperation
   b. 若是写操作 → 保存原 TrackRevisions 状态 → 设置 _doc.TrackRevisions = true
   c. 调用 tool.Execute(parameters, _doc, _app)
   d. 若是写操作 → 追加 "（已使用修订模式标记所有更改）"
   e. 恢复 TrackRevisions 原始状态（finally 块保证）
3. 若 action == "finish" → 返回 "[系统] 任务将由调用方处理完成。"
4. 若找不到工具 → 返回 "[错误] 不存在的工具：{action}"
5. 任何异常 → 返回 "[执行异常] 工具 {action} 报错: {ex.Message}"
```

**关键设计**：
- 工具只接收 `doc` 和 `app`，不持有任何状态
- 写操作自动开启修订模式，完成后自动恢复
- 错误不中断循环，而是返回错误文本让 AI 自行决策

### 结果大小控制

执行后检查结果长度：

```csharp
int maxResultLen = GetToolMaxResultSize(action);  // 从工具获取，0 则用默认 2000
if (maxResultLen > 0 && toolResult.Length > maxResultLen)
{
    displayResult = toolResult.Substring(0, maxResultLen) 
        + "\n...[结果过长已截断，请使用分段工具查看更多]";
}
```

这是仿照 Claude Code 的 `maxResultSizeChars` 设计，保护上下文窗口。

---

## 7. 安全机制

### 7.1 循环检测

存储最近 5 步 `(Action, Params)`，当最近 2 步完全相同时注入警告：

```csharp
if (last.Action == prev.Action && last.Params == prev.Params && last.Action != "finish")
{
    prompt += "\n\n[系统警告] 你已经连续两次调用...";
}
```

### 7.2 错误升级（渐进式）

参照 Claude Code 的权限拒绝渐进升级设计：

```csharp
if (toolFailed)
    consecutiveFailures++;
else
    consecutiveFailures = 0;

if (consecutiveFailures >= 3)
{
    AskHumanQuestion("Agent 已连续 3 次执行失败。是否继续？");
    string userDecision = await _pendingResponse.Task;
    if (userDecision == "stop") break;
    consecutiveFailures = 0;
    _lastObservation = $"用户建议：{userDecision}";
}
```

### 7.3 COM 线程安全

所有 Word COM 操作通过 `_mainDispatcher.Invoke()` 强制在主线程执行。

### 7.4 修订模式保护

写操作自动开启 `TrackRevisions`，通过 try/finally 保证即使异常也能恢复原始状态。

### 7.5 取消机制

`CancellationTokenSource` 贯穿整个流程：
- UI 层：点击取消 → `_cts.Cancel()`
- Agent 层：每轮开头 `_ct.ThrowIfCancellationRequested()`
- HITL 层：取消时 `_agent.SetUserResponse("[用户取消]")` 唤醒挂起的 TCS

### 7.6 外部工具授权

- `execute_command`、`run_python_script`、`http_request` 和 MCP 工具默认不注册，也不会连接 MCP 服务器。
- 任务开始时可仅为当前任务启用外部工具；设置页的“默认允许”是用户主动选择的持久化开关。
- 每次外部调用仍需单独确认，确认内容包含命令、URL、Skill 或 MCP 服务/工具及脱敏后的参数。
- API key、token、secret、password、authorization 和 cookie 等字段不会出现在确认提示中。

---

## 8. 上下文管理

### 8.1 Action History

`_actionHistory` 是注入到 prompt 中的执行历史，格式：

```
[1] 调用 get_doc_overview 参数: {} -> 文档共 3 段，约 150 词...
[2] 调用 read_paragraphs 参数: {"start":1,"count":5} -> 段落1: 甲方（出租方）...
[3] 调用 insert_text_at_end 参数: {"text":"..."} -> 成功在文档末尾插入...
```

prompt 中只展示最近 5 条：
```csharp
string.Join("\n", _actionHistory.Skip(Math.Max(0, _actionHistory.Count - 5)))
```

### 8.2 上下文压缩

当 `_actionHistory.Count > 15` 时触发压缩：

```csharp
private void CompressActionHistoryIfNeeded()
{
    const int maxHistory = 15;
    const int keepRecent = 8;
    
    if (_actionHistory.Count <= maxHistory) return;
    
    int removeCount = _actionHistory.Count - keepRecent;
    // 统计被压缩条目的工具使用频次
    var toolCounts = CountToolUsage(0, removeCount);
    
    // 生成摘要行
    string summary = $"[第 1-{removeCount} 步已压缩] 使用的工具: read_paragraphs×3, replace_paragraph×2";
    
    // 替换
    _actionHistory.RemoveRange(0, removeCount);
    _actionHistory.Insert(0, summary);
}
```

压缩后 prompt 中的执行历史变为：
```
[第 1-7 步已压缩] 使用的工具: read_paragraphs×3, replace_paragraph×2, add_comment×2
[8] 调用 search_text 参数: {"keyword":"租金"} -> 找到 2 处...
[9] 调用 add_comment 参数: {...} -> 已在段落 5 添加批注...
...
```

### 8.3 Observation 截断

`_lastObservation` 注入到 prompt 的"上一步执行结果"位置。超长时被截断到 `MaxResultSizeChars`（默认 2000 字符）。

### 8.4 Prompt 中的 token 分配

一个典型 prompt 的 token 分布：

| 部分 | 估计 token 数 | 说明 |
|------|-------------|------|
| 系统指令 | ~500 | 角色定义、规则、输出格式 |
| 用户需求 | ~100 | 原始需求文本 |
| 执行历史（5条） | ~500 | 最近 5 步摘要 |
| 上次 Observation | ~500 | 上一步工具结果（截断后） |
| 工具列表 | ~800 | 所有工具描述 |
| 技能目录 | ~200 | 已安装技能 |
| **总计** | **~2600** | 不含 AI 响应 |

---

## 9. HITL 人机交互

### 9.1 机制

使用 `TaskCompletionSource<string>` 挂起主循环：

```
Agent 主循环                          UI 线程
    │                                   │
    ├─ AskHumanQuestion("问题")         │
    │   ├─ 创建 TCS                     │
    │   └─ 触发 QuestionAsked 事件 ───→ │ 显示问题气泡
    │                                   │ 用户输入回答
    │   await _pendingResponse.Task     │
    │   （挂起，不消耗 CPU）             │
    │                                   ├─ _agent.SetUserResponse("回答")
    │   ← TCS 被 SetResult 唤醒 ────────┘
    │   toolResult = "用户回答：..."
    └─ 继续循环
```

### 9.2 相关方法

```csharp
// Agent 侧 — 挂起
internal void AskHumanQuestion(string question)
{
    _pendingQuestion = question;
    _pendingResponse = new TaskCompletionSource<string>();
    QuestionAsked?.Invoke(question);
}

// UI 侧 — 唤醒
public void SetUserResponse(string response)
{
    var tcs = _pendingResponse;
    _pendingResponse = null;
    tcs?.TrySetResult(response ?? "（用户未回答）");
}
```

### 9.3 触发场景

1. **AI 主动提问**：`Action == "ask_human"`，AI 判断需要用户输入
2. **错误升级**：连续失败 ≥ 3 次，系统主动询问是否继续
3. **取消处理**：取消时 `SetUserResponse("[用户取消]")` 释放挂起的 TCS

---

## 10. UI 层

### AgentTaskPaneUI 架构

```
┌──────────────────────────────┐
│ 头部栏（标题 + 静默审查开关）   │
├──────────────────────────────┤
│ 状态栏（运行时显示当前阶段）    │
├──────────────────────────────┤
│                              │
│  对话流（ScrollViewer）        │
│  ┌──────────────────────┐    │
│  │ Agent: 💭 思考...     │    │
│  ├──────────────────────┤    │
│  │ Agent: 🔧 行动: xxx   │    │
│  ├──────────────────────┤    │
│  │ Agent: 👀 观察: ...   │    │
│  ├──────────────────────┤    │
│  │ 用户: 回答问题        │ ←右 │
│  ├──────────────────────┤    │
│  │ Agent: ✅ 任务完成    │    │
│  └──────────────────────┘    │
│                              │
├──────────────────────────────┤
│ 快捷指令 + 输入框 + 发送按钮   │
└──────────────────────────────┘
```

### 气泡模型

```csharp
public class ChatItem : INotifyPropertyChanged
{
    public bool IsUser { get; set; }
    public bool IsAgent { get; set; }
    public string Message { get; set; }  // 每个气泡只有一条消息
}
```

**每条 Agent 日志创建一个新气泡**：
```csharp
private void AddAgentBubble(string message)
{
    _chatItems.Add(new ChatItem { IsAgent = true, Message = message });
}
```

### 事件驱动

Agent 通过事件通知 UI 更新，所有 UI 操作通过 `Dispatcher.Invoke` 保证线程安全：

```csharp
_agent.LogEntry += msg => Dispatcher.Invoke(() => {
    AddAgentBubble(msg);           // 每条日志 → 新气泡
    UpdateStatusBarFromLog(msg);   // 更新状态栏
    ScrollToEnd();                 // 自动滚动
});

_agent.ResultReady += summary => Dispatcher.Invoke(() => {
    AddAgentBubble(summary);       // 最终结果 → 新气泡
});

_agent.QuestionAsked += question => Dispatcher.Invoke(() => {
    AddAgentBubble($"❓ AI 提问：{question}");
    AddAgentBubble("💬 请在下方输入您的回答并发送...");
});
```

### 状态栏

根据日志内容前缀自动更新状态：

| 日志前缀 | 状态栏显示 |
|---------|-----------|
| 🚀 | Agent 启动中... |
| 💭 | Agent 思考中... |
| 🔧 | Agent 执行行动中... |
| 👀 | Agent 观察结果中... |

---

## 11. 技能系统

### 目录结构

```
%USERPROFILE%/Documents/WordAgentSkills/
├── 合同审查/
│   └── SKILL.md
├── 法律翻译/
│   └── SKILL.md
└── 格式排版/
    └── SKILL.md
```

### 加载流程

1. `SkillManager.LoadSkills(skillsPath)` 扫描目录下所有 `SKILL.md`
2. 解析 YAML frontmatter（name, description）和 markdown body（instructions）
3. 生成技能目录文本注入到 prompt 中

### 使用方式

AI 在 prompt 中看到技能目录后，如果任务匹配某个技能，会第一步调用 `read_skill(skillName)` 加载它，然后按照技能指令编排工具操作。

---

## 12. MCP 集成

### 架构

```
DocumentAgent
    └── McpToolManager
            └── McpBridge (stdio 进程通信)
                    └── MCP Server (外部进程)
```

### 流程

1. `McpBridge` 启动 MCP Server 子进程
2. 通过 stdin/stdout 发送 JSON-RPC 请求
3. `McpToolManager` 发现可用工具
4. 每个 MCP 工具包装为 `McpToolWrapper`（实现 `IAgentTool`）
5. 注册到 `ToolRegistry`，与其他工具无差别调用

### 命名约定

MCP 工具名称格式：`mcp__{serverName}__{toolName}`

---

## 13. 配置与设置

### SharedSettings 关键属性

| 属性 | 类型 | 说明 |
|------|------|------|
| `AIMode` | string | "Ollama" 或 "ChatGPT" |
| `LanguageMode` | string | ""(默认) / "Chinese" / "English" / "Auto" |
| `AgentMaxIterations` | int | 最大轮数（0=无限制） |
| `AgentExternalToolsEnabled` | bool | 是否默认注册高风险外部工具（默认 false） |
| `OllamaEndPoint` | string | Ollama API 地址 |
| `OllamaModel` | string | Ollama 模型名 |
| `ChatGPTEndPoint` | string | 在线模型 API 地址 |
| `ChatGPTKey` | string | API Key |
| `themeType` | string | "Light" / "Dark" |

### Agent 设置 UI

`NewSetting.xaml` 中有专门的 "Agent 设置" 页签：
- 最大任务轮数（支持无上限复选框）
- 文档写入模式
- 外部工具默认权限
- 设置持久化到 `%LocalAppData%\\WordOllama\\appconfig.json`

---

## 14. 完整工具列表

### 文档操作工具（DocumentTools.cs）

| 工具名 | 类型 | 说明 |
|--------|------|------|
| `search_text` | 读 | 搜索关键词，返回段落编号和上下文预览 |
| `get_doc_overview` | 读 | 文档总字数、段落数、前2段预览 |
| `get_selection` | 读 | 获取用户当前选中的文本 |
| `read_paragraphs` | 读 | 按段落范围读取内容（start + count） |
| `insert_text_at_end` | 写 | 在文档末尾插入文本 |
| `apply_style` | 写 | 给段落应用 Word 样式（标题1、正文等） |
| `replace_paragraph` | 写 | 替换指定段落的内容 |
| `add_comment` | 写 | 在指定位置添加批注 |

### 法律工具（LegalTools.cs）

| 工具名 | 类型 | 说明 |
|--------|------|------|
| `get_document_outline` | 读 | 生成文档大纲（标题层级） |
| `read_clause` | 读 | 读取指定标题下的条款全文 |
| `extract_definitions` | 读 | 提取文档中的定义条款 |
| `check_cross_references` | 读 | 检查交叉引用完整性 |
| `replace_exact_text` | 写 | 精确替换指定文本 |
| `insert_clause_after` | 写 | 在指定标题后插入条款 |
| `highlight_risk` | 写 | 高亮标记风险文本并添加批注 |

### 系统工具

| 工具名 | 类型 | 说明 |
|--------|------|------|
| `read_skill` | 读 | 加载技能文件内容 |
| `ask_human` | 读 | 向用户提问（HITL） |
| `finish` | 系统 | 任务完成（不在注册表中，特殊处理） |

### MCP 工具

通过 MCP 协议动态注册，命名格式 `mcp__{server}__{tool}`。

---

## 附录 A：完整执行流程示例

**用户输入**："帮我审查一下合同里有没有风险条款"

```
轮次 1:
  Prompt → AI 返回: { thought: "先了解文档结构", action: "get_doc_overview", params: {} }
  执行 → "文档共 12 段，约 800 词。前两段预览：..."
  观察 → 更新 _lastObservation

轮次 2:
  Prompt → AI 返回: { thought: "逐段审查", action: "read_paragraphs", params: {start:1, count:5} }
  执行 → "段落1: 甲方（出租方）...段落2: 租金为每月..."
  观察 → 更新

轮次 3:
  Prompt → AI 返回: { thought: "发现租金条款无递增约定，标记风险", action: "highlight_risk", params: {paragraph:2, reason:"..."} }
  执行 → "已高亮段落 2 并添加批注：租金条款无递增约定..."
  （自动开启修订模式，完成后恢复）

...继续审查其他段落...

轮次 N:
  Prompt → AI 返回: { thought: "审查完毕", action: "finish", params: {message: "审查完成，共发现 3 处风险..."} }
  任务结束
```

## 附录 B：与 Claude Code 的设计对照

| 设计点 | Claude Code | WordOllama Agent |
|--------|-------------|------------------|
| 核心模式 | ReAct（隐式） | ReAct（显式 JSON） |
| 工具发现 | 运行时动态 | 启动时程序集扫描 |
| 工具并发 | isConcurrencySafe 分批 | 声明了属性，尚未实现批量 |
| 结果大小 | maxResultSizeChars | 同名属性，per-tool + 全局默认 |
| 上下文压缩 | 5 层机制（microCompact 等） | 单层：action history 摘要 |
| 循环检测 | 无内置 | recentActions 双步对比 |
| 错误升级 | 渐进式（3次连续/20次累计） | 连续 3 次 → ask_human |
| 权限系统 | 4 级模式 + AI 分类器 | 修订模式自动开关 |
| HITL | AskUserQuestion | ask_human + TaskCompletionSource |
| 记忆 | CLAUDE.md 多层注入 | 无（prompt 中无持久记忆） |
| Hooks | Pre/Post ToolUse 等 24 种 | 无 |
