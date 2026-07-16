using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
using WordOllama.lib;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.NewUI
{
    /// <summary>
    /// 静默 Linter — 定时轮询审查用户最近输入的内容（同时保留保存时触发作为保底）。
    ///
    /// 防爆 Token 策略：
    /// 1. 定时轮询（默认 30 秒），而非每次保存触发
    /// 2. 只审查有变化的段落（Diff-only），不查全文
    /// 3. 使用小模型（Qwen-2.5-7B 等本地快速模型）而非昂贵的大模型
    /// 4. 间隔由用户通过 LinterDebounceSeconds 自行设置
    /// 5. 结果写入侧边栏日志，不弹通知框
    /// </summary>
    public class SilentLinter
    {
        private readonly Word.Application _app;
        private readonly Func<string, Task<string>> _smallModelCall;
        private readonly Action _showPane;
        private AgentTaskPaneUI _taskPane;

        /// <summary>定时轮询审查计时器。</summary>
        private readonly System.Windows.Forms.Timer _reviewTimer;

        /// <summary>段落索引快照：用于定位光标附近的变更段落。
        /// Key = 段落索引，Value = 段落文本。</summary>
        private Dictionary<int, string> _snapshotByIndex = new Dictionary<int, string>();

        /// <summary>保存时用更短的节流（5s），避免和定时器互相干扰。</summary>
        private DateTime _lastSaveInspection = DateTime.MinValue;

        // 并发保护
        private readonly SemaphoreSlim _inspectionLock = new SemaphoreSlim(1, 1);

        /// <summary>是否启用静默审查。</summary>
        public bool IsEnabled { get; private set; }

        /// <summary>
        /// 初始化静默 Linter。
        /// </summary>
        public SilentLinter(Word.Application app, Func<string, Task<string>> smallModelCall, AgentTaskPaneUI taskPane = null, Action showPane = null)
        {
            _app = app;
            _smallModelCall = smallModelCall;
            _taskPane = taskPane;
            _showPane = showPane;
            _reviewTimer = new System.Windows.Forms.Timer();
            _reviewTimer.Tick += OnTimerTick;
        }

        /// <summary>更新任务窗格引用（任务窗格创建后调用）。</summary>
        public void SetTaskPane(AgentTaskPaneUI taskPane) => _taskPane = taskPane;

        /// <summary>启动定时轮询审查。从 ThisAddIn 中调用。</summary>
        public void Start()
        {
            IsEnabled = true;
            try { CaptureSnapshot(); } catch { }
            ApplyInterval();
            _reviewTimer.Start();
        }

        /// <summary>停止定时轮询审查。</summary>
        public void Stop()
        {
            IsEnabled = false;
            _reviewTimer.Stop();
        }

        /// <summary>设置更改后调用，使新间隔即时生效。</summary>
        public void ApplyInterval()
        {
            int seconds = Math.Max(3, AppConfig.Instance.LinterDebounceSeconds);
            _reviewTimer.Interval = seconds * 1000;
        }

        /// <summary>定时器触发 — 执行变更段落审查。</summary>
        private async void OnTimerTick(object sender, EventArgs e)
        {
            if (!IsEnabled) return;
            var doc = _app.ActiveDocument;
            if (doc == null) return;
            await OnDocumentBeforeSaveAsync(doc);
        }

        /// <summary>
        /// 捕获当前文档的段落快照。
        /// 同时存储内容哈希集（用于快速 Diff）和索引映射（用于定位）。
        /// </summary>
        public void CaptureSnapshot()
        {
            _snapshotByIndex.Clear();
            try
            {
                var doc = _app.ActiveDocument;
                if (doc == null) return;
                int count = doc.Paragraphs.Count;
                for (int i = 1; i <= count; i++)
                {
                    Word.Range paraRange = null;
                    try
                    {
                        paraRange = doc.Paragraphs[i].Range;
                        string text = paraRange.Text?.Trim() ?? "";
                        if (!string.IsNullOrEmpty(text))
                        {
                            _snapshotByIndex[i] = text;
                        }
                    }
                    finally
                    {
                        ComHelper.ReleaseCom(ref paraRange);
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// 保存时触发 — 查找变化段落并用小模型审查。
        /// 从 ThisAddIn.Application_DocumentBeforeSave 调用（保底）。
        /// 定时器也通过此方法执行审查。
        /// </summary>
        public async Task OnDocumentBeforeSaveAsync(Word.Document doc)
        {
            if (!IsEnabled) return;

            // 保存触发用短节流（5s），定时器不经过这里所以不影响
            if (DateTime.Now - _lastSaveInspection < TimeSpan.FromSeconds(5)) return;
            _lastSaveInspection = DateTime.Now;

            // 防止并发
            if (!await _inspectionLock.WaitAsync(0)) return;

            try
            {
                var changedParagraphs = FindChangedParagraphs(doc);
                if (changedParagraphs.Count == 0) return;

                // 构建审查上下文（只包含变化段落，限制 5 段）
                var sb = new StringBuilder();
                int count = 0;
                foreach (var paragraph in changedParagraphs)
                {
                    if (count >= 5) break;
                    sb.AppendLine($"[Paragraph {paragraph.ParagraphIndex}] {paragraph.Text}");
                    count++;
                }

                string context = sb.ToString().Trim();
                if (string.IsNullOrEmpty(context)) return;

                // 通知任务窗格
                _taskPane?.StartAgentReply();
                _taskPane?.SetWorkflowStatus(WorkflowState.Reviewing, Loc.T("I18n_SilentReviewRunning"));
                _taskPane?.AppendLog("🔍 " + Loc.F("I18n_SilentReviewChanged", changedParagraphs.Count));

                // 用小模型审查（防爆 Token）
                string prompt = PromptTemplateService.BuildIssueReviewPrompt(context, concise: true);

                string result = await _smallModelCall(prompt);
                var issues = ParseReviewIssues(result);

                if (issues.Count > 0)
                {
                    // 发现问题 → 自动弹出侧边栏
                    _showPane?.Invoke();

                    // 写入任务窗格
                    _taskPane?.AddReviewIssuesFromExternal(issues);
                    _taskPane?.AppendLog("⚠️ " + Loc.F("I18n_SilentReviewFound", issues.Count));
                    _taskPane?.SetFinalMessage(Loc.F("I18n_SilentReviewCompletedFound", issues.Count));
                    _taskPane?.SetWorkflowStatus(WorkflowState.Completed, Loc.F("I18n_SilentReviewCompletedFound", issues.Count));
                }
                else if (!string.IsNullOrWhiteSpace(result) && result.Trim() != "OK" && result.Trim() != "[]")
                {
                    _showPane?.Invoke();
                    _taskPane?.AppendLog($"⚠️ 发现问题：\n{result}");
                    _taskPane?.SetFinalMessage(Loc.F("I18n_SilentReviewCompletedFound", CountIssues(result)));
                    _taskPane?.SetWorkflowStatus(WorkflowState.Completed, Loc.F("I18n_SilentReviewCompletedFound", CountIssues(result)));
                }
                else
                {
                    _taskPane?.AppendLog("✅ " + Loc.T("I18n_SilentReviewPassed"));
                    _taskPane?.SetFinalMessage(Loc.T("I18n_SilentReviewPassedShort"));
                    _taskPane?.SetWorkflowStatus(WorkflowState.Completed, Loc.T("I18n_SilentReviewPassedShort"));
                }

                // 更新快照
                CaptureSnapshot();
            }
            catch (Exception ex)
            {
                _taskPane?.AppendLog("❌ " + Loc.F("I18n_SilentReviewException", ex.Message));
                _taskPane?.SetFinalMessage(Loc.T("I18n_SilentReviewExceptionShort"));
                _taskPane?.SetWorkflowStatus(WorkflowState.Failed, Loc.T("I18n_SilentReviewExceptionShort"));
            }
            finally
            {
                _inspectionLock.Release();
            }
        }

        /// <summary>
        /// 查找与快照相比发生变化的段落。
        ///
        /// 使用索引映射对比（而非内容哈希集），避免因其他段落含相同文本而漏检。
        /// 索引偏移场景（插入/删除段落）会标记受影响的段落为已变更，不遗漏。
        /// </summary>
        private List<ChangedParagraph> FindChangedParagraphs(Word.Document doc)
        {
            var changedTexts = new List<ChangedParagraph>();
            try
            {
                int count = doc.Paragraphs.Count;
                if (count == 0) return changedTexts;

                // 首次运行无快照 → 建立基线，下次再检
                if (_snapshotByIndex.Count == 0)
                {
                    CaptureSnapshot();
                    return changedTexts;
                }

                // 获取光标所在段落的大致位置
                int cursorPara = 1;
                try
                {
                    int cursorPos = _app.Selection.Range.Start;
                    // 二分查找光标所在段落（比遍历快）
                    int lo = 1, hi = count;
                    while (lo < hi)
                    {
                        int mid = (lo + hi) / 2;
                        if (doc.Paragraphs[mid].Range.End <= cursorPos)
                            lo = mid + 1;
                        else
                            hi = mid;
                    }
                    cursorPara = lo;
                }
                catch { cursorPara = 1; }

                // 只检查光标附近 ±15 段（而非全篇 500 段）
                int startPara = Math.Max(1, cursorPara - 15);
                int endPara = Math.Min(count, cursorPara + 15);

                for (int i = startPara; i <= endPara; i++)
                {
                    string currentText = doc.Paragraphs[i].Range.Text?.Trim() ?? "";
                    if (string.IsNullOrEmpty(currentText)) continue;

                    // 按段落索引对比：同索引下文本变了才算变化
                    if (_snapshotByIndex.TryGetValue(i, out string oldText))
                    {
                        if (currentText != oldText)
                            changedTexts.Add(new ChangedParagraph { ParagraphIndex = i, Text = currentText });
                    }
                    else
                    {
                        // 新增段落（用户按回车插入新段，或索引偏移）
                        changedTexts.Add(new ChangedParagraph { ParagraphIndex = i, Text = currentText });
                    }
                }
            }
            catch { }
            return changedTexts;
        }

        private List<ReviewIssue> ParseReviewIssues(string raw)
        {
            var issues = new List<ReviewIssue>();
            if (string.IsNullOrWhiteSpace(raw)) return issues;

            string json = ExtractJsonArray(raw);
            if (string.IsNullOrWhiteSpace(json)) return issues;

            try
            {
                var array = JArray.Parse(json);
                foreach (var item in array.OfType<JObject>())
                {
                    int paragraph = item["paragraph"]?.Value<int?>()
                        ?? item["paragraphIndex"]?.Value<int?>()
                        ?? item["段落"]?.Value<int?>()
                        ?? 0;
                    string description = item["description"]?.ToString() ?? item["问题说明"]?.ToString();
                    string title = item["title"]?.ToString() ?? item["标题"]?.ToString();
                    if (paragraph <= 0 && string.IsNullOrWhiteSpace(description) && string.IsNullOrWhiteSpace(title))
                        continue;

                    issues.Add(new ReviewIssue
                    {
                        Source = Loc.T("I18n_SilentReviewSource"),
                        ParagraphIndex = paragraph,
                        Category = ReviewIssue.DisplayCategory(item["category"]?.ToString() ?? item["类型"]?.ToString()),
                        Severity = ReviewIssue.DisplaySeverity(item["severity"]?.ToString() ?? item["严重程度"]?.ToString()),
                        Title = string.IsNullOrWhiteSpace(title) ? Loc.T("I18n_UiSilentReviewIssue") : title,
                        Description = description,
                        Suggestion = item["suggestion"]?.ToString() ?? item["建议"]?.ToString(),
                        Excerpt = item["excerpt"]?.ToString() ?? item["原文"]?.ToString()
                    });
                }
            }
            catch { }

            return issues;
        }

        private static string ExtractJsonArray(string raw)
        {
            int start = raw.IndexOf('[');
            int end = raw.LastIndexOf(']');
            if (start < 0 || end <= start) return "";
            return raw.Substring(start, end - start + 1);
        }

        /// <summary>统计审查结果中的问题数量。</summary>
        private int CountIssues(string result)
        {
            if (string.IsNullOrEmpty(result)) return 0;
            return result.Split('\n').Count(line => line.Contains(":") || line.Contains("："));
        }

        private class ChangedParagraph
        {
            public int ParagraphIndex { get; set; }
            public string Text { get; set; }
        }
    }
}
