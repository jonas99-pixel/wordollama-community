using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.Drawing;
using Newtonsoft.Json.Linq;
using DiffPlex;
using Word = Microsoft.Office.Interop.Word;
using System.Threading.Channels;
using WordOllama.lib;
using WordOllama.NewUI;
using System.Runtime.InteropServices;
using WordOllama.Forms;
using System.Threading;
using WordOllama.lib.Providers;
using System.Text.RegularExpressions;
using System.Windows.Controls;
using System.Globalization;
using Microsoft.Office.Core;



namespace WordOllama
{
    public partial class ThisAddIn
    {
        public string txtLang = null;
        public string themeType = SharedSettings.themeType;

        // [多文档支持] thinkingPane 字段由 EnsureThinkingPane() 按当前窗口动态更新
        // OllamaFunctions / OpenAIFunctions 通过此字段写入思考内容
        public ThinkingPane thinkingPane;

        // 静默审查（定时轮询 + 保存时保底）
        private SilentLinter _silentLinter;

        /// <summary>供设置界面控制静默审查启停。</summary>
        internal void StartSilentLinter()
        {
            try { _silentLinter?.Start(); } catch { }
        }
        internal void StopSilentLinter()
        {
            try { _silentLinter?.Stop(); } catch { }
        }
        internal void ApplyLinterInterval()
        {
            try { _silentLinter?.ApplyInterval(); } catch { }
        }


        private void OnStartupIdle(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Idle -= OnStartupIdle;
            try { Globals.Ribbons.Ribbon1.ChangeRibbonLabels(); } catch { }
            if (!AppConfig.Instance.SetupAssistantCompleted)
            {
                try { new SetupAssistantWindow().ShowDialog(); }
                catch (Exception ex) { DebugLogger.Log("Setup", ex.Message); }
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try { SaveAgentPaneWidth(); } catch { }
            try { Globals.Ribbons.Ribbon1.DisposeNotifications(); } catch { }

            // 停止静默审查
            try { _silentLinter?.Stop(); } catch { }

            // 关闭所有由 HybridWindowManager 托管的 WPF 窗口，防止进程残留
            try { HybridWindowManager.CloseAllWindows(); } catch { }

            try
            {
                // 保存统一配置
                AppConfig.Save();
            }
            catch (Exception ex)
            {
                // 记录异常信息
                Console.WriteLine($"Error during shutdown: {ex.Message}");
            }
        }


        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // UI resources must be resolved before any runtime windows or controls are created.
            // Chinese Word uses zh-CN; every other Word UI language falls back to English.
            UiCulture.Initialize(Application);

            // 全局未捕获异常处理 — 防止 async void 等未处理异常直接崩溃 Word
            AppDomain.CurrentDomain.UnhandledException += (_, args) =>
            {
                try
                {
                    var ex = args.ExceptionObject as Exception;
                    DebugLogger.Log("CRASH", $"未捕获异常: {ex?.Message}\n{ex?.StackTrace}");
                }
                catch { }
            };

            // 初始化统一配置（JSON 单一配置源）
            AppConfig.Load();

            // 初始化静默审查（定时轮询审查 + 保存时保底）
            try
            {
                _silentLinter = new SilentLinter(Application, CallLinterModelAsync,
                    showPane: () =>
                    {
                        try
                        {
                            var ctp = EnsureAgentTaskPane();
                            if (ctp != null)
                            {
                                AttachAgentPaneToLinter(ctp);
                                ctp.Visible = true;
                            }
                        }
                        catch { }
                    });

                // 保存时触发（保底）
                this.Application.DocumentBeforeSave += OnDocumentBeforeSave;

                if (AppConfig.Instance.LinterEnabled)
                    _silentLinter.Start();
            }
            catch (Exception ex)
            {
                DebugLogger.Log("SilentLinter", $"初始化失败: {ex.Message}");
            }

            // 应用就绪后刷新功能区标签，确保显示上次保存的 AI 模式和模型名
            System.Windows.Forms.Application.Idle += OnStartupIdle;

            // AppConfig.Load above is the only runtime configuration load.
        }

        /// <summary>
        /// 确保当前活动窗口有 Agent 侧边栏，并返回它。支持多文档。
        /// 遍历所有已创建的 CustomTaskPane，查找绑定到当前 ActiveWindow 的实例。
        /// </summary>
        public Microsoft.Office.Tools.CustomTaskPane EnsureAgentTaskPane()
        {
            try
            {
                Word.Window activeWindow = this.Application.ActiveWindow;

                // 1. 遍历现有侧边栏，查找属于当前窗口的
                foreach (Microsoft.Office.Tools.CustomTaskPane ctp in this.CustomTaskPanes)
                {
                    if (ctp.Window == activeWindow && ctp.Title == NewUI.UiText.AgentTitle)
                    {
                        ctp.Width = AppConfig.Instance.AgentPaneWidthUserSet
                            ? ClampAgentPaneWidth(ctp.Width)
                            : DefaultAgentPaneWidth;
                        AttachAgentPaneToLinter(ctp);
                        return ctp;
                    }
                }

                // 2. 当前窗口没有 Agent 侧边栏，新建一个
                var newPane = new AgentTaskPane();
                var newCtp = this.CustomTaskPanes.Add(newPane, NewUI.UiText.AgentTitle, activeWindow);
                int configuredWidth = AppConfig.Instance.AgentPaneWidthUserSet
                    ? AppConfig.Instance.AgentPaneWidth
                    : DefaultAgentPaneWidth;
                newCtp.Width = ClampAgentPaneWidth(configuredWidth);
                newCtp.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                AttachAgentPaneToLinter(newCtp);
                return newCtp;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"EnsureAgentTaskPane error: {ex.Message}");
                return null;
            }
        }

        public void ShowAgentNotification(string message)
        {
            if (SharedSettings.SuppressAgentNotifications || string.IsNullOrWhiteSpace(message))
                return;

            foreach (Microsoft.Office.Tools.CustomTaskPane ctp in this.CustomTaskPanes)
            {
                if (ctp.Control is NewUI.AgentTaskPane agentPane)
                    agentPane.Panel.ShowExternalNotification(message);
                else if (ctp.Control is NewUI.ReviewTaskPane reviewPane)
                    reviewPane.Panel.ShowExternalNotification(message);
            }
        }

        private const int DefaultAgentPaneWidth = 900;
        private const int MinAgentPaneWidth = 420;
        private const int MaxAgentPaneWidth = 900;

        private const int DefaultReviewPaneWidth = 900;
        private const int MinReviewPaneWidth = 680;
        private const int MaxReviewPaneWidth = 1100;
        private const int LegacyReviewPaneWidth = 640;

        private static int ClampAgentPaneWidth(int width)
        {
            if (width <= 0) width = DefaultAgentPaneWidth;
            return Math.Max(MinAgentPaneWidth, Math.Min(MaxAgentPaneWidth, width));
        }

        private static int ClampReviewPaneWidth(int width)
        {
            if (width <= 0) width = DefaultReviewPaneWidth;
            return Math.Max(MinReviewPaneWidth, Math.Min(MaxReviewPaneWidth, width));
        }

        /// <summary>
        /// 确保当前活动窗口有独立的审阅工作台。
        /// 审阅与 Agent 使用不同的 CustomTaskPane，避免两个工作流互相遮挡。
        /// </summary>
        public Microsoft.Office.Tools.CustomTaskPane EnsureReviewTaskPane()
        {
            try
            {
                Word.Window activeWindow = this.Application.ActiveWindow;
                string title = NewUI.UiText.ReviewPaneTitle;

                foreach (Microsoft.Office.Tools.CustomTaskPane ctp in this.CustomTaskPanes)
                {
                    if (ctp.Window == activeWindow && ctp.Title == title)
                    {
                        ctp.Width = ClampReviewPaneWidth(ctp.Width);
                        return ctp;
                    }
                }

                var newPane = new ReviewTaskPane();
                var newCtp = this.CustomTaskPanes.Add(newPane, title, activeWindow);
                // Migrate the former 640px default so existing installations get a usable review workspace.
                int configuredWidth = AppConfig.Instance.ReviewPaneWidthUserSet &&
                                      AppConfig.Instance.ReviewPaneWidth > LegacyReviewPaneWidth
                    ? AppConfig.Instance.ReviewPaneWidth
                    : DefaultReviewPaneWidth;
                newCtp.Width = ClampReviewPaneWidth(configuredWidth);
                newCtp.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                return newCtp;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"EnsureReviewTaskPane error: {ex.Message}");
                return null;
            }
        }

        public void OpenReviewWorkspace(bool loadSelection = true, string instruction = null, bool replaceExisting = false)
        {
            var ctp = EnsureReviewTaskPane();
            if (ctp == null) return;

            ctp.Visible = true;
            if (ctp.Control is ReviewTaskPane pane)
                pane.Panel.OpenReviewWorkspace(loadSelection, instruction, replaceExisting);
        }

        private void SaveAgentPaneWidth()
        {
            foreach (Microsoft.Office.Tools.CustomTaskPane ctp in this.CustomTaskPanes)
            {
                if (ctp.Title == NewUI.UiText.AgentTitle)
                {
                    int width = ClampAgentPaneWidth(ctp.Width);
                    // Preserve an explicitly resized pane, while migrating the historical default to 900px.
                    if (AppConfig.Instance.AgentPaneWidthUserSet ||
                        AppConfig.Instance.AgentPaneWidth > MinAgentPaneWidth ||
                        width > MinAgentPaneWidth)
                    {
                        AppConfig.Instance.AgentPaneWidth = width;
                        AppConfig.Instance.AgentPaneWidthUserSet = true;
                    }
                }
                else if (ctp.Title == NewUI.UiText.ReviewPaneTitle)
                {
                    AppConfig.Instance.ReviewPaneWidth = ClampReviewPaneWidth(ctp.Width);
                    AppConfig.Instance.ReviewPaneWidthUserSet = true;
                }
            }

            AppConfig.Save();
        }

        private void AttachAgentPaneToLinter(Microsoft.Office.Tools.CustomTaskPane ctp)
        {
            try
            {
                if (_silentLinter != null && ctp?.Control is AgentTaskPane atp)
                    _silentLinter.SetTaskPane(atp.Panel);
            }
            catch { }
        }

        /// <summary>
        /// 确保当前活动窗口有 ThinkingPane（思维链浏览），并更新 thinkingPane 字段。
        /// OllamaFunctions / OpenAIFunctions 通过 thinkingPane 字段写入思考内容。
        /// </summary>
        public void EnsureThinkingPane()
        {
            try
            {
                Word.Window activeWindow = this.Application.ActiveWindow;
                string thinkTitle = WordOllama.Properties.Resources.思维链浏览;

                // 1. 遍历现有侧边栏，查找属于当前窗口的
                foreach (Microsoft.Office.Tools.CustomTaskPane ctp in this.CustomTaskPanes)
                {
                    if (ctp.Window == activeWindow && ctp.Title == thinkTitle)
                    {
                        thinkingPane = ctp.Control as ThinkingPane;
                        return;
                    }
                }

                // 2. 当前窗口没有 ThinkingPane，新建一个
                var newThinkingPane = new ThinkingPane();
                this.CustomTaskPanes.Add(newThinkingPane, thinkTitle, activeWindow);
                thinkingPane = newThinkingPane;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"EnsureThinkingPane error: {ex.Message}");
            }
        }

        /// <summary>
        /// 以注释方式插入文本
        /// </summary>
        /// <param name="selectionRange">文本范围</param>
        /// <param name="text">插入的文本内容</param>
        /// 
        public void InsertTextBelowSelection(Word.Range selectionRange, string text, bool manageTrackRevisions = true)
        {
            Word.Document document = selectionRange.Document;
            bool enabledTrackRevisions = manageTrackRevisions && !document.TrackRevisions;

            try
            {
                if (enabledTrackRevisions)
                    document.TrackRevisions = true;
                selectionRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                selectionRange.InsertParagraphAfter();
                selectionRange.InsertAfter(text + "\n");
            }
            finally
            {
                if (enabledTrackRevisions)
                    document.TrackRevisions = false;
            }
        }

        public void InsertTextBelowSelectionWithDiff(Word.Range selectionRange, string newText, bool manageTrackRevisions = true)
        {
            Word.Document document = selectionRange.Document;
            bool enabledTrackRevisions = manageTrackRevisions && !document.TrackRevisions;
            try
            {
                if (enabledTrackRevisions)
                    document.TrackRevisions = true;

                // 获取原文
                string originalText = selectionRange.Text.Trim();

                // 将范围折叠到结尾
                selectionRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                // 在选择范围之后插入段落
                selectionRange.InsertParagraphAfter();
                selectionRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                // 插入新文本
                selectionRange.Text = newText;

                // 对比原文和新文本，获取标记的差异
                var differences = GetDifferences(originalText, newText);

                // 在新插入的文本上添加批注，显示原文和新文本的差异
                if (differences.Count > 0)
                {
                    // 创建评论的锚点范围
                    Word.Range commentRange = selectionRange.Duplicate;
                    commentRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    commentRange.MoveStart(Word.WdUnits.wdCharacter, -newText.Length);

                    // 添加评论，初始文本为“原文与插入文本的区别:”
                    Word.Comment comment = commentRange.Comments.Add(commentRange, WordOllama.Properties.Resources.原文与插入文本的区别);

                    // 获取评论的文本范围
                    Word.Range commentTextRange = comment.Range;

                    // 在初始文本后插入一个段落标记，确保换行
                    commentTextRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    commentTextRange.InsertParagraphAfter();
                    commentTextRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    // 遍历所有差异，逐一插入并格式化
                    foreach (var diff in differences)
                    {
                        // 将范围折叠到评论文本的末尾
                        commentTextRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                        // 插入差异文本
                        commentTextRange.Text = diff.Text;

                        // 根据差异类型应用格式
                        switch (diff.Type)
                        {
                            case DiffType.Deleted:
                                commentTextRange.Font.Color = Word.WdColor.wdColorRed;
                                commentTextRange.Font.StrikeThrough = -1; // 删除线
                                commentTextRange.Font.Bold = 0;           // 非粗体
                                break;
                            case DiffType.Inserted:
                                commentTextRange.Font.Color = Word.WdColor.wdColorGreen;
                                commentTextRange.Font.Bold = -1;          // 粗体
                                commentTextRange.Font.StrikeThrough = 0;  // 无删除线
                                break;
                            case DiffType.Unchanged:
                                commentTextRange.Font.Color = Word.WdColor.wdColorBlack;
                                commentTextRange.Font.Bold = 0;           // 非粗体
                                commentTextRange.Font.StrikeThrough = 0;  // 无删除线
                                break;
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(WordOllama.Properties.Resources.遇到错误 + ex.Message);
            }
            finally
            {
                if (enabledTrackRevisions)
                    document.TrackRevisions = false;
            }
        }

        /// <summary>
        /// Applies only changed character spans, preserving unchanged Word formatting and
        /// producing focused tracked revisions instead of replacing the whole paragraph.
        /// </summary>
        public void ReplaceTextPrecisely(Word.Range sourceRange, string revisedText, bool manageTrackRevisions = true)
        {
            if (sourceRange == null) throw new ArgumentNullException(nameof(sourceRange));

            Word.Document document = sourceRange.Document;
            string originalText = sourceRange.Text ?? "";
            string targetText = TextDiffService.PreserveStructuralSuffix(originalText, revisedText);
            var edits = TextDiffService.CreateEdits(originalText, targetText);
            if (edits.Count == 0) return;

            bool enabledTrackRevisions = manageTrackRevisions && !document.TrackRevisions;
            int rangeStart = sourceRange.Start;
            try
            {
                if (enabledTrackRevisions)
                    document.TrackRevisions = true;

                for (int i = edits.Count - 1; i >= 0; i--)
                {
                    TextEdit edit = edits[i];
                    Word.Range editRange = null;
                    try
                    {
                        int editStart = rangeStart + edit.OriginalStart;
                        editRange = document.Range(editStart, editStart + edit.DeleteLength);
                        editRange.Text = edit.InsertText ?? "";
                    }
                    finally
                    {
                        ComHelper.ReleaseCom(ref editRange);
                    }
                }
            }
            finally
            {
                if (enabledTrackRevisions)
                    document.TrackRevisions = false;
            }
        }
        public async Task GenerateAndInsertTextBelowSelection(Word.Range selectionRange, string prompt)
        {
            // 获取当前文档对象
            Word.Document document = selectionRange.Document;
            bool originalTrackRevisions = document.TrackRevisions;

            try
            {
                // 启用修订模式
                document.TrackRevisions = true;

                // 将范围折叠到结尾并插入新段落
                selectionRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                selectionRange.InsertParagraphAfter();
                selectionRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                var pendingText = new StringBuilder();
                var pendingLock = new object();
                var dispatcher = System.Windows.Application.Current?.Dispatcher
                    ?? System.Windows.Threading.Dispatcher.FromThread(Thread.CurrentThread);

                Action<string> processChunk = chunk =>
                {
                    if (string.IsNullOrEmpty(chunk)) return;
                    lock (pendingLock)
                    {
                        pendingText.Append(chunk);
                    }
                };

                Action flushPendingText = () =>
                {
                    string text;
                    lock (pendingLock)
                    {
                        if (pendingText.Length == 0) return;
                        text = pendingText.ToString();
                        pendingText.Clear();
                    }

                    Action insert = () =>
                    {
                        selectionRange.InsertAfter(text);
                        selectionRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    };
                    if (dispatcher != null && !dispatcher.CheckAccess())
                        dispatcher.Invoke(insert);
                    else
                        insert();
                };

                Task streamTask = GenerateModelSeletorStream(prompt, processChunk);
                while (!streamTask.IsCompleted)
                {
                    await Task.WhenAny(streamTask, Task.Delay(120));
                    flushPendingText();
                }
                await streamTask;
                flushPendingText();

                // 在生成的文本后插入换行
                selectionRange.InsertAfter("\n");
            }
            finally
            {
                // 恢复原始的修订模式设置
                document.TrackRevisions = originalTrackRevisions;
            }
        }






        public List<DiffResult> GetDifferences(string originalText, string newText)
        {
            var differ = new Differ();
            var diffResult = differ.CreateCharacterDiffs(originalText, newText, ignoreWhitespace: false, ignoreCase: false);

            var differences = new List<DiffResult>();
            int originalIndex = 0;
            int newIndex = 0;

            foreach (var diffPiece in diffResult.DiffBlocks)
            {
                // 添加未改变的部分
                if (diffPiece.InsertStartB > newIndex)
                {
                    differences.Add(new DiffResult
                    {
                        Type = DiffType.Unchanged,
                        Text = newText.Substring(newIndex, diffPiece.InsertStartB - newIndex)
                    });
                }

                // 添加删除的部分
                if (diffPiece.DeleteCountA > 0)
                {
                    differences.Add(new DiffResult
                    {
                        Type = DiffType.Deleted,
                        Text = originalText.Substring(diffPiece.DeleteStartA, diffPiece.DeleteCountA)
                    });
                }

                // 添加插入的部分
                if (diffPiece.InsertCountB > 0)
                {
                    differences.Add(new DiffResult
                    {
                        Type = DiffType.Inserted,
                        Text = newText.Substring(diffPiece.InsertStartB, diffPiece.InsertCountB)
                    });
                }

                originalIndex = diffPiece.DeleteStartA + diffPiece.DeleteCountA;
                newIndex = diffPiece.InsertStartB + diffPiece.InsertCountB;
            }

            // 添加最后未改变的部分（如果有）
            if (newIndex < newText.Length)
            {
                differences.Add(new DiffResult
                {
                    Type = DiffType.Unchanged,
                    Text = newText.Substring(newIndex)
                });
            }

            return differences;
        }
        public enum DiffType
        {
            Unchanged,
            Deleted,
            Inserted
        }

        public class DiffResult
        {
            public DiffType Type { get; set; }
            public string Text { get; set; }
        }






        public void AddCommentToSelection(Word.Range selectedRange, string text)
        {
            // 如果选择的范围为空或仅包含空白字符，则退出
            if (selectedRange == null || string.IsNullOrWhiteSpace(selectedRange.Text))
            {
                MessageBox.Show(WordOllama.Properties.Resources.请选择有效的文本进行批注);
                return;
            }

            string formattedText = ConvertMarkdownToWordFormat(text, selectedRange.Document);
            Word.Comment comment = selectedRange.Comments.Add(selectedRange, formattedText);

            // 添加批注
           // Word.Comment comment = selectedRange.Comments.Add(selectedRange, text);
        }


        private string ConvertMarkdownToWordFormat(string markdown, Word.Document document)
        {
            if (string.IsNullOrEmpty(markdown)) return markdown;

            // 存储转换后的文本
            Word.Range tempRange = document.Range(0, 0);
            tempRange.Text = markdown;

            // 处理标题
            // H1
            markdown = Regex.Replace(markdown, @"^# (.+)$", match =>
            {
                tempRange.Text = match.Groups[1].Value;
                tempRange.Font.Size = 16;
                tempRange.Font.Bold = 1;
                return tempRange.Text;
            }, RegexOptions.Multiline);

            // H2
            markdown = Regex.Replace(markdown, @"^## (.+)$", match =>
            {
                tempRange.Text = match.Groups[1].Value;
                tempRange.Font.Size = 14;
                tempRange.Font.Bold = 1;
                return tempRange.Text;
            }, RegexOptions.Multiline);

            // H3
            markdown = Regex.Replace(markdown, @"^### (.+)$", match =>
            {
                tempRange.Text = match.Groups[1].Value;
                tempRange.Font.Size = 12;
                tempRange.Font.Bold = 1;
                return tempRange.Text;
            }, RegexOptions.Multiline);

            // 处理粗体
            markdown = Regex.Replace(markdown, @"\*\*(.+?)\*\*", match =>
            {
                tempRange.Text = match.Groups[1].Value;
                tempRange.Font.Bold = 1;
                return tempRange.Text;
            });

            // 处理斜体
            markdown = Regex.Replace(markdown, @"\*(.+?)\*", match =>
            {
                tempRange.Text = match.Groups[1].Value;
                tempRange.Font.Italic = 1;
                return tempRange.Text;
            });

            // 处理无序列表
            markdown = Regex.Replace(markdown, @"^- (.+)$", match =>
            {
                return "• " + match.Groups[1].Value;
            }, RegexOptions.Multiline);

            // 处理有序列表
            int listCounter = 1;
            markdown = Regex.Replace(markdown, @"^(\d+)\. (.+)$", match =>
            {
                return $"{listCounter++}. {match.Groups[2].Value}";
            }, RegexOptions.Multiline);

            // 清理临时范围
            tempRange.Text = "";

            return markdown;
        }

        public async Task<String> GenerateModelSeletor(
            string inputPrompt,
            string imageBase64 = null,
            string SPMode = null,
            bool requireJsonOutput = false,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            Globals.Ribbons.Ribbon1.AIStatus.Label = WordOllama.Properties.Resources.AI工作中;
            if (!DocumentAgent.IsAgentRunning)
            {
                Globals.Ribbons.Ribbon1.ShowNotification(WordOllama.Properties.Resources.AI开始工作了请稍等, 2000);
                ShowAgentNotification(WordOllama.Properties.Resources.AI开始工作了请稍等);
            }
            string responseLanguage = PromptTemplateService.ResolveLanguage(
                AppConfig.Instance.LanguageMode,
                inputPrompt);
            inputPrompt = PromptTemplateService.BuildReplyLanguagePrefix(responseLanguage) + inputPrompt;

            string SPMode_static = null;
            if (SPMode == "IMG")
            {
                SPMode_static = SharedSettings.IMGMode;
            }
            else if (SPMode == "HTML")
            {
                SPMode_static = SharedSettings.HTMLMode;
            }



            try
            {
                if (!string.IsNullOrEmpty(imageBase64) && SPMode == null && AppConfig.Instance.AIMode != "Ollama" && !ProviderRequestPolicy.GetActiveCapabilities().SupportsVision)
                    throw new InvalidOperationException(Loc.T("I18n_ProfileVisionDisabled"));
                if (SPMode == null || SPMode_static == "Global")
                {
                    if (AppConfig.Instance.AIMode == "Ollama")
                    {

                        string response = await OllamaFunctions.OllamaGenerateAsync(
                            inputPrompt,
                            imageBase64,
                            cancellationToken);
                        return response;
                    }
                    else
                    {
                        bool requestJsonOutput = requireJsonOutput && ProviderRequestPolicy.GetActiveCapabilities().SupportsJsonOutput;
                        var provider = ProviderFactory.CreateFromConfig(requestJsonOutput);
                        string response = provider is ICancellableOnlineProvider cancellable
                            ? await cancellable.GenerateAsync(inputPrompt, imageBase64, null, cancellationToken)
                            : await provider.GenerateAsync(inputPrompt, imageBase64);
                        return response;
                    }
                }
                else {

                    string endpoint = SPMode == "IMG" ? SharedSettings.IMGEndPint : SharedSettings.HTMLEndPint;
                    string apiKey = SPMode == "IMG" ? SharedSettings.IMGAPIKEY : SharedSettings.HTMLAPIKEY;
                    string modelName = SPMode == "IMG" ? SharedSettings.IMGModel : SharedSettings.HTMLModel;

                    var openAIProvider = new OpenAIProvider(endpoint, apiKey, modelName);
                    string response = await openAIProvider.GenerateAsync(
                        inputPrompt,
                        imageBase64,
                        null,
                        cancellationToken);
                    return response;

                }
                
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Globals.Ribbons.Ribbon1.ShowNotification(WordOllama.Properties.Resources.发生错误 + ex.Message, 3000, ToolTipIcon.Error);
                string message = ex.Message;
                return message;
            }
            finally
            {
                Globals.Ribbons.Ribbon1.AIStatus.Label = WordOllama.Properties.Resources.AI空闲中;
            }
        }

        /// <summary>
        /// 基于聊天历史的 AI 调用 — DocumentAgent 多轮对话架构使用。
        /// 将完整的消息列表发送给 Ollama /api/chat 或 OpenAI /chat/completions，
        /// 由 API 原生管理上下文窗口和历史记忆。
        /// </summary>
        public async Task<string> GenerateModelSeletorChat(List<Dictionary<string, string>> chatHistory)
        {
            Globals.Ribbons.Ribbon1.AIStatus.Label = WordOllama.Properties.Resources.AI工作中;
            if (!DocumentAgent.IsAgentRunning)
                Globals.Ribbons.Ribbon1.ShowNotification(WordOllama.Properties.Resources.AI开始工作了请稍等, 2000);

            try
            {
                chatHistory = ProviderRequestPolicy.TrimChatHistoryToContextWindow(chatHistory);
                if (AppConfig.Instance.AIMode == "Ollama")
                {
                    string response = await ExecuteProviderChatWithRetryAsync(
                        ct => OllamaFunctions.OllamaChatAsync(chatHistory, null, ct),
                        CancellationToken.None);
                    return response;
                }
                else
                {
                    var provider = ProviderFactory.CreateFromConfig();
                    string response = await ExecuteProviderChatWithRetryAsync(
                        ct => provider is ICancellableOnlineProvider cancellable
                            ? cancellable.ChatAsync(chatHistory, null, ct)
                            : provider.ChatAsync(chatHistory),
                        CancellationToken.None);
                    return response;
                }
            }
            catch (Exception ex)
            {
                Globals.Ribbons.Ribbon1.ShowNotification(WordOllama.Properties.Resources.发生错误 + ex.Message, 3000, ToolTipIcon.Error);
                return AgentToolText.F("I18n_AgentResultApiError", ex.Message);
            }
            finally
            {
                Globals.Ribbons.Ribbon1.AIStatus.Label = WordOllama.Properties.Resources.AI空闲中;
            }
        }

        /// <summary>
        /// 带原生 Tool Calling 支持的 AI 调用 — 用于 DocumentAgent。
        /// 与 GenerateModelSeletorChat 的区别在于会传递 tools 参数，
        /// 并在 API 返回 tool_calls 时自动解析为 [TOOL_CALL] 标记格式。
        /// 若 API 不支持 tools，返回 [TOOL_CALL_NOT_SUPPORTED] 供调用方降级。
        /// </summary>
        public async Task<string> GenerateModelSeletorChatWithTools(
            List<Dictionary<string, string>> chatHistory,
            JArray tools,
            CancellationToken ct)
        {
            Globals.Ribbons.Ribbon1.AIStatus.Label = WordOllama.Properties.Resources.AI工作中;
            if (!DocumentAgent.IsAgentRunning)
                Globals.Ribbons.Ribbon1.ShowNotification(WordOllama.Properties.Resources.AI开始工作了请稍等, 2000);

            try
            {
                chatHistory = ProviderRequestPolicy.TrimChatHistoryToContextWindow(chatHistory);
                if (AppConfig.Instance.AIMode == "Ollama")
                {
                    if (ProviderCompatibility.ParseToolCallingMode(AppConfig.Instance.ProviderToolCallingMode) == ToolCallingMode.ReAct)
                        return "[TOOL_CALL_NOT_SUPPORTED]";
                    return await ExecuteProviderChatWithRetryAsync(
                        requestCt => OllamaFunctions.OllamaChatAsync(chatHistory, tools, requestCt), ct);
                }
                else
                {
                    if (ProviderCompatibility.ParseToolCallingMode(AppConfig.Instance.ProviderToolCallingMode) == ToolCallingMode.ReAct)
                        return "[TOOL_CALL_NOT_SUPPORTED]";
                    var provider = ProviderFactory.CreateFromConfig();
                    return await ExecuteProviderChatWithRetryAsync(
                        requestCt => provider is ICancellableOnlineProvider cancellableProvider
                            ? cancellableProvider.ChatAsync(chatHistory, tools, requestCt)
                            : provider.ChatAsync(chatHistory, tools), ct);
                }
            }
            catch (Exception ex)
            {
                Globals.Ribbons.Ribbon1.ShowNotification(WordOllama.Properties.Resources.发生错误 + ex.Message, 3000, ToolTipIcon.Error);
                return AgentToolText.F("I18n_AgentResultApiError", ex.Message);
            }
            finally
            {
                Globals.Ribbons.Ribbon1.AIStatus.Label = WordOllama.Properties.Resources.AI空闲中;
            }
        }

        private async Task<string> ExecuteProviderChatWithRetryAsync(
            Func<CancellationToken, Task<string>> request,
            CancellationToken cancellationToken)
        {
            int maxRetries = Math.Max(0, Math.Min(3, AppConfig.Instance.ProviderMaxRetries));
            for (int retryIndex = 0; ; retryIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                string response = await request(cancellationToken);
                if (!ProviderRequestPolicy.IsTransientError(response) || retryIndex >= maxRetries)
                    return response;

                await Task.Delay(ProviderRequestPolicy.GetRetryDelay(retryIndex), cancellationToken);
            }
        }

        string GetSharedSetting(string fieldName)
        {
            // 使用反射获取 SharedSettings 中的字段值
            var fieldInfo = typeof(SharedSettings).GetField(fieldName,
                          System.Reflection.BindingFlags.Public);
            if (fieldInfo != null)
            {
                return fieldInfo.GetValue(null)?.ToString();
            }
            throw new Exception($"未找到名为 {fieldName} 的静态字段。");
        }

        public async Task GenerateModelSeletorStream(string inputPrompt, Action<string> processChunk)
        {
            try
            {
                if (AppConfig.Instance.AIMode != "Ollama" && !ProviderRequestPolicy.GetActiveCapabilities().SupportsStreaming)
                {
                    processChunk(await GenerateModelSeletor(inputPrompt));
                    return;
                }
                if (AppConfig.Instance.AIMode == "Ollama")
                {
                    var cts = new CancellationTokenSource();
                    try
                    {
                        await foreach (var chunk in OllamaFunctions.OllamaStreamAsync(inputPrompt, cts.Token))
                        {
                            try
                            {
                                processChunk(chunk);
                            }
                            catch (OperationCanceledException)
                            {
                                // 捕获取消操作，确保响应取消请求
                                Debug.WriteLine("流式处理已被取消。");
                                break;  // 跳出流式处理
                            }
                            catch (Exception ex)
                            {
                                // 处理其他异常
                                Debug.WriteLine($"处理流块时发生错误: {ex.Message}");
                            }
                        }

                    }
                    catch (OperationCanceledException)
                    {
                        // 取消操作的处理逻辑
                        Debug.WriteLine("请求已被取消");
                    }
                }

                else
                {
                    var provider = ProviderFactory.CreateFromConfig();
                    await provider.StreamAsync(inputPrompt, processChunk);
                }
            }
            catch (Exception ex)
            {
                processChunk($"Error: {ex.Message}");
            }
        }

        public void OpenWebPage(string url)
        {
            try
            {
                // 使用默认浏览器打开指定的URL
                Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                // 处理可能的异常，例如浏览器未能启动
                MessageBox.Show(WordOllama.Properties.Resources.无法打开网页 + ex.Message);
            }
        }

        public async Task LanguageCheck(string txt)
        {
            if (AppConfig.Instance.LanguageMode == "Auto")
            {
                string prompt = PromptTemplateService.BuildLanguageDetectionPrompt(txt);
                try
                {
                    txtLang = await GenerateModelSeletor(prompt);
                }
                catch (Exception ex)
                {
                    Globals.Ribbons.Ribbon1.ShowNotification(Loc.T("I18n_GenericErrorPrefix") + ex.Message, 1000, ToolTipIcon.Error);
                    txtLang = null;
                }
            }
        }

        public async Task<Word.Document> CompareAndAnalyzeDifferences(string path1, string path2, string whoAmI, IProgress<string> progress)
        {
            Word.Application wordApp = new();
            Word.Document doc1 = null;
            Word.Document doc2 = null;
            Word.Document comparisonDoc = null;
            Word.Document outputDoc = null;
            bool keepOutputDocumentOpen = false;
            try
            {
                progress.Report(WordOllama.Properties.Resources.正在打开文档);
                doc1 = wordApp.Documents.Open(path1);
                doc2 = wordApp.Documents.Open(path2);

                doc1.TrackRevisions = true;
                doc1.AcceptAllRevisions();

                doc2.TrackRevisions = true;
                doc2.AcceptAllRevisions();


                progress.Report(WordOllama.Properties.Resources.对比文档中);
                comparisonDoc = wordApp.CompareDocuments(
                    OriginalDocument: doc1,
                    RevisedDocument: doc2,
                    CompareWhitespace: false
                );

                progress.Report(WordOllama.Properties.Resources.创建对比文档);
                outputDoc = wordApp.Documents.Add();
                doc1?.Close(SaveChanges: false);
                doc2?.Close(SaveChanges: false);

                progress.Report(WordOllama.Properties.Resources.分析中);

                int totalRevisions = comparisonDoc.Revisions.Count;
                Debug.WriteLine($"发现{totalRevisions}处修订");
                int processedItems = 0;

                List<string> original_text = new();
                List<string> new_text = new();

                foreach (Word.Revision rev in comparisonDoc.Revisions)
                {
                    try
                    {
                        Word.Range revRange = rev.Range.Duplicate;
                        // 获取修订所在段落的全部文本
                        Word.Range paragraphRange = revRange.Paragraphs[1].Range;


                        // 扩展范围以确保涵盖整个段落
                        paragraphRange.Expand(Word.WdUnits.wdParagraph);

                        foreach (Word.Revision thisParagraphRev in paragraphRange.Revisions)
                        {
                            Debug.WriteLine(thisParagraphRev.Type.ToString());
                            thisParagraphRev.Accept();
                        }
                        Debug.WriteLine(paragraphRange.Text);
                        if (!string.IsNullOrWhiteSpace(paragraphRange.Text))
                        {
                            Debug.WriteLine($"NT:{paragraphRange.Text}");
                            new_text.Add(paragraphRange.Text.ToString());

                        }
                        else
                        {
                            new_text.Add(WordOllama.Properties.Resources.对方删除了这条文本);
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"错误：{ex.Message}");
                    }
                }

                comparisonDoc.Undo(totalRevisions);

                foreach (Word.Revision rev in comparisonDoc.Revisions)
                {
                    try
                    {
                        Word.Range revRange = rev.Range.Duplicate;
                        // 获取修订所在段落的全部文本
                        Word.Range paragraphRange = revRange.Paragraphs[1].Range;

                        // 扩展范围以确保涵盖整个段落
                        paragraphRange.Expand(Word.WdUnits.wdParagraph);

                        foreach (Word.Revision thisParagraphRev in paragraphRange.Revisions)
                        {
                            thisParagraphRev.Reject();
                        }

                        if (!string.IsNullOrWhiteSpace(paragraphRange.Text))
                        {
                            Debug.WriteLine($"OT:{paragraphRange.Text}");
                            original_text.Add(paragraphRange.Text.ToString());

                        }
                        else
                        {
                            original_text.Add(WordOllama.Properties.Resources.无原文为对方新增文本);
                        }


                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"错误：{ex.Message}");
                    }
                }

                int totalItems = new_text.Count;

                for (int i = 0; i < totalItems; i++)
                {
                    try
                    {
                        string originalText = original_text[i];
                        string newText = new_text[i];

                        if (originalText != newText)
                        {
                            Debug.WriteLine("开始AI分析...");
                            string analysis = await AnalyzeRevisionWithAI(originalText, newText, whoAmI);
                            Debug.WriteLine(analysis);

                            outputDoc.Content.InsertAfter($"{WordOllama.Properties.Resources.原文本} \n{originalText}\n\n");
                            outputDoc.Content.InsertAfter($"{WordOllama.Properties.Resources.新文本} \n{newText}\n\n");
                            outputDoc.Content.InsertAfter($"{WordOllama.Properties.Resources.AI分析} \n{analysis}\n\n");
                            outputDoc.Content.InsertAfter(new string('-', 50) + "\n");

                            processedItems++;
                            progress.Report($"{WordOllama.Properties.Resources.已分析} {processedItems} / {totalItems} {WordOllama.Properties.Resources.处修订}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"处理修订时出错: {ex.Message}");
                    }
                }

                progress.Report($"{WordOllama.Properties.Resources.评估结束共处理} {processedItems} {WordOllama.Properties.Resources.处有效修订}");
                keepOutputDocumentOpen = true;
                return outputDoc;

            }
            catch (Exception)
            {
                outputDoc?.Close(SaveChanges: false);
                throw;
            }
            finally
            {
                comparisonDoc?.Close(SaveChanges: false);
                if (!keepOutputDocumentOpen)
                {
                    try { doc1?.Close(SaveChanges: false); } catch { }
                    try { doc2?.Close(SaveChanges: false); } catch { }
                    try
                    {
                        object saveChanges = false;
                        object missing = Type.Missing;
                        wordApp?.Quit(ref saveChanges, ref missing, ref missing);
                    }
                    catch { }
                }
            }


        }

        public Word.Document OpenInteractiveComparison(string originalPath, string revisedPath)
        {
            Word.Document original = null;
            Word.Document revised = null;
            try
            {
                original = Application.Documents.Open(originalPath, ReadOnly: true, Visible: false);
                revised = Application.Documents.Open(revisedPath, ReadOnly: true, Visible: false);
                Word.Document comparison = Application.CompareDocuments(
                    OriginalDocument: original,
                    RevisedDocument: revised,
                    Destination: Word.WdCompareDestination.wdCompareDestinationNew,
                    Granularity: Word.WdGranularity.wdGranularityWordLevel,
                    CompareFormatting: true,
                    CompareCaseChanges: true,
                    CompareWhitespace: false,
                    CompareTables: true,
                    CompareHeaders: true,
                    CompareFootnotes: true,
                    CompareTextboxes: true,
                    CompareFields: true,
                    CompareComments: true,
                    CompareMoves: true);
                original.Close(SaveChanges: false);
                revised.Close(SaveChanges: false);
                original = revised = null;
                comparison.Activate();
                return comparison;
            }
            finally
            {
                try { original?.Close(SaveChanges: false); } catch { }
                try { revised?.Close(SaveChanges: false); } catch { }
            }
        }

        private async Task<string> AnalyzeRevisionWithAI(string originalText, string newText, string whoAmI)
        {
            string finalPrompt = PromptTemplateService.BuildRevisionAnalysisPrompt(whoAmI, originalText, newText);
            string analysis = await GenerateModelSeletor(finalPrompt);

            return analysis;
        }

        private string GetRevisionTypeDescription(Word.WdRevisionType revisionType)
        {
            Debug.WriteLine(revisionType);
            try
            {
                switch (revisionType)
                {
                    case Word.WdRevisionType.wdRevisionInsert:
                        return Loc.T("I18n_RevisionInsertText");
                    case Word.WdRevisionType.wdRevisionDelete:
                        return Loc.T("I18n_RevisionDeleteText");
                    case Word.WdRevisionType.wdRevisionProperty:
                        return Loc.T("I18n_RevisionPropertyChanged");
                    case Word.WdRevisionType.wdRevisionStyle:
                        return Loc.T("I18n_RevisionStyleChanged");
                    case Word.WdRevisionType.wdRevisionReplace:
                        return Loc.T("I18n_RevisionReplaceText");
                    case Word.WdRevisionType.wdRevisionParagraphProperty:
                        return Loc.T("I18n_RevisionParagraphPropertyChanged");
                    case Word.WdRevisionType.wdRevisionTableProperty:
                        return Loc.T("I18n_RevisionTablePropertyChanged");
                    case Word.WdRevisionType.wdRevisionSectionProperty:
                        return Loc.T("I18n_RevisionSectionPropertyChanged");
                    case Word.WdRevisionType.wdRevisionMovedFrom:
                        return Loc.T("I18n_RevisionMovedFrom");
                    case Word.WdRevisionType.wdRevisionMovedTo:
                        return Loc.T("I18n_RevisionMovedTo");
                    // 可以根据需要添加更多的类型
                    default:
                        return Loc.T("I18n_RevisionOther");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"处理修订类型时出错: {ex.Message}");
                return Loc.T("I18n_RevisionUnknown");
            }
        }







        #region 静默审查

        /// <summary>保存时触发审查（作为定时轮询的保底）。</summary>
        private void OnDocumentBeforeSave(Word.Document doc, ref bool saveAsUi, ref bool cancel)
        {
            if (_silentLinter == null || !_silentLinter.IsEnabled) return;
            try
            {
                var task = _silentLinter.OnDocumentBeforeSaveAsync(doc);
                task.ContinueWith(t =>
                {
                    if (t.IsFaulted)
                        DebugLogger.Log("SilentLinter", $"保存审查异常: {t.Exception?.InnerException?.Message}");
                }, TaskContinuationOptions.OnlyOnFaulted);
            }
            catch (Exception ex)
            {
                DebugLogger.Log("SilentLinter", $"保存审查失败: {ex.Message}");
            }
        }

        /// <summary>静默审查专用的模型调用 — 若配置了独立审查模型则使用，否则回退到主模型。</summary>
        private async Task<string> CallLinterModelAsync(string prompt)
        {
            string linterModel = AppConfig.Instance.LinterModel;

            // 未配置独立审查模型时，用主模型名走专有路径（确保 temperature=0.1 的确定性输出）
            if (string.IsNullOrEmpty(linterModel))
                linterModel = AppConfig.Instance.AIMode == "Ollama"
                    ? AppConfig.Instance.OllamaModel
                    : AppConfig.Instance.ChatGPTModel;

            if (AppConfig.Instance.AIMode == "Ollama")
            {
                string endpoint = (string.IsNullOrEmpty(AppConfig.Instance.LinterEndPoint)
                    ? AppConfig.Instance.OllamaEndPoint
                    : AppConfig.Instance.LinterEndPoint) + "/api/generate";

                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromMinutes(5);
                    var requestData = new Dictionary<string, object>
                    {
                        ["model"] = linterModel,
                        ["prompt"] = prompt,
                        ["stream"] = false,
                        ["options"] = new Dictionary<string, object>
                        {
                            ["temperature"] = 0.1,
                            ["num_ctx"] = 4096
                        }
                    };

                    var postData = new StringContent(
                        JsonConvert.SerializeObject(requestData),
                        Encoding.UTF8,
                        "application/json"
                    );

                    var response = await client.PostAsync(endpoint, postData);
                    if (response.IsSuccessStatusCode)
                    {
                        string json = await response.Content.ReadAsStringAsync();
                        return JObject.Parse(json)["response"]?.ToString() ?? "";
                    }
                    return "";
                }
            }
            else
            {
                string endpoint = string.IsNullOrEmpty(AppConfig.Instance.LinterEndPoint)
                    ? AppConfig.Instance.ChatGPTEndPoint
                    : AppConfig.Instance.LinterEndPoint;

                if (string.IsNullOrWhiteSpace(endpoint)) return "";

                var chatClient = new OpenAIProvider(
                    endpoint,
                    AppConfig.Instance.ChatGPTKey,
                    linterModel
                );
                return await chatClient.GenerateAsync(prompt, null, 0.1);
            }
        }

        #endregion

        #region VSTO generated code

        /// <summary>  
        /// Designer support required method - do not modify  
        /// the contents of this method with the code editor.  
        /// </summary>  
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
