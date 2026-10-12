using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WordOllama.lib;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.NewUI
{
    public partial class AgentTaskPaneUI
    {
        private bool _agentReviewWorkflowActive;
        private CancellationToken _agentReviewToken;
        private string _issueReviewInstruction;

        private Task<string> ExecuteReviewAgentToolAsync(string action, JObject args, Word.Document document, CancellationToken token)
        {
            if (Dispatcher.CheckAccess()) return ExecuteReviewAgentToolOnUiAsync(action, args, document, token);
            return Dispatcher.InvokeAsync(() => ExecuteReviewAgentToolOnUiAsync(action, args, document, token)).Task.Unwrap();
        }

        private async Task<string> ExecuteReviewAgentToolOnUiAsync(string action, JObject args, Word.Document document, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            if (_agentReviewWorkflowActive || _reviewCts != null || _isLoadingReviewScope)
                throw new InvalidOperationException(UiText.T("I18n_UiAgentRunningReview"));
            var active = Globals.ThisAddIn.Application.ActiveDocument;
            if (active == null || document == null || !string.Equals(active.FullName, document.FullName, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(UiText.T("I18n_ReviewSourceChanged"));
            _agentReviewWorkflowActive = true; _agentReviewToken = token;
            try
            {
                FlushReviewSuggestionEdits();
                if (action == "list_review_sessions")
                    return JsonConvert.SerializeObject(_reviewSessionStore.List(document.FullName, null)
                        .Select(m => new { session_id = m.Id, kind = m.Kind, state = m.State.ToString(), m.CreatedUtc,
                            m.CompletedItems, m.PendingItems, m.FailedItems, document_path = m.DocumentIdentity,
                            copy_path = m.WorkingCopyPath, reviewer = m.ReviewerIdentity }));
                if (action == "review_document")
                {
                    if (args["create_copy"]?.Value<bool>() != false)
                    {
                        string extension = ReviewCopyService.ExtensionFor(ReviewCopyService.SaveFormatFor(document));
                        string folder = string.IsNullOrEmpty(document.Path) ? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) : document.Path;
                        string path = args["copy_path"]?.Value<string>();
                        if (string.IsNullOrWhiteSpace(path)) path = Path.Combine(folder, Path.GetFileNameWithoutExtension(document.Name)
                            + ".review." + DateTime.Now.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N").Substring(0, 6) + extension);
                        if (!Path.IsPathRooted(path)) throw new ArgumentException("copy_path must be an absolute path.");
                        _reviewCopyOriginalPath = ReviewCopyService.SaveAsCopy(document, path);
                        _reviewCopyPath = document.FullName;
                    }
                    else { _reviewCopyOriginalPath = null; _reviewCopyPath = null; }
                    _issueRecord = null; _suggestionRecord = null;
                    _reviewIssues.Clear(); _editSuggestions.Clear();
                    if (args["kind"]?.Value<string>() == "issues")
                    {
                        _issueReviewInstruction = args["instruction"]?.Value<string>();
                        await RunManualReviewAsync(false);
                    }
                    else
                    {
                        txtReviewInstruction.Text = args["instruction"]?.Value<string>() ?? Loc.T("I18n_ReviewAgentDefaultInstruction");
                        await LoadDocumentForReviewAsync();
                        await GenerateAllSuggestionsAsync();
                    }
                }
                else if (action == "resume_review")
                {
                    var record = _reviewSessionStore.Load(args["session_id"]?.Value<string>());
                    if (!string.Equals(record.Metadata.DocumentIdentity, document.FullName, StringComparison.OrdinalIgnoreCase))
                        return JsonConvert.SerializeObject(new { session_id = record.Metadata.Id, requires_open_document = record.Metadata.DocumentIdentity,
                            error = UiText.T("I18n_ReviewSessionSourceMismatch") });
                    if (!ReviewSessionStore.CanResume(record, document.FullName, ReviewDocumentHash(document, record.Metadata.SchemaVersion)))
                        throw new InvalidOperationException(UiText.T("I18n_ReviewSessionSourceMismatch"));
                    RestoreReviewRecord(record);
                    if (record.Metadata.Kind == "Issues") await RunManualReviewAsync(record.Metadata.Scope == "Selection", resume: true);
                    else await GenerateAllSuggestionsAsync(resume: true);
                }
                else if (action == "apply_review_revisions")
                {
                    if (_suggestionRecord == null) throw new InvalidOperationException(UiText.T("I18n_ReviewSessionNotLoaded"));
                    ValidateReviewResume(_suggestionRecord, requireConfiguration: false);
                    if (_suggestionRecord.Items.Any(i => i.NeedsProcessing)) throw new InvalidOperationException(Loc.T("I18n_ReviewAgentIncomplete"));
                    // Only a separate, write-classified Agent tool can apply changes.
                    ApplyAllSuggestions("replace");
                }
                else if (action != "save_review_document") throw new ArgumentException("Unknown review action.");

                token.ThrowIfCancellationRequested();
                var current = _suggestionRecord ?? _issueRecord;
                if (current == null) throw new InvalidOperationException(UiText.T("I18n_ReviewSessionNotLoaded"));
                ValidateReviewResume(current, requireConfiguration: false);
                if (action == "save_review_document")
                {
                    if (string.IsNullOrEmpty(document.Path)) throw new InvalidOperationException(UiText.T("I18n_ReviewSessionSaveDocumentFirst"));
                    document.Save();
                    SaveReviewRecord(current);
                    if (args["export_report"]?.Value<bool>() != false)
                    {
                        string report = Path.Combine(document.Path, Path.GetFileNameWithoutExtension(document.Name) + "." + current.Metadata.Id + ".review");
                        ReviewReportExporter.Write(current, report + ".md");
                        ReviewReportExporter.Write(current, report + ".json");
                    }
                }
                else if (!string.IsNullOrEmpty(current.Metadata.WorkingCopyPath)) document.Save();
                var metadata = current.Metadata; current.RefreshSummary();
                return JsonConvert.SerializeObject(new { session_id = metadata.Id, kind = metadata.Kind,
                    state = metadata.State.ToString(), completed = metadata.CompletedItems, total = metadata.TotalItems,
                    failed = metadata.FailedItems, pending = metadata.PendingItems, document_path = metadata.DocumentIdentity,
                    copy_path = metadata.WorkingCopyPath, reviewer = metadata.ReviewerIdentity,
                    applied = current.Items.Count(i => i.UserAction == "replace"),
                    suggestions_require_apply = metadata.Kind == "Suggestions" && current.Items.Any(i => i.State == ReviewItemState.Succeeded && string.IsNullOrEmpty(i.UserAction)),
                    coverage = ReviewReportExporter.ScopeDescription(metadata), error = metadata.Error });
            }
            finally { _agentReviewWorkflowActive = false; _agentReviewToken = default(CancellationToken); }
        }

        private void RestoreReviewRecord(ReviewSessionRecord record)
        {
            _reviewPersistenceFailed = false;
            _reviewCopyOriginalPath = record.Metadata.OriginalDocumentPath; _reviewCopyPath = record.Metadata.WorkingCopyPath;
            if (record.Metadata.Kind == "Issues")
            {
                _suggestionRecord = null; _issueRecord = record; _reviewIssues.Clear();
                _issueReviewInstruction = record.Metadata.ReviewTaskInstruction;
                AddReviewIssues(record.Items.SelectMany(i => i.Findings).Where(i => i.UserAction != "ignored"));
            }
            else
            {
                _issueRecord = null; _suggestionRecord = record;
                txtReviewInstruction.Text = record.Metadata.Instruction ?? ""; txtWritingProfile.Text = record.Metadata.WritingProfile ?? "";
                _editSuggestions.Clear();
                foreach (var item in record.Items.Where(i => string.IsNullOrEmpty(i.UserAction)))
                    _editSuggestions.Add(new EditSuggestionSession { Id = item.Id, Index = item.Index, Story = item.Story,
                        ParagraphIndex = item.FirstParagraph, RangeStart = item.RangeStart, RangeEnd = item.RangeEnd,
                        BookmarkName = item.BookmarkName, OriginalText = item.OriginalText, OriginalHash = item.OriginalHash,
                        DocumentIdentity = record.Metadata.DocumentIdentity, SuggestedText = item.SuggestedText,
                        GenerationSucceeded = item.State == ReviewItemState.Succeeded, FailureReason = item.Error });
                UpdateSuggestionEmptyState(); SelectReviewTab();
            }
        }
    }
}
