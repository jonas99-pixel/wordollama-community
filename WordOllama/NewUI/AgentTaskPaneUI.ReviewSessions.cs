using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Threading;
using Microsoft.Win32;
using Newtonsoft.Json;
using WordOllama.lib;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.NewUI
{
    public partial class AgentTaskPaneUI
    {
        private readonly ReviewSessionStore _reviewSessionStore = new ReviewSessionStore();
        private ReviewSessionRecord _issueRecord;
        private ReviewSessionRecord _suggestionRecord;
        private string _reviewCopyOriginalPath;
        private string _reviewCopyPath;
        private bool _reviewPersistenceFailed;
        private bool _deferReviewCopySave;
        private bool _bulkReviewSourceValidated;
        private ReviewSessionRecord _deferredReviewCopyRecord;
        private readonly HashSet<EditSuggestionSession> _observedReviewSuggestions = new HashSet<EditSuggestionSession>();
        private readonly HashSet<string> _pendingReviewEdits = new HashSet<string>();
        private DispatcherTimer _reviewEditSaveTimer;

        private void InitializeReviewSessions()
        {
            _reviewEditSaveTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(600) };
            _reviewEditSaveTimer.Tick += (_, __) => FlushReviewSuggestionEdits();
            _editSuggestions.CollectionChanged += (_, change) =>
            {
                if (change.Action == NotifyCollectionChangedAction.Reset)
                {
                    foreach (var session in _observedReviewSuggestions) session.PropertyChanged -= ReviewSuggestionEdited;
                    _observedReviewSuggestions.Clear();
                }
                if (change.OldItems != null) foreach (EditSuggestionSession session in change.OldItems)
                {
                    session.PropertyChanged -= ReviewSuggestionEdited;
                    _observedReviewSuggestions.Remove(session);
                }
                if (change.NewItems != null) foreach (EditSuggestionSession session in change.NewItems)
                {
                    session.PropertyChanged += ReviewSuggestionEdited;
                    _observedReviewSuggestions.Add(session);
                }
            };
            SetReviewSessionActionsEnabled(true);
        }

        private void ReviewSuggestionEdited(object sender, PropertyChangedEventArgs change)
        {
            if (change.PropertyName != nameof(EditSuggestionSession.SuggestedText)
                || !(sender is EditSuggestionSession session) || session.IsBusy || !session.GenerationSucceeded
                || _reviewCts != null || _suggestionRecord?.FindItem(session.Id) == null) return;
            _pendingReviewEdits.Add(session.Id);
            _reviewEditSaveTimer.Stop();
            _reviewEditSaveTimer.Start();
        }

        private void FlushReviewSuggestionEdits()
        {
            _reviewEditSaveTimer?.Stop();
            try
            {
                foreach (string id in _pendingReviewEdits)
                {
                    var item = _suggestionRecord?.FindItem(id);
                    var session = _observedReviewSuggestions.FirstOrDefault(s => s.Id == id);
                    if (item == null || session == null) continue;
                    item.SuggestedText = session.SuggestedText;
                    SaveReviewRecord(_suggestionRecord, item);
                }
            }
            catch (Exception ex) { AddAgentBubble(ex.Message, ChatMessageType.System); }
            finally { _pendingReviewEdits.Clear(); }
        }

        private static string ReviewDocumentHash(Word.Document document, int schemaVersion = 2)
        {
            if (schemaVersion >= 2) return ReviewDocumentScope.Fingerprint(document);
            Word.Range content = null;
            Word.Paragraphs paragraphs = null;
            try
            {
                content = document.Content; paragraphs = content.Paragraphs;
                return ReviewSessionStore.DocumentFingerprint(content.Text, content.End, paragraphs.Count);
            }
            finally { ComHelper.ReleaseCom(ref paragraphs); ComHelper.ReleaseCom(ref content); }
        }

        private string ReviewIdentity(ReviewSessionRecord record)
        {
            var c = AppConfig.Instance;
            DateTime date = c.ReviewReviewerIdentityMode == "Session" && record != null
                ? record.Metadata.CreatedUtc.ToLocalTime() : DateTime.Now;
            string identity = ReviewRevisionIdentity.Build(c.ReviewReviewerName, c.ReviewReviewerIdentityMode, date, record?.Metadata.Id);
            if (record != null) record.Metadata.ReviewerIdentity = identity;
            return identity;
        }

        private ReviewModelSettings CaptureReviewModelSettings()
        {
            var c = AppConfig.Instance;
            bool local = c.AIMode == "Ollama";
            string provider = local ? "Ollama" : c.OnlineProviderType;
            string model = local ? c.OllamaModel : provider == "Claude" ? c.ClaudeModel : provider == "Gemini" ? c.GeminiModel : c.ChatGPTModel;
            return new ReviewModelSettings
            {
                Mode = c.AIMode, Provider = provider, Model = model,
                ContextTokens = local ? c.OllamaMaxToken : c.OnlineContextWindow,
                OutputTokens = local ? Math.Min(Math.Max(256, c.ReviewOllamaOutputTokens), c.OllamaMaxToken / 2) : c.ChatGPTMaxToken,
                Temperature = local ? c.OllamaTemperature : c.ChatGPTTemperature,
                ThinkingEnabled = c.ThinkingEnabled, ReasoningEffort = c.ReasoningEffort,
                IssueParagraphs = c.ReviewIssueBatchMaxParagraphs, IssueCharacters = c.ReviewIssueBatchMaxCharacters,
                SelectionCharacters = c.ReviewSelectionBatchMaxCharacters, SuggestionItems = c.ReviewSuggestionBatchMaxItems,
                SuggestionCharacters = c.ReviewSuggestionBatchMaxCharacters, ConcurrentRequests = c.ReviewMaxConcurrentRequests,
                TimeoutSeconds = local ? c.OllamaTimeoutSeconds : c.ProviderTimeoutSeconds
            };
        }

        private string ReviewConfigurationFingerprint(string kind)
        {
            string instruction = kind == "Suggestions" ? (txtReviewInstruction.Text?.Trim() ?? "") : BuildIssueReviewPrompt("");
            string profile = kind == "Suggestions" ? (txtWritingProfile.Text?.Trim() ?? "") : "";
            return ReviewSessionStore.Fingerprint(GetReviewProviderKey() + "\u001f" + JsonConvert.SerializeObject(CaptureReviewModelSettings())
                + "\u001f" + instruction + "\u001f" + profile);
        }

        private ReviewSessionRecord NewReviewRecord(Word.Document document, string kind, string scope, int start, int end)
        {
            _reviewPersistenceFailed = false;
            bool copy = string.Equals(document.FullName, _reviewCopyPath, StringComparison.OrdinalIgnoreCase);
            var record = new ReviewSessionRecord { Metadata = new ReviewSessionMetadata
            {
                SchemaVersion = 2,
                Kind = kind, Scope = scope, DocumentIdentity = document.FullName, DocumentName = document.Name,
                IsSavedDocument = !string.IsNullOrEmpty(document.Path), ExpectedDocumentHash = ReviewDocumentHash(document),
                OriginalDocumentPath = copy ? _reviewCopyOriginalPath : document.FullName,
                WorkingCopyPath = copy ? _reviewCopyPath : null, RangeStart = start, RangeEnd = end,
                ConfigurationFingerprint = ReviewConfigurationFingerprint(kind), ModelSettings = CaptureReviewModelSettings(),
                Instruction = kind == "Suggestions" ? (txtReviewInstruction.Text?.Trim() ?? "") : BuildIssueReviewPrompt(""),
                ReviewTaskInstruction = kind == "Issues" ? _issueReviewInstruction : null,
                WritingProfile = kind == "Suggestions" ? (txtWritingProfile.Text?.Trim() ?? "") : ""
            } };
            ReviewIdentity(record);
            return record;
        }

        private async Task CreateSuggestionRecordAsync(Word.Document doc, Word.Range range, bool wholeDocument = false)
        {
            _suggestionRecord = null;
            if (_editSuggestions.Count == 0) return;
            Word.Range content = null;
            bool wholeBody;
            try { content = doc.Content; wholeBody = range.Start == content.Start && range.End >= content.End - 1; }
            finally { ComHelper.ReleaseCom(ref content); }
            var record = NewReviewRecord(doc, "Suggestions", wholeDocument ? "Document" : wholeBody ? "Body" : "Selection", range.Start, range.End);
            record.Items = _editSuggestions.Select(s => new ReviewSessionItem
            {
                Id = s.Id, Index = s.Index, FirstParagraph = s.ParagraphIndex, LastParagraph = s.ParagraphIndex,
                Story = s.Story, RangeStart = s.RangeStart, RangeEnd = s.RangeEnd, BookmarkName = s.BookmarkName,
                OriginalText = s.OriginalText, OriginalHash = s.OriginalHash, Characters = s.OriginalText?.Length ?? 0
            }).ToList();
            await Task.Run(() => _reviewSessionStore.Create(record));
            _suggestionRecord = record;
            ShowReviewRecordSaved(record);
        }

        private void ShowReviewRecordSaved(ReviewSessionRecord record)
        {
            AddAgentBubble(UiText.F("I18n_ReviewSessionSaved", record.Metadata.CreatedLabel, record.Metadata.ModelSettings.Model), ChatMessageType.System);
            if (!record.Metadata.IsSavedDocument) AddAgentBubble(UiText.T("I18n_ReviewSessionSaveDocumentFirst"), ChatMessageType.System);
            SetReviewSessionActionsEnabled(_reviewCts == null && !_isLoadingReviewScope && !_isRunning);
        }

        private void SaveReviewRecord(ReviewSessionRecord record, ReviewSessionItem item = null)
        {
            if (record == null) return;
            try
            {
                if (item == null) _reviewSessionStore.SaveMetadata(record);
                else _reviewSessionStore.SaveItem(record, item);
            }
            catch (Exception ex)
            {
                _reviewPersistenceFailed = true;
                _reviewCts?.Cancel();
                record.Metadata.Error = UiText.F("I18n_ReviewSessionSaveFailed", ex.Message);
                throw new IOException(record.Metadata.Error, ex);
            }
        }

        private void FinishReviewRecord(ReviewSessionRecord record, bool cancelled = false, string error = null)
        {
            if (record == null) return;
            record.Metadata.Error = error ?? (_reviewPersistenceFailed ? record.Metadata.Error : null);
            record.Metadata.State = _reviewPersistenceFailed ? ReviewRunState.Failed : cancelled ? ReviewRunState.Cancelled
                : !string.IsNullOrEmpty(error) ? ReviewRunState.Failed
                : record.Items.All(i => i.State == ReviewItemState.Succeeded) ? ReviewRunState.Completed : ReviewRunState.Partial;
            record.Metadata.FinishedUtc = DateTime.UtcNow;
            record.Metadata.RunLog.Add(new ReviewRunLogEntry { State = record.Metadata.State, Error = record.Metadata.Error });
            try { SaveReviewRecord(record); }
            catch (Exception ex) { AddAgentBubble(ex.Message, ChatMessageType.System); }
        }

        private void ValidateReviewResume(ReviewSessionRecord record, bool requireConfiguration = true)
        {
            var document = Globals.ThisAddIn.Application.ActiveDocument;
            // In-memory retry also works for a new document. Reopening a saved run
            // separately requires CanResume, including a stable file identity.
            if (document == null || record == null
                || !string.Equals(record.Metadata.DocumentIdentity, document.FullName, StringComparison.OrdinalIgnoreCase)
                || record.Metadata.ExpectedDocumentHash != ReviewDocumentHash(document, record.Metadata.SchemaVersion))
                throw new InvalidOperationException(UiText.T("I18n_ReviewSessionSourceMismatch"));
            if (requireConfiguration && record.Metadata.ConfigurationFingerprint != ReviewConfigurationFingerprint(record.Metadata.Kind))
                throw new InvalidOperationException(UiText.T("I18n_ReviewSessionConfigurationMismatch"));
        }

        private async Task PrepareSuggestionRunAsync(bool resume)
        {
            FlushReviewSuggestionEdits();
            if (resume)
            {
                if (_suggestionRecord == null) throw new InvalidOperationException(UiText.T("I18n_ReviewSessionNotLoaded"));
                ValidateReviewResume(_suggestionRecord);
            }
            else
            {
                var doc = Globals.ThisAddIn.Application.ActiveDocument;
                if (doc == null || _editSuggestions.Any(s => !string.Equals(s.DocumentIdentity, doc.FullName, StringComparison.OrdinalIgnoreCase))
                    || _suggestionRecord != null && _suggestionRecord.Metadata.ExpectedDocumentHash != ReviewDocumentHash(doc, _suggestionRecord.Metadata.SchemaVersion))
                    throw new InvalidOperationException(UiText.T("I18n_ReviewSourceChanged"));
                if (_suggestionRecord == null || _suggestionRecord.Metadata.StartedUtc.HasValue)
                {
                    var record = NewReviewRecord(doc, "Suggestions", _suggestionRecord?.Metadata.Scope ?? "Selection",
                        _editSuggestions.Min(s => s.RangeStart), _editSuggestions.Max(s => s.RangeEnd));
                    record.Items = _editSuggestions.Select(s => new ReviewSessionItem
                    {
                        Id = s.Id, Index = s.Index, FirstParagraph = s.ParagraphIndex, LastParagraph = s.ParagraphIndex,
                        Story = s.Story, RangeStart = s.RangeStart, RangeEnd = s.RangeEnd, BookmarkName = s.BookmarkName,
                        OriginalText = s.OriginalText, OriginalHash = s.OriginalHash, Characters = s.OriginalText?.Length ?? 0
                    }).ToList();
                    await Task.Run(() => _reviewSessionStore.Create(record));
                    _suggestionRecord = record;
                }
                _suggestionRecord.Metadata.ModelSettings = CaptureReviewModelSettings();
                _suggestionRecord.Metadata.Instruction = txtReviewInstruction.Text?.Trim() ?? "";
                _suggestionRecord.Metadata.WritingProfile = txtWritingProfile.Text?.Trim() ?? "";
                _suggestionRecord.Metadata.ConfigurationFingerprint = ReviewConfigurationFingerprint("Suggestions");
            }
            _reviewPersistenceFailed = false;
            _suggestionRecord.Metadata.StartedUtc = _suggestionRecord.Metadata.StartedUtc ?? DateTime.UtcNow;
            _suggestionRecord.Metadata.FinishedUtc = null;
            _suggestionRecord.Metadata.State = ReviewRunState.Running;
            _suggestionRecord.Metadata.Error = null;
            SaveReviewRecord(_suggestionRecord);
        }

        private void ValidateRunningReviewSource(ReviewSessionRecord record)
        {
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            if (doc == null || record == null || !string.Equals(record.Metadata.DocumentIdentity, doc.FullName, StringComparison.OrdinalIgnoreCase)
                || record.Metadata.ExpectedDocumentHash != ReviewDocumentHash(doc, record.Metadata.SchemaVersion))
            {
                _reviewCts?.Cancel();
                throw new InvalidOperationException(UiText.T("I18n_ReviewSourceChanged"));
            }
            if (record.Metadata.ConfigurationFingerprint != ReviewConfigurationFingerprint(record.Metadata.Kind))
            {
                _reviewCts?.Cancel();
                throw new InvalidOperationException(UiText.T("I18n_ReviewConfigurationChanged"));
            }
        }

        private void RecordSuggestionState(EditSuggestionSession session, ReviewItemState state, string error = null)
        {
            var item = _suggestionRecord?.FindItem(session.Id);
            if (item == null) return;
            item.State = state; item.Error = error;
            if (state == ReviewItemState.Running)
            {
                item.Attempts++; item.StartedUtc = DateTime.UtcNow; item.FinishedUtc = null; item.SuggestedText = null; item.GeneratedText = null;
                item.BeginAttempt();
            }
            else
            {
                item.FinishAttempt(state, error);
                if (state == ReviewItemState.Succeeded || state == ReviewItemState.Failed) item.FinishedUtc = DateTime.UtcNow;
            }
            if (state == ReviewItemState.Succeeded) item.GeneratedText = item.SuggestedText = session.SuggestedText;
            if (!_reviewPersistenceFailed) SaveReviewRecord(_suggestionRecord, item);
        }

        private void RecordSuggestionAction(EditSuggestionSession session, string action, bool changedText = false, bool sourceWasConsistent = true, Word.Document updatedDocument = null)
        {
            var item = _suggestionRecord?.FindItem(session.Id);
            if (item == null) return;
            item.SuggestedText = session.SuggestedText;
            item.ReviewerIdentity = ReviewIdentity(_suggestionRecord);
            item.UserAction = action; item.ActionUtc = DateTime.UtcNow;
            var doc = updatedDocument ?? Globals.ThisAddIn.Application.ActiveDocument;
            if (changedText && doc != null && string.Equals(doc.FullName, _suggestionRecord.Metadata.DocumentIdentity, StringComparison.OrdinalIgnoreCase))
            {
                // A bulk application validates once, then updates the full stamp at its end.
                // Null during the write prevents a crashed, unsaved batch from being resumed as intact.
                _suggestionRecord.Metadata.ExpectedDocumentHash = sourceWasConsistent && !_bulkReviewSourceValidated
                    ? ReviewDocumentHash(doc, _suggestionRecord.Metadata.SchemaVersion) : null;
            }
            // The Word edit has already happened. Record it even if saving the copy
            // later fails, so restoring the journal cannot silently apply it twice.
            try { SaveReviewRecord(_suggestionRecord, item); }
            catch (Exception ex) { AddAgentBubble(ex.Message, ChatMessageType.System); }
            SaveWorkingReviewCopy(doc, _suggestionRecord, action == "replace" || action == "insert" || action == "comment");
            SetReviewSessionActionsEnabled(true);
        }

        private void RebaseReviewSuggestionRanges(string editedId, int oldEnd, int difference, string storyKey = null)
        {
            if (_suggestionRecord == null || difference == 0) return;
            _suggestionRecord.ShiftRangesAfterEdit(editedId, oldEnd, difference, storyKey);
            foreach (var card in _editSuggestions)
            {
                var item = _suggestionRecord.FindItem(card.Id);
                if (item == null) continue;
                card.RangeStart = item.RangeStart; card.RangeEnd = item.RangeEnd;
            }
        }

        private void RecordIssueAction(ReviewIssue issue, string action)
        {
            if (_issueRecord == null || issue.ReviewSessionId != _issueRecord.Metadata.Id) return;
            var item = _issueRecord.FindItem(issue.ReviewItemId);
            if (item == null) return;
            issue.UserAction = action; issue.ActionUtc = DateTime.UtcNow;
            item.ReviewerIdentity = ReviewIdentity(_issueRecord);
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            try { SaveReviewRecord(_issueRecord, item); }
            catch (Exception ex) { AddAgentBubble(ex.Message, ChatMessageType.System); }
            SaveWorkingReviewCopy(doc, _issueRecord, action == "commented");
        }

        private void SaveWorkingReviewCopy(Word.Document document, ReviewSessionRecord record, bool modified)
        {
            if (!modified || document == null || !string.Equals(document.FullName, record.Metadata.WorkingCopyPath, StringComparison.OrdinalIgnoreCase)) return;
            if (_deferReviewCopySave) { _deferredReviewCopyRecord = record; return; }
            try { document.Save(); }
            catch (Exception ex) { AddAgentBubble(UiText.F("I18n_ReviewCopySaveFailed", ex.Message), ChatMessageType.System); }
        }

        private void FlushWorkingReviewCopy(Word.Document document)
        {
            if (_deferReviewCopySave || _deferredReviewCopyRecord == null) return;
            var record = _deferredReviewCopyRecord;
            _deferredReviewCopyRecord = null;
            SaveWorkingReviewCopy(document, record, true);
        }

        private static int ReviewDocumentEnd(Word.Document document)
        {
            Word.Range content = null;
            try { content = document.Content; return content.End; }
            finally { ComHelper.ReleaseCom(ref content); }
        }

        private void SetReviewSessionActionsEnabled(bool enabled)
        {
            foreach (Button button in new[] { btnCreateReviewCopy, btnCreateIssueReviewCopy, btnRestoreReviewSession, btnRestoreIssueSession,
                btnExportReviewReport, btnExportIssueReport }) if (button != null) button.IsEnabled = enabled;
            if (btnRetryReviewSuggestions != null) btnRetryReviewSuggestions.IsEnabled = enabled && _suggestionRecord?.Items.Any(i => i.NeedsProcessing) == true;
            if (btnRetryReviewIssues != null) btnRetryReviewIssues.IsEnabled = enabled && _issueRecord?.Items.Any(i => i.NeedsProcessing) == true;
        }

        private async void BtnCreateReviewCopy_Click(object sender, RoutedEventArgs e)
        {
            if (_reviewCts != null || _isRunning || _isLoadingReviewScope) return;
            try
            {
                var doc = Globals.ThisAddIn.Application.ActiveDocument;
                if (doc == null) throw new InvalidOperationException(UiText.T("I18n_UiOpenDocumentFirst"));
                FlushReviewSuggestionEdits();
                string extension = ReviewCopyService.ExtensionFor(ReviewCopyService.SaveFormatFor(doc));
                var dialog = new SaveFileDialog
                {
                    Title = UiText.T("I18n_ReviewCreateCopy"), DefaultExt = extension, AddExtension = true,
                    Filter = UiText.T("I18n_ReviewCopyFileType") + " (*" + extension + ")|*" + extension,
                    FileName = Path.GetFileNameWithoutExtension(doc.Name) + ".review." + DateTime.Now.ToString("yyyyMMdd-HHmmss") + extension,
                    InitialDirectory = string.IsNullOrEmpty(doc.Path) ? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) : doc.Path
                };
                if (dialog.ShowDialog() != true) return;
                _reviewCopyOriginalPath = ReviewCopyService.SaveAsCopy(doc, dialog.FileName);
                _reviewCopyPath = doc.FullName;
                _issueRecord = null; _suggestionRecord = null;
                _editSuggestions.Clear(); UpdateSuggestionEmptyState();
                _reviewIssues.Clear(); _issueReviewCoverage = null; _issueReviewIncomplete = false; UpdateIssueSummary();
                AddAgentBubble(UiText.F("I18n_ReviewCopyCreated", _reviewCopyPath), ChatMessageType.System);
                if ((sender as FrameworkElement)?.Tag?.ToString() == "Issues") SelectIssuesTab();
                else { SelectReviewTab(); await LoadDocumentForReviewAsync(); }
                SetReviewSessionActionsEnabled(true);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, UiText.ReviewTab, MessageBoxButton.OK, MessageBoxImage.Error); }
        }

        private void BtnRestoreReviewSession_Click(object sender, RoutedEventArgs e)
        {
            if (_reviewCts != null || _isRunning || _isLoadingReviewScope) return;
            try
            {
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;
                if (doc == null) throw new InvalidOperationException(UiText.T("I18n_UiOpenDocumentFirst"));
                string kind = (sender as FrameworkElement)?.Tag?.ToString() == "Issues" ? "Issues" : "Suggestions";
                var saved = _reviewSessionStore.List(doc.FullName, kind);
                if (saved.Count == 0) throw new InvalidOperationException(UiText.T("I18n_ReviewSessionNone"));
                var dialog = new ReviewSessionWindow(saved);
                new WindowInteropHelper(dialog).Owner = new IntPtr(app.ActiveWindow.Hwnd);
                if (dialog.ShowDialog() != true) return;
                var record = _reviewSessionStore.Load(dialog.SelectedSession.Id);
                if (dialog.ExportRequested) { ExportReviewRecord(record); return; }
                if (!ReviewSessionStore.CanResume(record, doc.FullName, ReviewDocumentHash(doc, record.Metadata.SchemaVersion)))
                    throw new InvalidOperationException(UiText.T("I18n_ReviewSessionSourceMismatch"));
                _reviewPersistenceFailed = false;
                FlushReviewSuggestionEdits();
                _reviewCopyOriginalPath = record.Metadata.OriginalDocumentPath;
                _reviewCopyPath = record.Metadata.WorkingCopyPath;
                if (kind == "Issues")
                {
                    _issueReviewInstruction = record.Metadata.ReviewTaskInstruction;
                    _issueRecord = record;
                    _reviewIssues.Clear();
                    AddReviewIssues(record.Items.SelectMany(i => i.Findings).Where(i => i.UserAction != "ignored"));
                    var m = record.Metadata;
                    _issueReviewCoverage = BuildReviewCoverage(m.CompletedItems, m.FailedItems, m.TotalItems,
                        m.ReviewedCharacters, m.FailedCharacters, m.TotalCharacters);
                    _issueReviewIncomplete = m.CompletedItems != m.TotalItems;
                    UpdateIssueSummary(); SelectIssuesTab();
                }
                else
                {
                    _suggestionRecord = record;
                    txtReviewInstruction.Text = record.Metadata.Instruction ?? "";
                    txtWritingProfile.Text = record.Metadata.WritingProfile ?? "";
                    _editSuggestions.Clear();
                    foreach (var item in record.Items.Where(i => string.IsNullOrEmpty(i.UserAction)))
                        _editSuggestions.Add(new EditSuggestionSession
                        {
                            Id = item.Id, Index = item.Index, ParagraphIndex = item.FirstParagraph,
                            Story = item.Story, RangeStart = item.RangeStart, RangeEnd = item.RangeEnd, BookmarkName = item.BookmarkName,
                            OriginalText = item.OriginalText, OriginalHash = item.OriginalHash, DocumentIdentity = record.Metadata.DocumentIdentity,
                            SuggestedText = item.State == ReviewItemState.Succeeded ? item.SuggestedText : null,
                            GenerationSucceeded = item.State == ReviewItemState.Succeeded, FailureReason = item.Error,
                            Status = UiText.T(item.State == ReviewItemState.Succeeded ? "I18n_UiGenerated" : item.State == ReviewItemState.Failed ? "I18n_UiGenerationFailed" : "I18n_UiReady")
                        });
                    UpdateSuggestionEmptyState(); SelectReviewTab();
                    SetReviewScopeStatus(UiText.F("I18n_ReviewSuggestionSummary", record.Metadata.CompletedItems, record.Metadata.FailedItems, record.Metadata.PendingItems));
                }
                AddAgentBubble(UiText.F("I18n_ReviewSessionRestored", record.Metadata.RunTimeLabel, record.Metadata.ModelSettings.Model), ChatMessageType.System);
                SetReviewSessionActionsEnabled(true);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, UiText.ReviewTab, MessageBoxButton.OK, MessageBoxImage.Error); }
        }

        private async void BtnRetryReviewSession_Click(object sender, RoutedEventArgs e)
        {
            if (_reviewCts != null || _isRunning || _isLoadingReviewScope) return;
            if ((sender as FrameworkElement)?.Tag?.ToString() == "Issues")
                await RunManualReviewAsync(_issueRecord?.Metadata.Scope == "Selection", resume: true);
            else await GenerateAllSuggestionsAsync(resume: true);
        }

        private void BtnExportReviewReport_Click(object sender, RoutedEventArgs e)
        {
            if (_reviewCts != null || _isRunning || _isLoadingReviewScope) return;
            var record = (sender as FrameworkElement)?.Tag?.ToString() == "Issues" ? _issueRecord : _suggestionRecord;
            try
            {
                if (record == null) throw new InvalidOperationException(UiText.T("I18n_ReviewSessionNotLoaded"));
                FlushReviewSuggestionEdits();
                ExportReviewRecord(record);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, UiText.ReviewTab, MessageBoxButton.OK, MessageBoxImage.Error); }
        }

        private void ExportReviewRecord(ReviewSessionRecord record)
        {
            var dialog = new SaveFileDialog
            {
                Title = UiText.T("I18n_ReviewExportReport"), DefaultExt = ".md", AddExtension = true,
                Filter = UiText.T("I18n_ReviewReportFilter"),
                FileName = Path.GetFileNameWithoutExtension(record.Metadata.DocumentName) + ".review-report."
                    + (record.Metadata.StartedUtc ?? record.Metadata.CreatedUtc).ToLocalTime().ToString("yyyyMMdd-HHmmss") + ".md"
            };
            if (dialog.ShowDialog() != true) return;
            ReviewReportExporter.Write(record, dialog.FileName);
            AddAgentBubble(UiText.F("I18n_ReviewReportExported", dialog.FileName), ChatMessageType.System);
        }
    }
}
