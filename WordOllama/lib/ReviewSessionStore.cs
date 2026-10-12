using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using WordOllama.NewUI;

namespace WordOllama.lib
{
    [JsonConverter(typeof(StringEnumConverter))]
    public enum ReviewItemState { Pending, Running, Succeeded, Failed }

    [JsonConverter(typeof(StringEnumConverter))]
    public enum ReviewRunState { Ready, Running, Completed, Partial, Cancelled, Failed }

    public sealed class ReviewModelSettings
    {
        public string Mode { get; set; }
        public string Provider { get; set; }
        public string Model { get; set; }
        public int ContextTokens { get; set; }
        public int OutputTokens { get; set; }
        public double Temperature { get; set; }
        public bool ThinkingEnabled { get; set; }
        public string ReasoningEffort { get; set; }
        public int IssueParagraphs { get; set; }
        public int IssueCharacters { get; set; }
        public int SelectionCharacters { get; set; }
        public int SuggestionItems { get; set; }
        public int SuggestionCharacters { get; set; }
        public int ConcurrentRequests { get; set; }
        public int TimeoutSeconds { get; set; }
    }

    public sealed class ReviewSessionMetadata
    {
        public int SchemaVersion { get; set; } = 2;
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        public string Kind { get; set; }
        public string Scope { get; set; }
        public string DocumentIdentity { get; set; }
        public string DocumentName { get; set; }
        public bool IsSavedDocument { get; set; }
        public string ExpectedDocumentHash { get; set; }
        public string OriginalDocumentPath { get; set; }
        public string WorkingCopyPath { get; set; }
        public string ReviewerIdentity { get; set; }
        public int RangeStart { get; set; }
        public int RangeEnd { get; set; }
        public string ConfigurationFingerprint { get; set; }
        public ReviewModelSettings ModelSettings { get; set; } = new ReviewModelSettings();
        public List<ReviewSessionRangeShift> RangeShifts { get; set; } = new List<ReviewSessionRangeShift>();
        public string Instruction { get; set; }
        public string ReviewTaskInstruction { get; set; }
        public string WritingProfile { get; set; }
        public DateTime CreatedUtc { get; set; } = DateTime.UtcNow;
        public DateTime UpdatedUtc { get; set; } = DateTime.UtcNow;
        public DateTime? StartedUtc { get; set; }
        public DateTime? FinishedUtc { get; set; }
        public ReviewRunState State { get; set; } = ReviewRunState.Ready;
        public string Error { get; set; }
        public List<ReviewRunLogEntry> RunLog { get; set; } = new List<ReviewRunLogEntry>();
        public int TotalItems { get; set; }
        public int CompletedItems { get; set; }
        public int FailedItems { get; set; }
        public int PendingItems { get; set; }
        public long TotalCharacters { get; set; }
        public long ReviewedCharacters { get; set; }
        public long FailedCharacters { get; set; }
        public long UnreviewedCharacters { get; set; }
        [JsonIgnore] public string CreatedLabel => CreatedUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.CurrentCulture);
        [JsonIgnore] public string RunTimeLabel => (StartedUtc ?? CreatedUtc).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.CurrentCulture);
        [JsonIgnore] public string KindLabel => Loc.T(Kind == "Issues" ? "I18n_UiIssuesTab" : "I18n_UiReviewTab");
        [JsonIgnore] public string StateLabel => Loc.T("I18n_ReviewSessionState" + State);
        [JsonIgnore] public string ProgressLabel => CompletedItems + "/" + TotalItems + " · " + FailedItems;
    }

    public sealed class ReviewSessionFragment
    {
        public int Start { get; set; }
        public string Source { get; set; }
        public string Suggestion { get; set; }
        public ReviewItemState State { get; set; }
        public string Error { get; set; }
        public int Attempts { get; set; }
    }

    public sealed class ReviewSessionAttempt
    {
        public DateTime StartedUtc { get; set; } = DateTime.UtcNow;
        public DateTime? FinishedUtc { get; set; }
        public int? FragmentStart { get; set; }
        public ReviewItemState State { get; set; } = ReviewItemState.Running;
        public string Error { get; set; }
    }

    public sealed class ReviewRunLogEntry
    {
        public DateTime AtUtc { get; set; } = DateTime.UtcNow;
        public ReviewRunState State { get; set; }
        public string Error { get; set; }
    }

    public sealed class ReviewSessionRangeShift
    {
        public string StoryKey { get; set; }
        public string EditedItemId { get; set; }
        public int Boundary { get; set; }
        public int Delta { get; set; }
    }

    public sealed class ReviewSessionItem
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        public int Index { get; set; }
        public int FirstParagraph { get; set; }
        public int LastParagraph { get; set; }
        public int RangeStart { get; set; }
        public int RangeEnd { get; set; }
        public int PositionRevision { get; set; }
        public string BookmarkName { get; set; }
        public ReviewStoryAddress Story { get; set; }
        public Dictionary<int, ReviewParagraphLocation> ParagraphLocations { get; set; } = new Dictionary<int, ReviewParagraphLocation>();
        public string OriginalText { get; set; }
        public string OriginalHash { get; set; }
        public string SuggestedText { get; set; }
        public string GeneratedText { get; set; }
        public string Context { get; set; }
        public int Characters { get; set; }
        public Dictionary<int, string> ParagraphHashes { get; set; } = new Dictionary<int, string>();
        public List<ReviewIssue> Findings { get; set; } = new List<ReviewIssue>();
        public List<ReviewSessionFragment> Fragments { get; set; } = new List<ReviewSessionFragment>();
        public List<ReviewSessionAttempt> AttemptLog { get; set; } = new List<ReviewSessionAttempt>();
        public ReviewItemState State { get; set; }
        public string Error { get; set; }
        public int Attempts { get; set; }
        public DateTime? StartedUtc { get; set; }
        public DateTime? FinishedUtc { get; set; }
        public string ReviewerIdentity { get; set; }
        public string UserAction { get; set; }
        public DateTime? ActionUtc { get; set; }
        [JsonIgnore] public bool NeedsProcessing => State != ReviewItemState.Succeeded && string.IsNullOrEmpty(UserAction);

        public void BeginAttempt(int? fragmentStart = null)
        {
            AttemptLog.Add(new ReviewSessionAttempt { FragmentStart = fragmentStart });
        }

        public void FinishAttempt(ReviewItemState state, string error = null, int? fragmentStart = null)
        {
            var attempt = AttemptLog.LastOrDefault(a => a.State == ReviewItemState.Running && a.FragmentStart == fragmentStart);
            if (attempt == null) return;
            attempt.State = state; attempt.Error = error; attempt.FinishedUtc = DateTime.UtcNow;
        }
    }

    public sealed class ReviewSessionRecord
    {
        public ReviewSessionMetadata Metadata { get; set; } = new ReviewSessionMetadata();
        public List<ReviewSessionItem> Items { get; set; } = new List<ReviewSessionItem>();
        private Dictionary<string, ReviewSessionItem> _index;
        public ReviewSessionItem FindItem(string id)
        {
            if (_index == null || _index.Count != Items.Count) _index = Items.ToDictionary(item => item.Id, StringComparer.Ordinal);
            return id != null && _index.TryGetValue(id, out var item) ? item : null;
        }

        public void RefreshSummary()
        {
            Metadata.TotalItems = Items.Count;
            Metadata.CompletedItems = Items.Count(i => i.State == ReviewItemState.Succeeded);
            Metadata.FailedItems = Items.Count(i => i.State == ReviewItemState.Failed);
            Metadata.PendingItems = Items.Count - Metadata.CompletedItems - Metadata.FailedItems;
            Metadata.TotalCharacters = Items.Sum(i => (long)i.Characters);
            Metadata.ReviewedCharacters = Items.Sum(i => i.State == ReviewItemState.Succeeded ? (long)i.Characters
                : i.Fragments.Where(f => f.State == ReviewItemState.Succeeded).Sum(f => (long)(f.Source?.Length ?? 0)));
            Metadata.FailedCharacters = Items.Sum(i => i.State == ReviewItemState.Failed && i.Fragments.Count == 0 ? (long)i.Characters
                : i.Fragments.Where(f => f.State == ReviewItemState.Failed).Sum(f => (long)(f.Source?.Length ?? 0)));
            Metadata.UnreviewedCharacters = Math.Max(0, Metadata.TotalCharacters - Metadata.ReviewedCharacters - Metadata.FailedCharacters);
        }

        public void ShiftRangesAfterEdit(string editedItemId, int boundary, int delta, string storyKey = null)
        {
            if (delta == 0) return;
            Metadata.RangeShifts.Add(new ReviewSessionRangeShift { EditedItemId = editedItemId, Boundary = boundary, Delta = delta, StoryKey = storyKey });
            ReplayRangeShifts();
        }

        public void ReplayRangeShifts()
        {
            if (Metadata.RangeShifts == null) throw new InvalidDataException("Missing review position journal.");
            foreach (var item in Items)
            {
                if (item.PositionRevision < 0 || item.PositionRevision > Metadata.RangeShifts.Count)
                    throw new InvalidDataException("Incomplete review position journal.");
                for (int i = item.PositionRevision; i < Metadata.RangeShifts.Count; i++)
                {
                    var shift = Metadata.RangeShifts[i];
                    if (shift == null || shift.Boundary < 0) throw new InvalidDataException("Invalid review position change.");
                    if ((item.Story?.Key ?? "") == (shift.StoryKey ?? "") && item.Id != shift.EditedItemId && string.IsNullOrEmpty(item.UserAction) && item.RangeStart >= shift.Boundary)
                    {
                        item.RangeStart = checked(item.RangeStart + shift.Delta);
                        item.RangeEnd = checked(item.RangeEnd + shift.Delta);
                        if (item.RangeStart < 0 || item.RangeEnd < item.RangeStart) throw new InvalidDataException("Invalid review position.");
                    }
                }
                item.PositionRevision = Metadata.RangeShifts.Count;
            }
        }
    }

    /// <summary>Encrypted per-item journal: completing one batch never rewrites all document text.</summary>
    public sealed class ReviewSessionStore
    {
        private readonly string _root;
        public ReviewSessionStore(string root = null)
        {
            _root = root ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WordOllama", "ReviewSessions");
        }

        public void Create(ReviewSessionRecord record)
        {
            Validate(record.Metadata.Id);
            if (record.Items.Select(i => i.Id).Distinct(StringComparer.Ordinal).Count() != record.Items.Count)
                throw new InvalidDataException("Duplicate review item identifiers.");
            string directory = DirectoryFor(record.Metadata.Id);
            if (File.Exists(Path.Combine(directory, "metadata.dat"))) throw new IOException("Review session already exists.");
            Directory.CreateDirectory(directory);
            foreach (var item in record.Items) WriteProtected(ItemPath(record.Metadata.Id, item.Id), item);
            WriteProtected(Path.Combine(directory, "plan.dat"), record.Items.Select(i => i.Id).ToList());
            SaveMetadata(record);
        }

        public void SaveItem(ReviewSessionRecord record, ReviewSessionItem item)
        {
            if (!ReferenceEquals(record.FindItem(item.Id), item)) throw new InvalidDataException("Item is not part of this review session.");
            WriteProtected(ItemPath(record.Metadata.Id, item.Id), item);
            SaveMetadata(record);
        }

        public void SaveMetadata(ReviewSessionRecord record)
        {
            record.RefreshSummary();
            record.Metadata.UpdatedUtc = DateTime.UtcNow;
            WriteProtected(Path.Combine(DirectoryFor(record.Metadata.Id), "metadata.dat"), record.Metadata);
        }

        public List<ReviewSessionMetadata> List(string documentIdentity, string kind)
        {
            var result = new List<ReviewSessionMetadata>();
            if (!Directory.Exists(_root)) return result;
            foreach (string directory in Directory.EnumerateDirectories(_root))
            {
                try
                {
                    string id = Path.GetFileName(directory);
                    Validate(id);
                    var metadata = ReadProtected<ReviewSessionMetadata>(Path.Combine(directory, "metadata.dat"));
                    if ((metadata.SchemaVersion == 1 || metadata.SchemaVersion == 2) && metadata.Id == id && (string.IsNullOrEmpty(kind) || metadata.Kind == kind)
                        && (string.Equals(metadata.DocumentIdentity, documentIdentity, StringComparison.OrdinalIgnoreCase)
                            || string.Equals(metadata.OriginalDocumentPath, documentIdentity, StringComparison.OrdinalIgnoreCase))) result.Add(metadata);
                }
                catch (Exception ex) when (ex is IOException || ex is InvalidDataException || ex is JsonException || ex is CryptographicException || ex is ArgumentException) { }
            }
            return result.OrderByDescending(m => m.UpdatedUtc).ToList();
        }

        public ReviewSessionRecord Load(string id)
        {
            string directory = DirectoryFor(id);
            var metadata = ReadProtected<ReviewSessionMetadata>(Path.Combine(directory, "metadata.dat"));
            if (metadata.Id != id || (metadata.SchemaVersion != 1 && metadata.SchemaVersion != 2) || metadata.ModelSettings == null || metadata.RunLog == null)
                throw new InvalidDataException("Unsupported review session.");
            var ids = ReadProtected<List<string>>(Path.Combine(directory, "plan.dat"));
            if (ids.Distinct(StringComparer.Ordinal).Count() != ids.Count) throw new InvalidDataException("Duplicate review plan identifiers.");
            var record = new ReviewSessionRecord { Metadata = metadata };
            foreach (string itemId in ids)
            {
                var item = ReadProtected<ReviewSessionItem>(ItemPath(id, itemId));
                if (item.Id != itemId || item.Characters < 0 || item.Fragments == null || item.Findings == null || item.ParagraphHashes == null || item.AttemptLog == null
                    || item.Fragments.Any(f => f == null || f.Source == null || f.Start < 0 || f.State == ReviewItemState.Succeeded && f.Suggestion == null)
                    || item.AttemptLog.Any(a => a == null))
                    throw new InvalidDataException("Invalid review item.");
                if (item.State == ReviewItemState.Running) item.State = ReviewItemState.Pending;
                foreach (var fragment in item.Fragments)
                    if (fragment.State == ReviewItemState.Running) fragment.State = ReviewItemState.Pending;
                foreach (var attempt in item.AttemptLog)
                    if (attempt.State == ReviewItemState.Running) attempt.State = ReviewItemState.Pending;
                record.Items.Add(item);
            }
            record.ReplayRangeShifts();
            record.RefreshSummary();
            if (metadata.State == ReviewRunState.Running)
                metadata.State = record.Items.All(i => i.State == ReviewItemState.Succeeded) ? ReviewRunState.Completed : ReviewRunState.Partial;
            return record;
        }

        public static bool CanResume(ReviewSessionRecord record, string documentIdentity, string documentHash)
        {
            return record?.Metadata.IsSavedDocument == true
                && string.Equals(record.Metadata.DocumentIdentity, documentIdentity, StringComparison.OrdinalIgnoreCase)
                && !string.IsNullOrEmpty(documentHash) && record.Metadata.ExpectedDocumentHash == documentHash;
        }

        public static string Fingerprint(string text)
        {
            using (var hash = SHA256.Create()) return Convert.ToBase64String(hash.ComputeHash(Encoding.UTF8.GetBytes(text ?? "")));
        }

        public static string DocumentFingerprint(string text, int end, int paragraphs)
        {
            return Fingerprint((text ?? "") + "\u001f" + end.ToString(CultureInfo.InvariantCulture) + "\u001f" + paragraphs.ToString(CultureInfo.InvariantCulture));
        }

        private string DirectoryFor(string id) { Validate(id); return Path.Combine(_root, id); }
        private string ItemPath(string sessionId, string itemId) { Validate(itemId); return Path.Combine(DirectoryFor(sessionId), "item-" + itemId + ".dat"); }
        private static void Validate(string id)
        {
            if (!Guid.TryParseExact(id, "N", out _)) throw new ArgumentException("Invalid review identifier.");
        }

        private static void WriteProtected(string path, object value)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                File.WriteAllText(temporary, SecretProtector.Protect(JsonConvert.SerializeObject(value)), new UTF8Encoding(false));
                if (File.Exists(path)) File.Replace(temporary, path, path + ".bak");
                else File.Move(temporary, path);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }

        private static T ReadProtected<T>(string path) where T : class
        {
            try { return ReadOne<T>(path); }
            catch (Exception ex) when (ex is IOException || ex is InvalidDataException || ex is JsonException || ex is CryptographicException)
            {
                if (!File.Exists(path + ".bak")) throw;
                return ReadOne<T>(path + ".bak");
            }
        }

        private static T ReadOne<T>(string path) where T : class
        {
            string protectedText = File.ReadAllText(path, Encoding.UTF8);
            if (!SecretProtector.IsProtected(protectedText)) throw new InvalidDataException("Review session is not protected.");
            string json = SecretProtector.UnprotectOrLegacy(protectedText);
            if (string.IsNullOrEmpty(json)) throw new InvalidDataException("Review session cannot be decrypted.");
            var value = JsonConvert.DeserializeObject<T>(json);
            if (value == null) throw new InvalidDataException("Empty review session data.");
            return value;
        }
    }
}
