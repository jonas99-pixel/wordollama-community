using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace WordOllama.lib
{
    public static class ReviewReportExporter
    {
        public static string ScopeDescription(ReviewSessionMetadata metadata)
        {
            return Loc.T(metadata.Scope == "Selection" ? "I18n_ReviewSelectedArea"
                : metadata.SchemaVersion >= 2 && metadata.Scope == "Document" ? "I18n_ReviewAllStories" : "I18n_ReviewMainStoryOnly");
        }
        public static void Write(ReviewSessionRecord record, string path)
        {
            record.RefreshSummary();
            string report = string.Equals(Path.GetExtension(path), ".json", StringComparison.OrdinalIgnoreCase)
                ? BuildJson(record) : BuildMarkdown(record);
            File.WriteAllText(path, report, new UTF8Encoding(true));
        }

        public static string BuildJson(ReviewSessionRecord record)
        {
            record.RefreshSummary();
            var report = JObject.FromObject(record);
            ((JObject)report["Metadata"]).Remove("ConfigurationFingerprint");
            return report.ToString(Formatting.Indented);
        }

        public static string BuildMarkdown(ReviewSessionRecord record)
        {
            record.RefreshSummary();
            var m = record.Metadata;
            var sb = new StringBuilder();
            sb.AppendLine("# " + Loc.T("I18n_ReviewReportTitle"));
            sb.AppendLine();
            sb.AppendLine("| " + Loc.T("I18n_ReviewReportField") + " | " + Loc.T("I18n_ReviewReportValue") + " |");
            sb.AppendLine("| --- | --- |");
            Row(sb, Loc.T("I18n_ReviewSessionId"), m.Id);
            Row(sb, Loc.T("I18n_ReviewSessionDocument"), m.DocumentIdentity);
            Row(sb, Loc.T("I18n_ReviewCopyOriginal"), m.OriginalDocumentPath);
            Row(sb, Loc.T("I18n_ReviewCopyWorking"), m.WorkingCopyPath);
            Row(sb, Loc.T("I18n_ReviewSessionCreated"), m.CreatedUtc.ToLocalTime().ToString("O"));
            Row(sb, Loc.T("I18n_ReviewSessionUpdated"), m.UpdatedUtc.ToLocalTime().ToString("O"));
            if (m.StartedUtc.HasValue) Row(sb, Loc.T("I18n_ReviewSessionStarted"), m.StartedUtc.Value.ToLocalTime().ToString("O"));
            if (m.FinishedUtc.HasValue) Row(sb, Loc.T("I18n_ReviewSessionFinished"), m.FinishedUtc.Value.ToLocalTime().ToString("O"));
            Row(sb, Loc.T("I18n_ReviewSessionState"), m.StateLabel);
            Row(sb, Loc.T("I18n_ReviewSessionKind"), m.KindLabel + " / " + m.Scope);
            Row(sb, Loc.T("I18n_ReviewReviewerIdentity"), m.ReviewerIdentity);
            Row(sb, Loc.T("I18n_ReviewSessionModel"), m.ModelSettings.Provider + " / " + m.ModelSettings.Model);
            Row(sb, Loc.T("I18n_ReviewReportCoverage"), Loc.F("I18n_ReviewCoverage", m.CompletedItems, m.FailedItems,
                m.PendingItems, m.ReviewedCharacters, m.TotalCharacters, m.FailedCharacters, m.UnreviewedCharacters));
            if (!string.IsNullOrEmpty(m.Error)) Row(sb, Loc.T("I18n_ReviewSessionError"), m.Error);
            sb.AppendLine();
            sb.AppendLine(ScopeDescription(m));
            sb.AppendLine();
            foreach (var run in m.RunLog)
            {
                sb.AppendLine("- " + run.AtUtc.ToLocalTime().ToString("O") + " · " + Loc.T("I18n_ReviewSessionState" + run.State));
                if (!string.IsNullOrEmpty(run.Error)) Code(sb, run.Error);
            }
            sb.AppendLine();
            sb.AppendLine("## " + Loc.T("I18n_ReviewReportSettings"));
            Code(sb, JsonConvert.SerializeObject(m.ModelSettings, Formatting.Indented));
            sb.AppendLine("## " + Loc.T("I18n_ReviewReportInstruction"));
            Code(sb, m.Instruction);
            if (!string.IsNullOrEmpty(m.WritingProfile)) { sb.AppendLine("## " + Loc.T("I18n_ReviewReportProfile")); Code(sb, m.WritingProfile); }
            sb.AppendLine("## " + Loc.T("I18n_ReviewReportItems"));
            foreach (var item in record.Items)
            {
                sb.AppendLine("### #" + item.Index + " · P" + item.FirstParagraph + "–" + item.LastParagraph);
                sb.AppendLine();
                sb.AppendLine("| " + Loc.T("I18n_ReviewReportField") + " | " + Loc.T("I18n_ReviewReportValue") + " |");
                sb.AppendLine("| --- | --- |");
                Row(sb, Loc.T("I18n_ReviewSessionState"), item.State.ToString());
                Row(sb, Loc.T("I18n_ReviewReportAttempts"), item.Attempts.ToString(CultureInfo.InvariantCulture));
                Row(sb, Loc.T("I18n_ReviewReportCharacters"), item.Characters.ToString(CultureInfo.InvariantCulture));
                Row(sb, Loc.T("I18n_ReviewReportAction"), item.UserAction);
                if (item.Story != null) Row(sb, Loc.T("I18n_ReviewReportArea"), item.Story.Label);
                if (!string.IsNullOrEmpty(item.ReviewerIdentity)) Row(sb, Loc.T("I18n_ReviewReviewerIdentity"), item.ReviewerIdentity);
                if (item.ActionUtc.HasValue) Row(sb, Loc.T("I18n_ReviewReportActionTime"), item.ActionUtc.Value.ToLocalTime().ToString("O"));
                if (!string.IsNullOrEmpty(item.Error)) Row(sb, Loc.T("I18n_ReviewSessionError"), item.Error);
                sb.AppendLine();
                foreach (var attempt in item.AttemptLog)
                {
                    sb.AppendLine("- " + attempt.StartedUtc.ToLocalTime().ToString("O") + " · " + attempt.State
                        + (attempt.FragmentStart.HasValue ? " · " + Loc.T("I18n_ReviewReportFragment") + " " + attempt.FragmentStart : ""));
                    if (!string.IsNullOrEmpty(attempt.Error)) Code(sb, attempt.Error);
                }
                sb.AppendLine();
                sb.AppendLine("**" + Loc.T("I18n_ReviewReportOriginal") + "**");
                Code(sb, item.OriginalText ?? item.Context);
                if (item.GeneratedText != null && item.GeneratedText != item.SuggestedText)
                {
                    sb.AppendLine("**" + Loc.T("I18n_ReviewReportGeneratedSuggestion") + "**");
                    Code(sb, item.GeneratedText);
                }
                if (item.SuggestedText != null) { sb.AppendLine("**" + Loc.T("I18n_ReviewReportSuggestion") + "**"); Code(sb, item.SuggestedText); }
                foreach (var fragment in item.Fragments)
                {
                    sb.AppendLine("- " + Loc.T("I18n_ReviewReportFragment") + " " + fragment.Start + " / " + (fragment.Source?.Length ?? 0)
                        + " · " + fragment.State + (string.IsNullOrEmpty(fragment.Error) ? "" : " · " + Cell(fragment.Error)));
                    if (fragment.Suggestion != null && item.State != ReviewItemState.Succeeded) Code(sb, fragment.Suggestion);
                }
                foreach (var issue in item.Findings)
                {
                    sb.AppendLine("#### P" + issue.ParagraphIndex + " · " + issue.TitleText);
                    if (issue.Location?.Story != null) sb.AppendLine(issue.LocationLabel);
                    Code(sb, issue.DescriptionText);
                    if (!string.IsNullOrWhiteSpace(issue.Suggestion)) Code(sb, issue.Suggestion);
                    if (!string.IsNullOrEmpty(issue.UserAction)) sb.AppendLine(Loc.T("I18n_ReviewReportAction") + ": " + issue.UserAction);
                    if (issue.ActionUtc.HasValue) sb.AppendLine(Loc.T("I18n_ReviewReportActionTime") + ": " + issue.ActionUtc.Value.ToLocalTime().ToString("O"));
                    sb.AppendLine();
                }
            }
            return sb.ToString();
        }

        private static string Cell(string value) => (value ?? "").Replace("|", "\\|").Replace("\r", " ").Replace("\n", " ");
        private static void Row(StringBuilder sb, string key, string value) => sb.AppendLine("| " + Cell(key) + " | " + Cell(value) + " |");
        private static void Code(StringBuilder sb, string text)
        {
            text = text ?? "";
            int longest = Regex.Matches(text, "`+").Cast<Match>().Select(m => m.Length).DefaultIfEmpty(2).Max();
            string fence = new string('`', Math.Max(3, longest + 1));
            sb.AppendLine(); sb.AppendLine(fence); sb.AppendLine(text); sb.AppendLine(fence); sb.AppendLine();
        }
    }
}
