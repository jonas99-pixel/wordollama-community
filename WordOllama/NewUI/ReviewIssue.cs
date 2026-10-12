using System;
using System.ComponentModel;
using WordOllama.lib;

namespace WordOllama.NewUI
{
    /// <summary>
    /// 审查问题面板中的单条问题。
    /// </summary>
    public class ReviewIssue : INotifyPropertyChanged
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        public string Source { get; set; } = Loc.T("I18n_UiReviewSource");
        public string Category { get; set; } = Loc.T("I18n_UiCategoryText");
        public string Severity { get; set; } = Loc.T("I18n_UiSeverityMedium");
        public int ParagraphIndex { get; set; }
        public string DocumentIdentity { get; set; }
        public string SourceParagraphHash { get; set; }
        public ReviewParagraphLocation Location { get; set; }
        public string ReviewSessionId { get; set; }
        public string ReviewItemId { get; set; }
        public string UserAction { get; set; }
        public DateTime? ActionUtc { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
        public string Suggestion { get; set; }
        public string Excerpt { get; set; }
        public DateTime CreatedAt { get; set; } = DateTime.Now;

        public string LocationLabel => Location != null ? Location.Story.Label + " / " + Loc.F("I18n_UiParagraph", Location.Paragraph) : ParagraphIndex > 0
            ? Loc.F("I18n_UiParagraph", ParagraphIndex)
            : Loc.T("I18n_UiUnlocated");
        public string SeverityLabel => string.IsNullOrWhiteSpace(Severity) ? Loc.T("I18n_UiSeverityMedium") : Severity;
        public string TitleText => string.IsNullOrWhiteSpace(Title) ? Category : Title;
        public string DescriptionText => string.IsNullOrWhiteSpace(Description) ? Excerpt : Description;
        public string SuggestionText => string.IsNullOrWhiteSpace(Suggestion) ? Loc.T("I18n_UiNoSuggestion") : Suggestion;
        public string SourceLabel => string.IsNullOrWhiteSpace(Source) ? Loc.T("I18n_UiReviewSource") : Source;

        public event PropertyChangedEventHandler PropertyChanged;

        public static string DisplaySeverity(string severity)
        {
            string value = (severity ?? "").Trim().ToLowerInvariant();
            switch (value)
            {
                case "high":
                case "高":
                    return Loc.T("I18n_UiSeverityHigh");
                case "low":
                case "低":
                    return Loc.T("I18n_UiSeverityLow");
                case "medium":
                case "中":
                default:
                    return Loc.T("I18n_UiSeverityMedium");
            }
        }

        public static string DisplayCategory(string category)
        {
            string value = (category ?? "").Trim().ToLowerInvariant();
            switch (value)
            {
                case "typo":
                    return Loc.T("I18n_UiCategoryTypo");
                case "grammar":
                    return Loc.T("I18n_UiCategoryGrammar");
                case "format":
                    return Loc.T("I18n_UiCategoryFormat");
                case "reference":
                    return Loc.T("I18n_UiCategoryReference");
                case "terminology":
                    return Loc.T("I18n_UiCategoryTerminology");
                case "legal":
                    return Loc.T("I18n_UiCategoryLegal");
                case "clarity":
                    return Loc.T("I18n_UiCategoryClarity");
                default:
                    return string.IsNullOrWhiteSpace(category) ? Loc.T("I18n_UiCategoryText") : category;
            }
        }

        public void RefreshComputed()
        {
            OnPropertyChanged(nameof(LocationLabel));
            OnPropertyChanged(nameof(SeverityLabel));
            OnPropertyChanged(nameof(TitleText));
            OnPropertyChanged(nameof(DescriptionText));
            OnPropertyChanged(nameof(SuggestionText));
            OnPropertyChanged(nameof(SourceLabel));
        }

        protected void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
