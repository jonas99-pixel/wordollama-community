using System;
using System.ComponentModel;
using System.Security.Cryptography;
using System.Text;

namespace WordOllama.NewUI
{
    public class EditSuggestionSession : INotifyPropertyChanged
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        public string BookmarkName { get; set; }
        public int Index { get; set; }
        public int ParagraphIndex { get; set; }
        public int RangeStart { get; set; }
        public int RangeEnd { get; set; }
        public string OriginalHash { get; set; }
        public string DocumentIdentity { get; set; }
        public WordOllama.lib.ReviewStoryAddress Story { get; set; }
        public bool GenerationSucceeded { get; set; }

        private string _failureReason;
        public string FailureReason
        {
            get => _failureReason;
            set { _failureReason = value; OnPropertyChanged(nameof(FailureReason)); }
        }

        private string _originalText;
        public string OriginalText
        {
            get => _originalText;
            set { _originalText = value; OnPropertyChanged(nameof(OriginalText)); }
        }

        private string _suggestedText;
        public string SuggestedText
        {
            get => _suggestedText;
            set
            {
                _suggestedText = value;
                OnPropertyChanged(nameof(SuggestedText));
                OnPropertyChanged(nameof(HasSuggestion));
            }
        }

        private string _status;
        public string Status
        {
            get => _status;
            set { _status = value; OnPropertyChanged(nameof(Status)); }
        }

        private bool _isBusy;
        public bool IsBusy
        {
            get => _isBusy;
            set { _isBusy = value; OnPropertyChanged(nameof(IsBusy)); }
        }

        public bool HasSuggestion => !string.IsNullOrWhiteSpace(SuggestedText);
        public string LocationLabel => (Story == null ? "" : Story.Label + " / ") + (ParagraphIndex > 0 ? $"P{ParagraphIndex}" : $"#{Index}");

        public static string NormalizeWordText(string text)
        {
            return (text ?? "").TrimEnd('\r', '\a');
        }

        public static string Hash(string text)
        {
            using (var sha = SHA256.Create())
            {
                var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(NormalizeWordText(text)));
                return Convert.ToBase64String(bytes);
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
