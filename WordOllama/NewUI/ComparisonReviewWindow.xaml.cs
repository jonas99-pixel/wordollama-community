using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using WordOllama.lib;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.NewUI
{
    public sealed class ComparisonRevisionItem
    {
        public int OriginalIndex { get; set; }
        public int Start { get; set; }
        public Word.WdRevisionType RevisionType { get; set; }
        public string Type { get; set; }
        public string Text { get; set; }
        public string Author { get; set; }
        public DateTime Date { get; set; }
    }

    public partial class ComparisonReviewWindow : Window
    {
        private readonly Word.Document _document;
        private readonly ObservableCollection<ComparisonRevisionItem> _items = new ObservableCollection<ComparisonRevisionItem>();
        public ComparisonReviewWindow(Word.Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            InitializeComponent();
            ThemeManager.ApplyTheme(Resources);
            RevisionGrid.ItemsSource = _items;
            RefreshItems();
        }

        private void RefreshItems()
        {
            _items.Clear();
            for (int index = 1; index <= _document.Revisions.Count; index++)
            {
                Word.Revision revision = _document.Revisions[index];
                string text = (revision.Range.Text ?? "").Trim();
                if (text.Length > 500) text = text.Substring(0, 500) + "…";
                _items.Add(new ComparisonRevisionItem { OriginalIndex = index, Start = revision.Range.Start, RevisionType = revision.Type, Type = revision.Type.ToString(), Text = text, Author = revision.Author, Date = revision.Date });
            }
            Summary.Text = Loc.F("I18n_RevisionCount", _items.Count);
        }

        private Word.Revision Resolve(ComparisonRevisionItem item)
        {
            if (item == null) return null;
            for (int i = 1; i <= _document.Revisions.Count; i++)
            {
                Word.Revision revision = _document.Revisions[i];
                if (revision.Type == item.RevisionType && revision.Range.Start == item.Start && string.Equals(revision.Author, item.Author, StringComparison.Ordinal)) return revision;
            }
            return null;
        }

        private ComparisonRevisionItem Selected => RevisionGrid.SelectedItem as ComparisonRevisionItem;
        private void Locate_Click(object sender, RoutedEventArgs e) { Word.Revision revision = Resolve(Selected); if (revision == null) { RefreshItems(); return; } _document.Activate(); revision.Range.Select(); }
        private void Accept_Click(object sender, RoutedEventArgs e) { Word.Revision revision = Resolve(Selected); if (revision == null) { RefreshItems(); return; } revision.Accept(); RefreshItems(); }
        private void Reject_Click(object sender, RoutedEventArgs e) { Word.Revision revision = Resolve(Selected); if (revision == null) { RefreshItems(); return; } revision.Reject(); RefreshItems(); }
        private void AcceptAll_Click(object sender, RoutedEventArgs e) { _document.AcceptAllRevisions(); RefreshItems(); }
        private void RejectAll_Click(object sender, RoutedEventArgs e) { _document.RejectAllRevisions(); RefreshItems(); }
    }
}
