using System.Collections.Generic;
using System.Windows;
using WordOllama.lib;

namespace WordOllama.NewUI
{
    public partial class ReviewSessionWindow : Window
    {
        public ReviewSessionMetadata SelectedSession { get; private set; }
        public bool ExportRequested { get; private set; }
        public ReviewSessionWindow(IEnumerable<ReviewSessionMetadata> sessions)
        {
            InitializeComponent();
            ThemeManager.ApplyTheme(Resources);
            SessionGrid.ItemsSource = sessions;
        }
        private void Restore_Click(object sender, RoutedEventArgs e) { Select(false); }
        private void Export_Click(object sender, RoutedEventArgs e) { Select(true); }
        private void Select(bool export)
        {
            if (!(SessionGrid.SelectedItem is ReviewSessionMetadata metadata)) return;
            SelectedSession = metadata;
            ExportRequested = export;
            DialogResult = true;
        }
    }
}
