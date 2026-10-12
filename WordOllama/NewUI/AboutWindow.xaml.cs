using System.Reflection;
using System.Diagnostics;
using System.Windows;
using System.Windows.Navigation;

namespace WordOllama.NewUI
{
    public partial class AboutWindow : Window
    {
        public AboutWindow()
        {
            InitializeComponent();
            var ver = Assembly.GetExecutingAssembly().GetName().Version;
            VersionText.Text = ver != null ? $"v{ver.Major}.{ver.Minor}.{ver.Build}" : "v2.5";
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
            e.Handled = true;
        }
    }
}
