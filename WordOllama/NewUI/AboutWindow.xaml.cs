using System.Reflection;
using System.Windows;

namespace WordOllama.NewUI
{
    public partial class AboutWindow : Window
    {
        public AboutWindow()
        {
            InitializeComponent();
            var ver = Assembly.GetExecutingAssembly().GetName().Version;
            VersionText.Text = ver != null ? $"v{ver.Major}.{ver.Minor}.{ver.Build}.{ver.Revision}" : "v2.0.1.0";
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
