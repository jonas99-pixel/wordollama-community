using System.Windows;

namespace WordOllama.NewUI
{
    public partial class UpdateLogWindow : Window
    {
        public UpdateLogWindow()
        {
            InitializeComponent();
        }

        public void SetContent(string content)
        {
            txtUpdateLog.Text = content;
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
