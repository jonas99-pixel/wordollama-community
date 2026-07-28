using System.Windows;
using WordOllama.lib;

namespace WordOllama.NewUI
{
    public partial class AgentPermissionDialog : Window
    {
        public AgentPermissionDecision Decision { get; private set; } = AgentPermissionDecision.Deny;

        public AgentPermissionDialog(string message, bool allowSession = true)
        {
            InitializeComponent();
            PermissionMessage.Text = message ?? "";
            AllowSessionButton.Visibility = allowSession ? Visibility.Visible : Visibility.Collapsed;
        }

        private void AllowOnce_Click(object sender, RoutedEventArgs e)
        {
            Decision = AgentPermissionDecision.AllowOnce;
            DialogResult = true;
        }

        private void AllowSession_Click(object sender, RoutedEventArgs e)
        {
            Decision = AgentPermissionDecision.AllowSession;
            DialogResult = true;
        }

        private void Deny_Click(object sender, RoutedEventArgs e)
        {
            Decision = AgentPermissionDecision.Deny;
            DialogResult = false;
        }
    }
}
