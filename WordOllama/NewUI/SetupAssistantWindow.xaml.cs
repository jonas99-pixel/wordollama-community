using System;
using System.Net.Http;
using System.Windows;
using System.Windows.Controls;
using WordOllama.lib;
using WordOllama.lib.Providers;

namespace WordOllama.NewUI
{
    public partial class SetupAssistantWindow : Window
    {
        public SetupAssistantWindow()
        {
            InitializeComponent();
            ThemeManager.ApplyTheme(Resources);
        }

        private void Back_Click(object sender, RoutedEventArgs e)
        {
            if (Steps.SelectedIndex > 0) Steps.SelectedIndex--;
            UpdateButtons();
        }

        private void Next_Click(object sender, RoutedEventArgs e)
        {
            if (Steps.SelectedIndex == 1) SaveProvider();
            if (Steps.SelectedIndex < 2) Steps.SelectedIndex++;
            else Finish();
            UpdateButtons();
        }

        private void Skip_Click(object sender, RoutedEventArgs e) => Finish();

        private async void Test_Click(object sender, RoutedEventArgs e)
        {
            TestStatus.Text = Loc.T("I18n_TestingConnection");
            try
            {
                if (SetupOllama.IsChecked == true)
                {
                    using (var client = new HttpClient { Timeout = TimeSpan.FromSeconds(10) })
                    {
                        var response = await client.GetAsync(SetupEndpoint.Text.TrimEnd('/') + "/api/tags");
                        response.EnsureSuccessStatusCode();
                        TestStatus.Text = "Ollama OK";
                    }
                }
                else
                {
                    string type = (SetupProviderType.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "OpenAI";
                    var provider = ProviderFactory.Create(type, SetupEndpoint.Text, SetupApiKey.Password, SetupModel.Text);
                    CapabilityProbeResult result = await ProviderCapabilityProbe.RunAsync(provider);
                    TestStatus.Text = Loc.F("I18n_CapabilityProbeResult", result.Passed, result.Total, result.LatencyMs, result.Summary);
                }
            }
            catch (Exception ex) { TestStatus.Text = ex.Message; }
        }

        private void SaveProvider()
        {
            if (SetupOllama.IsChecked == true)
            {
                AppConfig.Instance.AIMode = "Ollama";
                AppConfig.Instance.OllamaEndPoint = SetupEndpoint.Text;
                AppConfig.Instance.OllamaModel = SetupModel.Text;
            }
            else
            {
                AppConfig.Instance.AIMode = "ChatGPT";
                string type = (SetupProviderType.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "OpenAI";
                AppConfig.Instance.OnlineProviderType = type;
                if (type == "Claude") { AppConfig.Instance.ClaudeEndPoint = SetupEndpoint.Text; AppConfig.Instance.ClaudeKey = SetupApiKey.Password; AppConfig.Instance.ClaudeModel = SetupModel.Text; }
                else if (type == "Gemini") { AppConfig.Instance.GeminiEndPoint = SetupEndpoint.Text; AppConfig.Instance.GeminiKey = SetupApiKey.Password; AppConfig.Instance.GeminiModel = SetupModel.Text; }
                else { AppConfig.Instance.ChatGPTEndPoint = SetupEndpoint.Text; AppConfig.Instance.ChatGPTKey = SetupApiKey.Password; AppConfig.Instance.ChatGPTModel = SetupModel.Text; }
            }
            AppConfig.Save();
        }

        private void Finish()
        {
            AppConfig.Instance.SetupAssistantCompleted = true;
            AppConfig.SaveNow();
            DialogResult = true;
            Close();
        }

        private void UpdateButtons()
        {
            BackButton.IsEnabled = Steps.SelectedIndex > 0;
            NextButton.Content = Steps.SelectedIndex == 2 ? Loc.T("I18n_Finish") : Loc.T("I18n_Next");
        }
    }
}
