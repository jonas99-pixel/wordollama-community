using System;
using System.Windows;
using System.Windows.Controls;
using WordOllama.lib;

namespace WordOllama.Forms
{
    /// <summary>
    /// translatorUI.xaml 的交互逻辑
    /// </summary>
    public partial class translatorUI : UserControl
    {
        public translatorUI()
        {
            InitializeComponent();
            ThemeManager.ApplyTheme(this.Resources);
        }

        private async void TranslateBegin(object sender, RoutedEventArgs e)
        {
            string lang = LangComboBox.Text;
            string inputText = InputText.Text;
            if (!string.IsNullOrEmpty(lang) && !string.IsNullOrEmpty(inputText))
            {
                string systemPrompt = PromptTemplateService.BuildTranslationWindowPrompt(
                    lang,
                    inputText,
                    ComparisonModeToggle.IsChecked == true);
                string result = await Globals.ThisAddIn.GenerateModelSeletor(systemPrompt);

                if (!string.IsNullOrEmpty(result))
                {
                        OutputText.Text = result;
                }
                else
                {
                        MessageBox.Show(WordOllama.Properties.Resources.未能返回结果请检查);
                }
            }
            else
            {
                MessageBox.Show(WordOllama.Properties.Resources.请确认已经输入文本以及已选择翻译语言);
            }

        }
    }
}
