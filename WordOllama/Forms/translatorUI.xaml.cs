using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using WordOllama.lib;

namespace WordOllama.Forms
{
    /// <summary>Free translation workspace with reusable terminology and style schemes.</summary>
    public partial class translatorUI : UserControl
    {
        private static readonly string SchemeStorePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "WordOllama", "translation-schemes.json");

        private readonly List<TranslationScheme> _schemes = new List<TranslationScheme>();
        private TranslationSchemeSettings _schemeSettings;
        private bool _initializing = true;

        public translatorUI()
        {
            InitializeComponent();
            // The host can raise theme notifications while Word is still creating panes.
            // The bundled theme remains a safe fallback if a listener is not ready yet.
            try { ThemeManager.ApplyTheme(Resources); } catch { }
            InitializeLanguages();
            LoadSchemes();
            _initializing = false;
        }

        private void InitializeLanguages()
        {
            var sourceLanguages = new[]
            {
                Loc.T("I18n_TranslationAutoDetect"),
                Properties.Resources.汉语,
                Properties.Resources.英语,
                Properties.Resources.日语,
                Properties.Resources.法语,
                Properties.Resources.德语
            };
            var targetLanguages = sourceLanguages.Skip(1).ToArray();
            SourceLangComboBox.ItemsSource = sourceLanguages;
            LangComboBox.ItemsSource = targetLanguages;
            SourceLangComboBox.SelectedIndex = 0;

            bool chineseUi = string.Equals(SharedSettings.LanguageMode, "Chinese", StringComparison.OrdinalIgnoreCase);
            LangComboBox.SelectedItem = chineseUi ? Properties.Resources.英语 : Properties.Resources.汉语;
            if (LangComboBox.SelectedIndex < 0) LangComboBox.SelectedIndex = 0;
        }

        private void LoadSchemes()
        {
            try
            {
                _schemeSettings = File.Exists(SchemeStorePath)
                    ? JsonConvert.DeserializeObject<TranslationSchemeSettings>(File.ReadAllText(SchemeStorePath))
                    : null;
            }
            catch
            {
                _schemeSettings = null;
            }

            _schemeSettings = _schemeSettings ?? new TranslationSchemeSettings();
            _schemeSettings.CustomSchemes = _schemeSettings.CustomSchemes ?? new List<TranslationScheme>();

            _schemes.Clear();
            _schemes.Add(new TranslationScheme("builtin-default", Loc.T("I18n_TranslationBuiltInScheme"),
                Loc.T("I18n_TranslationBuiltInInstruction"), true));
            _schemes.Add(new TranslationScheme("builtin-legal", Loc.T("I18n_TranslationLegalScheme"),
                Loc.T("I18n_TranslationLegalInstruction"), true));
            _schemes.Add(new TranslationScheme("builtin-academic", Loc.T("I18n_TranslationAcademicScheme"),
                Loc.T("I18n_TranslationAcademicInstruction"), true));
            _schemes.Add(new TranslationScheme("builtin-business", Loc.T("I18n_TranslationBusinessScheme"),
                Loc.T("I18n_TranslationBusinessInstruction"), true));
            _schemes.Add(new TranslationScheme("builtin-technical", Loc.T("I18n_TranslationTechnicalScheme"),
                Loc.T("I18n_TranslationTechnicalInstruction"), true));
            _schemes.AddRange(_schemeSettings.CustomSchemes.Where(item => item != null && !string.IsNullOrWhiteSpace(item.Name)));

            SchemeComboBox.ItemsSource = null;
            SchemeComboBox.ItemsSource = _schemes;
            string selectedId = _schemeSettings.UseDefaultScheme ? _schemeSettings.DefaultSchemeId : "builtin-default";
            SchemeComboBox.SelectedValue = selectedId;
            if (SchemeComboBox.SelectedIndex < 0) SchemeComboBox.SelectedIndex = 0;
            DefaultSchemeToggle.IsChecked = _schemeSettings.UseDefaultScheme;
            LoadSelectedScheme();
        }

        private void LoadSelectedScheme()
        {
            var scheme = SchemeComboBox.SelectedItem as TranslationScheme;
            if (scheme == null) return;
            SchemeInstruction.Text = scheme.Instruction ?? string.Empty;
            SchemeName.Text = scheme.IsBuiltIn ? string.Empty : scheme.Name;
        }

        private void SaveSchemeSettings()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(SchemeStorePath));
                File.WriteAllText(SchemeStorePath, JsonConvert.SerializeObject(_schemeSettings, Formatting.Indented));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Loc.T("I18n_TranslationScheme"), MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private async void TranslateBegin(object sender, RoutedEventArgs e)
        {
            string targetLanguage = LangComboBox.Text?.Trim();
            string sourceLanguage = SourceLangComboBox.Text?.Trim();
            string inputText = InputText.Text;
            if (string.IsNullOrWhiteSpace(targetLanguage) || string.IsNullOrWhiteSpace(inputText))
            {
                MessageBox.Show(Properties.Resources.请确认已经输入文本以及已选择翻译语言);
                return;
            }

            TranslateButton.IsEnabled = false;
            StatusText.Text = Loc.T("I18n_Translating");
            try
            {
                string prompt = PromptTemplateService.BuildTranslationWindowPrompt(
                    targetLanguage,
                    inputText,
                    ComparisonModeToggle.IsChecked == true,
                    sourceLanguage == Loc.T("I18n_TranslationAutoDetect") ? null : sourceLanguage,
                    SchemeInstruction.Text?.Trim());
                string result = await Globals.ThisAddIn.GenerateModelSeletor(prompt);
                if (string.IsNullOrWhiteSpace(result))
                {
                    MessageBox.Show(Properties.Resources.未能返回结果请检查);
                    StatusText.Text = string.Empty;
                    return;
                }

                OutputText.Text = result.Trim();
                StatusText.Text = Loc.T("I18n_TranslationReady");
            }
            catch (Exception ex)
            {
                StatusText.Text = string.Empty;
                MessageBox.Show(ex.Message, Properties.Resources.错误, MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                TranslateButton.IsEnabled = true;
            }
        }

        private void SwapLanguages_Click(object sender, RoutedEventArgs e)
        {
            string source = SourceLangComboBox.Text;
            string target = LangComboBox.Text;
            if (source == Loc.T("I18n_TranslationAutoDetect"))
                source = target;

            SourceLangComboBox.Text = target;
            string nextTarget = source;
            if (string.Equals(nextTarget, target, StringComparison.OrdinalIgnoreCase))
                nextTarget = string.Equals(target, Properties.Resources.汉语, StringComparison.OrdinalIgnoreCase)
                    ? Properties.Resources.英语
                    : Properties.Resources.汉语;
            LangComboBox.Text = nextTarget;

            if (!string.IsNullOrWhiteSpace(OutputText.Text))
            {
                string oldInput = InputText.Text;
                InputText.Text = OutputText.Text;
                OutputText.Text = oldInput;
            }
        }

        private void LoadSelection_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string selected = UiThread.RunOfficeSync(() => Globals.ThisAddIn.Application.Selection?.Text ?? string.Empty);
                selected = (selected ?? string.Empty).TrimEnd('\r', '\a');
                if (string.IsNullOrWhiteSpace(selected))
                {
                    MessageBox.Show(Properties.Resources.请先选择文本);
                    return;
                }
                InputText.Text = selected;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Properties.Resources.错误, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ClearInput_Click(object sender, RoutedEventArgs e)
        {
            InputText.Clear();
        }

        private void CopyOutput_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(OutputText.Text)) Clipboard.SetText(OutputText.Text);
        }

        private void InputText_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (InputCount != null) InputCount.Text = (InputText.Text?.Length ?? 0).ToString();
        }

        private void OutputText_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (OutputCount != null) OutputCount.Text = (OutputText.Text?.Length ?? 0).ToString();
        }

        private void SchemeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_initializing) return;
            LoadSelectedScheme();
            if (DefaultSchemeToggle.IsChecked == true)
            {
                _schemeSettings.DefaultSchemeId = (SchemeComboBox.SelectedItem as TranslationScheme)?.Id;
                SaveSchemeSettings();
            }
        }

        private async void OptimizeScheme_Click(object sender, RoutedEventArgs e)
        {
            await Utils.OptimizeTextBoxAsync(SchemeInstruction, sender as Button);
        }

        private void SaveScheme_Click(object sender, RoutedEventArgs e)
        {
            string name = SchemeName.Text?.Trim();
            if (string.IsNullOrWhiteSpace(name))
            {
                MessageBox.Show(Loc.T("I18n_TranslationSchemeNameRequired"));
                return;
            }

            var current = SchemeComboBox.SelectedItem as TranslationScheme;
            TranslationScheme scheme = current != null && !current.IsBuiltIn
                ? _schemeSettings.CustomSchemes.FirstOrDefault(item => item.Id == current.Id)
                : null;
            if (scheme == null)
            {
                scheme = new TranslationScheme(Guid.NewGuid().ToString("N"), name, SchemeInstruction.Text?.Trim(), false);
                _schemeSettings.CustomSchemes.Add(scheme);
            }
            else
            {
                scheme.Name = name;
                scheme.Instruction = SchemeInstruction.Text?.Trim();
            }

            if (DefaultSchemeToggle.IsChecked == true)
                _schemeSettings.DefaultSchemeId = scheme.Id;
            SaveSchemeSettings();
            _initializing = true;
            LoadSchemes();
            SchemeComboBox.SelectedValue = scheme.Id;
            LoadSelectedScheme();
            _initializing = false;
            StatusText.Text = Loc.T("I18n_TranslationSchemeSaved");
        }

        private void ManageScheme_Click(object sender, RoutedEventArgs e)
        {
            var selected = SchemeComboBox.SelectedItem as TranslationScheme;
            if (selected == null || selected.IsBuiltIn)
            {
                MessageBox.Show(Loc.T("I18n_TranslationSelectCustomScheme"));
                return;
            }

            if (MessageBox.Show(Loc.T("I18n_TranslationDeleteSchemeConfirm"), Loc.T("I18n_TranslationScheme"),
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            _schemeSettings.CustomSchemes.RemoveAll(item => item.Id == selected.Id);
            if (_schemeSettings.DefaultSchemeId == selected.Id)
            {
                _schemeSettings.DefaultSchemeId = "builtin-default";
                _schemeSettings.UseDefaultScheme = false;
            }
            SaveSchemeSettings();
            _initializing = true;
            LoadSchemes();
            _initializing = false;
            StatusText.Text = Loc.T("I18n_TranslationSchemeDeleted");
        }

        private void DefaultSchemeToggle_Changed(object sender, RoutedEventArgs e)
        {
            if (_initializing || _schemeSettings == null) return;
            _schemeSettings.UseDefaultScheme = DefaultSchemeToggle.IsChecked == true;
            _schemeSettings.DefaultSchemeId = _schemeSettings.UseDefaultScheme
                ? (SchemeComboBox.SelectedItem as TranslationScheme)?.Id
                : null;
            SaveSchemeSettings();
        }
    }

    public sealed class TranslationScheme
    {
        public TranslationScheme() { }

        public TranslationScheme(string id, string name, string instruction, bool isBuiltIn)
        {
            Id = id;
            Name = name;
            Instruction = instruction;
            IsBuiltIn = isBuiltIn;
        }

        public string Id { get; set; }
        public string Name { get; set; }
        public string Instruction { get; set; }
        public bool IsBuiltIn { get; set; }
    }

    public sealed class TranslationSchemeSettings
    {
        public List<TranslationScheme> CustomSchemes { get; set; } = new List<TranslationScheme>();
        public bool UseDefaultScheme { get; set; }
        public string DefaultSchemeId { get; set; }
    }
}
