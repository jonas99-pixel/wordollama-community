using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Newtonsoft.Json;
using WordOllama.lib;
using System.Windows.Documents;
using System.Windows.Media;
using static WordOllama.ThisAddIn;
using WordOllama.Forms;

namespace WordOllama.NewUI
{
    public partial class ModWritting : Window, INotifyPropertyChanged
    {
        private string dataFilePath;
        private ObservableCollection<ModPromptTemplate> _templates;
        public ObservableCollection<ModPromptTemplate> Templates
        {
            get => _templates;
            set
            {
                _templates = value;
                OnPropertyChanged();
            }
        }

        public ModWritting()
        {
            InitializeComponent();
            dataFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WordOllama", "customModPrompt.json");
            LoadTemplates();
            InitializeTemplateComboBox();
            DataContext = this;
            ThemeManager.ApplyTheme(this.Resources);
            InputTextBox.Text = string.IsNullOrEmpty(GetText.SelectedText) ? string.Empty : GetText.SelectedText;
        }

        private void LoadTemplates()
        {
            var directory = Path.GetDirectoryName(dataFilePath);
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            if (!File.Exists(dataFilePath))
            {
                Templates = new ObservableCollection<ModPromptTemplate>();
                SaveTemplates();
            }
            else
            {
                var json = File.ReadAllText(dataFilePath);
                var loadedTemplates = JsonConvert.DeserializeObject<ObservableCollection<ModPromptTemplate>>(json) ?? new ObservableCollection<ModPromptTemplate>();
                Templates = new ObservableCollection<ModPromptTemplate>(loadedTemplates);
            }

            TemplateDataGrid.ItemsSource = Templates;
        }

        private void SaveTemplates()
        {
            var json = JsonConvert.SerializeObject(Templates, Formatting.Indented);
            File.WriteAllText(dataFilePath, json);
            InitializeTemplateComboBox();
        }

        private void InitializeTemplateComboBox()
        {
            TemplateComboBox.ItemsSource = Templates.Select(t => t.Name);
        }

        private void TemplateComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (TemplateComboBox.SelectedIndex != -1)
            {
                var selectedTemplate = Templates[TemplateComboBox.SelectedIndex];
                PromptBox.Text = selectedTemplate.Prompt;
            }
        }

        private async Task ModifyNow()
        {
            try
            {
                string prompt = PromptTemplateService.BuildModificationPrompt(InputTextBox.Text, PromptBox.Text);

                OutputTextBox.Clear();
                DiffBox.Document.Blocks.Clear();

                string response = await Globals.ThisAddIn.GenerateModelSeletor(prompt);

                OutputTextBox.Text = response;
                OutputTextBox.ScrollToEnd();

                // Display diff result
                var differences = Globals.ThisAddIn.GetDifferences(InputTextBox.Text, OutputTextBox.Text);

                Paragraph paragraph = new Paragraph();

                foreach (var diff in differences)
                {
                    Run run = new Run(diff.Text);
                    switch (diff.Type)
                    {
                        case DiffType.Deleted:
                            run.Foreground = Brushes.Red;
                            run.TextDecorations = TextDecorations.Strikethrough;
                            break;
                        case DiffType.Inserted:
                            run.Foreground = Brushes.Green;
                            run.FontWeight = FontWeights.Bold;
                            break;
                        case DiffType.Unchanged:
                            break;
                    }
                    paragraph.Inlines.Add(run); }

                DiffBox.Document.Blocks.Add(paragraph);

            }
            catch (Exception ex)
            {
                MessageBox.Show(Loc.F("I18n_ModifyError", ex.Message));
            }
        }




        private async void ModifyButton_Click(object sender, RoutedEventArgs e)
        {
            thisProgressBar.Show();
            await ModifyNow();
            thisProgressBar.Hide();
        }

        private void AddTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            if (CanAddNewRow())
            {
                Templates.Add(new ModPromptTemplate { Name = "", Prompt = "" });
            }
            else
            {
                MessageBox.Show(Loc.T("I18n_RowIncomplete"), Loc.T("I18n_Info"), MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private bool CanAddNewRow()
        {
            return !Templates.Any(p => string.IsNullOrWhiteSpace(p.Name) || string.IsNullOrWhiteSpace(p.Prompt));
        }


        private void SaveTemplatesButton_Click(object sender, RoutedEventArgs e)
        {
            SaveTemplates();
            MessageBox.Show(Loc.T("I18n_TemplateSaved"), Loc.T("I18n_Info"), MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private async void OptimizationButton_Click(object sender, RoutedEventArgs e)
        {
            await Utils.OptimizeTextBoxAsync(PromptBox, sender as Button);
        }

        private async void OptimizeTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            if (!(button?.DataContext is ModPromptTemplate template)) return;
            button.IsEnabled = false;
            try
            {
                string optimized = await Utils.PromptOptimizer(template.Prompt);
                if (!string.IsNullOrWhiteSpace(optimized)) template.Prompt = optimized;
            }
            finally { button.IsEnabled = true; }
        }

        private void DeleteTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var template = button.DataContext as ModPromptTemplate;

            if (template != null)
            {
                var result = MessageBox.Show(Loc.T("I18n_DeleteTemplateConfirm"), WordOllama.Properties.Resources.确认删除, MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
                {
                    Templates.Remove(template);
                    SaveTemplates();
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void Copy_Btn_Click(object sender, RoutedEventArgs e)
        {
            OutputTextBox.SelectAll();
            OutputTextBox.Copy();
        }
    }

    public class ModPromptTemplate : INotifyPropertyChanged
    {
        private string _name;
        public string Name
        {
            get => _name;
            set
            {
                _name = value;
                OnPropertyChanged();
            }
        }

        private string _prompt;
        public string Prompt
        {
            get => _prompt;
            set
            {
                _prompt = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
