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

namespace WordOllama.NewUI
{
    public partial class WritingUI : Window, INotifyPropertyChanged
    {
        private string dataFilePath;
        private ObservableCollection<PromptTemplate> _templates;
        public ObservableCollection<PromptTemplate> Templates
        {
            get => _templates;
            set
            {
                _templates = value;
                OnPropertyChanged();
            }
        }

        public WritingUI()
        {
            InitializeComponent();
            dataFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WordOllama", "customPrompt.json");
            LoadTemplates();
            InitializeTemplateComboBox();
            DataContext = this;
            ThemeManager.ApplyTheme(this.Resources);
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
                Templates = new ObservableCollection<PromptTemplate>();
                SaveTemplates();
            }
            else
            {
                var json = File.ReadAllText(dataFilePath);
                var loadedTemplates = JsonConvert.DeserializeObject<ObservableCollection<PromptTemplate>>(json) ?? new ObservableCollection<PromptTemplate>();
                Templates = new ObservableCollection<PromptTemplate>(loadedTemplates);
            }

            TemplateDataGrid.ItemsSource = Templates;
        }
        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            if (CanAddNewRow())
            {
                Templates.Add(new PromptTemplate { Name = "", Prompt = ""});
            }
            else
            {
                MessageBox.Show(WordOllama.Properties.Resources.请先填写完当前行再添加新行, WordOllama.Properties.Resources.提示, MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }


        private bool CanAddNewRow()
        {
            return !Templates.Any(p => string.IsNullOrWhiteSpace(p.Name) || string.IsNullOrWhiteSpace(p.Prompt));
        }
        private async void OptimizationButton_Click(object sender, RoutedEventArgs e)
        {
            RequirementTextBox.Text = await Utils.PromptOptimizer(RequirementTextBox.Text);
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
                RequirementTextBox.Text = selectedTemplate.Prompt;
            }
        }

        private async Task GenerateNow()
        {
            try
            {
                string prompt = PromptTemplateService.BuildFreeWritingPrompt(ThemeTextBox.Text, RequirementTextBox.Text);
                OutputTextBox.Clear();
                GenerateButton.IsEnabled = false;
                GenerateButton.Content = WordOllama.Properties.Resources.生成中;

                // Assuming Globals.ThisAddIn.GenerateModelSeletor is accessible
                string response = await Globals.ThisAddIn.GenerateModelSeletor(prompt);

                OutputTextBox.Text = response;
                OutputTextBox.ScrollToEnd();

                GenerateButton.IsEnabled = true;
                GenerateButton.Content = WordOllama.Properties.Resources.开始生成;
            }
            catch (Exception ex)
            {
                MessageBox.Show(WordOllama.Properties.Resources.遇到错误 +  ex.Message);
            }
        }






        private async void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            thisProgressBar.Show();
            await GenerateNow();
            thisProgressBar.Hide();
        }

        private void SaveTemplatesButton_Click(object sender, RoutedEventArgs e)
        {
            SaveTemplates();
            MessageBox.Show(WordOllama.Properties.Resources.模板已保存, WordOllama.Properties.Resources.提示, MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void DeleteTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var template = button.DataContext as PromptTemplate;

            if (template != null)
            {
                var result = MessageBox.Show(WordOllama.Properties.Resources.确定要删除这个模板吗, WordOllama.Properties.Resources.确认删除, MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
                {
                    Templates.Remove(template);
                    SaveTemplates();
                }
            }
        }

        private void AddTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            Templates.Add(new PromptTemplate { Name = WordOllama.Properties.Resources.新模板, Prompt = "" });
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

    public class PromptTemplate : INotifyPropertyChanged
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
