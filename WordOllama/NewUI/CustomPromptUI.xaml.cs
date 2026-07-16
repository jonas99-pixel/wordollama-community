using MaterialDesignThemes.Wpf;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using WordOllama.lib;

namespace WordOllama.NewUI
{
    public partial class CustomPromptUI : Window, INotifyPropertyChanged
    {
        private ObservableCollection<CustomPrompt> _customPrompts;
        public ObservableCollection<CustomPrompt> CustomPrompts
        {
            get => _customPrompts;
            set
            {
                _customPrompts = value;
                OnPropertyChanged();
            }
        }

        public List<string> FunctionOptions { get; } = new List<string>
        {
            Loc.T("I18n_CustomPromptInsert"),
            Loc.T("I18n_CustomPromptInsertWithDiff"),
            Loc.T("I18n_CustomPromptComment")
        };

        private string jsonFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WordOllama", "OneClickCustomPrompts.json");

        private List<DispatcherOperation> _pendingOperations = new List<DispatcherOperation>();
        public CustomPromptUI()
        {
            InitializeComponent();
            LoadCustomPrompts();
            InitializeControls();
            DataContext = this;
            //Closing += CustomPromptUI_Closing;
            ThemeManager.ApplyTheme(this.Resources);
        }

        private void LoadCustomPrompts()
        {
            if (File.Exists(jsonFilePath))
            {
                string json = File.ReadAllText(jsonFilePath);
                var loadedPrompts = JsonConvert.DeserializeObject<List<dynamic>>(json);

                CustomPrompts = new ObservableCollection<CustomPrompt>(loadedPrompts.Select(prompt => new CustomPrompt
                {
                    Name = prompt.Name,
                    Prompt = prompt.Prompt,
                    Function = GetOriginalFunction((string)prompt.Function) // 显式转换为 string
                }));
            }
            else
            {
                CustomPrompts = new ObservableCollection<CustomPrompt>();
            }

            

            // 确保 CustomPrompts 集合中至少有一个非空值
            if (CustomPrompts.Count == 0 || CustomPrompts.All(prompt => string.IsNullOrEmpty(prompt.Name) && string.IsNullOrEmpty(prompt.Prompt) && string.IsNullOrEmpty(prompt.Function)))
            {
                CustomPrompts.Add(new CustomPrompt { Name = "", Prompt = "", Function = Loc.T("I18n_CustomPromptInsert") });
            }
        }

        private string GetOriginalFunction(string transformedFunction)
        {
            switch (transformedFunction)
            {
                case "In":
                case "插入":
                    return Loc.T("I18n_CustomPromptInsert");
                case "InDiff":
                case "插入并对比":
                    return Loc.T("I18n_CustomPromptInsertWithDiff");
                case "CM":
                case "批注":
                    return Loc.T("I18n_CustomPromptComment");
                default:
                    return transformedFunction;
            }
        }

        private void InitializeControls()
        {
            string[] prefixes = { "C1", "C2", "C3", "C4" };
            foreach (string prefix in prefixes)
            {
                TextBox promptBox = (TextBox)FindName($"{prefix}_Prompt");
                TextBox nameBox = (TextBox)FindName($"{prefix}_Name");
                ComboBox functionComboBox = (ComboBox)FindName($"{prefix}_Function");

                if (promptBox != null)
                {
                    promptBox.Text = AppConfig.Instance.GetSetting($"{prefix}_Prompt").ToString();
                }
                if (nameBox != null)
                {
                    nameBox.Text = AppConfig.Instance.GetSetting($"{prefix}_label").ToString();
                }
                if (functionComboBox != null)
                {
                    string value = GetOriginalFunction(AppConfig.Instance.GetSetting($"{prefix}_Funt").ToString());
                    foreach (ComboBoxItem item in functionComboBox.Items)
                    {
                        if (item.Content.ToString() == value)
                        {
                            functionComboBox.SelectedItem = item;
                            break;
                        }
                    }
                }
            }

            CustomDataGrid.ItemsSource = CustomPrompts;
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            SaveControls();
            SaveCustomPrompts();
            Globals.Ribbons.Ribbon1.ChangeRibbonLabels();
            Globals.Ribbons.Ribbon1.LoadCustomPrompts();
            Globals.Ribbons.Ribbon1.ShowNotification(WordOllama.Properties.Resources.保存成功, 1000);
        }

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var customPrompt = button.DataContext as CustomPrompt;

            if (customPrompt != null)
            {
                var result = MessageBox.Show(WordOllama.Properties.Resources.确定要删除这一行吗, WordOllama.Properties.Resources.确认删除, MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
                {
                    CustomPrompts.Remove(customPrompt);
                }
            }
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            if (CanAddNewRow())
            {
                CustomPrompts.Add(new CustomPrompt { Name = "", Prompt = "", Function = Loc.T("I18n_CustomPromptInsert") });
            }
            else
            {
                MessageBox.Show(Loc.T("I18n_RowIncomplete"), WordOllama.Properties.Resources.提示, MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private bool CanAddNewRow()
        {
            return !CustomPrompts.Any(p => string.IsNullOrWhiteSpace(p.Name) || string.IsNullOrWhiteSpace(p.Prompt));
        }

        private void SaveControls()
        {
            string[] prefixes = { "C1", "C2", "C3", "C4" };
            foreach (string prefix in prefixes)
            {
                TextBox promptBox = (TextBox)FindName($"{prefix}_Prompt");
                TextBox nameBox = (TextBox)FindName($"{prefix}_Name");
                ComboBox functionComboBox = (ComboBox)FindName($"{prefix}_Function");

                if (promptBox != null)
                    AppConfig.Instance.SetSetting($"{prefix}_Prompt", promptBox.Text);
                if (nameBox != null)
                    AppConfig.Instance.SetSetting($"{prefix}_label", nameBox.Text);
                if (functionComboBox != null && functionComboBox.SelectedItem != null)
                {
                    // 获取 ComboBoxItem 的内容
                    string selectedFunction = ((ComboBoxItem)functionComboBox.SelectedItem).Content.ToString();
                    AppConfig.Instance.SetSetting($"{prefix}_Funt", GetTransformedFunction(selectedFunction));
                }
            }
            AppConfig.Save();
        }

        //private void CustomPromptUI_Closing(object sender, CancelEventArgs e)
        //{
        //    // Cancel all pending dispatcher operations
        //    foreach (var operation in _pendingOperations)
        //    {
        //        operation.Abort();
        //    }
        //    _pendingOperations.Clear();
        //}

        private string GetTransformedFunction(string Function)
        {
            switch (Function)
            {
                case var _ when Function == Loc.T("I18n_CustomPromptInsert") || Function == "插入":
                    return "In";
                case var _ when Function == Loc.T("I18n_CustomPromptInsertWithDiff") || Function == "插入并对比":
                    return "InDiff";
                case var _ when Function == Loc.T("I18n_CustomPromptComment") || Function == "批注":
                    return "CM";
                default:
                    return Function;
            }
        }

        private void SaveCustomPrompts()
        {
            var transformedPrompts = CustomPrompts
                .Where(prompt => !string.IsNullOrEmpty(prompt.Name) || !string.IsNullOrEmpty(prompt.Prompt) || !string.IsNullOrEmpty(prompt.Function))
                .Select(prompt => new
                {
                    Name = prompt.Name,
                    Prompt = prompt.Prompt,
                    Function = prompt.GetTransformedFunction()
                }).ToList();

            string json = JsonConvert.SerializeObject(transformedPrompts, Formatting.Indented);
            Directory.CreateDirectory(Path.GetDirectoryName(jsonFilePath));
            File.WriteAllText(jsonFilePath, json);

        }

        private void GetPromptTextLength_And_SetHelperText(object sender, TextChangedEventArgs e)
        {
            try
            {
                if (sender is TextBox textBox)
                {
                    var operation = textBox.Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(() =>
                    {
                        MaterialDesignThemes.Wpf.HintAssist.SetHelperText(
                            textBox, Loc.F("I18n_CustomPromptLength", textBox.Text.Length));
                    }));
                    _pendingOperations.Add(operation);
                   operation.Completed += (s, args) => _pendingOperations.Remove(operation);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"在设置 HelperText 时发生错误: {ex.Message}");
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
         {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

    public class CustomPrompt : INotifyPropertyChanged
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

        private string _function = Loc.T("I18n_CustomPromptInsert");
        public string Function
        {
            get => _function;
            set
            {
                _function = value;
                OnPropertyChanged();
            }
        }

        public string GetTransformedFunction()
        {
            switch (Function)
            {
                case var _ when Function == Loc.T("I18n_CustomPromptInsert") || Function == "插入":
                    return "In";
                case var _ when Function == Loc.T("I18n_CustomPromptInsertWithDiff") || Function == "插入并对比":
                    return "InDiff";
                case var _ when Function == Loc.T("I18n_CustomPromptComment") || Function == "批注":
                    return "CM";
                default:
                    return Function;
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
