using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using MaterialDesignThemes.Wpf;
using System.Text.Json;
using Style = System.Windows.Style;
using Task = System.Threading.Tasks.Task;
using System.Text.Json.Serialization;
using WordOllama.lib;
using static System.Net.Mime.MediaTypeNames;
using WordOllama.Forms;
using System.Collections.ObjectModel;
using Newtonsoft.Json;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace WordOllama.NewUI
{
    /// <summary>
    /// TextToTableWindow.xaml 的交互逻辑
    /// </summary>
    public partial class TextToTableWindow : System.Windows.Window, INotifyPropertyChanged
    {
        private string dataFilePath;
        private System.Data.DataTable tableData;
        private TableResponse currentResponse;

        private ObservableCollection<TableTemplate> _templates;
        public ObservableCollection<TableTemplate> Templates
        {
            get => _templates;
            set
            {
                _templates = value;
                OnPropertyChanged();
            }
        }

        public TextToTableWindow()
        {
            dataFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WordOllama", "TablePrompt.json");
            InitializeComponent();
            LoadTemplates();
            InitializeTemplateComboBox();
            ThemeManager.ApplyTheme(this.Resources);
            InputTextBox.Text = string.IsNullOrEmpty(GetText.SelectedText) ? string.Empty : GetText.SelectedText;
            
            DataContext = this;
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
                Templates = new ObservableCollection<TableTemplate>();
                SaveTemplates();
            }
            else
            {
                var json = File.ReadAllText(dataFilePath);
                var loadedTemplates = JsonConvert.DeserializeObject<ObservableCollection<TableTemplate>>(json) ?? new ObservableCollection<TableTemplate>();
                Templates = new ObservableCollection<TableTemplate>(loadedTemplates);
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

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void TemplateComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (TemplateComboBox.SelectedIndex != -1)
            {
                var selectedTemplate = Templates[TemplateComboBox.SelectedIndex];
                RequirementTextBox.Text = selectedTemplate.Prompt;
            }
        }
        private void DeleteTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var template = button.DataContext as TableTemplate;

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
            if (CanAddNewRow())
            {
                Templates.Add(new TableTemplate { Name = "", Prompt = "" });
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

        private void SaveTemplatesButton_Click(object sender, RoutedEventArgs e)
        {
            SaveTemplates();
            MessageBox.Show(WordOllama.Properties.Resources.模板已保存, WordOllama.Properties.Resources.提示, MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private async void OptimizeTextBox_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            await Utils.OptimizeTextBoxAsync(button?.Tag as TextBox, button);
        }

        private async void OptimizeTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            if (!(button?.DataContext is TableTemplate template)) return;
            button.IsEnabled = false;
            try
            {
                string optimized = await Utils.PromptOptimizer(template.Prompt);
                if (!string.IsNullOrWhiteSpace(optimized)) template.Prompt = optimized;
            }
            finally { button.IsEnabled = true; }
        }

        // 表格数据结构
        private class TableResponse
        {
            [JsonPropertyName("headers")]
            public string[] Headers { get; set; }

            [JsonPropertyName("rows")]
            public string[][] Rows { get; set; }
        }


        private async 

        Task
GenerateTable()
        {
            try
            {
                // 显示进度条
                Dispatcher.Invoke(() =>
                {
                    ProgressBar.Visibility = Visibility.Visible;
                    GenerateButton.IsEnabled = false;
                    InsertToWordButton.IsEnabled = false;

                });


                // 获取输入文本
                string inputText = InputTextBox.Text;
                if (string.IsNullOrWhiteSpace(inputText))
                {
                    throw new Exception(WordOllama.Properties.Resources.请输入需要转换的文本);
                }

                // 构建提示词
                var promptBuilder = new TablePromptBuilder()
                    .WithText(inputText)
                    .WithRequirement(RequirementTextBox.Text)
                    .Build();

                // 调用LLM服务
                string response = await Globals.ThisAddIn.GenerateModelSeletor(promptBuilder);

                // 去除开头的 ```json\n 和结尾的 ```
                if (response.StartsWith("```json\n") && response.EndsWith("```"))
                {
                    response = response.Substring("```json\n".Length, response.Length - "```json\n".Length - "```".Length);
                }

                // 解析JSON响应
                currentResponse = System.Text.Json.JsonSerializer.Deserialize<TableResponse>(response);

                // 创建DataTable
                tableData = new System.Data.DataTable();

                // 添加列
                foreach (string header in currentResponse.Headers)
                {
                    tableData.Columns.Add(header);
                }

                // 添加行
                foreach (string[] row in currentResponse.Rows)
                {
                    tableData.Rows.Add(row);
                }

                // 更新DataGrid
                PreviewDataGrid.ItemsSource = tableData.DefaultView;

                // 启用插入按钮
                InsertToWordButton.IsEnabled = true;
            }
            catch (Exception ex)
            {
                await ShowErrorDialog(WordOllama.Properties.Resources.生成表格失败, ex.Message);
            }
            finally
            {
                // 隐藏进度条
                Dispatcher.Invoke(() =>
                {
                    ProgressBar.Visibility = Visibility.Collapsed;
                    GenerateButton.IsEnabled = true;
                });



            }
        }

        // 生成表格按钮点击事件
        private async void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            await GenerateTable();
        }

        // 插入到Word按钮点击事件
        private async void InsertToWordButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {

                // 获取当前Word文档
                var wordApp = Globals.ThisAddIn.Application;
                var doc = wordApp.ActiveDocument;

                if (doc == null)
                {
                    throw new Exception(WordOllama.Properties.Resources.未找到活动的Word文档);
                }

                // 获取当前光标位置
                var range = wordApp.Selection.Range;

                // 创建表格
                var table = doc.Tables.Add(
                    range,
                    currentResponse.Rows.Length + 1,  // 行数(包含表头)
                    currentResponse.Headers.Length,    // 列数
                    WdDefaultTableBehavior.wdWord9TableBehavior,
                    WdAutoFitBehavior.wdAutoFitContent);

                // 设置表格样式
                table.Borders.Enable = 1;
                table.set_Style("网格型");

                // 添加表头
                for (int i = 0; i < currentResponse.Headers.Length; i++)
                {
                    table.Cell(1, i + 1).Range.Text = currentResponse.Headers[i];
                }

                // 设置表头样式
                table.Rows[1].Range.Bold = 1;
                table.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 添加数据行
                for (int row = 0; row < currentResponse.Rows.Length; row++)
                {
                    for (int col = 0; col < currentResponse.Rows[row].Length; col++)
                    {
                        table.Cell(row + 2, col + 1).Range.Text = currentResponse.Rows[row][col];
                    }
                }

                // 自动调整表格
                table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);

                await ShowSuccessDialog(WordOllama.Properties.Resources.插入成功, WordOllama.Properties.Resources.表格已成功插入到Word文档中);
            }
            catch (Exception ex)
            {
                await ShowErrorDialog(WordOllama.Properties.Resources.插入表格失败, ex.Message);
            }
            finally
            {
            }
        }

        // 显示错误对话框
        private async Task ShowErrorDialog(string title, string message)
        {
            var content = new StackPanel { Margin = new Thickness(16) };
            content.Children.Add(new TextBlock
            {
                Text = title,
                Style = (Style)FindResource("MaterialDesignHeadline6TextBlock"),
                Margin = new Thickness(0, 0, 0, 8)
            });
            content.Children.Add(new TextBlock
            {
                Text = message,
                TextWrapping = TextWrapping.Wrap
            });

            var button = new Button
            {
                Content = WordOllama.Properties.Resources.确定,
                Style = (Style)FindResource("MaterialDesignFlatButton"),
                Command = DialogHost.CloseDialogCommand,
                Margin = new Thickness(0, 16, 0, 0)
            };
            content.Children.Add(button);

            await DialogHost.Show(content, "MainDialogHost");
        }


        // 显示成功对话框
        private async Task ShowSuccessDialog(string title, string message)
        {
            var content = new StackPanel { Margin = new Thickness(16) };
            content.Children.Add(new TextBlock
            {
                Text = title,
                Style = (Style)FindResource("MaterialDesignHeadline6TextBlock"),
                Margin = new Thickness(0, 0, 0, 8)
            });
            content.Children.Add(new TextBlock
            {
                Text = message,
                TextWrapping = TextWrapping.Wrap
            });
            var button = new Button
            {
                Content = WordOllama.Properties.Resources.确定,
                Style = (Style)FindResource("MaterialDesignFlatButton"),
                Command = DialogHost.CloseDialogCommand,
                Margin = new Thickness(0, 16, 0, 0)
            };
            content.Children.Add(button);

            await DialogHost.Show(content, "MainDialogHost");
        }
        public class TablePromptBuilder
        {
            private readonly List<string> requirements;
            private string sourceText;

            public TablePromptBuilder()
            {
                requirements = new List<string>();
            }

            // 设置源文本
            public TablePromptBuilder WithText(string text)
            {
                this.sourceText = text;
                return this;
            }

            // 添加指定列名
            public TablePromptBuilder WithColumns(params string[] columns)
            {
                if (columns != null && columns.Length > 0)
                {
                    requirements.Add(WordOllama.Properties.Resources.必须包含以下列 + string.Join("、", columns));
                }
                return this;
            }

            // 添加数据格式要求
            public TablePromptBuilder WithDataFormat(string format)
            {
                requirements.Add(WordOllama.Properties.Resources.数据格式要求 + format);
                return this;
            }

            // 添加自定义要求
            public TablePromptBuilder WithRequirement(string requirement)
            {
                requirements.Add(requirement);
                return this;
            }

            // 生成最终的提示词
            public string Build()
            {
                return PromptTemplateService.BuildTablePrompt(sourceText, requirements);
            }
        }
        
    }
    public class TableTemplate : INotifyPropertyChanged
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
