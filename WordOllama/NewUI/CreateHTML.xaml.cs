using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Office.Tools;
using System.Diagnostics;
using MaterialDesignThemes.Wpf;
using static WordOllama.NewUI.NewSetting;
using WordOllama.Forms;
using WordOllama.lib;
using System.Net;

namespace WordOllama.NewUI
{

    public partial class CreateHTML : Window
    {
        private readonly string _htmlsFolder;
        private ObservableCollection<HtmlFileInfo> _htmlFiles;
        public ObservableCollection<OllamaModelInfo> OllamaModels { get; set; }
        public ObservableCollection<ChatGPTModelInfo> OnlineModels { get; set; }

        public CreateHTML()
        {
            InitializeComponent();
            ThemeManager.ApplyTheme(this.Resources);
            OllamaModels = new ObservableCollection<OllamaModelInfo>();
            OnlineModels = new ObservableCollection<ChatGPTModelInfo>();
            this.DataContext = this;
            LoadModels();



            // 初始化HTML文件夹路径
            _htmlsFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "WordOllama",
                "HTMLS");

            if (!Directory.Exists(_htmlsFolder))
            {
                Directory.CreateDirectory(_htmlsFolder);
            }

            // 初始化控件引用
            InitializeControls();

            // 加载已保存的HTML文件
            LoadHtmlFiles();

            // 绑定事件处理
            BindEvents();
        }

        private void LoadModels()
        {
            // 从 NewSetting 类中获取 OllamaModels 和 ChatGPTModels
            NewSetting newSetting = new NewSetting();
            OllamaModels = newSetting.OllamaModels;
            OnlineModels = newSetting.ChatGPTModels;
            HTMLAPIKEY.Password = SharedSettings.HTMLAPIKEY;
            HTMLEndPoint.Text = SharedSettings.HTMLEndPint;
            HTMLModel.Text = SharedSettings.HTMLModel;

            if (SharedSettings.HTMLMode == "Global")
            {
                HTMLGlobalModel.IsChecked = true;
                HTMLModelNow.Text = WordOllama.Properties.Resources.当前模型全局模式;
            }
            else
            {
                HTMLSelectModel.IsChecked = true;
                HTMLModelNow.Text = WordOllama.Properties.Resources.当前模型 + SharedSettings.HTMLModel;
            }
        }

        private void OllamaModelSeletor_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (OllamaModelSeletor.SelectedItem is OllamaModelInfo selectedModel)
            {
                HTMLModel.Text = selectedModel.Name;
                HTMLEndPoint.Text = AppConfig.Instance.OllamaEndPoint + "/v1/chat/completions";
                HTMLAPIKEY.Password = "ollama";
                OnlineModelSeletor.SelectedItem = null;
                HTMLModelNow.Text = WordOllama.Properties.Resources.当前模型 + selectedModel.Name;
            }
        }

        private void OnlineModelSeletor_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (OnlineModelSeletor.SelectedItem is ChatGPTModelInfo selectedModel)
            {
                HTMLModel.Text = selectedModel.ModelName;
                HTMLEndPoint.Text = selectedModel.EndPoint;
                HTMLAPIKEY.Password = selectedModel.RealApiKey;
                OllamaModelSeletor.SelectedItem = null;
                HTMLModelNow.Text = WordOllama.Properties.Resources.当前模型 + selectedModel.ModelName;
            }
        }

        private void Model_Save_btn_Click(object sender, RoutedEventArgs e)
        {
            SharedSettings.HTMLModel = HTMLModel.Text;
            SharedSettings.HTMLEndPint = HTMLEndPoint.Text;
            SharedSettings.HTMLAPIKEY = HTMLAPIKEY.Password;
            SharedSettings.HTMLMode = HTMLGlobalModel.IsChecked == true ? "Global" : "Select";
            HTMLModelNow.Text = HTMLGlobalModel.IsChecked == true ? WordOllama.Properties.Resources.当前模型全局模式 : HTMLModel.Text;
        }


        private void InitializeControls()
        {
        }

        private void BindEvents()
        {
            // 为ListBox绑定选择改变事件
            htmlFileListBox.SelectionChanged += HtmlFileListBox_SelectionChanged;
        }

        private void LoadHtmlFiles()
        {
            try
            {
                var files = Directory.GetFiles(_htmlsFolder, "*.html");
                _htmlFiles = new ObservableCollection<HtmlFileInfo>();

                foreach (var file in files)
                {
                    _htmlFiles.Add(new HtmlFileInfo
                    {
                        Name = Path.GetFileNameWithoutExtension(file),
                        Path = file
                    });
                }

                htmlFileListBox.ItemsSource = _htmlFiles;
            }
            catch (Exception ex)
            {
                MessageBox.Show(WordOllama.Properties.Resources.当前模型 + ex.Message , WordOllama.Properties.Resources.错误, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void HtmlFileListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (htmlFileListBox.SelectedItem is HtmlFileInfo selectedFile)
            {
                try
                {
                    // 使用同步方法替代异步方法
                    string content = File.ReadAllText(selectedFile.Path);
                    selectedHtmlContentTextBox.Text = content;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(WordOllama.Properties.Resources.读取HTML文件失败 + ex.Message, WordOllama.Properties.Resources.错误, MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private async void SendPrompt_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Begin_Btn.Dispatcher.Invoke(() =>
                {
                    Begin_Btn.IsEnabled = false;
                    Begin_Btn.Content = Properties.Resources.Common_Creating;
                    thisProgressBar.Show();
                });

                if (string.IsNullOrWhiteSpace(promptTextBox.Text))
                {
                    MessageBox.Show(WordOllama.Properties.Resources.请输入提示词, WordOllama.Properties.Resources.提示, MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                string prompt = PromptTemplateService.BuildHtmlAppPrompt(promptTextBox.Text);
                string response = await Globals.ThisAddIn.GenerateModelSeletor(prompt,null,"HTML");

                if (!string.IsNullOrWhiteSpace(response))
                {
                    // 去除开头的 ```html 和结尾的 ```

                        response = response.Replace("```html", "").Replace("```", "");

                    htmlCodeTextBox.Text = response;
                }
                else
                {
                    htmlCodeTextBox.Text = Properties.Resources.Error_NoAIResponse;
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(WordOllama.Properties.Resources.发送提示词失败 + ex.Message, WordOllama.Properties.Resources.错误, MessageBoxButton.OK, MessageBoxImage.Error);
            }

            finally {
                Begin_Btn.Dispatcher.Invoke(() =>
                {
                    Begin_Btn.IsEnabled = true;
                    Begin_Btn.Content = Properties.Resources.Common_Begin;
                    thisProgressBar.Hide();
                });
            }
        }

        private async void OptimizationButton_Click(object sender, RoutedEventArgs e)
        {
            promptTextBox.Text = await Utils.PromptOptimizer(promptTextBox.Text);
        }

        private void OpenPopup_Click(object sender, RoutedEventArgs e)
        {
            AddHTMLPopup.Dispatcher.Invoke(() =>
            {
                AddHTMLPopup.IsOpen = true;
            });
        }

        private void AddHtml_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(HTMLNameTextBox.Text))
            {
                throw new Exception(WordOllama.Properties.Resources.请填入有效的文件名);
            }
            try
            {
                string fileName = $"{HTMLNameTextBox.Text}.html";
                string filePath = Path.Combine(_htmlsFolder, fileName);

                string contentToSave = htmlCodeTextBox.Text;

                // 使用同步方法替代异步方法
                File.WriteAllText(filePath, contentToSave);
                LoadHtmlFiles();

                MessageBox.Show(WordOllama.Properties.Resources.保存成功, WordOllama.Properties.Resources.提示, MessageBoxButton.OK, MessageBoxImage.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show(WordOllama.Properties.Resources.保存HTML文件失败 + ex.Message, WordOllama.Properties.Resources.错误, MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                AddHTMLPopup.Dispatcher.Invoke(() =>
                {
                    AddHTMLPopup.IsOpen = false;
                    HTMLNameTextBox.Clear();
                });
            }
        }
        private void DeleteHtml_Click(object sender, RoutedEventArgs e)
        {
            if (htmlFileListBox.SelectedItem is HtmlFileInfo selectedFile)
            {
                var result = MessageBox.Show(WordOllama.Properties.Resources.确定删除 + selectedFile.Name, WordOllama.Properties.Resources.确认删除, MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        File.Delete(selectedFile.Path);
                        _htmlFiles.Remove(selectedFile);
                        selectedHtmlContentTextBox.Text = string.Empty;
                        MessageBox.Show(WordOllama.Properties.Resources.文件删除成功, WordOllama.Properties.Resources.提示, MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(WordOllama.Properties.Resources.删除文件失败 + ex.Message, WordOllama.Properties.Resources.错误, MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            else
            {
                MessageBox.Show(WordOllama.Properties.Resources.请选择要删除的文件, WordOllama.Properties.Resources.提示, MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        private void SaveHtml_Click(object sender, RoutedEventArgs e)
        {
            if (htmlFileListBox.SelectedItem is HtmlFileInfo selectedFile)
            {
                string filePath = selectedFile.Path;
                string contentToSave = selectedHtmlContentTextBox.Text;

                if (File.Exists(filePath))
                {
                    var result = MessageBox.Show(selectedFile.Name + WordOllama.Properties.Resources.已存在是否覆盖保存, WordOllama.Properties.Resources.确认覆盖, MessageBoxButton.YesNo, MessageBoxImage.Warning);
                    if (result == MessageBoxResult.No)
                    {
                        return;
                    }
                }

                try
                {
                    File.WriteAllText(filePath, contentToSave);
                    MessageBox.Show(WordOllama.Properties.Resources.保存成功, WordOllama.Properties.Resources.提示, MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(WordOllama.Properties.Resources.保存文件失败 + ex.Message, WordOllama.Properties.Resources.错误, MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                MessageBox.Show(WordOllama.Properties.Resources.请选择要保存的文件, WordOllama.Properties.Resources.提示, MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void CancelAddHTML_Click(object sender, RoutedEventArgs e)
        {
            AddHTMLPopup.IsOpen = false;
            HTMLNameTextBox.Clear();
        }

        private async void RunHtml_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Button clickedButton = sender as Button;
                string contentToRun = clickedButton?.Tag.ToString() == "NewAPP" ? htmlCodeTextBox.Text : selectedHtmlContentTextBox.Text;

                if (string.IsNullOrWhiteSpace(contentToRun))
                {
                    MessageBox.Show(WordOllama.Properties.Resources.没有可运行的HTML内容, WordOllama.Properties.Resources.提示, MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                var webViewWindow = new WebViewWindow();
                await webViewWindow.UpdateWebViewAsync(contentToRun);
                webViewWindow.Show();

            }
            catch (Exception ex)
            {
                MessageBox.Show(WordOllama.Properties.Resources.运行HTML失败 + ex.Message, WordOllama.Properties.Resources.错误, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CancelAddModel_Click(object sender, RoutedEventArgs e)
        {

        }
    }

    public class HtmlFileInfo
    {
        public string Name { get; set; }
        public string Path { get; set; }
    }

}
