using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.IO;
using Microsoft.Win32;
using WordOllama.lib;
using WordOllama.NewUI;
using System.Collections.ObjectModel;
using static WordOllama.NewUI.NewSetting;
using System.Windows.Controls;
using WordOllama.Forms;

namespace WordOllama.NewUI
{
    /// <summary>
    /// ChatWithIMG.xaml 的交互逻辑
    /// </summary>
    public partial class ChatWithIMG : Window
    {
        private string _selectedImagePath;

        public ObservableCollection<OllamaModelInfo> OllamaModels { get; set; }
        public ObservableCollection<ChatGPTModelInfo> OnlineModels { get; set; }

        public ChatWithIMG()
        {
            InitializeComponent();
            ThemeManager.ApplyTheme(this.Resources);
            OllamaModels = new ObservableCollection<OllamaModelInfo>();
            OnlineModels = new ObservableCollection<ChatGPTModelInfo>();
            this.DataContext = this;
            LoadModels();
        }

        private void LoadModels()
        {
            // 从 NewSetting 类中获取 OllamaModels 和 ChatGPTModels
            NewSetting newSetting = new NewSetting();
            OllamaModels = newSetting.OllamaModels;
            OnlineModels = newSetting.ChatGPTModels;
            IMGAPIKEY.Password = SharedSettings.IMGAPIKEY;
            IMGEndPoint.Text = SharedSettings.IMGEndPint;
            IMGModel.Text = SharedSettings.IMGModel;

            if (SharedSettings.IMGMode == "Global")
            {
                IMGGlobalModel.IsChecked = true;
                IMGModelNow.Text = WordOllama.Properties.Resources.当前模型全局模式;
            }
            else
            {
                IMGSelectModel.IsChecked = true;
                IMGModelNow.Text = WordOllama.Properties.Resources.当前模型 + SharedSettings.IMGModel;
            }
        }

        private void btnSelectImage_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Image files (*.png;*.jpeg;*.jpg)|*.png;*.jpeg;*.jpg|All files (*.*)|*.*"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _selectedImagePath = openFileDialog.FileName;
                txtImagePath.Text = _selectedImagePath;

                // 显示图片预览
                try
                {
                    imgPreview.Source = new System.Windows.Media.Imaging.BitmapImage(new Uri(_selectedImagePath));
                }
                catch (Exception ex)
                {
                    MessageBox.Show(WordOllama.Properties.Resources.当前模型 + ex.Message , WordOllama.Properties.Resources.错误, MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private async void btnProcess_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_selectedImagePath))
            {
                MessageBox.Show(WordOllama.Properties.Resources.请先选择图片, WordOllama.Properties.Resources.提示, MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtPrompt.Text))
            {
                MessageBox.Show(WordOllama.Properties.Resources.请输入提示词, WordOllama.Properties.Resources.提示, MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                btnProcess.IsEnabled = false;
                string base64String = ConvertImageToBase64(_selectedImagePath);
                string result = await SendToBackend(base64String, txtPrompt.Text);
                txtResult.Text = result;
            }
            catch (Exception ex)
            {
                MessageBox.Show(WordOllama.Properties.Resources.当前模型 + ex.Message , WordOllama.Properties.Resources.错误, MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnProcess.IsEnabled = true;
            }
        }

        private string ConvertImageToBase64(string imagePath)
        {
            byte[] imageBytes = File.ReadAllBytes(imagePath);
            return Convert.ToBase64String(imageBytes);
        }

        private async Task<string> SendToBackend(string base64Image, string prompt)
        {
            if (!string.IsNullOrEmpty(base64Image) && !string.IsNullOrEmpty(prompt))
            {
                var response = await Globals.ThisAddIn.GenerateModelSeletor(prompt, base64Image,"IMG");
                if (response != null)
                {
                    return response;
                }
                else
                {
                    return WordOllama.Properties.Resources.错误AI没有返回结果请检查;
                }
            }
            else
            {
                return WordOllama.Properties.Resources.错误请确认图片和提示词不为空;
            }
        }

        private async void OptimizationButton_Click(object sender, RoutedEventArgs e)
        {
            await Utils.OptimizeTextBoxAsync(txtPrompt, sender as Button);
        }

        private void Copy_Btn_Click(object sender, RoutedEventArgs e)
        {
            txtResult.SelectAll();
            txtResult.Copy();
        }

        private void OllamaModelSeletor_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (OllamaModelSeletor.SelectedItem is OllamaModelInfo selectedModel)
            {
                IMGModel.Text = selectedModel.Name;
                IMGEndPoint.Text = AppConfig.Instance.OllamaEndPoint + "/v1/chat/completions";
                IMGAPIKEY.Password = "ollama";
                OnlineModelSeletor.SelectedItem = null;
                IMGModelNow.Text = WordOllama.Properties.Resources.当前模型 + selectedModel.Name;
            }
        }

        private void OnlineModelSeletor_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (OnlineModelSeletor.SelectedItem is ChatGPTModelInfo selectedModel)
            {
                IMGModel.Text = selectedModel.ModelName;
                IMGEndPoint.Text = selectedModel.EndPoint;
                IMGAPIKEY.Password = selectedModel.RealApiKey;
                OllamaModelSeletor.SelectedItem = null;
                IMGModelNow.Text = WordOllama.Properties.Resources.当前模型 + selectedModel.ModelName;
            }
        }

        private void Save_btn_Click(object sender, RoutedEventArgs e)
        {
            SharedSettings.IMGModel = IMGModel.Text;
            SharedSettings.IMGEndPint = IMGEndPoint.Text;
            SharedSettings.IMGAPIKEY = IMGAPIKEY.Password;
            SharedSettings.IMGMode = IMGGlobalModel.IsChecked == true ? "Global" : "Select";
            IMGModelNow.Text = IMGGlobalModel.IsChecked == true ? WordOllama.Properties.Resources.当前模型全局模式 : IMGModel.Text;
        }
    }
}
