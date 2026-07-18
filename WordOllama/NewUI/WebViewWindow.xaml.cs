using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.Wpf;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;

namespace WordOllama.NewUI
{
    public partial class WebViewWindow : Window
    {
        private WebView2 _webView;
        private readonly string _userDataFolder;
        private readonly Task _initializationTask;

        public WebViewWindow()
        {
            InitializeComponent();
            _webView = webView;

            // 设置用户数据文件夹
            _userDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "WordOllama",
                "WebView2UserData");

            if (!Directory.Exists(_userDataFolder))
            {
                Directory.CreateDirectory(_userDataFolder);
            }

            _initializationTask = InitializeWebViewAsync();

            this.Closed += WebViewWindow_Closed;
        }

        private async Task InitializeWebViewAsync()
        {
            var env = await CoreWebView2Environment.CreateAsync(null, _userDataFolder);
            await _webView.EnsureCoreWebView2Async(env);
        }

        public async Task UpdateWebViewAsync(string htmlContent)
        {
            await _initializationTask;
            if (_webView?.CoreWebView2 == null)
                throw new InvalidOperationException("WebView2 initialization did not complete.");
            _webView.CoreWebView2.NavigateToString(htmlContent);
        }

        private void WebViewWindow_Closed(object sender, EventArgs e)
        {
            // 释放 WebView2 控件资源
            _webView?.Dispose();
            _webView = null;
        }


    }
}
