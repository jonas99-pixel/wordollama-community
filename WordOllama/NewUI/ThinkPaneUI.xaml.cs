using System;
using System.Windows;
using System.Windows.Controls;
using MaterialDesignThemes.Wpf;
using WordOllama.Forms;
using WordOllama.lib;

namespace WordOllama.NewUI
{
    /// <summary>
    /// ThinkPaneUI.xaml 的交互逻辑
    /// </summary>
    public partial class ThinkPaneUI : UserControl
    {
        public ThinkPaneUI()
        {
            InitializeComponent();
            ApplyTheme();
            ThemeManager.ThemeChanged += (_, _) => ApplyTheme();
        }

        private void ApplyTheme()
        {
            try
            {
                if (!Dispatcher.CheckAccess())
                {
                    Dispatcher.Invoke(() => ApplyTheme());
                    return;
                }
                var theme = Resources.GetTheme();
                theme.SetBaseTheme(SharedSettings.themeType == "Dark" ? BaseTheme.Dark : BaseTheme.Light);
                Resources.SetTheme(theme);
            }
            catch { }
        }

        // 更新文本内容
        public void UpdateContent(string content)
        {
            if (ContentTextBox != null)
            {
                ContentTextBox.Text = content;
                ContentTextBox.ScrollToEnd();
            }
        }

        // 清除内容的方法
        public void ClearContent()
        {
            if (ContentTextBox != null)
            {
                ContentTextBox.Clear();
            }
        }

        // 追加内容的方法
        public void AppendContent(string content)
        {
            if (ContentTextBox != null)
            {
                ContentTextBox.AppendText(content);
                ContentTextBox.ScrollToEnd();
            }
        }

        // 获取当前内容的方法
        public string GetContent()
        {
            return ContentTextBox?.Text ?? string.Empty;
        }
    }
}
