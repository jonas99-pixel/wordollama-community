using System;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using WordOllama.NewUI;

namespace WordOllama.NewUI
{
    public partial class ThinkingPane : UserControl
    {
        private ElementHost elementHost;
        private ThinkPaneUI thinkPaneUI;

        public ThinkingPane()
        {
            InitializeComponent();
            InitializeWpfControl();
        }

        private void InitializeWpfControl()
        {
            // 创建 ElementHost
            elementHost = new ElementHost();
            elementHost.Dock = DockStyle.Fill;

            // 创建 WPF 控件
            thinkPaneUI = new ThinkPaneUI();
            elementHost.Child = thinkPaneUI;

            // 添加到当前 UserControl
            this.Controls.Add(elementHost);
        }

        // 提供公共方法来更新内容
        public void UpdateThinkText(string content)
        {
            if (thinkPaneUI != null)
            {
                // 确保在UI线程上执行
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => thinkPaneUI.UpdateContent(content)));
                }
                else
                {
                    thinkPaneUI.UpdateContent(content);
                }
            }
        }

        // 清除内容
        public void ClearContent()
        {
            if (thinkPaneUI != null)
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => thinkPaneUI.ClearContent()));
                }
                else
                {
                    thinkPaneUI.ClearContent();
                }
            }
        }

        // 追加内容
        public void AppendContent(string content)
        {
            if (thinkPaneUI != null)
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => thinkPaneUI.AppendContent(content)));
                }
                else
                {
                    thinkPaneUI.AppendContent(content);
                }
            }
        }

        // 获取内容
        public string GetContent()
        {
            if (thinkPaneUI != null)
            {
                if (this.InvokeRequired)
                {
                    return (string)this.Invoke(new Func<string>(() => thinkPaneUI.GetContent()));
                }
                else
                {
                    return thinkPaneUI.GetContent();
                }
            }
            return string.Empty;
        }

        protected override void OnHandleDestroyed(EventArgs e)
        {
            // 清理资源
            if (elementHost != null)
            {
                elementHost.Child = null;
                elementHost.Dispose();
                elementHost = null;
            }
            base.OnHandleDestroyed(e);
        }
    }
}