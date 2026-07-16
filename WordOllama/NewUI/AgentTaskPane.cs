using System.Windows.Forms;

namespace WordOllama.NewUI
{
    /// <summary>
    /// Agent 任务窗格 — 终端风格的侧边栏交互界面。
    /// 包装 WPF UserControlAgentPanel 以适配 VSTO CustomTaskPane。
    /// 链: UserControl → ElementHost → AgentTaskPaneUI (WPF) → CustomTaskPane
    /// </summary>
    public partial class AgentTaskPane : UserControl
    {
        private readonly System.Windows.Forms.Integration.ElementHost _host;
        private readonly AgentTaskPaneUI _panel;

        public AgentTaskPaneUI Panel => _panel;

        public AgentTaskPane()
        {
            _panel = new AgentTaskPaneUI();
            _host = new System.Windows.Forms.Integration.ElementHost
            {
                Dock = DockStyle.Fill,
                Child = _panel
            };
            Controls.Add(_host);
        }
    }
}
