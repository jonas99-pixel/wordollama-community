using System.Windows.Forms;

namespace WordOllama.NewUI
{
    /// <summary>
    /// 独立审阅工作台，复用 AgentTaskPaneUI 的审阅引擎但只显示审阅界面。
    /// </summary>
    public sealed class ReviewTaskPane : UserControl
    {
        private readonly System.Windows.Forms.Integration.ElementHost _host;
        private readonly AgentTaskPaneUI _panel;

        public AgentTaskPaneUI Panel => _panel;

        public ReviewTaskPane()
        {
            _panel = new AgentTaskPaneUI(reviewOnly: true);
            _host = new System.Windows.Forms.Integration.ElementHost
            {
                Dock = DockStyle.Fill,
                Child = _panel
            };
            Controls.Add(_host);
        }
    }
}
