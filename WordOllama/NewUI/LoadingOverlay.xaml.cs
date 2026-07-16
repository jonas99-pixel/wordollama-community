using System.Windows;
using System.Windows.Controls;

namespace WordOllama.NewUI
{
    /// <summary>
    /// 通用加载遮罩层 — 被 CreateHTML、ModWritting、MootCourt、WritingUI 等窗口复用。
    /// 替代手写的 &lt;Grid Background="#7F000000"&gt; 遮罩层。
    /// </summary>
    public partial class LoadingOverlay : UserControl
    {
        public LoadingOverlay()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 显示遮罩层，默认提示文字。
        /// </summary>
        public void Show()
        {
            Show(WordOllama.lib.Loc.T("I18n_LoadingPleaseWait"));
        }

        /// <summary>
        /// 显示遮罩层，自定义提示文字。
        /// </summary>
        public void Show(string message)
        {
            MessageText.Text = message;
            OverlayRoot.Visibility = Visibility.Visible;
        }

        /// <summary>
        /// 隐藏遮罩层。
        /// </summary>
        public void Hide()
        {
            OverlayRoot.Visibility = Visibility.Collapsed;
        }

        /// <summary>
        /// 当前是否显示。
        /// </summary>
        public new bool IsVisible => OverlayRoot.Visibility == Visibility.Visible;
    }
}
