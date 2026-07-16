using System.Windows;

namespace WordOllama.NewUI
{
    /// <summary>
    /// 通用确认对话框 — 支持"不再提示"选项（会话级/永久级）。
    /// 使用方式：
    ///   var dlg = new ConfirmationDialog("消息内容", "标题", "确认点Key");
    ///   if (dlg.ShowDialog() == true) { /* 用户确认 */ }
    ///   通过 dlg.IsOptedOutSession / dlg.IsOptedOutPermanent 检查用户是否选择了免打扰。
    /// </summary>
    public partial class ConfirmationDialog : Window
    {
        /// <summary>用户是否点击了「确认」。</summary>
        public bool IsConfirmed { get; private set; }

        /// <summary>是否选择了会话级免打扰。</summary>
        public bool IsOptedOutSession { get; private set; }

        /// <summary>是否选择了永久免打扰。</summary>
        public bool IsOptedOutPermanent { get; private set; }

        /// <summary>关联的确认点 Key（用于 ConfirmationManager）。</summary>
        public string ConfirmationKey { get; }

        /// <param name="message">对话框显示的消息内容。</param>
        /// <param name="title">窗口标题。</param>
        /// <param name="confirmationKey">确认点 Key（ConfirmationManager.KeyPlanConfirmation 等），为 null 时不显示免打扰选项。</param>
        public ConfirmationDialog(string message, string title = null, string confirmationKey = null)
        {
            InitializeComponent();
            txtMessage.Text = message;
            Title = string.IsNullOrWhiteSpace(title) ? WordOllama.lib.Loc.T("I18n_Confirm") : title;
            ConfirmationKey = confirmationKey;

            // 如果没有确认点 Key，不显示免打扰选项
            if (string.IsNullOrEmpty(confirmationKey))
            {
                chkSession.Visibility = Visibility.Collapsed;
                chkPermanent.Visibility = Visibility.Collapsed;
            }
        }

        private void BtnConfirm_Click(object sender, RoutedEventArgs e)
        {
            IsConfirmed = true;
            if (!string.IsNullOrEmpty(ConfirmationKey))
            {
                if (chkSession.IsChecked == true)
                {
                    IsOptedOutSession = true;
                    lib.ConfirmationManager.OptOutSession(ConfirmationKey);
                }
                if (chkPermanent.IsChecked == true)
                {
                    IsOptedOutPermanent = true;
                    lib.ConfirmationManager.OptOutPermanently(ConfirmationKey);
                }
            }
            DialogResult = true;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            IsConfirmed = false;
            DialogResult = false;
            Close();
        }

        private void ChkSession_Checked(object sender, RoutedEventArgs e)
        {
            // 取消勾选永久时自动取消会话级
        }

        private void ChkPermanent_Checked(object sender, RoutedEventArgs e)
        {
            // 勾选永久时自动勾选会话级（永久包含会话级）
            chkSession.IsChecked = true;
        }
    }
}
