using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using WordOllama.lib;

namespace WordOllama.NewUI
{
    public partial class DiffReviewWindow : Window
    {
        private readonly DocumentAgent _agent;
        private readonly ObservableCollection<DiffChange> _changes;

        public DiffReviewWindow(DocumentAgent agent, System.Collections.Generic.List<DiffChange> changes)
        {
            InitializeComponent();

            _agent = agent;
            _changes = new ObservableCollection<DiffChange>(changes);
            changeList.ItemsSource = _changes;

            UpdateSummary();

            // 没有变更时禁用操作按钮
            if (changes.Count == 0)
            {
                btnAcceptAll.IsEnabled = false;
                btnRejectAll.IsEnabled = false;
                btnAcceptAll.Content = Loc.T("I18n_DiffNoChangesAccept");
                btnRejectAll.Content = Loc.T("I18n_DiffNoChangesReject");
            }
        }

        private void UpdateSummary()
        {
            int added = _changes.Count(c => c.Type == "新增");
            int modified = _changes.Count(c => c.Type == "修改");
            int deleted = _changes.Count(c => c.Type == "删除");
            lblSummary.Text = Loc.F("I18n_DiffSummary", added, modified, deleted, _changes.Count);
        }

        // ===== 全部操作 =====

        private void BtnAcceptAll_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _agent.AcceptAllRevisions();
                MessageBox.Show(Loc.T("I18n_DiffAcceptedAll"), Loc.T("I18n_Completed"), MessageBoxButton.OK, MessageBoxImage.Information);
                Close();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(Loc.F("I18n_DiffAcceptError", ex.Message), Loc.T("I18n_Error"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnRejectAll_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var result = MessageBox.Show(
                    Loc.T("I18n_DiffRejectConfirm"),
                    Loc.T("I18n_DiffRejectConfirmTitle"),
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);
                if (result == MessageBoxResult.Yes)
                {
                    _agent.RejectAllRevisions();
                    MessageBox.Show(Loc.T("I18n_DiffRejectedAll"), Loc.T("I18n_Completed"), MessageBoxButton.OK, MessageBoxImage.Information);
                    Close();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(Loc.F("I18n_DiffRejectError", ex.Message), Loc.T("I18n_Error"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ===== 单项操作 =====

        private void BtnAcceptChange_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement fe && fe.Tag is DiffChange change)
            {
                _agent.AcceptChange(change);
                _changes.Remove(change);
                UpdateSummary();
            }
        }

        private void BtnRejectChange_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement fe && fe.Tag is DiffChange change)
            {
                _agent.RejectChange(change);
                _changes.Remove(change);
                UpdateSummary();
            }
        }

        private void BtnApplyEdit_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement fe && fe.Tag is DiffChange change && change.Type == "修改")
            {
                // TwoWay binding 已同步 NewText，直接应用
                if (string.IsNullOrWhiteSpace(change.NewText))
                {
                    MessageBox.Show(Loc.T("I18n_DiffEmptyEditedContent"), Loc.T("I18n_Info"), MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                _agent.ModifyChange(change, change.NewText);
                _changes.Remove(change);
                UpdateSummary();
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
