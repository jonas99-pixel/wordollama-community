using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;
using WordOllama.lib;

namespace WordOllama.NewUI
{
    /// <summary>
    /// DocxDiffAnalyze.xaml 的交互逻辑
    /// </summary>
    public partial class DocxDiffAnalyze : System.Windows.Window
    {
        private Word.Document _outputDoc;
        public DocxDiffAnalyze()
        {
            InitializeComponent();
            ThemeManager.ApplyTheme(this.Resources);
        }

        private void SelectFile(TextBox textBox)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = Loc.T("I18n_WordDocumentsFilter"),
                Title = Loc.T("I18n_SelectDocumentTitle")
            };

            if (openFileDialog.ShowDialog() == true)
            {
                textBox.Text = openFileDialog.FileName;
            }
        }

        private void BtnSelectFile1_Click(object sender, RoutedEventArgs e)
        {
            SelectFile(Path1TextBox);
        }

        private void BtnSelectFile2_Click(object sender, RoutedEventArgs e)
        {
            SelectFile(Path2TextBox);
        }

        private async void BtnCompare_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(Path1TextBox.Text) || string.IsNullOrEmpty(Path2TextBox.Text))
            {
                MessageBox.Show(Loc.T("I18n_DocxDiffSelectDocsError"), Loc.T("I18n_Error"), MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (string.IsNullOrEmpty(WhoComboBox.Text))
            {
                MessageBox.Show(Loc.T("I18n_DocxDiffSelectRoleError"), Loc.T("I18n_Error"), MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            BtnCompare.IsEnabled = false;
            lblStatus.Text = Loc.T("I18n_DocxDiffRunning");

            try
            {
                // 在 UI 线程上捕获所有必要的值
                string path1 = Path1TextBox.Text;
                string path2 = Path2TextBox.Text;
                string who = WhoComboBox.Text;

                // 创建一个进度报告器，它在 UI 线程上更新状态
                var progress = new Progress<string>(s =>
                {
                    lblStatus.Text = s;
                });

                // 在后台线程上运行耗时操作
                _outputDoc = await Task.Run(() =>
                {
                    return Globals.ThisAddIn.CompareAndAnalyzeDifferences(path1, path2, who, progress);
                });

                // 回到 UI 线程处理结果
                await Dispatcher.InvokeAsync(() =>
                {
                    _outputDoc.Application.Visible = true;
                    _outputDoc.Activate();
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(Loc.T("I18n_Error") + ": " + ex.Message, Loc.T("I18n_Error"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                BtnCompare.IsEnabled = true;
                lblStatus.Text = Loc.T("I18n_DocxDiffComplete");
            }
        }
    }
}
