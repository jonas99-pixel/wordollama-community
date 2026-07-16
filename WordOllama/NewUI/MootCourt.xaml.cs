using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using WordOllama.Forms;
using WordOllama.lib;

namespace WordOllama.NewUI
{
    /// <summary>
    /// MootCourt.xaml 的交互逻辑
    /// </summary>
    public partial class MootCourt : UserControl
    {
        public MootCourt()
        {
            InitializeComponent();
            Loaded += MootCourt_Loaded;
        }

        private void MootCourt_Loaded(object sender, RoutedEventArgs e)
        {
            LoadDocText();
            ThemeManager.ApplyTheme(this.Resources);
        }

        private void LoadDocText()
        {
            docText.Text = GetText.FullText;
        }

        private async void GetResponse_Click(object sender, RoutedEventArgs e)
        {
            thisProgressBar.Show();
            await MootIt();
            thisProgressBar.Hide();
        }

        private async Task MootIt()
        {
            bool english = string.Equals(SharedSettings.LanguageMode, "English", StringComparison.OrdinalIgnoreCase);
            string DocType;

            if (IndictmentRadio.IsChecked == true)
            {
                DocType = english ? "indictment" : "起诉状";
            }
            else
            {
                DocType = english ? "statement of defence" : "答辩状";
            }

            try
            {
                string prompt = PromptTemplateService.BuildMootCourtInvestigationPrompt(DocType, docText.Text, english);

                string response = await Globals.ThisAddIn.GenerateModelSeletor(prompt);
                if (!string.IsNullOrEmpty(response))
                {
                    ResponseBox.Text = Loc.T("I18n_MootCourtResultPrefix") + "\n" + response;
                }
                else
                {
                    ResponseBox.Text = Loc.T("I18n_ModelReturnedEmpty");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(Loc.F("I18n_ModifyError", ex.Message));
            }
        }



    }

}
