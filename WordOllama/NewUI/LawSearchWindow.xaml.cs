using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Newtonsoft.Json.Linq;
using WordOllama.lib;

namespace WordOllama.NewUI
{
    public partial class LawSearchWindow : Window
    {
        private static readonly HttpClient _httpClient = new HttpClient();
        private const string ApiBase = "https://lawapi.lslby.com";
        private const string ArticleApiPath = "/api/v1/article";
        private readonly ObservableCollection<LawArticle> _results = new ObservableCollection<LawArticle>();

        public LawSearchWindow()
        {
            InitializeComponent();
            ThemeManager.ApplyTheme(this.Resources);
            resultsPanel.ItemsSource = _results;
            txtLawName.Focus();
        }

        private void Input_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                BtnSearch_Click(sender, e);
                e.Handled = true;
            }
        }

        private async void BtnSearch_Click(object sender, RoutedEventArgs e)
        {
            string lawName = txtLawName.Text?.Trim();
            string articleRaw = txtArticleNum.Text?.Trim();
            if (string.IsNullOrEmpty(lawName) || string.IsNullOrEmpty(articleRaw)) return;

            // 容错：自动把阿拉伯数字转换为 "第一百二十三条" 这种格式
            string articleNum = FormatArticleNumber(articleRaw);
            string url = $"{ApiBase}{ArticleApiPath}?law={Uri.EscapeDataString(lawName)}&article={Uri.EscapeDataString(articleNum)}";
            string logText = Loc.F("I18n_LawLogText", lawName, articleNum);

            btnSearch.IsEnabled = false;
            lblStatus.Text = Loc.T("I18n_LawSearching");
            _results.Clear();

            try
            {
                using (var response = await _httpClient.GetAsync(url))
                {
                    string responseBody = await response.Content.ReadAsStringAsync();
                    JObject json = null;

                    try
                    {
                        json = JObject.Parse(responseBody);
                    }
                    catch (Newtonsoft.Json.JsonException)
                    {
                        // Preserve the HTTP status when the service returns non-JSON content.
                    }

                    if (!response.IsSuccessStatusCode)
                    {
                        string apiError = json?["error"]?.ToString();
                        string detail = string.IsNullOrWhiteSpace(apiError)
                            ? $"HTTP {(int)response.StatusCode} ({response.ReasonPhrase})"
                            : $"HTTP {(int)response.StatusCode}: {apiError}";
                        lblStatus.Text = Loc.F("I18n_LawRequestFailed", detail);
                        return;
                    }

                    if (json == null)
                    {
                        lblStatus.Text = Loc.F("I18n_LawRequestFailed", "The service returned invalid JSON.");
                        return;
                    }

                    if (json["error"] != null)
                    {
                        lblStatus.Text = Loc.F("I18n_LawQueryFailed", json["error"]);
                        return;
                    }

                    var dataToken = json["data"];
                    if (dataToken is JObject singleData)
                    {
                        string exactLawName = json["law_name"]?.ToString();
                        _results.Add(new LawArticle
                        {
                            LawName = singleData["law_name"]?.ToString() ?? exactLawName ?? "",
                            ArticleNumber = singleData["article_number"]?.ToString() ?? "",
                            Content = singleData["content"]?.ToString() ?? "",
                            Category = singleData["category"]?.ToString() ?? ""
                        });
                        lblStatus.Text = Loc.F("I18n_LawQuerySuccess", logText);
                    }
                    else
                    {
                        lblStatus.Text = Loc.F("I18n_LawNotFound", logText);
                    }
                }
            }
            catch (Exception ex)
            {
                lblStatus.Text = Loc.F("I18n_LawRequestFailed", ex.Message);
            }
            finally
            {
                btnSearch.IsEnabled = true;
            }
        }

        #region 数字到中文的容错转换引擎

        /// <summary>
        /// 将 "14", "第14条", "一百条" 等非标准输入转化为 "第十四条"
        /// </summary>
        private string FormatArticleNumber(string input)
        {
            input = input.Trim();
            // 提取出其中的阿拉伯数字部分
            var match = Regex.Match(input, @"\d+");
            if (match.Success)
            {
                int num = int.Parse(match.Value);
                return "第" + NumberToChinese(num) + "条";
            }

            // 如果用户已经全用中文输入，但没加“第”和“条”，做个简单的补全（可选）
            if (!input.StartsWith("第")) input = "第" + input;
            if (!input.EndsWith("条")) input = input + "条";

            return input;
        }

        /// <summary>
        /// 标准阿拉伯数字转中文（最大支持到 9999）
        /// 14 -> 十四, 102 -> 一百零二, 114 -> 一百一十四
        /// </summary>
        private static string NumberToChinese(int number)
        {
            if (number == 0) return "零";
            string[] cnNums = { "零", "一", "二", "三", "四", "五", "六", "七", "八", "九" };
            string[] cnIntRadice = { "", "十", "百", "千" };

            string result = "";
            string numStr = number.ToString();
            int len = numStr.Length;

            for (int i = 0; i < len; i++)
            {
                int n = numStr[i] - '0';
                int j = len - 1 - i;

                if (n != 0)
                {
                    // 法律里的 10~19 通常叫做 "第十条"、"第十四条"，不叫 "第一十四条"
                    if (n == 1 && j == 1 && len == 2)
                    {
                        result += cnIntRadice[j];
                    }
                    else
                    {
                        result += cnNums[n] + cnIntRadice[j];
                    }
                }
                else
                {
                    // 处理连续的零，比如 101 -> 一百零一
                    if (i < len - 1 && numStr[i + 1] != '0')
                    {
                        result += cnNums[0];
                    }
                }
            }
            return result;
        }
        #endregion

        private void CopySingle_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement btn && btn.Tag is LawArticle article)
            {
                string text = FormatArticle(article);
                Clipboard.SetText(text);
                ShowCopyHint(article.ArticleNumber);
            }
        }

        private void BtnCopyAll_Click(object sender, RoutedEventArgs e)
        {
            if (_results.Count == 0)
            {
                lblStatus.Text = Loc.T("I18n_NoCopyableResult");
                return;
            }

            var all = string.Join("\n\n---\n\n", _results.Select(FormatArticle));
            Clipboard.SetText(all);
            lblStatus.Text = Loc.F("I18n_CopiedLawArticles", _results.Count);
        }

        private string FormatArticle(LawArticle article)
        {
            return $"【{article.LawName}】{article.ArticleNumber}\n{article.Content}";
        }

        private void ShowCopyHint(string articleNumber)
        {
            var orig = lblStatus.Text;
            lblStatus.Text = Loc.F("I18n_CopiedArticle", articleNumber);
            Dispatcher.InvokeAsync(async () =>
            {
                await Task.Delay(2000);
                lblStatus.Text = orig;
            });
        }
    }

    public class LawArticle
    {
        public string LawName { get; set; }
        public string ArticleNumber { get; set; }
        public string Content { get; set; }
        public string Category { get; set; }

        public Visibility CategoryVisibility =>
            string.IsNullOrEmpty(Category) ? Visibility.Collapsed : Visibility.Visible;
    }
}
