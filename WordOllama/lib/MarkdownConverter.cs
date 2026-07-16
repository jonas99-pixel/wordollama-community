using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Controls;

namespace WordOllama.lib
{
    /// <summary>
    /// Markdown → WPF FlowDocument 转换器。
    /// 支持：##/### 标题、**粗体**、*斜体*、`行内代码`、```代码块```、- 无序列表、1. 有序列表。
    /// 用于 Agent 侧边栏的聊天气泡渲染。
    /// </summary>
    public class MarkdownToFlowDocumentConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string text = value as string;
            if (string.IsNullOrWhiteSpace(text))
                return null;

            var doc = new FlowDocument
            {
                PagePadding = new Thickness(0),
                TextAlignment = TextAlignment.Left,
                FontFamily = new FontFamily("Segoe UI, Microsoft YaHei UI"),
                FontSize = 13
            };

            ParseToDocument(doc, text);
            return doc;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return Binding.DoNothing;
        }

        private void ParseToDocument(FlowDocument doc, string text)
        {
            string[] lines = text.Split('\n');
            int i = 0;

            while (i < lines.Length)
            {
                string line = lines[i];

                // 代码块
                if (line.TrimStart().StartsWith("```"))
                {
                    i = ParseCodeBlock(doc, lines, i);
                    continue;
                }

                // 标题
                var headerMatch = Regex.Match(line, @"^(#{2,3})\s+(.+)$");
                if (headerMatch.Success)
                {
                    int level = headerMatch.Groups[1].Value.Length;
                    string content = headerMatch.Groups[2].Value.Trim();
                    doc.Blocks.Add(CreateHeader(content, level));
                    i++;
                    continue;
                }

                // 空行 → 段落间距
                if (string.IsNullOrWhiteSpace(line))
                {
                    i++;
                    continue;
                }

                // 无序列表
                var ulMatch = Regex.Match(line, @"^[-*]\s+(.+)$");
                if (ulMatch.Success)
                {
                    doc.Blocks.Add(CreateListItem(ulMatch.Groups[1].Value, "•"));
                    i++;
                    continue;
                }

                // 有序列表
                var olMatch = Regex.Match(line, @"^(\d+)\.\s+(.+)$");
                if (olMatch.Success)
                {
                    doc.Blocks.Add(CreateListItem(olMatch.Groups[2].Value, $"{olMatch.Groups[1].Value}."));
                    i++;
                    continue;
                }

                // 表格行
                if (line.TrimStart().StartsWith("|"))
                {
                    i = ParseTable(doc, lines, i);
                    continue;
                }

                // 普通段落（可能跨多行）
                i = ParseParagraph(doc, lines, i);
            }
        }

        private int ParseCodeBlock(FlowDocument doc, string[] lines, int start)
        {
            var codeLines = new List<string>();
            int i = start + 1;
            while (i < lines.Length && !lines[i].TrimStart().StartsWith("```"))
            {
                codeLines.Add(lines[i]);
                i++;
            }
            // 跳过结束的 ```
            if (i < lines.Length) i++;

            if (codeLines.Count > 0)
            {
                var codeBlock = new Paragraph
                {
                    Margin = new Thickness(0, 4, 0, 6),
                    Padding = new Thickness(10, 6, 10, 6),
                    Background = new SolidColorBrush(Color.FromRgb(245, 245, 245)),
                    FontFamily = new FontFamily("Consolas, Courier New"),
                    FontSize = 12
                };
                codeBlock.Inlines.Add(new Run(string.Join("\n", codeLines)));
                doc.Blocks.Add(codeBlock);
            }

            return i;
        }

        private int ParseTable(FlowDocument doc, string[] lines, int start)
        {
            if (start + 1 >= lines.Length || string.IsNullOrWhiteSpace(lines[start + 1]))
            {
                // 只有一行表头 → 当普通文本处理
                return ParseParagraph(doc, lines, start);
            }

            // 跳过可能的分隔行（|---|---|）
            int dataStart = start + 1;
            if (Regex.IsMatch(lines[dataStart].Trim(), @"^[\|\-\:\s]+$"))
                dataStart++;

            var rows = new List<string[]> { ParseTableRow(lines[start]) };
            for (int i = dataStart; i < lines.Length && !string.IsNullOrWhiteSpace(lines[i]); i++)
            {
                if (lines[i].TrimStart().StartsWith("|"))
                    rows.Add(ParseTableRow(lines[i]));
                else
                    break;
            }

            // 用固定宽度文本格式显示表格
            var tableBlock = new Paragraph
            {
                Margin = new Thickness(0, 4, 0, 6),
                FontSize = 12,
                FontFamily = new FontFamily("Consolas, Courier New")
            };

            foreach (var row in rows)
            {
                tableBlock.Inlines.Add(new Run(string.Join(" | ", row) + "\n"));
            }
            doc.Blocks.Add(tableBlock);

            return start + rows.Count;
        }

        private string[] ParseTableRow(string line)
        {
            string trimmed = line.Trim().TrimStart('|').TrimEnd('|');
            return trimmed.Split('|');
        }

        private int ParseParagraph(FlowDocument doc, string[] lines, int start)
        {
            var paraLines = new List<string> { lines[start] };
            int i = start + 1;
            while (i < lines.Length)
            {
                string next = lines[i];
                if (string.IsNullOrWhiteSpace(next))
                    break;
                if (next.TrimStart().StartsWith("```") || next.TrimStart().StartsWith("#") ||
                    next.TrimStart().StartsWith("- ") || next.TrimStart().StartsWith("* ") ||
                    Regex.IsMatch(next, @"^\d+\.\s+") || next.TrimStart().StartsWith("|"))
                    break;
                paraLines.Add(next);
                i++;
            }

            string paraText = string.Join("\n", paraLines);
            doc.Blocks.Add(CreateParagraph(paraText));
            return i;
        }

        private Paragraph CreateHeader(string text, int level)
        {
            var para = new Paragraph();
            double fontSize = level == 2 ? 15 : 14;
            para.FontSize = fontSize;
            para.FontWeight = FontWeights.SemiBold;
            para.Margin = new Thickness(0, 6, 0, 3);
            para.Inlines.AddRange(ParseInlines(text));
            return para;
        }

        private Paragraph CreateListItem(string text, string bullet)
        {
            var para = new Paragraph
            {
                Margin = new Thickness(0, 0, 0, 3),
                TextIndent = -14,
                Padding = new Thickness(18, 0, 0, 0)
            };
            para.Inlines.Add(new Run(bullet + " "));
            para.Inlines.AddRange(ParseInlines(text));
            return para;
        }

        private Paragraph CreateParagraph(string text)
        {
            var para = new Paragraph
            {
                Margin = new Thickness(0, 0, 0, 4)
            };
            para.Inlines.AddRange(ParseInlines(text));
            return para;
        }

        /// <summary>
        /// 解析行内格式：**粗体**、*斜体*、`行内代码`、[链接文本](url)
        /// </summary>
        private List<Inline> ParseInlines(string text)
        {
            var inlines = new List<Inline>();

            // 按顺序解析：代码 > 粗体 > 斜体 > 链接 > 文本
            var patterns = new (string pattern, Func<Match, int, int, Inline> factory)[]
            {
                // 行内代码 `...` — 优先匹配，避免 `**` 被误解析
                (@"`([^`]+)`", (m, s, l) => new Run(m.Groups[1].Value)
                {
                    FontFamily = new FontFamily("Consolas, Courier New"),
                    Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)),
                    FontSize = 12
                }),
                // **粗体**
                (@"\*\*(.+?)\*\*", (m, s, l) => new Bold(new Run(m.Groups[1].Value))),
                // *斜体*（不跨行）
                (@"\*(.+?)\*", (m, s, l) => new Italic(new Run(m.Groups[1].Value))),
                // ![alt](url) — 图片行内渲染（支持 data: URI 和 http 链接）
                (@"!\[([^\]]*)\]\(([^)]+)\)", (m, s, l) =>
                {
                    string alt = m.Groups[1].Value;
                    string url = m.Groups[2].Value;
                    var img = new Image { MaxWidth = 240, MaxHeight = 320, ToolTip = alt };

                    try
                    {
                        if (url.StartsWith("data:"))
                        {
                            string base64Data = url;
                            int commaIndex = base64Data.IndexOf(',');
                            if (commaIndex >= 0)
                                base64Data = base64Data.Substring(commaIndex + 1);
                            byte[] bytes = System.Convert.FromBase64String(base64Data);
                            var ms = new System.IO.MemoryStream(bytes);
                            var bitmap = new BitmapImage();
                            bitmap.BeginInit();
                            bitmap.StreamSource = ms;
                            bitmap.CacheOption = BitmapCacheOption.OnLoad;
                            bitmap.EndInit();
                            img.Source = bitmap;
                        }
                        else
                        {
                            var bitmap = new BitmapImage(new Uri(url, UriKind.RelativeOrAbsolute));
                            img.Source = bitmap;
                        }
                    }
                    catch { }

                    return new InlineUIContainer(img);
                }),
                // [链接](url)
                (@"\[([^\]]+)\]\(([^)]+)\)", (m, s, l) =>
                {
                    var link = new Hyperlink(new Run(m.Groups[1].Value))
                    {
                        NavigateUri = new Uri(m.Groups[2].Value, UriKind.RelativeOrAbsolute),
                        IsEnabled = true
                    };
                    link.RequestNavigate += (sender, args) =>
                    {
                        try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(args.Uri.ToString()) { UseShellExecute = true }); }
                        catch { }
                    };
                    return link;
                })
            };

            int pos = 0;
            while (pos < text.Length)
            {
                // 在每个位置依次尝试所有模式
                bool matched = false;
                foreach (var (pattern, factory) in patterns)
                {
                    var m = Regex.Match(text.Substring(pos), pattern);
                    if (m.Success && m.Index == 0)
                    {
                        inlines.Add(factory(m, pos, text.Length));
                        pos += m.Length;
                        matched = true;
                        break;
                    }
                }
                if (matched) continue;

                // 当前字符是普通文本
                int nextSpecial = text.Length;
                foreach (var (pattern, _) in patterns)
                {
                    var m = Regex.Match(text.Substring(pos), pattern);
                    if (m.Success && m.Index > 0)
                        nextSpecial = Math.Min(nextSpecial, pos + m.Index);
                }

                // 取到行尾或下一个特殊字符
                int take = Math.Min(nextSpecial - pos, text.Length - pos);
                if (take > 0)
                {
                    string plain = text.Substring(pos, take);
                    // 硬换行
                    if (plain.EndsWith("  \n"))
                        plain = plain.Substring(0, plain.Length - 3);
                    inlines.Add(new Run(plain));
                    pos += take;
                }
                else
                {
                    pos++;
                }
            }

            return inlines;
        }
    }
}
