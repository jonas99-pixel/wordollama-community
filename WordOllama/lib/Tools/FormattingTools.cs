using Newtonsoft.Json.Linq;
using System;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.lib.Tools
{
    /// <summary>
    /// 段落格式化 — 对齐、缩进、行距、段间距。
    /// 满足公文格式规范：首行缩进2字符、固定行距等。
    /// </summary>
    public class FormatParagraphTool : IAgentTool
    {
        public string Name => "format_paragraph";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => true;
        public bool IsConcurrencySafe => false;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["paragraph_index"] = new JObject { ["type"] = "integer" },
                ["alignment"] = new JObject { ["type"] = "string" },
                ["first_line_indent_chars"] = new JObject { ["type"] = "number" },
                ["line_spacing"] = new JObject { ["type"] = "number" },
                ["space_before"] = new JObject { ["type"] = "number" },
                ["space_after"] = new JObject { ["type"] = "number" }
            },
            ["required"] = new JArray("paragraph_index")
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            int index = parameters["paragraph_index"]?.Value<int>() ?? 1;
            if (index < 1 || index > doc.Paragraphs.Count)
                return AgentToolText.F("I18n_AgentResultParagraphIndexOutOfRange", index, doc.Paragraphs.Count);

            try
            {
                Word.Paragraph para = doc.Paragraphs[index];
                var fmt = para.Format;
                var changes = new System.Collections.Generic.List<string>();

                // 对齐方式
                string alignment = parameters["alignment"]?.ToString()?.ToLower();
                if (!string.IsNullOrEmpty(alignment))
                {
                    switch (alignment)
                    {
                        case "left":
                            fmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                            changes.Add(AgentToolText.T("I18n_AgentResultAlignLeft"));
                            break;
                        case "center":
                            fmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            changes.Add(AgentToolText.T("I18n_AgentResultAlignCenter"));
                            break;
                        case "right":
                            fmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                            changes.Add(AgentToolText.T("I18n_AgentResultAlignRight"));
                            break;
                        case "justify":
                            fmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                            changes.Add(AgentToolText.T("I18n_AgentResultAlignJustify"));
                            break;
                    }
                }

                // 首行缩进（字符数 × 当前字号 ≈ 磅值）
                if (parameters["first_line_indent_chars"] != null)
                {
                    double chars = parameters["first_line_indent_chars"].Value<double>();
                    if (chars > 0)
                    {
                        float fontSize = para.Range.Font.Size > 0 ? para.Range.Font.Size : 12;
                        fmt.FirstLineIndent = (float)(chars * fontSize);
                        changes.Add(AgentToolText.F("I18n_AgentResultFirstLineIndent", chars));
                    }
                    else
                    {
                        fmt.FirstLineIndent = 0;
                        changes.Add(AgentToolText.T("I18n_AgentResultFirstLineIndentCleared"));
                    }
                }

                // 行距
                if (parameters["line_spacing"] != null)
                {
                    float spacing = parameters["line_spacing"].Value<float>();
                    if (spacing > 0)
                    {
                        fmt.LineSpacing = spacing;
                        changes.Add(AgentToolText.F("I18n_AgentResultLineSpacing", spacing));
                    }
                    else
                    {
                        fmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                        changes.Add(AgentToolText.T("I18n_AgentResultSingleLineSpacing"));
                    }
                }

                // 段前间距
                if (parameters["space_before"] != null)
                {
                    fmt.SpaceBefore = parameters["space_before"].Value<float>();
                    changes.Add(AgentToolText.F("I18n_AgentResultSpaceBefore", fmt.SpaceBefore));
                }

                // 段后间距
                if (parameters["space_after"] != null)
                {
                    fmt.SpaceAfter = parameters["space_after"].Value<float>();
                    changes.Add(AgentToolText.F("I18n_AgentResultSpaceAfter", fmt.SpaceAfter));
                }

                DocumentAgent.FocusOnRange(para.Range, app);
                return AgentToolText.F("I18n_AgentResultParagraphFormatUpdated", index, string.Join(", ", changes));
            }
            catch (Exception ex)
            {
                return AgentToolText.F("I18n_AgentResultFormatParagraphFailed", ex.Message);
            }
        }
    }

    /// <summary>
    /// 字体格式化 — 设置字体、字号、加粗、斜体、下划线、颜色。
    /// </summary>
    public class FormatTextTool : IAgentTool
    {
        public string Name => "format_text";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => true;
        public bool IsConcurrencySafe => false;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["paragraph_index"] = new JObject { ["type"] = "integer" },
                ["text_to_format"] = new JObject { ["type"] = "string" },
                ["font_name"] = new JObject { ["type"] = "string" },
                ["font_size"] = new JObject { ["type"] = "number" },
                ["bold"] = new JObject { ["type"] = "boolean" },
                ["italic"] = new JObject { ["type"] = "boolean" },
                ["underline"] = new JObject { ["type"] = "boolean" },
                ["color"] = new JObject { ["type"] = "string" }
            },
            ["required"] = new JArray("paragraph_index")
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            int index = parameters["paragraph_index"]?.Value<int>() ?? 1;
            if (index < 1 || index > doc.Paragraphs.Count)
                return AgentToolText.F("I18n_AgentResultParagraphIndexOutOfRange", index, doc.Paragraphs.Count);

            try
            {
                string specificText = parameters["text_to_format"]?.ToString() ?? "";
                Word.Range targetRange;

                if (!string.IsNullOrEmpty(specificText))
                {
                    // 在段落中查找特定文字
                    Word.Range paraRange = doc.Paragraphs[index].Range;
                    paraRange.Find.ClearFormatting();
                    paraRange.Find.Text = specificText;
                    paraRange.Find.Forward = true;
                    if (!paraRange.Find.Execute())
                        return AgentToolText.F("I18n_AgentResultTextNotFoundInParagraph", index, specificText);

                    targetRange = paraRange.Duplicate;
                }
                else
                {
                    targetRange = doc.Paragraphs[index].Range;
                }

                Word.Font font = targetRange.Font;
                var changes = new System.Collections.Generic.List<string>();

                // 字体
                string fontName = parameters["font_name"]?.ToString();
                if (!string.IsNullOrEmpty(fontName))
                {
                    font.Name = fontName;
                    changes.Add(AgentToolText.F("I18n_AgentResultFontName", fontName));
                }

                // 字号
                if (parameters["font_size"] != null)
                {
                    font.Size = parameters["font_size"].Value<float>();
                    changes.Add(AgentToolText.F("I18n_AgentResultFontSize", font.Size));
                }

                // 加粗
                if (parameters["bold"] != null)
                {
                    font.Bold = parameters["bold"].Value<int>() == 1 ? 1 : 0;
                    changes.Add(parameters["bold"].Value<bool>() ? AgentToolText.T("I18n_AgentResultBold") : AgentToolText.T("I18n_AgentResultBoldCleared"));
                }

                // 斜体
                if (parameters["italic"] != null)
                {
                    font.Italic = parameters["italic"].Value<int>() == 1 ? 1 : 0;
                    changes.Add(parameters["italic"].Value<bool>() ? AgentToolText.T("I18n_AgentResultItalic") : AgentToolText.T("I18n_AgentResultItalicCleared"));
                }

                // 下划线
                if (parameters["underline"] != null)
                {
                    font.Underline = parameters["underline"].Value<bool>()
                        ? Word.WdUnderline.wdUnderlineSingle
                        : Word.WdUnderline.wdUnderlineNone;
                    changes.Add(parameters["underline"].Value<bool>() ? AgentToolText.T("I18n_AgentResultUnderline") : AgentToolText.T("I18n_AgentResultUnderlineCleared"));
                }

                // 颜色
                string color = parameters["color"]?.ToString()?.ToLower();
                if (!string.IsNullOrEmpty(color))
                {
                    font.Color = ParseColor(color);
                    changes.Add(AgentToolText.F("I18n_AgentResultFontColor", color));
                }

                string target = string.IsNullOrEmpty(specificText)
                    ? AgentToolText.T("I18n_AgentResultWholeParagraph")
                    : AgentToolText.F("I18n_AgentResultQuotedText", specificText);
                DocumentAgent.FocusOnRange(doc.Paragraphs[index].Range, app);
                return AgentToolText.F("I18n_AgentResultTextFormatUpdated", index, target, string.Join(", ", changes));
            }
            catch (Exception ex)
            {
                return AgentToolText.F("I18n_AgentResultFormatTextFailed", ex.Message);
            }
        }

        private static Word.WdColor ParseColor(string color)
        {
            // 十六进制格式 #RRGGBB
            if (color.StartsWith("#") && color.Length == 7)
            {
                try
                {
                    int r = Convert.ToInt32(color.Substring(1, 2), 16);
                    int g = Convert.ToInt32(color.Substring(3, 2), 16);
                    int b = Convert.ToInt32(color.Substring(5, 2), 16);
                    return (Word.WdColor)(r + (g << 8) + (b << 16));
                }
                catch { }
            }

            return color switch
            {
                "red" => Word.WdColor.wdColorRed,
                "blue" => Word.WdColor.wdColorBlue,
                "green" => Word.WdColor.wdColorGreen,
                "black" => Word.WdColor.wdColorBlack,
                "white" => Word.WdColor.wdColorWhite,
                "yellow" => Word.WdColor.wdColorYellow,
                "gray" or "grey" => Word.WdColor.wdColorGray50,
                "darkred" => Word.WdColor.wdColorDarkRed,
                "darkblue" => Word.WdColor.wdColorDarkBlue,
                "darkgreen" => Word.WdColor.wdColorDarkGreen,
                "orange" => Word.WdColor.wdColorOrange,
                "pink" => Word.WdColor.wdColorPink,
                _ => Word.WdColor.wdColorBlack
            };
        }
    }

    /// <summary>
    /// 编号/项目符号列表 — 设置多级编号或项目符号。
    /// 法律公文常用：一、(一) 1. (1) ① 等多级编号。
    /// </summary>
    public class FormatListTool : IAgentTool
    {
        public string Name => "format_list";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => true;
        public bool IsConcurrencySafe => false;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["paragraph_index"] = new JObject { ["type"] = "integer" },
                ["range_end"] = new JObject { ["type"] = "integer" },
                ["list_type"] = new JObject { ["type"] = "string" }
            },
            ["required"] = new JArray("paragraph_index", "list_type")
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            int startIndex = parameters["paragraph_index"]?.Value<int>() ?? 1;
            int endIndex = parameters["range_end"]?.Value<int>() ?? startIndex;
            string listType = parameters["list_type"]?.ToString()?.ToLower() ?? "number";

            if (startIndex < 1 || startIndex > doc.Paragraphs.Count)
                return AgentToolText.F("I18n_AgentResultParagraphIndexOutOfRangeSimple", startIndex);

            try
            {
                if (listType == "none")
                {
                    for (int i = startIndex; i <= endIndex && i <= doc.Paragraphs.Count; i++)
                    {
                        doc.Paragraphs[i].Range.ListFormat.RemoveNumbers();
                    }
                    return AgentToolText.F("I18n_AgentResultListRemoved", startIndex, endIndex);
                }

                Word.Range rng = doc.Range(
                    doc.Paragraphs[startIndex].Range.Start,
                    doc.Paragraphs[Math.Min(endIndex, doc.Paragraphs.Count)].Range.End);
                Word.WdListType wdType;

                switch (listType)
                {
                    case "bullet":
                        wdType = Word.WdListType.wdListBullet;
                        break;
                    case "number":
                        wdType = Word.WdListType.wdListSimpleNumbering;
                        break;
                    case "legal":
                        wdType = Word.WdListType.wdListListNumOnly;
                        break;
                    default:
                        return AgentToolText.F("I18n_AgentResultUnsupportedListType", listType);
                }

                rng.ListFormat.ApplyListTemplateWithLevel(
                    ListTemplate: app.ListGalleries[wdType].ListTemplates[1],
                    ContinuePreviousList: true,
                    ApplyTo: Word.WdListApplyTo.wdListApplyToWholeList);

                string typeName = listType switch
                {
                    "bullet" => AgentToolText.T("I18n_AgentResultListBullet"),
                    "number" => AgentToolText.T("I18n_AgentResultListNumber"),
                    "legal" => AgentToolText.T("I18n_AgentResultListLegal"),
                    _ => listType
                };
                return AgentToolText.F("I18n_AgentResultListApplied", startIndex, endIndex, typeName);
            }
            catch (Exception ex)
            {
                return AgentToolText.F("I18n_AgentResultFormatListFailed", ex.Message);
            }
        }
    }
}
