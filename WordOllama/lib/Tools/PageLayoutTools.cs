using Newtonsoft.Json.Linq;
using System;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.lib.Tools
{
    /// <summary>
    /// 页面设置 — 页边距、纸张方向、纸张大小等。
    /// 满足公文国标页面要求。
    /// </summary>
    public class PageSetupTool : IAgentTool
    {
        public string Name => "page_setup";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => true;
        public bool IsConcurrencySafe => false;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["margin_top"] = new JObject { ["type"] = "number" },
                ["margin_bottom"] = new JObject { ["type"] = "number" },
                ["margin_left"] = new JObject { ["type"] = "number" },
                ["margin_right"] = new JObject { ["type"] = "number" },
                ["orientation"] = new JObject { ["type"] = "string" }
            },
            ["required"] = new JArray()
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            try
            {
                Word.PageSetup ps = doc.PageSetup;
                var changes = new System.Collections.Generic.List<string>();

                // 各单位：传入的是厘米，Word COM 需要磅 (1cm = 28.35pt)
                if (parameters["margin_top"] != null)
                {
                    ps.TopMargin = CentimetersToPoints(parameters["margin_top"].Value<double>());
                    changes.Add(AgentToolText.F("I18n_AgentResultMarginTop", parameters["margin_top"].Value<double>()));
                }
                if (parameters["margin_bottom"] != null)
                {
                    ps.BottomMargin = CentimetersToPoints(parameters["margin_bottom"].Value<double>());
                    changes.Add(AgentToolText.F("I18n_AgentResultMarginBottom", parameters["margin_bottom"].Value<double>()));
                }
                if (parameters["margin_left"] != null)
                {
                    ps.LeftMargin = CentimetersToPoints(parameters["margin_left"].Value<double>());
                    changes.Add(AgentToolText.F("I18n_AgentResultMarginLeft", parameters["margin_left"].Value<double>()));
                }
                if (parameters["margin_right"] != null)
                {
                    ps.RightMargin = CentimetersToPoints(parameters["margin_right"].Value<double>());
                    changes.Add(AgentToolText.F("I18n_AgentResultMarginRight", parameters["margin_right"].Value<double>()));
                }

                string orientation = parameters["orientation"]?.ToString()?.ToLower();
                if (orientation == "landscape")
                {
                    ps.Orientation = Word.WdOrientation.wdOrientLandscape;
                    changes.Add(AgentToolText.T("I18n_AgentResultLandscape"));
                }
                else if (orientation == "portrait")
                {
                    ps.Orientation = Word.WdOrientation.wdOrientPortrait;
                    changes.Add(AgentToolText.T("I18n_AgentResultPortrait"));
                }

                return changes.Count > 0
                    ? AgentToolText.F("I18n_AgentResultPageSetupUpdated", string.Join(", ", changes))
                    : AgentToolText.T("I18n_AgentResultNoPageSetupParameters");
            }
            catch (Exception ex)
            {
                return AgentToolText.F("I18n_AgentResultPageSetupFailed", ex.Message);
            }
        }

        private static float CentimetersToPoints(double cm) => (float)(cm * 28.35);
    }

    /// <summary>
    /// 插入分页符。
    /// </summary>
    public class InsertPageBreakTool : IAgentTool
    {
        public string Name => "insert_page_break";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => true;
        public bool IsConcurrencySafe => false;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject(),
            ["required"] = new JArray()
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            try
            {
                app.Selection.InsertBreak(Word.WdBreakType.wdPageBreak);
                return AgentToolText.T("I18n_AgentResultPageBreakInserted");
            }
            catch (Exception ex)
            {
                return AgentToolText.F("I18n_AgentResultInsertPageBreakFailed", ex.Message);
            }
        }
    }

    /// <summary>
    /// 设置页眉/页脚/页码。
    /// </summary>
    public class HeaderFooterTool : IAgentTool
    {
        public string Name => "header_footer";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => true;
        public bool IsConcurrencySafe => false;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["element"] = new JObject { ["type"] = "string" },
                ["text"] = new JObject { ["type"] = "string" },
                ["alignment"] = new JObject { ["type"] = "string" }
            },
            ["required"] = new JArray("element")
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            string element = parameters["element"]?.ToString()?.ToLower() ?? "";
            string text = parameters["text"]?.ToString() ?? "";
            string alignment = parameters["alignment"]?.ToString()?.ToLower() ?? "center";

            try
            {
                // 判断节是否存在，不存在则创建
                if (doc.Sections.Count == 0)
                    return AgentToolText.T("I18n_AgentResultNoSections");

                Word.HeaderFooter hf;
                if (element == "header")
                    hf = doc.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                else if (element == "footer" || element == "page_number")
                    hf = doc.Sections[1].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                else
                    return AgentToolText.F("I18n_AgentResultUnsupportedHeaderFooterElement", element);

                hf.LinkToPrevious = false;
                Word.Range hfRange = hf.Range;

                // 设置对齐
                switch (alignment)
                {
                    case "left":
                        hfRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        break;
                    case "right":
                        hfRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                        break;
                    default:
                        hfRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        break;
                }

                if (element == "page_number")
                {
                    // 插入页码域
                    hfRange.Text = "";
                    doc.Fields.Add(hfRange, Word.WdFieldType.wdFieldEmpty, @"PAGE \* MERGEFORMAT", false);
                    hfRange.InsertAfter(" / ");
                    doc.Fields.Add(hfRange, Word.WdFieldType.wdFieldEmpty, @"NUMPAGES \* MERGEFORMAT", false);
                    return AgentToolText.T("I18n_AgentResultPageNumberInserted");
                }
                else
                {
                    hfRange.Text = text;
                    string elementName = element == "header" ? AgentToolText.T("I18n_AgentResultHeader") : AgentToolText.T("I18n_AgentResultFooter");
                    return AgentToolText.F("I18n_AgentResultHeaderFooterSet", elementName, text);
                }
            }
            catch (Exception ex)
            {
                return AgentToolText.F("I18n_AgentResultHeaderFooterFailed", element, ex.Message);
            }
        }
    }
}
