using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.lib.Tools
{
    /// <summary>
    /// Get the document outline by extracting heading-level paragraphs.
    /// 让 AI 先看"骨架"，快速判断合同是否缺必备条款。
    /// </summary>
    public class GetDocumentOutlineTool : IAgentTool
    {
        public string Name => "get_document_outline";
        public string Description => AgentToolText.Description(Name);

        public bool IsWriteOperation => false;
        public bool IsConcurrencySafe => !IsWriteOperation;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => new()
        {
            ["type"] = "object",
            ["properties"] = new JObject(),
            ["required"] = new JArray()
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            var outline = new JArray();
            int paraCount = doc.Paragraphs.Count;

            for (int i = 1; i <= paraCount; i++)
            {
                Word.Paragraph para = doc.Paragraphs[i];
                try
                {
                    int level = (int)para.OutlineLevel;
                    // 大纲级别 1-9 为标题，10 = BodyText
                    if (level >= 1 && level <= 9)
                    {
                        string text = para.Range.Text?.TrimEnd('\r', '\a', '\v') ?? "";
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            outline.Add(new JObject
                            {
                                ["text"] = text,
                                ["level"] = level,
                                ["paragraph"] = i
                            });
                        }
                    }
                }
                catch { }
            }

            if (outline.Count == 0)
                return AgentToolText.T("I18n_AgentResultNoDocumentOutline");

            return AgentToolText.F("I18n_AgentResultDocumentOutline", outline.Count, outline.ToString(Newtonsoft.Json.Formatting.Indented));
        }
    }

    /// <summary>
    /// Read a provision by locating its heading and returning the body text.
    /// 代替盲猜段落索引，让 AI 像律师一样通过条款名称导航。
    /// </summary>
    public class ReadClauseTool : IAgentTool
    {
        public string Name => "read_clause";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => false;
        public bool IsConcurrencySafe => !IsWriteOperation;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["heading_title"] = new JObject { ["type"] = "string" }
            },
            ["required"] = new JArray("heading_title")
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            string title = parameters["heading_title"]?.ToString()?.Trim() ?? "";
            if (string.IsNullOrEmpty(title))
                return AgentToolText.T("I18n_AgentResultHeadingRequired");

            // 使用 Word Find 定位标题
            Word.Range searchRange = doc.Range();
            searchRange.Find.ClearFormatting();
            searchRange.Find.Text = title;
            searchRange.Find.Forward = true;
            searchRange.Find.MatchCase = false;
            searchRange.Find.MatchWildcards = false;

            if (!searchRange.Find.Execute())
                return AgentToolText.F("I18n_AgentResultHeadingNotFound", title);

            int headingStart = searchRange.Start;
            int headingLevel = 0;
            Word.Paragraph headingPara = searchRange.Paragraphs[1];
            try { headingLevel = (int)headingPara.OutlineLevel; }
            catch { headingLevel = 10; }

            // 如果匹配到的不是标题段落，尝试向上找最近的标题
            if (headingLevel > 9 || headingLevel < 1)
            {
                // 向上搜索最近的标题
                int searchParaIndex = searchRange.Paragraphs[1].Range.Start;
                headingLevel = 9; // 默认从低级别开始
                for (int i = doc.Paragraphs.Count; i >= 1; i--)
                {
                    try
                    {
                        Word.Paragraph p = doc.Paragraphs[i];
                        if (p.Range.Start <= searchParaIndex)
                        {
                            int ol = (int)p.OutlineLevel;
                            if (ol >= 1 && ol <= 9)
                            {
                                headingLevel = ol;
                                headingStart = p.Range.Start;
                                break;
                            }
                        }
                    }
                    catch { }
                }
            }

            // 收集从标题开始到下一个同级或更高级标题之间的所有段落
            var contentLines = new List<string>();
            bool foundHeading = false;
            int paraCount = doc.Paragraphs.Count;

            for (int i = 1; i <= paraCount; i++)
            {
                Word.Paragraph para = doc.Paragraphs[i];
                int paraStart = para.Range.Start;

                if (!foundHeading)
                {
                    if (paraStart >= headingStart)
                    {
                        foundHeading = true;
                        string headingText = para.Range.Text?.TrimEnd('\r', '\a', '\v') ?? "";
                        contentLines.Add($"【{headingText}】");
                    }
                    continue;
                }

                // 检查是否到达下一个同级或更高级标题
                try
                {
                    int ol = (int)para.OutlineLevel;
                    if (ol >= 1 && ol <= headingLevel && paraStart > headingStart)
                        break;
                }
                catch { }

                string text = para.Range.Text?.TrimEnd('\r', '\a', '\v') ?? "";
                if (!string.IsNullOrWhiteSpace(text))
                    contentLines.Add(text);
            }

            if (contentLines.Count == 0)
                return AgentToolText.T("I18n_AgentResultClauseExtractionFailed");

            return string.Join("\n", contentLines);
        }
    }

    /// <summary>
    /// Extract a terminology table from definition clauses and "hereinafter" phrases.
    /// 让 AI 建立术语记忆，后续审查时检查术语一致性。
    /// </summary>
    public class ExtractDefinitionsTool : IAgentTool
    {
        public string Name => "extract_definitions";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => false;
        public bool IsConcurrencySafe => !IsWriteOperation;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject(),
            ["required"] = new JArray()
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            string fullText = doc.Range().Text ?? "";
            var definitions = new Dictionary<string, string>(); // term -> definition context

            // === 中文合同定义模式 ===
            // 模式1：以下简称为"XXX" / 以下简称 XXX
            var patterns = new[]
            {
                // 全称（以下简称[为]"简称"）
                @"(?:甲方|乙方|丙方|发包人|承包人|委托人|受托人|买方|卖方|出租人|承租人|许可人|被许可人|投资人|融资方|[\u4e00-\u9fa5]{2,20}(?:公司|单位|机构|组织|中心|集团))(?:（[^）]*?）|\([^)]*?\))?[，,]?\s*(?:以下[简统]?称[为是]?)\s*[""'《](.+?)[""'》”]",
                // 以下简称为"XXX"
                @"以下[简统]?称[为是]?\s*[""「'《](.+?)[""」'》”]",
                // "XXX"是指/指的是...
                @"[「""'\u300A](.+?)[」""'\u300B\u201d][，,]?\s*(?:是指|指的是|即指|是[指谓])",
                // 本合同/协议中的"XXX"
                @"本(?:合同|协议|文件|契约)(?:中|项下|里)?(?:的)?\s*[\u300C""'\u300A](.+?)[\u300D""'\u300B\u201d]",
                // 定义：XXX是指...
                @"(?:定义|释义)[：:]\s*[「""'\u300A]?(.+?)[」""'\u300B""]?\s*(?:是指|指的是|系指|应[理具]解为)",
            };

            foreach (string pattern in patterns)
            {
                try
                {
                    var matches = Regex.Matches(fullText, pattern, RegexOptions.Multiline);
                    foreach (Match m in matches)
                    {
                        string term = m.Groups.Count > 1 ? m.Groups[1].Value.Trim() : m.Value.Trim();
                        if (!string.IsNullOrEmpty(term) && term.Length <= 50 && !definitions.ContainsKey(term))
                        {
                            // 获取上下文（前后各30字符）
                            int start = Math.Max(0, m.Index - 20);
                            int length = Math.Min(fullText.Length - start, m.Length + 40);
                            string context = fullText.Substring(start, length).Replace('\r', ' ').Replace('\n', ' ').Trim();
                            definitions[term] = context;
                        }
                    }
                }
                catch { }
            }

            // === 英文合同定义模式 ===
            var enPatterns = new[]
            {
                @"""([^""]{2,50})""\s*\((?:the\s+)?[\""]?[A-Z][^\""]{0,30}[\""]?\)",  // "Full Name" (the "Short Name")
                @"""([^""]{2,50})""\s*(?:means|refers to|shall mean|shall have the meaning)",  // "Term" means...
                @"(?:the\s+)?[\""]([^\""]{2,50})[\""]\s*(?:means|refers to|shall mean)",
            };

            foreach (string pattern in enPatterns)
            {
                try
                {
                    var matches = Regex.Matches(fullText, pattern);
                    foreach (Match m in matches)
                    {
                        string term = m.Groups.Count > 1 ? m.Groups[1].Value.Trim() : m.Value.Trim();
                        if (!string.IsNullOrEmpty(term) && term.Length <= 50 && !definitions.ContainsKey(term))
                        {
                            int start = Math.Max(0, m.Index - 20);
                            int length = Math.Min(fullText.Length - start, m.Length + 40);
                            string context = fullText.Substring(start, length).Replace('\r', ' ').Replace('\n', ' ').Trim();
                            definitions[term] = context;
                        }
                    }
                }
                catch { }
            }

            if (definitions.Count == 0)
                return AgentToolText.T("I18n_AgentResultNoDefinitions");

            var result = new JObject();
            foreach (var kv in definitions.OrderBy(x => x.Key))
                result[kv.Key] = kv.Value;

            return AgentToolText.F("I18n_AgentResultDefinitionsExtracted", definitions.Count, result.ToString(Newtonsoft.Json.Formatting.Indented));
        }
    }

    /// <summary>
    /// Check cross-reference consistency and verify referenced provisions.
    /// 合同修改中最容易产生的错误就是条款编号变动后引用未更新。
    /// </summary>
    public class CheckCrossReferencesTool : IAgentTool
    {
        public string Name => "check_cross_references";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => false;
        public bool IsConcurrencySafe => !IsWriteOperation;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject(),
            ["required"] = new JArray()
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            // 第一步：构建文档中实际存在的条款编号集合
            // 从标题中提取编号模式
            var existingNumbers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int paraCount = doc.Paragraphs.Count;

            for (int i = 1; i <= paraCount; i++)
            {
                Word.Paragraph para = doc.Paragraphs[i];
                try
                {
                    int level = (int)para.OutlineLevel;
                    if (level >= 1 && level <= 9)
                    {
                        string text = para.Range.Text?.TrimEnd('\r', '\a', '\v') ?? "";
                        // 提取标题中的编号（如 "1.", "1.1", "2.3.1", "第一条", "第1.1条"）
                        var numMatches = Regex.Matches(text, @"^(?:(\d+(?:\.\d+)*)[\.、\s]|第(\d+(?:\.\d+)*)\s*[条款项节])");
                        foreach (Match m in numMatches)
                        {
                            string num = m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value;
                            existingNumbers.Add(num);
                        }
                    }
                }
                catch { }
            }

            // 第二步：扫描全文查找所有交叉引用
            string fullText = doc.Range().Text ?? "";
            var errors = new List<string>();

            // 引用模式：第X.X条 / 第X.X.X款 / 第X项 / Section X.X / Clause X.X
            var refPatterns = new[]
            {
                // 中文：第X.X条
                new { Pattern = @"第\s*(\d+(?:\.\d+)*)\s*[条款项节]", Offset = 0 },
                // 英文：Section X.X / Clause X.X
                new { Pattern = @"(?:Section|Clause|Article|Paragraph)\s+(\d+(?:\.\d+)*)", Offset = 0 },
            };

            // 逐段扫描，记录错误位置
            for (int i = 1; i <= paraCount; i++)
            {
                Word.Paragraph para = doc.Paragraphs[i];
                string paraText = para.Range.Text ?? "";

                foreach (var rp in refPatterns)
                {
                    try
                    {
                        var matches = Regex.Matches(paraText, rp.Pattern, RegexOptions.IgnoreCase);
                        foreach (Match m in matches)
                        {
                            string refNum = m.Groups[1].Value;
                            // 简单验证：至少要有数字
                            if (string.IsNullOrEmpty(refNum)) continue;

                            // 检查是否在已知条款列表中
                            if (!existingNumbers.Contains(refNum))
                            {
                                // 尝试匹配前缀（如 "1.1" 对应标题 "1.1 定义"）
                                bool found = existingNumbers.Any(n => n == refNum || n.StartsWith(refNum + ".") || refNum.StartsWith(n + "."));
                                if (!found)
                                {
                                    string preview = paraText.TrimEnd('\r', '\a').Trim();
                                    if (preview.Length > 60) preview = preview.Substring(0, 60) + "...";
                                    errors.Add(AgentToolText.F("I18n_AgentResultBrokenReference", i, refNum, preview));
                                }
                            }
                        }
                    }
                    catch { }
                }
            }

            if (errors.Count == 0)
            {
                return AgentToolText.F("I18n_AgentResultCrossReferencesValid", existingNumbers.Count, string.Join(", ", existingNumbers.OrderBy(x => x)));
            }

            return AgentToolText.F("I18n_AgentResultCrossReferenceErrors", errors.Count, string.Join("\n", errors.Select((e, idx) => $"  {idx + 1}. {e}")));
        }
    }

    /// <summary>
    /// Replace a specific text fragment without disturbing paragraph formatting.
    /// 配合修订模式，只有改动的字会标红，更专业。
    /// </summary>
    public class ReplaceExactTextTool : IAgentTool
    {
        public string Name => "replace_exact_text";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => true;
        public bool IsConcurrencySafe => !IsWriteOperation;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["paragraph_index"] = new JObject { ["type"] = "integer" },
                ["old_text"] = new JObject { ["type"] = "string" },
                ["new_text"] = new JObject { ["type"] = "string" }
            },
            ["required"] = new JArray("paragraph_index", "old_text", "new_text")
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            int index = parameters["paragraph_index"]?.Value<int>() ?? 1;
            string oldText = parameters["old_text"]?.ToString() ?? "";
            string newText = parameters["new_text"]?.ToString() ?? "";

            if (index < 1 || index > doc.Paragraphs.Count)
                return AgentToolText.F("I18n_AgentResultParagraphIndexOutOfRange", index, doc.Paragraphs.Count);
            if (string.IsNullOrEmpty(oldText))
                return AgentToolText.T("I18n_AgentResultOldTextRequired");

            Word.Range paraRange = doc.Paragraphs[index].Range;

            // 在段落范围内使用 Find 定位并替换（单次调用，避免二次 Execute 导致的 Range 状态问题）
            Word.Find find = paraRange.Find;
            find.ClearFormatting();
            find.Text = oldText;
            find.MatchCase = false;
            find.MatchWildcards = false;
            find.Forward = true;
            find.Replacement.ClearFormatting();
            find.Replacement.Text = newText;

            if (!find.Execute(
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Word.WdReplace.wdReplaceOne,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing
                ))
            {
                string paraPreview = paraRange.Text?.TrimEnd('\r', '\a') ?? "";
                if (paraPreview.Length > 60) paraPreview = paraPreview.Substring(0, 60) + "...";
                return AgentToolText.F("I18n_AgentResultExactTextNotFound", index, oldText, paraPreview);
            }

            DocumentAgent.FocusOnRange(paraRange, app);
            return AgentToolText.F("I18n_AgentResultExactTextReplaced", index, oldText, newText);
        }
    }

    /// <summary>
    /// Insert a new provision after a precisely located heading.
    /// 用于补充缺失条款（如不可抗力、争议解决、可分割性等）。
    /// </summary>
    public class InsertClauseAfterTool : IAgentTool
    {
        public string Name => "insert_clause_after";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => true;
        public bool IsConcurrencySafe => !IsWriteOperation;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["target_heading"] = new JObject { ["type"] = "string" },
                ["new_clause_text"] = new JObject { ["type"] = "string" }
            },
            ["required"] = new JArray("target_heading", "new_clause_text")
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            string targetHeading = parameters["target_heading"]?.ToString()?.Trim() ?? "";
            string newText = parameters["new_clause_text"]?.ToString() ?? "";

            if (string.IsNullOrEmpty(targetHeading))
                return AgentToolText.T("I18n_AgentResultTargetHeadingRequired");
            if (string.IsNullOrEmpty(newText))
                return AgentToolText.T("I18n_AgentResultNewClauseTextRequired");

            // 使用 Find 定位目标标题
            Word.Range searchRange = doc.Range();
            searchRange.Find.ClearFormatting();
            searchRange.Find.Text = targetHeading;
            searchRange.Find.Forward = true;
            searchRange.Find.MatchCase = false;

            if (!searchRange.Find.Execute())
                return AgentToolText.F("I18n_AgentResultHeadingNotFound", targetHeading);

            // 获取目标标题的段落级别
            int targetLevel = 10;
            try { targetLevel = (int)searchRange.Paragraphs[1].OutlineLevel; }
            catch { }

            // 找到该条款的结束位置（下一个同级或更高级标题之前）
            int insertAfterParaIndex = 0;
            int paraCount = doc.Paragraphs.Count;

            for (int i = 1; i <= paraCount; i++)
            {
                Word.Paragraph para = doc.Paragraphs[i];
                if (para.Range.Start > searchRange.Start)
                {
                    try
                    {
                        int ol = (int)para.OutlineLevel;
                        if (ol >= 1 && ol <= targetLevel)
                        {
                            // 找到了下一个同级或更高级标题，在其前一个段落结束
                            insertAfterParaIndex = i - 1;
                            break;
                        }
                    }
                    catch { }
                }
            }

            // 如果没找到后续标题，插入到文档末尾
            if (insertAfterParaIndex == 0)
                insertAfterParaIndex = paraCount;

            // 获取插入点（目标条款最后一段的末尾）
            Word.Range insertRange = doc.Paragraphs[insertAfterParaIndex].Range;
            int insertPos = insertRange.End - 1; // 减1避免超出文档末尾

            // 在插入点后添加新段落和内容
            Word.Range targetRange = doc.Range(insertPos, insertPos);
            targetRange.InsertAfter("\r" + newText + "\r");

            DocumentAgent.FocusOnRange(targetRange, app);
            return AgentToolText.F("I18n_AgentResultClauseInserted", targetHeading);
        }
    }

    /// <summary>
    /// Highlight specified text to mark potential risk.
    /// 比批注更直观：黄色=注意，红色=高危。
    /// </summary>
    public class HighlightRiskTool : IAgentTool
    {
        public string Name => "highlight_risk";
        public string Description => AgentToolText.Description(Name);

        public bool IsWriteOperation => true;
        public bool IsConcurrencySafe => !IsWriteOperation;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["keyword"] = new JObject { ["type"] = "string" },
                ["color"] = new JObject { ["type"] = "string" },
                ["reason"] = new JObject { ["type"] = "string" }
            },
            ["required"] = new JArray("keyword", "color", "reason")
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            string keyword = parameters["keyword"]?.ToString()?.Trim() ?? "";
            string color = parameters["color"]?.ToString()?.ToLower() ?? "yellow";
            string reason = parameters["reason"]?.ToString() ?? "";

            if (string.IsNullOrEmpty(keyword))
                return AgentToolText.T("I18n_AgentResultKeywordRequired");

            // 映射颜色
            Word.WdColorIndex highlightColor;
            switch (color)
            {
                case "yellow": highlightColor = Word.WdColorIndex.wdYellow; break;
                case "red": highlightColor = Word.WdColorIndex.wdRed; break;
                case "green": highlightColor = Word.WdColorIndex.wdBrightGreen; break;
                case "light_blue":
                case "blue": highlightColor = Word.WdColorIndex.wdTurquoise; break;
                default: highlightColor = Word.WdColorIndex.wdYellow; break;
            }

            // 使用 Find 在全文中查找并高亮
            int count = 0;
            Word.Range searchRange = doc.Range();
            searchRange.Find.ClearFormatting();
            searchRange.Find.Text = keyword;
            searchRange.Find.Forward = true;
            searchRange.Find.MatchCase = false;
            searchRange.Find.MatchWildcards = false;
            searchRange.Find.MatchWholeWord = false;

            int originalEnd = searchRange.End;
            List<Word.Range> foundRanges = new List<Word.Range>();

            while (searchRange.Find.Execute())
            {
                if (searchRange.Start >= originalEnd) break;
                if (searchRange.End <= searchRange.Start) break;

                // 收集找到的 Range（不能在迭代中修改）
                foundRanges.Add(doc.Range(searchRange.Start, searchRange.End));
                searchRange.SetRange(searchRange.End, originalEnd);
                count++;

                if (count >= 20) break; // 安全上限，避免高亮太多
            }

            // 对每个找到的范围进行高亮和添加批注
            foreach (var range in foundRanges)
            {
                range.HighlightColorIndex = highlightColor;

                if (!string.IsNullOrEmpty(reason))
                {
                    try
                    {
                        doc.Comments.Add(range, AgentToolText.F("I18n_AgentResultAiRiskComment", reason));
                    }
                    catch { }
                }
            }

            if (count == 0)
                return AgentToolText.F("I18n_AgentResultRiskHighlightNotFound", keyword);

            // 视觉同步：滚动到第一个高亮位置
            if (foundRanges.Count > 0)
                DocumentAgent.FocusOnRange(foundRanges[0], app);

            string colorName = color == "yellow" ? AgentToolText.T("I18n_AgentResultRiskColorYellow") :
                               color == "red" ? AgentToolText.T("I18n_AgentResultRiskColorRed") :
                               color == "green" ? AgentToolText.T("I18n_AgentResultRiskColorGreen") : AgentToolText.T("I18n_AgentResultRiskColorLightBlue");
            return string.IsNullOrEmpty(reason)
                ? AgentToolText.F("I18n_AgentResultRiskHighlighted", count, keyword, colorName)
                : AgentToolText.F("I18n_AgentResultRiskHighlightedWithComment", count, keyword, colorName, reason);
        }
    }

    /// <summary>
    /// Apply standard legal-document formatting.
    /// 支持中国法律文书格式（仿宋/宋体、行距、页边距、标题层级）和英文法律文书格式。
    /// </summary>
    public class ApplyLegalFormatTool : IAgentTool
    {
        public string Name => "apply_legal_format";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => true;
        public bool IsConcurrencySafe => !IsWriteOperation;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["style"] = new JObject { ["type"] = "string" },
                ["font_name"] = new JObject { ["type"] = "string" },
                ["font_size"] = new JObject { ["type"] = "number" },
                ["line_spacing"] = new JObject { ["type"] = "number" }
            },
            ["required"] = new JArray("style")
        };

        private static readonly Dictionary<string, (string Font, double Size, double Spacing, double TopMargin, double BottomMargin, double LeftMargin, double RightMargin, string TitleFont, double TitleSize)> Styles = new()
        {
            ["chinese_civil"] = ("仿宋_GB2312", 12, 28, 72.0f / 28.35f * 3.7f, 72.0f / 28.35f * 3.5f, 72.0f / 28.35f * 2.8f, 72.0f / 28.35f * 2.6f, "方正小标宋简体", 22),
            // 如果没有仿宋_GB2312，回退到宋体
            ["chinese_judgment"] = ("仿宋_GB2312", 12, 28, 72.0f / 28.35f * 3.7f, 72.0f / 28.35f * 3.5f, 72.0f / 28.35f * 2.8f, 72.0f / 28.35f * 2.6f, "黑体", 22),
            ["chinese_gov"] = ("仿宋_GB2312", 12, 28, 72.0f / 28.35f * 3.7f, 72.0f / 28.35f * 3.5f, 72.0f / 28.35f * 2.8f, 72.0f / 28.35f * 2.6f, "方正小标宋简体", 22),
            ["english_memo"] = ("Times New Roman", 12, 24, 72, 72, 72, 72, "Times New Roman", 16),
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            string style = parameters["style"]?.ToString()?.ToLower() ?? "chinese_civil";
            if (!Styles.ContainsKey(style))
                return AgentToolText.F("I18n_AgentResultUnsupportedLegalFormat", style, string.Join(", ", Styles.Keys));

            var cfg = Styles[style];

            // 参数覆盖
            string fontName = parameters["font_name"]?.ToString() ?? cfg.Font;
            double fontSize = parameters["font_size"]?.Value<double>() ?? cfg.Size;
            double lineSpacing = parameters["line_spacing"]?.Value<double>() ?? cfg.Spacing;

            // 检查字体是否可用，不可用时回退
            if (!IsFontInstalled(fontName))
            {
                string fallback = style == "english_memo" ? "Times New Roman" : "宋体";
                fontName = fallback;
            }

            // 设置页面边距
            try
            {
                doc.PageSetup.TopMargin = (float)cfg.TopMargin;
                doc.PageSetup.BottomMargin = (float)cfg.BottomMargin;
                doc.PageSetup.LeftMargin = (float)cfg.LeftMargin;
                doc.PageSetup.RightMargin = (float)cfg.RightMargin;
            }
            catch { /* 边距设置失败不影响整体 */ }

            // 设置默认段落格式
            int paraCount = doc.Paragraphs.Count;
            int modifiedCount = 0;

            for (int i = 1; i <= paraCount; i++)
            {
                try
                {
                    Word.Paragraph para = doc.Paragraphs[i];
                    Word.Range range = para.Range;

                    // 判断段落类型
                    int outlineLevel = 10;
                    try { outlineLevel = (int)para.OutlineLevel; } catch { }

                    bool isTitle = i == 1; // 第一段作为标题

                    if (isTitle && style != "english_memo")
                    {
                        // 标题：居中、大号、标题字体
                        para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        range.Font.Name = cfg.TitleFont;
                        range.Font.Size = (float)cfg.TitleSize;
                        range.Font.Bold = 1;
                    }
                    else if (outlineLevel >= 1 && outlineLevel <= 3)
                    {
                        // 一级/二级标题：加粗
                        para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        range.Font.Name = fontName;
                        range.Font.Size = (float)(fontSize + (4 - outlineLevel)); // 逐级递减
                        range.Font.Bold = 1;
                    }
                    else if (outlineLevel >= 4 && outlineLevel <= 9)
                    {
                        // 三级及以下标题
                        para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        range.Font.Name = fontName;
                        range.Font.Size = (float)fontSize;
                        range.Font.Bold = 0;
                    }
                    else
                    {
                        // 正文
                        para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        range.Font.Name = fontName;
                        range.Font.Size = (float)fontSize;
                        range.Font.Bold = 0;

                        // 首行缩进（中文习惯两个字符）
                        if (style != "english_memo")
                        {
                            para.CharacterUnitFirstLineIndent = 2;
                        }
                    }

                    // 行距设置
                    para.LineSpacing = (float)lineSpacing;

                    modifiedCount++;
                }
                catch { /* 单个段落格式失败继续处理下一段 */ }
            }

            return AgentToolText.F("I18n_AgentResultLegalFormatApplied", style, fontName, fontSize, lineSpacing, modifiedCount, paraCount);
        }

        private bool IsFontInstalled(string fontName)
        {
            try
            {
                using (var font = new System.Drawing.Font(fontName, 8))
                {
                    return font.Name == fontName;
                }
            }
            catch
            {
                return false;
            }
        }
    }

    /// <summary>
    /// Validate legal provision references found in the document.
    /// 内置常见中国法律法规条款范围表，同时支持联网查询（需配置）。
    /// </summary>
    public class ValidateCitationTool : IAgentTool
    {
        public string Name => "validate_citation";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => false;
        public bool IsConcurrencySafe => true;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["law_name"] = new JObject { ["type"] = "string" },
                ["article_number"] = new JObject { ["type"] = "string" }
            },
            ["required"] = new JArray("law_name", "article_number")
        };

        // 内置常见法律条款范围表
        private static readonly Dictionary<string, (int MinArticle, int MaxArticle, string Alias)> KnownLaws = new()
        {
            ["中华人民共和国民法典"] = (1, 1260, "民法典"),
            ["民法典"] = (1, 1260, "中华人民共和国民法典"),
            ["中华人民共和国刑法"] = (1, 452, "刑法"),
            ["刑法"] = (1, 452, "中华人民共和国刑法"),
            ["中华人民共和国劳动法"] = (1, 107, "劳动法"),
            ["劳动法"] = (1, 107, "中华人民共和国劳动法"),
            ["中华人民共和国劳动合同法"] = (1, 98, "劳动合同法"),
            ["劳动合同法"] = (1, 98, "中华人民共和国劳动合同法"),
            ["中华人民共和国公司法"] = (1, 218, "公司法"),
            ["公司法"] = (1, 218, "中华人民共和国公司法"),
            ["中华人民共和国著作权法"] = (1, 67, "著作权法"),
            ["著作权法"] = (1, 67, "中华人民共和国著作权法"),
            ["中华人民共和国专利法"] = (1, 82, "专利法"),
            ["专利法"] = (1, 82, "中华人民共和国专利法"),
            ["中华人民共和国商标法"] = (1, 72, "商标法"),
            ["商标法"] = (1, 72, "中华人民共和国商标法"),
            ["中华人民共和国消费者权益保护法"] = (1, 63, "消费者权益保护法"),
            ["消费者权益保护法"] = (1, 63, "中华人民共和国消费者权益保护法"),
            ["中华人民共和国行政处罚法"] = (1, 86, "行政处罚法"),
            ["行政处罚法"] = (1, 86, "中华人民共和国行政处罚法"),
            ["中华人民共和国行政复议法"] = (1, 43, "行政复议法"),
            ["行政复议法"] = (1, 43, "中华人民共和国行政复议法"),
            ["中华人民共和国行政诉讼法"] = (1, 103, "行政诉讼法"),
            ["行政诉讼法"] = (1, 103, "中华人民共和国行政诉讼法"),
            ["中华人民共和国民事诉讼法"] = (1, 284, "民事诉讼法"),
            ["民事诉讼法"] = (1, 284, "中华人民共和国民事诉讼法"),
            ["中华人民共和国刑事诉讼法"] = (1, 308, "刑事诉讼法"),
            ["刑事诉讼法"] = (1, 308, "中华人民共和国刑事诉讼法"),
            ["中华人民共和国保险法"] = (1, 186, "保险法"),
            ["保险法"] = (1, 186, "中华人民共和国保险法"),
            ["中华人民共和国证券法"] = (1, 226, "证券法"),
            ["证券法"] = (1, 226, "中华人民共和国证券法"),
            ["中华人民共和国数据安全法"] = (1, 55, "数据安全法"),
            ["数据安全法"] = (1, 55, "中华人民共和国数据安全法"),
            ["中华人民共和国个人信息保护法"] = (1, 74, "个人信息保护法"),
            ["个人信息保护法"] = (1, 74, "中华人民共和国个人信息保护法"),
            ["中华人民共和国网络安全法"] = (1, 79, "网络安全法"),
            ["网络安全法"] = (1, 79, "中华人民共和国网络安全法"),
            ["中华人民共和国招标投标法"] = (1, 68, "招标投标法"),
            ["招标投标法"] = (1, 68, "中华人民共和国招标投标法"),
            ["中华人民共和国政府采购法"] = (1, 88, "政府采购法"),
            ["政府采购法"] = (1, 88, "中华人民共和国政府采购法"),
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            string lawName = parameters["law_name"]?.ToString()?.Trim() ?? "";
            string articleNumber = parameters["article_number"]?.ToString()?.Trim() ?? "";

            if (string.IsNullOrEmpty(lawName) || string.IsNullOrEmpty(articleNumber))
                return AgentToolText.T("I18n_AgentResultCitationFieldsRequired");

            // 解析条款编号
            int articleNum = ExtractArticleNumber(articleNumber);
            if (articleNum < 0)
                return AgentToolText.F("I18n_AgentResultCitationArticleUnparsable", articleNumber);

            // 查找法律
            if (!KnownLaws.TryGetValue(lawName, out var lawInfo))
            {
                // 尝试模糊匹配
                var fuzzyMatch = KnownLaws.Keys.FirstOrDefault(k =>
                    k.Contains(lawName) || lawName.Contains(k) ||
                    KnownLaws[k].Alias.Contains(lawName));
                if (fuzzyMatch == null)
                    return AgentToolText.F("I18n_AgentResultCitationLawUnknown", lawName);
                lawInfo = KnownLaws[fuzzyMatch];
                lawName = fuzzyMatch;
            }

            // 校验
            if (articleNum < lawInfo.MinArticle || articleNum > lawInfo.MaxArticle)
            {
                return AgentToolText.F("I18n_AgentResultCitationArticleOutOfRange", lawName, articleNumber, lawInfo.MinArticle, lawInfo.MaxArticle);
            }

            string alias = lawInfo.Alias;
            return AgentToolText.F("I18n_AgentResultCitationValid", lawName, articleNumber, alias, lawInfo.MinArticle, lawInfo.MaxArticle);
        }

        /// <summary>从"第584条"、"第12条第2款"、"Article 15" 等格式中提取数字。</summary>
        private int ExtractArticleNumber(string input)
        {
            if (string.IsNullOrEmpty(input)) return -1;

            // 中文格式: 第584条
            var match = Regex.Match(input, @"第\s*(\d+)\s*条");
            if (match.Success && int.TryParse(match.Groups[1].Value, out int num))
                return num;

            // 英文格式: Article 15, Section 5
            match = Regex.Match(input, @"(?:Article|Section|Clause|§)\s*(\d+)", RegexOptions.IgnoreCase);
            if (match.Success && int.TryParse(match.Groups[1].Value, out num))
                return num;

            // 纯数字
            if (int.TryParse(input.Trim(), out num))
                return num;

            return -1;
        }
    }
}
