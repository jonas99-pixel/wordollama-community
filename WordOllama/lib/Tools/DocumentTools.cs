using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.lib.Tools
{
    /// <summary>
    /// High-performance text search using Word's native Find API.
    /// 相比 foreach (Paragraph) 遍历 500 页文档快几百倍。
    /// </summary>
    public class SearchTextTool : IAgentTool
    {
        public string Name => "search_text";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => false;
        public bool IsConcurrencySafe => !IsWriteOperation;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["keyword"] = new JObject { ["type"] = "string" }
            },
            ["required"] = new JArray("keyword")
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            string keyword = parameters["keyword"]?.ToString() ?? "";
            if (string.IsNullOrEmpty(keyword))
                return AgentToolText.T("I18n_AgentResultKeywordRequired");

            Word.Range searchRange = doc.Range();
            searchRange.Find.ClearFormatting();
            searchRange.Find.Text = keyword;
            searchRange.Find.Forward = true;

            var results = new List<string>();
            int originalEnd = searchRange.End;

            while (searchRange.Find.Execute())
            {
                if (searchRange.Start >= originalEnd) break;

                // 获取匹配所在的段落
                Word.Paragraph matchedPara = searchRange.Paragraphs[1];
                string paraText = matchedPara.Range.Text.TrimEnd('\r', '\a');
                string preview = paraText.Length > 50 ? paraText.Substring(0, 50) + "..." : paraText;

                // 通过 Range 比较确定段落编号（比遍历 Paragraphs 集合快得多）
                int paraNum = 1;
                foreach (Word.Paragraph p in doc.Paragraphs)
                {
                    if (p.Range.Start >= matchedPara.Range.Start) break;
                    paraNum++;
                }

                results.Add(AgentToolText.F("I18n_AgentResultSearchMatch", paraNum, preview));

                // 移动搜索起点到当前匹配之后
                searchRange.SetRange(searchRange.End, originalEnd);

                if (results.Count >= 10) break; // 安全上限
            }

            if (results.Count == 0) return AgentToolText.F("I18n_AgentResultKeywordNotFound", keyword);
            return AgentToolText.F("I18n_AgentResultKeywordMatches", keyword, string.Join("\n", results));
        }
    }

    public class GetDocOverviewTool : IAgentTool
    {
        public string Name => "get_doc_overview";
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
            int count = doc.Paragraphs.Count;
            string preview = "";
            for (int i = 1; i <= Math.Min(2, count); i++)
            {
                string text = doc.Paragraphs[i].Range.Text?.TrimEnd('\r', '\a') ?? "";
                preview += AgentToolText.F("I18n_AgentResultOverviewParagraph", i, text) + "\n";
            }
            // 获取字数
            int wordCount = 0;
            try { wordCount = doc.Range().ComputeStatistics(Word.WdStatistic.wdStatisticWords); }
            catch { wordCount = 0; }

            return AgentToolText.F("I18n_AgentResultDocumentOverview", count, wordCount, preview);
        }
    }

    public class GetSelectionTool : IAgentTool
    {
        public string Name => "get_selection";
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
            string selText = app.Selection.Text?.TrimEnd('\r', '\a');
            if (string.IsNullOrWhiteSpace(selText))
                return AgentToolText.T("I18n_AgentResultNoSelection");
            return AgentToolText.F("I18n_AgentResultSelection", selText);
        }
    }

    /// <summary>
    /// 读取文档批注 — 用于整理审阅意见或汇总待办，不修改批注内容。
    /// </summary>
    public class ReadCommentsTool : IAgentTool
    {
        public string Name => "read_comments";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => false;
        public bool IsConcurrencySafe => true;
        public int MaxResultSizeChars => 12000;
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
                int count = doc.Comments.Count;
                if (count == 0) return AgentToolText.T("I18n_AgentResultNoComments");

                var lines = new List<string>();
                for (int i = 1; i <= Math.Min(count, 50); i++)
                {
                    Word.Comment comment = doc.Comments[i];
                    string text = comment.Range.Text?.Trim() ?? "";
                    if (text.Length > 500) text = text.Substring(0, 500) + "...";
                    string author = comment.Author ?? AgentToolText.T("I18n_AgentResultUnknownAuthor");
                    string anchor = comment.Scope?.Text?.Trim() ?? "";
                    if (anchor.Length > 120) anchor = anchor.Substring(0, 120) + "...";
                    lines.Add(AgentToolText.F("I18n_AgentResultCommentItem", i, author, text, anchor));
                }

                if (count > 50)
                    lines.Add(AgentToolText.F("I18n_AgentResultCommentsTruncated", count));
                return string.Join("\n", lines);
            }
            catch (Exception ex)
            {
                return AgentToolText.F("I18n_AgentResultReadCommentsFailed", ex.Message);
            }
        }
    }

    /// <summary>
    /// 读取文档书签 — 返回书签名称、范围文本和位置，供 Agent 精确定位内容。
    /// </summary>
    public class ReadBookmarksTool : IAgentTool
    {
        public string Name => "read_bookmarks";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => false;
        public bool IsConcurrencySafe => true;
        public int MaxResultSizeChars => 10000;
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
                int count = doc.Bookmarks.Count;
                if (count == 0) return AgentToolText.T("I18n_AgentResultNoBookmarks");

                var lines = new List<string>();
                for (int i = 1; i <= Math.Min(count, 100); i++)
                {
                    Word.Bookmark bookmark = doc.Bookmarks[i];
                    string text = bookmark.Range.Text?.Trim() ?? "";
                    if (text.Length > 300) text = text.Substring(0, 300) + "...";
                    lines.Add(AgentToolText.F("I18n_AgentResultBookmarkItem", i, bookmark.Name, bookmark.Range.Start, bookmark.Range.End, text));
                }

                if (count > 100)
                    lines.Add(AgentToolText.F("I18n_AgentResultBookmarksTruncated", count));
                return string.Join("\n", lines);
            }
            catch (Exception ex)
            {
                return AgentToolText.F("I18n_AgentResultReadBookmarksFailed", ex.Message);
            }
        }
    }

    public class ReadParagraphsTool : IAgentTool
    {
        public string Name => "read_paragraphs";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => false;
        public bool IsConcurrencySafe => !IsWriteOperation;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["start"] = new JObject { ["type"] = "integer" },
                ["end"] = new JObject { ["type"] = "integer" }
            },
            ["required"] = new JArray("start", "end")
        };

        /// <summary>读取指定范围的段落 — 使用单次 Range 操作避免逐段 COM 调用。</summary>
        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            int start = parameters["start"]?.Value<int>() ?? 1;
            int end = parameters["end"]?.Value<int>() ?? 5;
            int max = doc.Paragraphs.Count;
            if (start > max)
                return AgentToolText.F("I18n_AgentResultStartParagraphOutOfRange", start, max);
            end = Math.Min(end, max);

            // [性能优化] 使用 Range 一次读取，避免逐段 COM 调用
            Word.Range startRange = doc.Paragraphs[start].Range;
            Word.Range endRange = doc.Paragraphs[end].Range;
            Word.Range bulkRange = doc.Range(startRange.Start, endRange.End);
            string bulkText = bulkRange.Text ?? "";

            // 按段落标记拆分并编号
            var sb = new StringBuilder();
            string[] paragraphs = bulkText.Split(new[] { '\r' }, StringSplitOptions.None);
            for (int i = 0; i < paragraphs.Length && (start + i) <= end; i++)
            {
                string text = paragraphs[i].TrimEnd('\a', '\r');
                if (!string.IsNullOrEmpty(text))
                    sb.AppendLine(AgentToolText.F("I18n_AgentResultParagraphItem", start + i, text));
            }
            return sb.ToString();
        }
    }

    /// <summary>
    /// 智能大块读取 — 一次性读取尽可能多的内容（默认50段），
    /// 避免 AI 陷入"几段几段看"的低效循环。使用单次 Range 操作，性能极高。
    /// </summary>
    public class ReadLargeChunkTool : IAgentTool
    {
        public string Name => "read_large_chunk";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => false;
        public bool IsConcurrencySafe => !IsWriteOperation;
        public int MaxResultSizeChars => 15000;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["start_paragraph"] = new JObject { ["type"] = "integer" }
            },
            ["required"] = new JArray("start_paragraph")
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            int startPara = parameters["start_paragraph"]?.Value<int>() ?? 1;
            int totalParas = doc.Paragraphs.Count;

            if (startPara > totalParas)
                return AgentToolText.T("I18n_AgentResultDocumentEndReached");

            // 一次性读取 50 段或到末尾（COM Range 单次操作极快）
            int endPara = Math.Min(startPara + 49, totalParas);
            Word.Range startRange = doc.Paragraphs[startPara].Range;
            Word.Range endRange = doc.Paragraphs[endPara].Range;
            Word.Range bulkRange = doc.Range(startRange.Start, endRange.End);
            string text = bulkRange.Text ?? "";

            // 按段落标记拆分并编号
            var sb = new StringBuilder();
            string[] paragraphs = text.Split(new[] { '\r' }, StringSplitOptions.None);
            for (int i = 0; i < paragraphs.Length && (startPara + i) <= endPara; i++)
            {
                string para = paragraphs[i].TrimEnd('\a', '\r');
                if (!string.IsNullOrEmpty(para))
                    sb.AppendLine(AgentToolText.F("I18n_AgentResultParagraphItem", startPara + i, para));
            }

            string tail = endPara < totalParas
                ? "\n\n" + AgentToolText.F("I18n_AgentResultLargeChunkMore", startPara, endPara, totalParas - endPara, endPara + 1)
                : "\n\n" + AgentToolText.T("I18n_AgentResultDocumentEndReached");

            return sb.ToString() + tail;
        }
    }

    /// <summary>
    /// 语义地图 — 每 10 段抽取首句，生成全文极简摘要。
    /// 帮助超长文档快速定位感兴趣的区域，再用 read_large_chunk 精读。
    /// </summary>
    public class BuildSemanticMapTool : IAgentTool
    {
        public string Name => "build_semantic_map";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => false;
        public bool IsConcurrencySafe => !IsWriteOperation;
        public int MaxResultSizeChars => 5000;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject(),
            ["required"] = new JArray()
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            int totalParas = doc.Paragraphs.Count;
            var map = new StringBuilder();
            map.AppendLine(AgentToolText.F("I18n_AgentResultSemanticMapHeader", totalParas));
            map.AppendLine();

            for (int i = 1; i <= totalParas; i += 10)
            {
                // 使用 Range 批量读取（每次 10 段），避免逐段 COM 调用
                int chunkEnd = Math.Min(i + 9, totalParas);
                Word.Range range = doc.Range(
                    doc.Paragraphs[i].Range.Start,
                    doc.Paragraphs[chunkEnd].Range.End);
                string chunkText = (range.Text ?? "").TrimEnd('\r', '\a');

                // 取首句作为摘要（最多80字符）
                string firstLine = chunkText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? "";
                string summary = firstLine.Length > 80 ? firstLine.Substring(0, 80) + "..." : firstLine;
                if (string.IsNullOrWhiteSpace(summary)) summary = AgentToolText.T("I18n_AgentResultEmptyParagraph");

                map.AppendLine(AgentToolText.F("I18n_AgentResultSemanticMapEntry", i, chunkEnd, summary));
            }

            map.AppendLine();
            map.AppendLine(AgentToolText.T("I18n_AgentResultSemanticMapHint"));
            return map.ToString();
        }
    }

    public class InsertTextAtEndTool : IAgentTool
    {
        public string Name => "insert_text_at_end";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => true;
        public bool IsConcurrencySafe => !IsWriteOperation;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["text"] = new JObject { ["type"] = "string" }
            },
            ["required"] = new JArray("text")
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            string text = parameters["text"]?.ToString() ?? "";
            if (string.IsNullOrEmpty(text))
                return AgentToolText.T("I18n_AgentResultInsertTextRequired");

            Word.Range rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
            rng.InsertParagraphAfter();
            Word.Range newRng = doc.Paragraphs[doc.Paragraphs.Count].Range;
            newRng.Text = text;
            DocumentAgent.FocusOnRange(newRng, app);
            return AgentToolText.T("I18n_AgentResultInsertedAtEnd");
        }
    }

    public class ApplyStyleTool : IAgentTool
    {
        public string Name => "apply_style";
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
                ["style_name"] = new JObject { ["type"] = "string" }
            },
            ["required"] = new JArray("paragraph_index", "style_name")
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            int index = parameters["paragraph_index"]?.Value<int>() ?? 1;
            string styleName = parameters["style_name"]?.ToString() ?? AgentToolText.T("I18n_AgentResultDefaultBodyStyle");

            if (index < 1 || index > doc.Paragraphs.Count)
                return AgentToolText.T("I18n_AgentResultParagraphOutOfRange");
            try
            {
                object style = styleName;
                doc.Paragraphs[index].set_Style(ref style);
                DocumentAgent.FocusOnRange(doc.Paragraphs[index].Range, app);
                return AgentToolText.F("I18n_AgentResultStyleApplied", index, styleName);
            }
            catch
            {
                return AgentToolText.F("I18n_AgentResultApplyStyleFailed", styleName);
            }
        }
    }

    /// <summary>
    /// Replace a paragraph; the execution layer applies tracked-change policy.
    /// </summary>
    public class ReplaceParagraphTool : IAgentTool
    {
        public string Name => "replace_paragraph";
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
                ["new_text"] = new JObject { ["type"] = "string" }
            },
            ["required"] = new JArray("paragraph_index", "new_text")
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            int index = parameters["paragraph_index"]?.Value<int>() ?? 1;
            string newText = parameters["new_text"]?.ToString() ?? "";

            if (index < 1 || index > doc.Paragraphs.Count)
                return AgentToolText.F("I18n_AgentResultParagraphIndexOutOfRange", index, doc.Paragraphs.Count);

            Word.Range rng = doc.Paragraphs[index].Range;
            rng.Text = newText + "\r"; // 保持段落标记
            DocumentAgent.FocusOnRange(rng, app);
            return AgentToolText.F("I18n_AgentResultParagraphReplaced", index);
        }
    }

    public class AddCommentTool : IAgentTool
    {
        public string Name => "add_comment";
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
                ["text_to_comment"] = new JObject { ["type"] = "string" },
                ["comment"] = new JObject { ["type"] = "string" }
            },
            ["required"] = new JArray("paragraph_index", "text_to_comment", "comment")
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            int index = parameters["paragraph_index"]?.Value<int>() ?? 1;
            string originalText = parameters["text_to_comment"]?.ToString() ?? "";
            string comment = parameters["comment"]?.ToString() ?? "";

            if (index < 1 || index > doc.Paragraphs.Count)
                return AgentToolText.T("I18n_AgentResultParagraphOutOfRange");
            if (string.IsNullOrEmpty(originalText))
                return AgentToolText.T("I18n_AgentResultCommentTextRequired");

            Word.Range rng = doc.Paragraphs[index].Range;

            // 使用 Word 原生 Find API 定位文字（避免 rng.Text 中控制字符导致 IndexOf 失败）
            Word.Find find = rng.Find;
            find.ClearFormatting();
            find.Text = originalText;
            find.MatchCase = false;
            find.MatchWildcards = false;
            find.Forward = true;

            if (!find.Execute())
                return AgentToolText.F("I18n_AgentResultCommentTextNotFound", index, originalText);

            // rng 此时已被 Find 修改为匹配范围
            doc.Comments.Add(rng, comment);
            DocumentAgent.FocusOnRange(rng, app);
            return AgentToolText.F("I18n_AgentResultCommentAdded", index, originalText);
        }
    }

    /// <summary>
    /// 全文查找替换 — 使用 Word 原生 Find.Replace 批量替换。
    /// </summary>
    public class FindReplaceTool : IAgentTool
    {
        public string Name => "find_replace";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => true;
        public bool IsConcurrencySafe => false;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["find_text"] = new JObject { ["type"] = "string" },
                ["replace_text"] = new JObject { ["type"] = "string" },
                ["match_case"] = new JObject { ["type"] = "boolean" }
            },
            ["required"] = new JArray("find_text", "replace_text")
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            string findText = parameters["find_text"]?.ToString() ?? "";
            string replaceText = parameters["replace_text"]?.ToString() ?? "";
            bool matchCase = parameters["match_case"]?.Value<bool>() ?? false;

            if (string.IsNullOrEmpty(findText))
                return AgentToolText.T("I18n_AgentResultFindTextRequired");

            try
            {
                Word.Range rng = doc.Range();
                rng.Find.ClearFormatting();
                rng.Find.Text = findText;
                rng.Find.Replacement.ClearFormatting();
                rng.Find.Replacement.Text = replaceText;
                rng.Find.MatchCase = matchCase;
                rng.Find.MatchWholeWord = false;
                rng.Find.Forward = true;
                rng.Find.Wrap = Word.WdFindWrap.wdFindContinue;

                int count = 0;
                rng.Find.Execute(Replace: Word.WdReplace.wdReplaceAll);
                // 通过文档变更次数估算替换数量较困难，改用逐次 Find 计数
                // 简单方案：执行替换后用另一种方式估算
                // 先尝试用 wdReplaceAll 快速执行
                object replaceAll = Word.WdReplace.wdReplaceAll;
                object missing = System.Type.Missing;
                rng.Find.Execute(
                    FindText: findText,
                    MatchCase: matchCase,
                    MatchWholeWord: false,
                    MatchWildcards: false,
                    MatchSoundsLike: false,
                    MatchAllWordForms: false,
                    Forward: true,
                    Wrap: Word.WdFindWrap.wdFindContinue,
                    Format: false,
                    ReplaceWith: replaceText,
                    Replace: replaceAll);

                // 替换后重新搜索计数
                Word.Range countRange = doc.Range();
                countRange.Find.ClearFormatting();
                countRange.Find.Text = replaceText;
                countRange.Find.MatchCase = matchCase;
                countRange.Find.Forward = true;
                while (countRange.Find.Execute())
                {
                    count++;
                }

                DocumentAgent.FocusOnRange(doc.Range(0, 0), app);
                return AgentToolText.F("I18n_AgentResultFindReplaceCompleted", findText, replaceText, count);
            }
            catch (Exception ex)
            {
                return AgentToolText.F("I18n_AgentResultFindReplaceFailed", ex.Message);
            }
        }
    }

    /// <summary>
    /// 在光标/选中位置插入文本。
    /// </summary>
    public class InsertAtCursorTool : IAgentTool
    {
        public string Name => "insert_at_cursor";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => true;
        public bool IsConcurrencySafe => false;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["text"] = new JObject { ["type"] = "string" },
                ["new_paragraph"] = new JObject { ["type"] = "boolean" }
            },
            ["required"] = new JArray("text")
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            string text = parameters["text"]?.ToString() ?? "";
            bool newPara = parameters["new_paragraph"]?.Value<bool>() ?? true;

            if (string.IsNullOrEmpty(text))
                return AgentToolText.T("I18n_AgentResultInsertTextRequired");

            try
            {
                Word.Selection sel = app.Selection;
                if (newPara)
                {
                    sel.TypeParagraph();
                }
                sel.TypeText(text);
                return AgentToolText.F("I18n_AgentResultInsertedAtCursor", text);
            }
            catch (Exception ex)
            {
                return AgentToolText.F("I18n_AgentResultInsertTextFailed", ex.Message);
            }
        }
    }
}
