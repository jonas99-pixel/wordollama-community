using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;

namespace WordOllama.lib
{
    public sealed class ReviewTextPart
    {
        public int Start { get; set; }
        public string Text { get; set; }
    }

    public sealed class ReviewParagraph
    {
        public int Index { get; set; }
        public string Text { get; set; }
        public string FullTextHash { get; set; }
        public ReviewParagraphLocation Location { get; set; }
    }

    public sealed class ReviewIssueBatch
    {
        public string Context { get; set; }
        public int Characters { get; set; }
        public HashSet<int> Paragraphs { get; set; }
        public int FirstParagraph => Paragraphs.Min();
        public int LastParagraph => Paragraphs.Max();
    }

    /// <summary>
    /// Conservative estimates, not a model tokenizer. Every request reserves room
    /// for its complete prompt, output and a safety margin before text is split.
    /// </summary>
    public sealed class ReviewRequestBudget
    {
        public int ContextTokens { get; }
        public int OutputTokens { get; }
        public int SafetyTokens { get; }
        public bool RewritesText { get; }
        public int TextTokens { get; }

        public ReviewRequestBudget(int contextTokens, int outputTokens, string emptyPrompt, bool rewritesText)
        {
            ContextTokens = contextTokens;
            OutputTokens = outputTokens;
            SafetyTokens = Math.Max(128, contextTokens / 10);
            RewritesText = rewritesText;
            int available = contextTokens - outputTokens - SafetyTokens - ReviewChunkPlanner.EstimateTokens(emptyPrompt) - 128;
            TextTokens = rewritesText ? Math.Min(available, outputTokens - 256) : available;
            if (contextTokens <= 0 || outputTokens <= 0 || TextTokens < 32)
                throw new InvalidOperationException(Loc.T("I18n_ReviewContextTooSmall"));
        }

        public bool Fits(string prompt, string text, int items = 1)
        {
            return ReviewChunkPlanner.EstimateTokens(prompt) + OutputTokens + SafetyTokens <= ContextTokens
                && (!RewritesText || ReviewChunkPlanner.EstimateTokens(text) + 128 + items * 32 <= OutputTokens);
        }

        public static ReviewRequestBudget Create(AppConfig config, string emptyPrompt, bool rewritesText)
        {
            bool local = string.Equals(config.AIMode, "Ollama", StringComparison.OrdinalIgnoreCase);
            int context = local ? config.OllamaMaxToken : config.OnlineContextWindow;
            int output = local ? Math.Min(Math.Max(256, config.ReviewOllamaOutputTokens), context / 2) : config.ChatGPTMaxToken;
            return new ReviewRequestBudget(context, output, emptyPrompt, rewritesText);
        }
    }

    public static class ReviewChunkPlanner
    {
        public static int EstimateTokens(string text)
        {
            long ascii = 0, unicode = 0;
            foreach (char c in text ?? "")
            {
                if (c < 128) ascii++;
                else unicode += 2;
            }
            return (int)Math.Min(int.MaxValue, (ascii + 2) / 3 + unicode);
        }

        public static List<ReviewTextPart> SplitText(string text, int maxCharacters, int maxTokens)
        {
            text = text ?? "";
            if (maxCharacters < 2 || maxTokens < 4) throw new ArgumentOutOfRangeException(nameof(maxCharacters));
            var result = new List<ReviewTextPart>();
            for (int start = 0; start < text.Length;)
            {
                int low = 0, high = Math.Min(maxCharacters, text.Length - start);
                while (low < high)
                {
                    int mid = low + (high - low + 1) / 2;
                    if (EstimateTokens(text.Substring(start, mid)) <= maxTokens) low = mid;
                    else high = mid - 1;
                }
                int length = low;
                if (length == 0) throw new InvalidOperationException(Loc.T("I18n_ReviewContextTooSmall"));
                if (start + length < text.Length)
                {
                    // Prefer a sentence/word boundary without dropping its whitespace.
                    for (int i = length - 1; i >= length / 2; i--)
                        if (char.IsWhiteSpace(text[start + i]) || ".!?。！？；;".IndexOf(text[start + i]) >= 0)
                        { length = i + 1; break; }
                    if (length > 0 && char.IsHighSurrogate(text[start + length - 1]) && char.IsLowSurrogate(text[start + length])) length--;
                    if (length > 0 && text[start + length - 1] == '\r' && text[start + length] == '\n') length--;
                }
                if (length == 0) throw new InvalidOperationException(Loc.T("I18n_ReviewContextTooSmall"));
                result.Add(new ReviewTextPart { Start = start, Text = text.Substring(start, length) });
                start += length;
            }
            return result;
        }

        public static List<ReviewIssueBatch> CreateIssueBatches(IEnumerable<ReviewParagraph> paragraphs,
            ReviewRequestBudget budget, int maxParagraphs, int maxCharacters, Func<string, string> buildPrompt = null)
        {
            buildPrompt = buildPrompt ?? (text => PromptTemplateService.BuildIssueReviewPrompt(text));
            maxParagraphs = Math.Max(1, Math.Min(1000, maxParagraphs));
            maxCharacters = Math.Max(128, Math.Min(1000000, maxCharacters));
            var result = new List<ReviewIssueBatch>();
            var context = new StringBuilder();
            var indexes = new HashSet<int>();
            int characters = 0, units = 0;
            foreach (var paragraph in paragraphs)
            {
                if (string.IsNullOrWhiteSpace(paragraph.Text)) continue;
                foreach (var part in SplitText(paragraph.Text, Math.Max(2, maxCharacters - 96), budget.TextTokens))
                {
                    string block = "[Paragraph " + paragraph.Index + ", offset " + part.Start
                        + (paragraph.Location?.Story == null ? "" : ", area " + paragraph.Location.Story.Label.Replace("\r", " ").Replace("\n", " "))
                        + "]\n" + part.Text + "\n";
                    string candidate = context.ToString() + block;
                    if (units > 0 && (units >= maxParagraphs || candidate.Length > maxCharacters
                        || !budget.Fits(buildPrompt(candidate), candidate)))
                    {
                        result.Add(new ReviewIssueBatch { Context = context.ToString(), Characters = characters, Paragraphs = indexes });
                        context.Clear(); indexes = new HashSet<int>(); characters = 0; units = 0;
                    }
                    if (!budget.Fits(buildPrompt(block), block))
                        throw new InvalidOperationException(Loc.T("I18n_ReviewContextTooSmall"));
                    context.Append(block); indexes.Add(paragraph.Index); characters += part.Text.Length; units++;
                }
            }
            if (units > 0) result.Add(new ReviewIssueBatch { Context = context.ToString(), Characters = characters, Paragraphs = indexes });
            return result;
        }

        public static string PreserveFragmentWhitespace(string source, string suggestion)
        {
            if (string.IsNullOrWhiteSpace(source)) return source;
            string text = (suggestion ?? "").Trim();
            int first = 0, last = source.Length;
            while (first < source.Length && char.IsWhiteSpace(source[first])) first++;
            while (last > first && char.IsWhiteSpace(source[last - 1])) last--;
            return source.Substring(0, first) + text + source.Substring(last);
        }

        public static bool TryReadIssues(string raw, ICollection<int> allowedParagraphs, out JArray issues)
        {
            issues = null;
            if (IsFailureResponse(raw)) return false;
            try
            {
                string json = Regex.Replace(raw.Trim(), @"^```(?:json)?\s*", "", RegexOptions.IgnoreCase);
                json = Regex.Replace(json, @"\s*```$", "");
                var array = JArray.Parse(json);
                foreach (var token in array)
                {
                    if (!(token is JObject item)) return false;
                    var index = item["paragraph"] ?? item["paragraphIndex"] ?? item["段落"];
                    int paragraph;
                    if (index == null || !int.TryParse(index.ToString(), out paragraph) || !allowedParagraphs.Contains(paragraph)) return false;
                    if (string.IsNullOrWhiteSpace((item["title"] ?? item["标题"])?.ToString())
                        && string.IsNullOrWhiteSpace((item["description"] ?? item["问题说明"] ?? item["issue"])?.ToString())) return false;
                }
                issues = array;
                return true;
            }
            catch (Exception ex) when (ex is Newtonsoft.Json.JsonException || ex is FormatException)
            { return false; }
        }

        public static bool IsFailureResponse(string response)
        {
            string text = response?.Trim() ?? "";
            return text.Length == 0 || text.StartsWith("[API", StringComparison.OrdinalIgnoreCase)
                || text == Loc.T("I18n_ReviewOutputTruncated");
        }

        public static Dictionary<string, string> ReadSuggestions(string raw, ICollection<string> allowedIds)
        {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            if (IsFailureResponse(raw)) return result;
            try
            {
                string json = Regex.Replace(raw.Trim(), @"^```(?:json)?\s*", "", RegexOptions.IgnoreCase);
                json = Regex.Replace(json, @"\s*```$", "");
                var seen = new HashSet<string>(StringComparer.Ordinal);
                foreach (var token in JArray.Parse(json))
                {
                    if (!(token is JObject item)) return new Dictionary<string, string>();
                    string id = item["id"]?.ToString();
                    var text = item["text"] ?? item["revised_text"] ?? item["suggestion"];
                    if (id == null || !allowedIds.Contains(id) || !seen.Add(id) || text?.Type != JTokenType.String)
                        return new Dictionary<string, string>();
                    if (!IsFailureResponse(text.ToString())) result[id] = text.ToString();
                }
                return result;
            }
            catch (Newtonsoft.Json.JsonException) { return new Dictionary<string, string>(); }
        }
    }
}
