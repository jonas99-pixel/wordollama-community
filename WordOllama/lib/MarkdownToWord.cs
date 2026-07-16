using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.lib
{
    public class MarkdownStyleMapping
    {
        public string MarkdownElement { get; set; }
        public string WordStyleName { get; set; }
    }

    public static class MarkdownToWord
    {
        private static readonly Dictionary<string, string> DefaultMappings = new()
        {
            ["h1"] = "Heading 1",
            ["h2"] = "Heading 2",
            ["h3"] = "Heading 3",
            ["h4"] = "Heading 4",
            ["p"] = "Normal",
            ["code"] = "Normal",
            ["blockquote"] = "Normal",
            ["ul"] = "List Bullet",
            ["ol"] = "List Number",
        };

        private static readonly Regex InlineRegex = new(
            @"\*\*(.+?)\*\*|\*(.+?)\*|`([^`]+)`|~~(.+?)~~|\[([^\]]+)\]\(([^)]+)\)",
            RegexOptions.Compiled);

        public static void InsertMarkdown(string markdown, List<MarkdownStyleMapping> styleMappings)
        {
            var wordApp = Globals.ThisAddIn.Application;
            var doc = wordApp.ActiveDocument;
            if (doc == null) { System.Windows.Forms.MessageBox.Show(Loc.T("I18n_NoActiveDocument")); return; }

            var mappingDict = new Dictionary<string, string>();
            foreach (var m in styleMappings)
                if (!string.IsNullOrEmpty(m.MarkdownElement) && !string.IsNullOrEmpty(m.WordStyleName))
                    mappingDict[m.MarkdownElement] = m.WordStyleName;
            foreach (var kv in DefaultMappings)
                if (!mappingDict.ContainsKey(kv.Key)) mappingDict[kv.Key] = kv.Value;

            var sel = wordApp.Selection;
            string[] lines = markdown.Split('\n');

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                string trimmed = line.Trim();

                // Code block
                if (trimmed.StartsWith("```"))
                { i = InsertCodeBlock(sel, lines, i); continue; }

                // Headings
                var hMatch = Regex.Match(trimmed, @"^(#{1,4})\s+(.+)$");
                if (hMatch.Success)
                {
                    int lvl = hMatch.Groups[1].Length;
                    string text = hMatch.Groups[2].Value.Trim();
                    string sn = mappingDict.ContainsKey($"h{lvl}") ? mappingDict[$"h{lvl}"] : "Normal";
                    sel.ParagraphFormat.Reset();
                    if (StyleExists(doc, sn)) sel.set_Style(sn);
                    InsertFormattedText(sel, text);
                    sel.TypeParagraph();
                    continue;
                }

                // Blockquote
                if (trimmed.StartsWith("> "))
                {
                    string text = trimmed.Substring(2).Trim();
                    sel.ParagraphFormat.Reset();
                    string sn = mappingDict["blockquote"];
                    if (StyleExists(doc, sn)) sel.set_Style(sn);
                    sel.ParagraphFormat.LeftIndent = 24f;
                    InsertFormattedText(sel, text);
                    sel.TypeParagraph();
                    continue;
                }

                // Unordered list
                var ulMatch = Regex.Match(trimmed, @"^[-*]\s+(.+)$");
                if (ulMatch.Success)
                {
                    string text = ulMatch.Groups[1].Value;
                    sel.ParagraphFormat.Reset();
                    string sn = mappingDict["ul"];
                    if (StyleExists(doc, sn)) sel.set_Style(sn);
                    else sel.ParagraphFormat.LeftIndent = 24f;
                    InsertFormattedText(sel, text);
                    sel.TypeParagraph();
                    continue;
                }

                // Ordered list
                var olMatch = Regex.Match(trimmed, @"^(\d+)\.\s+(.+)$");
                if (olMatch.Success)
                {
                    string text = olMatch.Groups[2].Value;
                    sel.ParagraphFormat.Reset();
                    string sn = mappingDict["ol"];
                    if (StyleExists(doc, sn)) sel.set_Style(sn);
                    else sel.ParagraphFormat.LeftIndent = 24f;
                    InsertFormattedText(sel, text);
                    sel.TypeParagraph();
                    continue;
                }

                // Empty line
                if (string.IsNullOrWhiteSpace(line))
                { sel.ParagraphFormat.Reset(); sel.TypeParagraph(); continue; }

                // Normal paragraph
                sel.ParagraphFormat.Reset();
                string pStyle = mappingDict["p"];
                if (StyleExists(doc, pStyle)) sel.set_Style(pStyle);
                InsertFormattedText(sel, line.Trim());
                sel.TypeParagraph();
            }
        }

        private static void InsertFormattedText(Word.Selection sel, string text)
        {
            int pos = 0;
            foreach (Match m in InlineRegex.Matches(text))
            {
                if (m.Index > pos)
                    sel.TypeText(text.Substring(pos, m.Index - pos));

                if (m.Groups[1].Success) // **bold**
                { sel.Font.Bold = 1; sel.TypeText(m.Groups[1].Value); sel.Font.Bold = 0; }
                else if (m.Groups[2].Success) // *italic*
                { sel.Font.Italic = 1; sel.TypeText(m.Groups[2].Value); sel.Font.Italic = 0; }
                else if (m.Groups[3].Success) // `inline code`
                {
                    var prevName = sel.Font.Name;
                    sel.Font.Name = "Consolas"; sel.Font.Size = 9;
                    sel.TypeText(m.Groups[3].Value);
                    sel.Font.Name = prevName; sel.Font.Reset();
                }
                else if (m.Groups[4].Success) // ~~strikethrough~~
                { sel.Font.StrikeThrough = 1; sel.TypeText(m.Groups[4].Value); sel.Font.StrikeThrough = 0; }
                else if (m.Groups[5].Success && m.Groups[6].Success) // [text](url)
                {
                    string linkText = m.Groups[5].Value;
                    string url = m.Groups[6].Value;
                    sel.Font.Reset();
                    sel.TypeText(linkText);
                    // Select the just-typed text and add hyperlink
                    Word.Range rng = sel.Range;
                    rng.MoveStart(Word.WdUnits.wdCharacter, -linkText.Length);
                    sel.Document.Hyperlinks.Add(rng, url, Type.Missing, Type.Missing, linkText);
                    sel.EndKey(Word.WdUnits.wdStory); // move cursor to end
                }

                pos = m.Index + m.Length;
            }
            if (pos < text.Length)
                sel.TypeText(text.Substring(pos));
        }

        private static int InsertCodeBlock(Word.Selection sel, string[] lines, int start)
        {
            var codeLines = new List<string>();
            int i = start + 1;
            while (i < lines.Length && !lines[i].TrimStart().StartsWith("```"))
            { codeLines.Add(lines[i]); i++; }
            if (i < lines.Length) i++;

            sel.ParagraphFormat.Reset();
            sel.Font.Name = "Consolas"; sel.Font.Size = 9;
            sel.ParagraphFormat.LeftIndent = 12f;
            foreach (var cl in codeLines)
            { sel.TypeText(cl); sel.TypeParagraph(); }
            sel.Font.Reset();
            return i;
        }

        private static bool StyleExists(Word.Document doc, string styleName)
        {
            try { return doc.Styles[styleName] != null; }
            catch { return false; }
        }
    }
}
