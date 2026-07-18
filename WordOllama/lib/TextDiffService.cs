using System;
using System.Collections.Generic;
using DiffPlex;

namespace WordOllama.lib
{
    public sealed class TextEdit
    {
        public int OriginalStart { get; set; }
        public int DeleteLength { get; set; }
        public string InsertText { get; set; }
    }

    /// <summary>
    /// Produces minimal character edits that can be applied from the end of a Word range.
    /// </summary>
    public static class TextDiffService
    {
        public static IReadOnlyList<TextEdit> CreateEdits(string originalText, string revisedText)
        {
            originalText = originalText ?? "";
            revisedText = revisedText ?? "";

            if (string.Equals(originalText, revisedText, StringComparison.Ordinal))
                return Array.Empty<TextEdit>();

            var differ = new Differ();
            var result = differ.CreateCharacterDiffs(
                originalText,
                revisedText,
                ignoreWhitespace: false,
                ignoreCase: false);
            var edits = new List<TextEdit>(result.DiffBlocks.Count);

            foreach (var block in result.DiffBlocks)
            {
                edits.Add(new TextEdit
                {
                    OriginalStart = block.DeleteStartA,
                    DeleteLength = block.DeleteCountA,
                    InsertText = block.InsertCountB > 0
                        ? revisedText.Substring(block.InsertStartB, block.InsertCountB)
                        : ""
                });
            }

            return edits;
        }

        public static string NormalizeWordLineBreaks(string text)
        {
            return (text ?? "")
                .Replace("\r\n", "\n")
                .Replace('\n', '\r');
        }

        public static string PreserveStructuralSuffix(string originalText, string revisedText)
        {
            originalText = originalText ?? "";
            revisedText = NormalizeWordLineBreaks(revisedText);

            int suffixStart = originalText.Length;
            while (suffixStart > 0)
            {
                char current = originalText[suffixStart - 1];
                if (current != '\r' && current != '\a') break;
                suffixStart--;
            }

            string suffix = originalText.Substring(suffixStart);
            if (suffix.Length == 0 || revisedText.EndsWith(suffix, StringComparison.Ordinal))
                return revisedText;

            return revisedText.TrimEnd('\r', '\a') + suffix;
        }
    }
}
