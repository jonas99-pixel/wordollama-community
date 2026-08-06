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

    /// <summary>
    /// Aligns paragraph snapshots with a real line diff so an insertion near the
    /// beginning of a document does not make every following paragraph look modified.
    /// </summary>
    public static class DocumentDiffPlanner
    {
        public static List<DiffChange> CreateChanges(
            IReadOnlyList<string> original,
            IReadOnlyList<string> revised)
        {
            original = original ?? Array.Empty<string>();
            revised = revised ?? Array.Empty<string>();
            var changes = new List<DiffChange>();
            var result = new Differ().CreateLineDiffs(
                string.Join("\n", original),
                string.Join("\n", revised),
                ignoreWhitespace: false,
                ignoreCase: false);

            foreach (var block in result.DiffBlocks)
            {
                int paired = Math.Min(block.DeleteCountA, block.InsertCountB);
                for (int i = 0; i < paired; i++)
                {
                    changes.Add(new DiffChange
                    {
                        Type = "修改",
                        ParagraphIndex = block.InsertStartB + i + 1,
                        PostParagraphIndex = block.InsertStartB + i + 1,
                        OldText = original[block.DeleteStartA + i],
                        NewText = revised[block.InsertStartB + i]
                    });
                }

                for (int i = paired; i < block.DeleteCountA; i++)
                {
                    changes.Add(new DiffChange
                    {
                        Type = "删除",
                        ParagraphIndex = block.DeleteStartA + i + 1,
                        PostParagraphIndex = block.InsertStartB + paired + 1,
                        OldText = original[block.DeleteStartA + i]
                    });
                }

                for (int i = paired; i < block.InsertCountB; i++)
                {
                    changes.Add(new DiffChange
                    {
                        Type = "新增",
                        ParagraphIndex = block.InsertStartB + i + 1,
                        PostParagraphIndex = block.InsertStartB + i + 1,
                        NewText = revised[block.InsertStartB + i]
                    });
                }
            }
            return changes;
        }
    }
}
