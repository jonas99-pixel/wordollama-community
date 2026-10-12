using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.lib
{
    // Addresses are independent of body offsets. Persist names, not COM objects.
    public sealed class ReviewStoryAddress
    {
        public string Kind { get; set; } = "Body";
        public int Section { get; set; }
        public int Variant { get; set; }
        public int Note { get; set; }
        public List<string> Shapes { get; set; } = new List<string>();
        public string Key => Kind + ":" + Section + ":" + Variant + ":" + Note
            + ":" + string.Join("/", Shapes.Select(Uri.EscapeDataString));
        public string Label => StoryLabel(Kind)
            + (Section > 0 ? " " + Section + "/" + Variant : Note > 0 ? " " + Note : "")
            + (Shapes.Count > 0 ? " / " + string.Join(" / ", Shapes) : "");
        private static string StoryLabel(string kind)
        {
            switch (kind)
            {
                case "Header": return Loc.T("I18n_ReviewStoryHeader");
                case "Footer": return Loc.T("I18n_ReviewStoryFooter");
                case "Footnote": return Loc.T("I18n_ReviewStoryFootnote");
                case "Endnote": return Loc.T("I18n_ReviewStoryEndnote");
                default: return Loc.T("I18n_ReviewStoryBody");
            }
        }
    }

    public sealed class ReviewParagraphLocation
    {
        public ReviewStoryAddress Story { get; set; }
        public int Paragraph { get; set; }
        public int Start { get; set; }
        public int End { get; set; }
    }

    public sealed class ReviewStoryRange : IDisposable
    {
        public ReviewStoryAddress Address { get; set; }
        public Word.Range Range;
        public void Dispose() { ComHelper.ReleaseCom(ref Range); }
    }

    public static class ReviewDocumentScope
    {
        public static List<ReviewStoryRange> Capture(Word.Document document)
        {
            var result = new List<ReviewStoryRange>();
            try
            {
                Add(result, new ReviewStoryAddress(), document.Content);
                AddShapes(result, document.Shapes, new ReviewStoryAddress(), new List<string>());
                Word.Sections sections = document.Sections;
                try
                {
                    for (int section = 1; section <= sections.Count; section++)
                    {
                        Word.Section item = sections[section];
                        try
                        {
                            Word.HeadersFooters headers = item.Headers, footers = item.Footers;
                            try
                            {
                                for (int variant = 1; variant <= 3; variant++)
                                {
                                    CaptureHeader(result, headers[(Word.WdHeaderFooterIndex)variant], "Header", section, variant);
                                    CaptureHeader(result, footers[(Word.WdHeaderFooterIndex)variant], "Footer", section, variant);
                                }
                            }
                            finally { ComHelper.ReleaseCom(ref headers); ComHelper.ReleaseCom(ref footers); }
                        }
                        finally { ComHelper.ReleaseCom(ref item); }
                    }
                }
                finally { ComHelper.ReleaseCom(ref sections); }
                Word.Footnotes footnotes = document.Footnotes;
                try
                {
                    for (int i = 1; i <= footnotes.Count; i++)
                    {
                        Word.Footnote note = footnotes[i];
                        try { AddNote(result, new ReviewStoryAddress { Kind = "Footnote", Note = i }, note.Range); }
                        finally { ComHelper.ReleaseCom(ref note); }
                    }
                }
                finally { ComHelper.ReleaseCom(ref footnotes); }
                Word.Endnotes endnotes = document.Endnotes;
                try
                {
                    for (int i = 1; i <= endnotes.Count; i++)
                    {
                        Word.Endnote note = endnotes[i];
                        try { AddNote(result, new ReviewStoryAddress { Kind = "Endnote", Note = i }, note.Range); }
                        finally { ComHelper.ReleaseCom(ref note); }
                    }
                }
                finally { ComHelper.ReleaseCom(ref endnotes); }
                return result;
            }
            catch { foreach (var story in result) story.Dispose(); throw; }
        }

        private static void CaptureHeader(List<ReviewStoryRange> result, Word.HeaderFooter header, string kind, int section, int variant)
        {
            try
            {
                if (!header.Exists || section > 1 && header.LinkToPrevious) return;
                var address = new ReviewStoryAddress { Kind = kind, Section = section, Variant = variant };
                Add(result, address, header.Range);
                AddShapes(result, header.Shapes, address, new List<string>());
            }
            finally { ComHelper.ReleaseCom(ref header); }
        }

        private static void AddNote(List<ReviewStoryRange> result, ReviewStoryAddress address, Word.Range range)
        {
            // The automatic note reference is structural, not proofreadable text.
            string text = range.Text ?? "";
            if (text.Length > 0 && (text[0] == '\u0002' || text[0] == '\u0003')) range.Start++;
            Add(result, address, range);
        }

        private static void Add(List<ReviewStoryRange> result, ReviewStoryAddress address, Word.Range range)
        {
            // A chain of linked text boxes has a single containing range.
            if (result.Any(s => s.Range.InStory(range) && s.Range.Start == range.Start && s.Range.End == range.End))
            { ComHelper.ReleaseCom(ref range); return; }
            result.Add(new ReviewStoryRange { Address = address, Range = range });
        }

        private static void AddShapes(List<ReviewStoryRange> result, Word.Shapes shapes, ReviewStoryAddress container, List<string> parents)
        {
            try
            {
                for (int i = 1; i <= shapes.Count; i++)
                {
                    Word.Shape shape = shapes[i];
                    try { AddShape(result, shape, container, parents); }
                    finally { ComHelper.ReleaseCom(ref shape); }
                }
            }
            finally { ComHelper.ReleaseCom(ref shapes); }
        }

        private static void AddShape(List<ReviewStoryRange> result, Word.Shape shape, ReviewStoryAddress container, List<string> parents)
        {
            var names = parents.Concat(new[] { shape.Name }).ToList();
            if ((int)shape.Type == 6) // msoGroup
            {
                var group = shape.GroupItems;
                try
                {
                    for (int i = 1; i <= group.Count; i++)
                    {
                        Word.Shape child = group[i];
                        try { AddShape(result, child, container, names); }
                        finally { ComHelper.ReleaseCom(ref child); }
                    }
                }
                finally { ComHelper.ReleaseCom(ref group); }
                return;
            }
            if ((int)shape.Type == 20) // msoCanvas
            {
                var canvas = shape.CanvasItems;
                try
                {
                    for (int i = 1; i <= canvas.Count; i++)
                    {
                        Word.Shape child = canvas[i];
                        try { AddShape(result, child, container, names); }
                        finally { ComHelper.ReleaseCom(ref child); }
                    }
                }
                finally { ComHelper.ReleaseCom(ref canvas); }
                return;
            }
            Word.TextFrame frame = null;
            try
            {
                // Pictures and lines have no text frame. Do not query their text.
                if ((int)shape.Type != 17 && (int)shape.Type != 1 && (int)shape.Type != 2) return;
                frame = shape.TextFrame;
                if (frame.HasText == 0) return;
                Add(result, new ReviewStoryAddress { Kind = container.Kind, Section = container.Section,
                    Variant = container.Variant, Shapes = names }, frame.ContainingRange);
            }
            finally { ComHelper.ReleaseCom(ref frame); }
        }

        public static Word.Range Resolve(Word.Document document, ReviewStoryAddress address)
        {
            if (address == null) return document.Content; // v2.4 body sessions
            Word.Range range = null;
            Word.HeaderFooter header = null;
            try
            {
                if (address.Kind == "Header" || address.Kind == "Footer")
                {
                    if (address.Section < 1 || address.Section > document.Sections.Count) return null;
                    header = address.Kind == "Header" ? document.Sections[address.Section].Headers[(Word.WdHeaderFooterIndex)address.Variant]
                        : document.Sections[address.Section].Footers[(Word.WdHeaderFooterIndex)address.Variant];
                    if (!header.Exists || address.Section > 1 && header.LinkToPrevious) return null;
                    range = header.Range;
                }
                else if (address.Kind == "Footnote") range = document.Footnotes[address.Note].Range;
                else if (address.Kind == "Endnote") range = document.Endnotes[address.Note].Range;
                else if (address.Kind == "Body") range = document.Content;
                else return null;
                if (address.Shapes.Count == 0) return range.Duplicate;
                Word.Shapes shapes = header != null ? header.Shapes : document.Shapes;
                Word.Shape shape = null;
                try
                {
                    object name = address.Shapes[0];
                    shape = shapes.get_Item(ref name);
                    for (int i = 1; i < address.Shapes.Count; i++)
                    {
                        Word.Shape previous = shape;
                        object childName = address.Shapes[i];
                        if ((int)previous.Type == 20)
                        {
                            var canvas = previous.CanvasItems;
                            try { shape = canvas.get_Item(ref childName); }
                            finally { ComHelper.ReleaseCom(ref canvas); }
                        }
                        else
                        {
                            var group = previous.GroupItems;
                            try { shape = group.get_Item(ref childName); }
                            finally { ComHelper.ReleaseCom(ref group); }
                        }
                        ComHelper.ReleaseCom(ref previous);
                    }
                    Word.TextFrame frame = null; Word.Range text = null;
                    try { frame = shape.TextFrame; text = frame.ContainingRange; return text.Duplicate; }
                    finally { ComHelper.ReleaseCom(ref text); ComHelper.ReleaseCom(ref frame); }
                }
                finally { ComHelper.ReleaseCom(ref shape); ComHelper.ReleaseCom(ref shapes); }
            }
            catch (System.Runtime.InteropServices.COMException) { return null; }
            finally { ComHelper.ReleaseCom(ref range); ComHelper.ReleaseCom(ref header); }
        }

        public static ReviewStoryAddress AddressFor(Word.Document document, Word.Range range)
        {
            var stories = Capture(document);
            try
            {
                var match = stories.FirstOrDefault(s => s.Range.InStory(range) && s.Range.Start <= range.Start && s.Range.End >= range.End);
                if (match == null) throw new InvalidOperationException(Loc.T("I18n_ReviewSourceChanged"));
                return match.Address;
            }
            finally { foreach (var story in stories) story.Dispose(); }
        }

        public static string Fingerprint(Word.Document document)
        {
            var stories = Capture(document);
            try
            {
                var text = new StringBuilder();
                foreach (var story in stories)
                    text.Append(story.Address.Key).Append(':').Append(story.Range.End).Append(':')
                        .Append(ReviewSessionStore.Fingerprint(story.Range.Text ?? "")).Append('\n');
                return ReviewSessionStore.Fingerprint(text.ToString());
            }
            finally { foreach (var story in stories) story.Dispose(); }
        }
    }
}
