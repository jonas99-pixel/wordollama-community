using System;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.lib
{
    public static class ReviewCopyService
    {
        public static int SaveFormatFor(Word.Document document)
        {
            // A new unsaved document can contain VBA while SaveFormat is docx.
            // Choose a macro-enabled copy instead of asking Word to discard VBA.
            int format = document.SaveFormat;
            return (format == 12 || format == 16 || format == 24) && document.HasVBProject ? 13 : format;
        }

        public static string ExtensionFor(int saveFormat)
        {
            switch (saveFormat)
            {
                case 0: return ".doc";
                case 1: return ".dot";
                case 6: return ".rtf";
                case 12: case 16: case 24: return ".docx";
                case 13: return ".docm";
                case 14: return ".dotx";
                case 15: return ".dotm";
                default: throw new InvalidOperationException(Loc.T("I18n_ReviewCopyUnsupported"));
            }
        }

        public static void ValidateTarget(string originalPath, string destination, string extension)
        {
            string target = Path.GetFullPath(destination);
            if (!string.Equals(Path.GetExtension(target), extension, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(Loc.T("I18n_ReviewCopyExtension"));
            if (Path.IsPathRooted(originalPath ?? "") && string.Equals(target, Path.GetFullPath(originalPath), StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(Loc.T("I18n_ReviewCopySamePath"));
            if (File.Exists(target)) throw new IOException(Loc.T("I18n_ReviewCopyExists"));
        }

        public static string SaveAsCopy(Word.Document document, string destination)
        {
            int format = SaveFormatFor(document);
            string extension = ExtensionFor(format);
            string original = document.FullName;
            ValidateTarget(original, destination, extension);
            // Native SaveAs keeps unsaved content, formatting, comments, and revision
            // history. The active document becomes the new copy; the source file stays.
            bool removePersonalInformation = document.RemovePersonalInformation;
            try
            {
                // Keep revision authors in the new copy; do not save this option into the source file.
                document.RemovePersonalInformation = false;
                document.SaveAs2(FileName: Path.GetFullPath(destination), FileFormat: format, AddToRecentFiles: false);
            }
            catch { document.RemovePersonalInformation = removePersonalInformation; throw; }
            return original;
        }
    }
}
