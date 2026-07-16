using System;
using System.IO;
using Newtonsoft.Json;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.NewUI
{
    internal class WritingProfile
    {
        public string Instructions { get; set; } = "";
    }

    internal static class WritingProfileManager
    {
        private static string ProfileDir => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "WordOllama",
            "Profiles");

        private static string DefaultProfilePath => Path.Combine(ProfileDir, "default-writing-profile.json");

        public static WritingProfile Load()
        {
            try
            {
                if (!File.Exists(DefaultProfilePath)) return new WritingProfile();
                var json = File.ReadAllText(DefaultProfilePath);
                return JsonConvert.DeserializeObject<WritingProfile>(json) ?? new WritingProfile();
            }
            catch
            {
                return new WritingProfile();
            }
        }

        public static void Save(WritingProfile profile)
        {
            try
            {
                Directory.CreateDirectory(ProfileDir);
                File.WriteAllText(DefaultProfilePath, JsonConvert.SerializeObject(profile ?? new WritingProfile(), Formatting.Indented));
            }
            catch { }
        }

        public static WritingProfile LoadForDocument(Word.Document document)
        {
            string sidecarPath = GetDocumentProfilePath(document);
            if (string.IsNullOrEmpty(sidecarPath)) return Load();

            try
            {
                if (!File.Exists(sidecarPath)) return Load();
                return JsonConvert.DeserializeObject<WritingProfile>(File.ReadAllText(sidecarPath)) ?? new WritingProfile();
            }
            catch
            {
                return Load();
            }
        }

        public static void SaveForDocument(Word.Document document, WritingProfile profile)
        {
            string sidecarPath = GetDocumentProfilePath(document);
            if (string.IsNullOrEmpty(sidecarPath))
            {
                Save(profile);
                return;
            }

            try
            {
                File.WriteAllText(sidecarPath, JsonConvert.SerializeObject(profile ?? new WritingProfile(), Formatting.Indented));
            }
            catch
            {
                Save(profile);
            }
        }

        private static string GetDocumentProfilePath(Word.Document document)
        {
            try
            {
                string fullName = document?.FullName;
                if (string.IsNullOrWhiteSpace(fullName) || !Path.IsPathRooted(fullName)) return null;
                return fullName + ".wordollama.json";
            }
            catch
            {
                return null;
            }
        }
    }
}
