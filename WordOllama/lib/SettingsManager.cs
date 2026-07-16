using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace WordOllama.lib
{
    /// <summary>
    /// One-time importer for pre-AppConfig installations. It is not a runtime settings service.
    /// </summary>
    internal static class SettingsManager
    {
        private static string LegacySettingsPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "WordOllama",
            "DefaultSettings.ini");

        internal static void LoadLegacyIniIntoProperties()
        {
            if (!File.Exists(LegacySettingsPath) || new FileInfo(LegacySettingsPath).Length == 0) return;
            var settings = Properties.Settings.Default;
            var settingProperties = settings.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.CanWrite && p.CanRead);

            // 使用 StreamReader 按行读取
            using (var reader = new StreamReader(LegacySettingsPath, Encoding.UTF8))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    var match = Regex.Match(line, @"^(?<key>.+?)=(?<value>.+)$");
                    if (match.Success)
                    {
                        var key = match.Groups["key"].Value;
                        var value = match.Groups["value"].Value;

                        var property = settingProperties.FirstOrDefault(p => p.Name == key);
                        if (property != null)
                        {
                            try
                            {
                                // 确保值可以转换为正确的类型
                                var convertedValue = Convert.ChangeType(value, property.PropertyType);
                                property.SetValue(settings, convertedValue);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error loading setting '{key}': {ex.Message}");
                            }
                        }
                    }
                }
            }
        }

        internal static void RemoveLegacyIni()
        {
            try
            {
                if (File.Exists(LegacySettingsPath))
                    File.Delete(LegacySettingsPath);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Unable to remove legacy settings: {ex.Message}");
            }
        }

    }
}
