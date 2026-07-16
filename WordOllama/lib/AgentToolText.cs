using System.Globalization;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using WordOllama.Forms;

namespace WordOllama.lib
{
    internal static class AgentToolText
    {
        private static string _currentAgentLanguage;

        internal static void SetCurrentAgentLanguage(string language)
        {
            _currentAgentLanguage = language;
        }

        private static bool UseEnglish => PromptTemplateService.ResolveUseEnglish(
            _currentAgentLanguage ?? SharedSettings.LanguageMode,
            null);

        public static string T(string key)
        {
            CultureInfo culture = UseEnglish ? CultureInfo.GetCultureInfo("en") : CultureInfo.GetCultureInfo("zh-CN");
            return Properties.Resources.ResourceManager.GetString(key, culture) ?? key;
        }

        public static string Description(string toolName)
        {
            return T("I18n_AgentToolDesc_" + toolName);
        }

        public static string F(string key, params object[] args)
        {
            return string.Format(T(key), args);
        }

        public static JObject LocalizeSchema(string toolName, JObject schema)
        {
            if (schema == null) return null;
            var localized = (JObject)schema.DeepClone();
            var properties = localized["properties"] as JObject;
            if (properties == null) return localized;

            foreach (var property in properties.Properties())
            {
                if (!(property.Value is JObject definition)) continue;
                string key = "I18n_AgentToolParam_" + toolName + "_" + property.Name;
                string description = T(key);
                if (description == key)
                {
                    description = T("I18n_AgentToolParam_" + property.Name);
                }
                if (description == "I18n_AgentToolParam_" + property.Name)
                    description = string.Format(T("I18n_AgentToolParameterFallback"), HumanizeParameterName(property.Name));
                definition["description"] = description;
            }
            return localized;
        }

        private static string HumanizeParameterName(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return "value";
            string spaced = Regex.Replace(value.Replace('_', ' '), "(?<!^)([A-Z])", " $1");
            return spaced.Trim();
        }
    }
}
