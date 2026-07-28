using System;
using System.Globalization;
using System.Threading;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.lib
{
    public static class UiCulture
    {
        private static readonly CultureInfo English = CultureInfo.GetCultureInfo("en");
        private static readonly CultureInfo Chinese = CultureInfo.GetCultureInfo("zh-CN");

        public static CultureInfo Current { get; private set; } = Resolve(CultureInfo.CurrentUICulture);

        public static void Initialize(Word.Application application = null)
        {
            CultureInfo source = null;
            try
            {
                int languageId = application == null
                    ? 0
                    : application.LanguageSettings.LanguageID[MsoAppLanguageID.msoLanguageIDUI];
                if (languageId > 0)
                    source = CultureInfo.GetCultureInfo(languageId);
            }
            catch
            {
                // Word may not be fully initialized while the Ribbon is being constructed.
            }

            Current = Resolve(source ?? CultureInfo.CurrentUICulture);
            Thread.CurrentThread.CurrentUICulture = Current;
            CultureInfo.DefaultThreadCurrentUICulture = Current;
            WordOllama.Properties.Resources.Culture = Current;
        }

        private static CultureInfo Resolve(CultureInfo culture)
        {
            return string.Equals(culture?.TwoLetterISOLanguageName, "zh", StringComparison.OrdinalIgnoreCase)
                ? Chinese
                : English;
        }
    }

    public static class Loc
    {
        public static string T(string key)
        {
            string value = WordOllama.Properties.Resources.ResourceManager.GetString(key, UiCulture.Current);
            return string.IsNullOrEmpty(value) ? key : value;
        }

        public static string F(string key, params object[] args)
        {
            return string.Format(T(key), args);
        }
    }
}
