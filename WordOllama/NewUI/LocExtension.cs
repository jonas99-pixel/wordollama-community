using System;
using System.Windows.Markup;
using WordOllama.lib;

namespace WordOllama.NewUI
{
    [MarkupExtensionReturnType(typeof(string))]
    public class LocExtension : MarkupExtension
    {
        public LocExtension()
        {
        }

        public LocExtension(string key)
        {
            Key = key;
        }

        public string Key { get; set; }

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return string.IsNullOrWhiteSpace(Key) ? "" : Loc.T(Key);
        }
    }
}
