using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace WordOllama.lib
{
    public sealed class AgentSource
    {
        public string Url { get; set; }
        public string Title { get; set; }
        public string Tool { get; set; }
        public DateTime RetrievedAt { get; set; } = DateTime.UtcNow;
    }

    public static class AgentSourceExtractor
    {
        private static readonly Regex UrlRegex = new Regex(@"https://[^\s\]\[\)\(<>\""']+", RegexOptions.IgnoreCase | RegexOptions.Compiled, TimeSpan.FromSeconds(1));

        public static IReadOnlyList<AgentSource> Extract(string tool, string content)
        {
            var found = new Dictionary<string, AgentSource>(StringComparer.OrdinalIgnoreCase);
            string title = "";
            try
            {
                JToken token = JToken.Parse(content ?? "");
                Walk(token, tool, found, ref title);
            }
            catch { }
            foreach (Match match in UrlRegex.Matches(content ?? ""))
            {
                if (TryPublicHttps(match.Value.TrimEnd('.', ',', ';', ':'), out Uri uri))
                    found[uri.AbsoluteUri] = new AgentSource { Url = uri.AbsoluteUri, Title = title, Tool = tool };
            }
            return found.Values.Take(50).ToList();
        }

        private static void Walk(JToken token, string tool, IDictionary<string, AgentSource> found, ref string title)
        {
            if (token is JObject obj)
            {
                string localTitle = obj["title"]?.ToString() ?? obj["name"]?.ToString() ?? title;
                foreach (JProperty property in obj.Properties())
                {
                    if ((property.Name.IndexOf("url", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         property.Name.IndexOf("link", StringComparison.OrdinalIgnoreCase) >= 0) &&
                        TryPublicHttps(property.Value?.ToString(), out Uri uri))
                    {
                        found[uri.AbsoluteUri] = new AgentSource { Url = uri.AbsoluteUri, Title = localTitle, Tool = tool };
                    }
                    Walk(property.Value, tool, found, ref localTitle);
                }
                title = localTitle;
            }
            else if (token is JArray array)
            {
                foreach (JToken child in array) Walk(child, tool, found, ref title);
            }
        }

        private static bool TryPublicHttps(string value, out Uri uri)
        {
            if (Uri.TryCreate(value, UriKind.Absolute, out uri) && uri.Scheme == Uri.UriSchemeHttps && string.IsNullOrEmpty(uri.UserInfo))
                return true;
            uri = null;
            return false;
        }
    }
}
