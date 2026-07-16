using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace WordOllama.lib
{
    /// <summary>
    /// Builds the public Agent tool catalog from the same registry used at runtime.
    /// The catalog documents safe wrappers and their effects; it never exposes raw COM objects.
    /// </summary>
    public static class AgentToolCatalogService
    {
        public static string BuildMarkdown(ToolRegistry registry, bool english)
        {
            var tools = (registry?.GetAllTools() ?? Enumerable.Empty<IAgentTool>())
                .OrderBy(tool => tool.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();
            var sb = new StringBuilder();

            sb.AppendLine(AgentToolText.T("I18n_AgentCatalogTitle"));
            sb.AppendLine();
            sb.AppendLine(AgentToolText.T("I18n_AgentCatalogIntro"));
            sb.AppendLine();
            sb.AppendLine(AgentToolText.T("I18n_AgentCatalogContractTitle"));
            sb.AppendLine(AgentToolText.T("I18n_AgentCatalogContract"));
            sb.AppendLine();
            sb.AppendLine(AgentToolText.T("I18n_AgentCatalogSafetyTitle"));
            sb.AppendLine(AgentToolText.T("I18n_AgentCatalogSafety"));
            sb.AppendLine();
            sb.AppendLine(AgentToolText.F("I18n_AgentCatalogToolsTitle", tools.Count));
            sb.AppendLine();

            foreach (var tool in tools)
            {
                ToolEffect effects = ToolExecutionPolicy.GetEffects(tool);
                bool usesWord = ToolExecutionPolicy.RequiresWordCom(effects);
                string effectText = ToolExecutionPolicy.DescribeEffects(effects, english);
                sb.AppendLine($"### `{tool.Name}`");
                sb.AppendLine(AgentToolText.F("I18n_AgentCatalogDescription", tool.Description));
                sb.AppendLine(AgentToolText.F("I18n_AgentCatalogEffects", effectText));
                sb.AppendLine(AgentToolText.F("I18n_AgentCatalogWordCom", usesWord ? AgentToolText.T("I18n_AgentCatalogYes") : AgentToolText.T("I18n_AgentCatalogNo")));
                sb.AppendLine(AgentToolText.F("I18n_AgentCatalogWrite", tool.IsWriteOperation ? AgentToolText.T("I18n_AgentCatalogYes") : AgentToolText.T("I18n_AgentCatalogNo")));
                sb.AppendLine(AgentToolText.F("I18n_AgentCatalogConcurrency", tool.IsConcurrencySafe ? AgentToolText.T("I18n_AgentCatalogParallelRead") : AgentToolText.T("I18n_AgentCatalogSerialized")));
                if (tool.MaxResultSizeChars > 0)
                    sb.AppendLine(AgentToolText.F("I18n_AgentCatalogResultLimit", tool.MaxResultSizeChars));
                sb.AppendLine();
                sb.AppendLine("```json");
                sb.AppendLine(AgentToolText.LocalizeSchema(tool.Name, tool.ParameterSchema)?.ToString(Formatting.Indented) ?? "{}");
                sb.AppendLine("```");
                sb.AppendLine();
            }

            sb.AppendLine(AgentToolText.T("I18n_AgentCatalogSkillRulesTitle"));
            sb.AppendLine(AgentToolText.T("I18n_AgentCatalogSkillRules"));
            return sb.ToString();
        }

        public static string BuildSkillAuthoringPrompt(string apiDocumentation, string skillName, string description, string goal, bool english)
        {
            return PromptTemplateService.BuildSkillAuthoringPrompt(
                TrimForPrompt(apiDocumentation), skillName, description, goal, english);
        }

        public static string NormalizeSkillDocument(string skillName, string description, string draft)
        {
            string safeName = MakeSafeFolderName(skillName);
            if (string.IsNullOrWhiteSpace(safeName))
                throw new InvalidDataException("Skill name is invalid.");
            string safeDescription = CleanFrontMatterValue(description);
            if (string.IsNullOrWhiteSpace(safeDescription))
                safeDescription = safeName;

            string body = StripFrontMatterAndFence(draft).Trim();
            if (body.Length == 0)
                throw new InvalidDataException("Skill body is empty.");

            return $"---\nname: {safeName}\ndescription: {safeDescription}\n---\n\n{body}\n";
        }

        public static bool TryGetSkillPath(string rootPath, string displayName, out string skillDirectory, out string error)
        {
            skillDirectory = null;
            error = null;
            string folderName = MakeSafeFolderName(displayName);
            if (folderName == null)
            {
                error = "Skill name must contain letters or numbers.";
                return false;
            }

            string fullRoot = Path.GetFullPath(rootPath);
            string candidate = Path.GetFullPath(Path.Combine(fullRoot, folderName));
            if (!candidate.StartsWith(fullRoot.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            {
                error = "Invalid Skill path.";
                return false;
            }

            skillDirectory = candidate;
            return true;
        }

        private static string MakeSafeFolderName(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;
            string folder = Regex.Replace(value.Trim(), @"[^\p{L}\p{Nd}_-]+", "-");
            folder = folder.Trim('-', '_', '.');
            if (folder.Length == 0 || folder == "." || folder == "..") return null;
            return folder.Length > 80 ? folder.Substring(0, 80).Trim('-', '_', '.') : folder;
        }

        private static string CleanFrontMatterValue(string value)
        {
            string clean = (value ?? "").Replace("\r", " ").Replace("\n", " ").Replace("---", "-").Trim();
            return clean.Length > 160 ? clean.Substring(0, 160) : clean;
        }

        private static string StripFrontMatterAndFence(string draft)
        {
            string text = (draft ?? "").Trim();
            Match fence = Regex.Match(text, @"^```(?:markdown|md)?\s*([\s\S]*?)\s*```$", RegexOptions.IgnoreCase);
            if (fence.Success) text = fence.Groups[1].Value.Trim();

            if (text.StartsWith("---", StringComparison.Ordinal))
            {
                int end = text.IndexOf("\n---", 3, StringComparison.Ordinal);
                if (end >= 0)
                    text = text.Substring(end + 4).Trim();
            }
            return text;
        }

        private static string TrimForPrompt(string value)
        {
            if (string.IsNullOrEmpty(value)) return "(no tools discovered)";
            const int maxLength = 60000;
            return value.Length <= maxLength ? value : value.Substring(0, maxLength) + "\n[Catalog truncated]";
        }
    }
}
