using System.Collections.Generic;
using System;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;

namespace WordOllama.lib
{
    public class AgentSkill
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string Instructions { get; set; }
        public string FilePath { get; set; }
    }

    /// <summary>
    /// 技能管理器 — 扫描 SKILL.md 文件，渐进式披露专业工作流。
    /// 支持两种来源：
    ///   1. 内置技能（嵌入在程序集中的 Resources/Skills/ 目录）
    ///   2. 外部技能（用户 ~/Documents/WordAgentSkills/ 目录）
    ///
    /// 内置技能在首次运行时自动提取到外部目录，用户可自由修改覆盖。
    /// </summary>
    public class SkillManager
    {
        private const string EMBEDDED_SKILLS_PREFIX = "WordOllama.Resources.Skills.";
        private const string SKILL_FILE_NAME = "SKILL.md";

        private readonly Dictionary<string, AgentSkill> _skills =
            new Dictionary<string, AgentSkill>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// 扫描指定目录下所有 SKILL.md 文件（含首次自动提取内置技能）。
        /// </summary>
        public void LoadSkills(string rootPath)
        {
            _skills.Clear();

            // 首次运行：将嵌入程序集中的内置 skill 提取到磁盘
            SeedEmbeddedSkills(rootPath);

            if (!Directory.Exists(rootPath)) return;

            foreach (var file in Directory.GetFiles(rootPath, SKILL_FILE_NAME, SearchOption.AllDirectories))
            {
                try
                {
                    string content = File.ReadAllText(file);
                    AgentSkill skill;
                    if (TryParseSkill(content, file, out skill))
                    {
                        _skills[skill.Name] = skill;
                    }
                }
                catch { }
            }
        }

        private static bool TryParseSkill(string content, string filePath, out AgentSkill skill)
        {
            skill = null;
            var frontMatter = Regex.Match(content ?? "",
                @"\A---\s*\r?\n(?<header>[\s\S]*?)\r?\n---\s*(?:\r?\n)?(?<body>[\s\S]*)\z");
            if (!frontMatter.Success) return false;

            string header = frontMatter.Groups["header"].Value;
            var nameMatch = Regex.Match(header, @"(?m)^name:\s*(?<value>.+?)\s*$", RegexOptions.IgnoreCase);
            var descriptionMatch = Regex.Match(header, @"(?m)^description:\s*(?<value>.+?)\s*$", RegexOptions.IgnoreCase);
            if (!nameMatch.Success || !descriptionMatch.Success) return false;

            string name = nameMatch.Groups["value"].Value.Trim().Trim('"', '\'');
            if (string.IsNullOrWhiteSpace(name)) return false;

            skill = new AgentSkill
            {
                Name = name,
                Description = descriptionMatch.Groups["value"].Value.Trim().Trim('"', '\''),
                Instructions = frontMatter.Groups["body"].Value.Trim(),
                FilePath = filePath
            };
            return true;
        }

        /// <summary>
        /// 从嵌入资源中提取内置 skill 到目标目录（仅当磁盘上不存在时）。
        /// 用户已自定义的 skill 不会被覆盖。
        /// </summary>
        private void SeedEmbeddedSkills(string rootPath)
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                foreach (string name in assembly.GetManifestResourceNames())
                {
                    if (!name.StartsWith(EMBEDDED_SKILLS_PREFIX, StringComparison.Ordinal))
                        continue;

                    string relativePath = name.Substring(EMBEDDED_SKILLS_PREFIX.Length)
                        .Replace('/', Path.DirectorySeparatorChar);
                    if (string.IsNullOrWhiteSpace(relativePath)) continue;

                    string fullRoot = Path.GetFullPath(rootPath)
                        .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                    string targetFile = Path.GetFullPath(Path.Combine(fullRoot, relativePath));
                    if (!targetFile.StartsWith(fullRoot + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                        continue;

                    // 磁盘上已存在 → 保留用户的版本，不覆盖
                    if (File.Exists(targetFile)) continue;

                    Directory.CreateDirectory(Path.GetDirectoryName(targetFile));
                    using (var stream = assembly.GetManifestResourceStream(name))
                    {
                        if (stream == null) continue;
                        using (var reader = new StreamReader(stream))
                        {
                            File.WriteAllText(targetFile, reader.ReadToEnd());
                        }
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// 生成技能目录（简短列表，注入到 Prompt 中）。
        /// </summary>
        public string GetSkillCatalog()
        {
            if (_skills.Count == 0) return AgentToolText.T("I18n_AgentSkillCatalogEmpty");
            var lines = new List<string>();
            foreach (var s in _skills.Values)
                lines.Add($"- {s.Name}: {s.Description}");
            return string.Join("\n", lines);
        }

        /// <summary>
        /// 获取指定技能的完整指令（Markdown 正文）。
        /// </summary>
        public string GetSkillInstructions(string name)
        {
            if (_skills.TryGetValue(name, out var skill))
                return AgentToolText.F("I18n_AgentSkillHeader", skill.Name, skill.Instructions);
            return AgentToolText.F("I18n_AgentResultSkillNotFound", name, string.Join(", ", _skills.Keys));
        }

        public string GetSkillReference(string name, string reference)
        {
            if (!_skills.TryGetValue(name, out var skill))
                return AgentToolText.F("I18n_AgentResultSkillNotFound", name, string.Join(", ", _skills.Keys));
            if (string.IsNullOrWhiteSpace(reference))
                return GetSkillInstructions(name);

            try
            {
                string root = Path.GetFullPath(Path.GetDirectoryName(skill.FilePath))
                    .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                string target = Path.GetFullPath(Path.Combine(root, reference));
                if (!target.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                    return AgentToolText.F("I18n_AgentResultSkillReferenceDenied", reference);
                if (!File.Exists(target))
                    return AgentToolText.F("I18n_AgentResultSkillReferenceNotFound", reference);

                string extension = Path.GetExtension(target)?.ToLowerInvariant();
                if (extension != ".md" && extension != ".txt")
                    return AgentToolText.F("I18n_AgentResultSkillReferenceDenied", reference);
                if (new FileInfo(target).Length > 128 * 1024)
                    return AgentToolText.F("I18n_AgentResultSkillReferenceTooLarge", reference);

                return AgentToolText.F("I18n_AgentSkillReferenceHeader", skill.Name, reference, File.ReadAllText(target));
            }
            catch
            {
                return AgentToolText.F("I18n_AgentResultSkillReferenceDenied", reference);
            }
        }

        public bool HasSkill(string name) => _skills.ContainsKey(name);
        public int Count => _skills.Count;
        public IEnumerable<AgentSkill> GetAllSkills() => _skills.Values;
    }
}
