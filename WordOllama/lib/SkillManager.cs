using System.Collections.Generic;
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

        private readonly Dictionary<string, AgentSkill> _skills = new Dictionary<string, AgentSkill>();

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
                    var match = Regex.Match(content,
                        @"^---\s*[\r\n]+name:\s*(.+?)[\r\n]+description:\s*(.+?)[\r\n]+---\s*[\r\n]+(.*)",
                        RegexOptions.Singleline);

                    if (match.Success)
                    {
                        var skill = new AgentSkill
                        {
                            Name = match.Groups[1].Value.Trim(),
                            Description = match.Groups[2].Value.Trim(),
                            Instructions = match.Groups[3].Value.Trim(),
                            FilePath = file
                        };
                        if (!string.IsNullOrEmpty(skill.Name))
                            _skills[skill.Name] = skill;
                    }
                }
                catch { }
            }
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
                string suffix = "." + SKILL_FILE_NAME; // ".SKILL.md"

                foreach (string name in assembly.GetManifestResourceNames())
                {
                    if (!name.StartsWith(EMBEDDED_SKILLS_PREFIX) || !name.EndsWith(suffix))
                        continue;

                    // 从 "WordOllama.Resources.Skills.{skillName}.SKILL.md" 提取 skill 名称
                    string skillName = name.Substring(
                        EMBEDDED_SKILLS_PREFIX.Length,
                        name.Length - EMBEDDED_SKILLS_PREFIX.Length - suffix.Length);
                    if (string.IsNullOrEmpty(skillName)) continue;

                    string targetDir = Path.Combine(rootPath, skillName);
                    string targetFile = Path.Combine(targetDir, SKILL_FILE_NAME);

                    // 磁盘上已存在 → 保留用户的版本，不覆盖
                    if (File.Exists(targetFile)) continue;

                    Directory.CreateDirectory(targetDir);
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
            if (_skills.Count == 0) return "（暂无扩展技能）";
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
                return $"# 技能: {skill.Name}\n\n{skill.Instructions}";
            return $"[错误] 未找到技能 \"{name}\"。可用技能: {string.Join(", ", _skills.Keys)}";
        }

        public bool HasSkill(string name) => _skills.ContainsKey(name);
        public int Count => _skills.Count;
        public IEnumerable<AgentSkill> GetAllSkills() => _skills.Values;
    }
}
