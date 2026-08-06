using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.lib.Tools
{
    public sealed class CreateSkillTool : IAgentTool
    {
        private readonly SkillManager _manager;
        public CreateSkillTool(SkillManager manager) { _manager = manager ?? throw new ArgumentNullException(nameof(manager)); }
        public string Name => "create_skill";
        public string Description => "Create a reusable local Word Agent Skill from a successful workflow. Never include document text, secrets, personal data, or raw conversation history.";
        public bool IsWriteOperation => false;
        public bool IsConcurrencySafe => false;
        public int MaxResultSizeChars => 4000;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["name"] = new JObject { ["type"] = "string" },
                ["description"] = new JObject { ["type"] = "string" },
                ["instructions"] = new JObject { ["type"] = "string" }
            },
            ["required"] = new JArray("name", "description", "instructions")
        };
        public string Execute(JObject p, Word.Document doc, Word.Application app)
        {
            string name = p?["name"]?.ToString()?.Trim() ?? "";
            string description = p?["description"]?.ToString()?.Trim() ?? "";
            string instructions = p?["instructions"]?.ToString()?.Trim() ?? "";
            string root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "WordAgentSkills");
            if (!AgentToolCatalogService.TryGetSkillPath(root, name, out string directory, out string error)) return "[Error] " + error;
            if (instructions.Length > 50000) return "[Error] Skill instructions are too long.";
            string lower = instructions.ToLowerInvariant();
            if (lower.Contains("api key") || lower.Contains("password") || lower.Contains("密码") || lower.Contains("密钥"))
                return "[Error] Skill content appears to contain sensitive data.";
            Directory.CreateDirectory(directory);
            string file = Path.Combine(directory, "SKILL.md");
            if (File.Exists(file)) return "[Error] A Skill with this name already exists.";
            File.WriteAllText(file, AgentToolCatalogService.NormalizeSkillDocument(name, description, instructions), new UTF8Encoding(false));
            _manager.LoadSkills(root);
            return "Skill saved: " + file;
        }
    }
}
