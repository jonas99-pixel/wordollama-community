using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.lib.Tools
{
    /// <summary>
    /// [技能] 读取技能说明书 — 当大模型判断任务需要某专业技能时，先调用此工具加载 SOP。
    /// 实现渐进式披露：Prompt 中只有简短目录，真正指令按需加载，避免上下文丢失。
    /// </summary>
    public class ReadSkillTool : IAgentTool
    {
        private SkillManager _skillManager;

        /// <summary>无参构造 — 供反射发现用，后续由 DocumentAgent 注入 SkillManager。</summary>
        public ReadSkillTool() { }

        public ReadSkillTool(SkillManager skillManager)
        {
            _skillManager = skillManager;
        }

        /// <summary>注入 SkillManager（无参构造后调用）。</summary>
        public void SetSkillManager(SkillManager manager)
        {
            _skillManager = manager;
        }

        public string Name => "read_skill";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => false;
        public bool IsConcurrencySafe => !IsWriteOperation;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
                ["properties"] = new JObject
                {
                ["skill_name"] = new JObject { ["type"] = "string" },
                ["reference"] = new JObject { ["type"] = "string" }
            },
            ["required"] = new JArray("skill_name")
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            string skillName = parameters["skill_name"]?.ToString()?.Trim() ?? "";
            if (string.IsNullOrEmpty(skillName))
                return AgentToolText.T("I18n_AgentResultSkillNameRequired");
            string reference = parameters["reference"]?.ToString()?.Trim();
            return string.IsNullOrWhiteSpace(reference)
                ? _skillManager.GetSkillInstructions(skillName)
                : _skillManager.GetSkillReference(skillName, reference);
        }
    }
}
