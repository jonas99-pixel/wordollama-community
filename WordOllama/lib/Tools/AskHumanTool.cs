using Newtonsoft.Json.Linq;
using System.Threading;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.lib.Tools
{
    /// <summary>
    /// [HITL] 人工确认工具 — 当 Agent 遇到不确定的情况时，主动向用户提问。
    /// 执行后挂起 ReAct 循环，直到用户通过 UI 提交回答。
    /// 实际的挂起逻辑由 DocumentAgent 主循环直接处理，此工具仅注册名称和 Schema。
    /// </summary>
    public class AskHumanTool : IAgentTool
    {
        public string Name => "ask_human";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => false;
        public bool IsConcurrencySafe => !IsWriteOperation;
        public int MaxResultSizeChars => 0;

        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["question"] = new JObject { ["type"] = "string" }
            },
            ["required"] = new JArray("question")
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            // 实际的 HITL 逻辑由 DocumentAgent 主循环直接处理（参见 ExecuteAsync 中 ask_human 分支）。
            // 此 Execute 仅在工具意外被直接调用时提供兜底。
            string question = parameters["question"]?.ToString() ?? "";
            return string.IsNullOrEmpty(question)
                ? AgentToolText.T("I18n_AgentResultQuestionRequired")
                : AgentToolText.F("I18n_AgentResultSubmitQuestion", question);
        }
    }
}
