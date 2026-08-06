using System;
using WordOllama.lib.Tools;

namespace WordOllama.lib
{
    [Flags]
    public enum ToolEffect
    {
        None = 0,
        DocumentRead = 1,
        DocumentWrite = 2,
        ExternalNetwork = 4,
        LocalProcess = 8,
        ExternalMcp = 16,
        LocalFileRead = 32,
        LocalFileWrite = 64
    }

    public enum AgentExecutionMode
    {
        ViewOnly,
        ProposeChanges,
        TrackedChanges
    }

    /// <summary>
    /// Centralizes Agent tool effects and execution rules. Tool descriptions are not a security boundary.
    /// </summary>
    public static class ToolExecutionPolicy
    {
        public static ToolEffect GetEffects(IAgentTool tool)
        {
            if (tool == null) return ToolEffect.None;

            if (tool is ExecuteCommandTool || tool is RunPythonScriptTool || tool is RunWorkspaceCodeTool)
                return ToolEffect.LocalProcess;
            if (tool is HttpRequestTool || tool is FetchUrlTool)
                return ToolEffect.ExternalNetwork;
            if (tool is WriteWorkspaceFileTool || tool is CreateSkillTool)
                return ToolEffect.LocalFileWrite;
            if (tool is WorkspaceToolBase)
                return ToolEffect.LocalFileRead;
            if (tool is McpToolWrapper)
                return ToolEffect.ExternalMcp;
            if (tool is GrepTool)
                return ToolEffect.DocumentRead | ToolEffect.LocalFileRead;

            return tool.IsWriteOperation ? ToolEffect.DocumentWrite : ToolEffect.DocumentRead;
        }

        public static bool RequiresWordCom(ToolEffect effects)
        {
            return (effects & (ToolEffect.DocumentRead | ToolEffect.DocumentWrite)) != 0;
        }

        public static bool RequiresExternalToolPermission(IAgentTool tool)
        {
            if (tool == null) return false;
            return (GetEffects(tool) & (ToolEffect.ExternalNetwork | ToolEffect.LocalProcess | ToolEffect.ExternalMcp | ToolEffect.LocalFileRead | ToolEffect.LocalFileWrite)) != 0;
        }

        public static AgentExecutionMode ParseMode(string value)
        {
            if (string.Equals(value, "ViewOnly", StringComparison.OrdinalIgnoreCase))
                return AgentExecutionMode.ViewOnly;
            if (string.Equals(value, "ProposeChanges", StringComparison.OrdinalIgnoreCase))
                return AgentExecutionMode.ProposeChanges;
            return AgentExecutionMode.TrackedChanges;
        }

        public static string DescribeEffects(ToolEffect effects, bool english)
        {
            if (effects == ToolEffect.None) return AgentToolText.T("I18n_AgentEffectNone");

            var parts = new System.Collections.Generic.List<string>();
            if ((effects & ToolEffect.DocumentRead) != 0) parts.Add(AgentToolText.T("I18n_AgentEffectDocumentRead"));
            if ((effects & ToolEffect.DocumentWrite) != 0) parts.Add(AgentToolText.T("I18n_AgentEffectDocumentWrite"));
            if ((effects & ToolEffect.ExternalNetwork) != 0) parts.Add(AgentToolText.T("I18n_AgentEffectExternalNetwork"));
            if ((effects & ToolEffect.LocalProcess) != 0) parts.Add(AgentToolText.T("I18n_AgentEffectLocalProcess"));
            if ((effects & ToolEffect.ExternalMcp) != 0) parts.Add(AgentToolText.T("I18n_AgentEffectExternalMcp"));
            if ((effects & ToolEffect.LocalFileRead) != 0) parts.Add(AgentToolText.T("I18n_AgentEffectLocalFileRead"));
            if ((effects & ToolEffect.LocalFileWrite) != 0) parts.Add(AgentToolText.T("I18n_AgentEffectLocalFileWrite"));
            return string.Join(english ? ", " : "、", parts);
        }
    }
}
