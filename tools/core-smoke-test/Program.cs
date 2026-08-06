using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Reflection;
using Newtonsoft.Json.Linq;
using WordOllama.lib;
using WordOllama.lib.Providers;
using WordOllama.lib.Tools;

internal static class Program
{
    private static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("Core smoke test failed: " + ex.GetType().FullName);
            Console.Error.WriteLine(ex.Message ?? "(no message)");
            if (ex.InnerException != null)
            {
                Console.Error.WriteLine("Inner: " + ex.InnerException.GetType().FullName);
                Console.Error.WriteLine(ex.InnerException.Message ?? "(no inner message)");
            }
            Environment.ExitCode = 1;
        }
    }

    private static void Run()
    {
        Assembly systemNetHttp = Assembly.Load(new AssemblyName(
            "System.Net.Http, Version=4.2.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"));
        Assert(systemNetHttp.GetName().Version == new Version(4, 0, 0, 0),
            "System.Net.Http binding redirects to the framework assembly");

        var registry = new ToolRegistry();
        registry.DiscoverAndRegister();
        string catalog = AgentToolCatalogService.BuildMarkdown(registry, true);
        Assert(catalog.Contains("# WordOllama Agent API"), "API catalog header");
        Assert(catalog.Contains("`search_text`"), "search_text is documented");
        Assert(catalog.Contains("`read_comments`"), "read_comments is documented");
        Assert(catalog.Contains("```json"), "tool schemas are documented");

        string skill = AgentToolCatalogService.NormalizeSkillDocument(
            "fiction/ polish", "A repeatable editing workflow", "```markdown\n## Workflow\nUse the catalog tools.\n```");
        Assert(skill.StartsWith("---\nname: fiction-polish\n"), "Skill front matter is normalized");
        Assert(!skill.Contains("```"), "Skill code fence is removed");

        string root = Path.Combine(Path.GetTempPath(), "WordOllamaCoreSmoke");
        string path;
        string pathError;
        Assert(AgentToolCatalogService.TryGetSkillPath(root, "../safe-skill", out path, out pathError), "Skill path is accepted safely");
        Assert(path.StartsWith(Path.GetFullPath(root) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase), "Skill path stays under root");

        string embeddedSkillRoot = Path.Combine(Path.GetTempPath(), "WordOllamaEmbeddedSkills-" + Guid.NewGuid().ToString("N"));
        var skillManager = new SkillManager();
        skillManager.LoadSkills(embeddedSkillRoot);
        Assert(skillManager.HasSkill("contract-reviewer"), "embedded contract-review skill is extracted");
        string riskReference = skillManager.GetSkillReference("contract-reviewer", "references/risk-guidelines.md");
        Assert(riskReference.Contains("contract-reviewer") && !riskReference.Contains("not found"),
            "embedded skill reference is extracted and readable");
        string deniedReference = skillManager.GetSkillReference("contract-reviewer", "../SKILL.md");
        Assert(!deniedReference.Contains("# contract-reviewer: ../SKILL.md"), "skill references cannot escape the skill directory");

        string rateLimit = ProviderRequestPolicy.NormalizeError("OpenAI", (HttpStatusCode)429, "rate limited");
        Assert(rateLimit.Contains(":rate_limit]"), "429 is classified as rate limit");
        Assert(ProviderRequestPolicy.IsTransientError(rateLimit), "rate limit is retryable");
        Assert(!ProviderRequestPolicy.IsTransientError("[API 错误] [OpenAI:authentication] denied"), "authentication is not retryable");

        AppConfig migratedConfig = AppConfig.Deserialize("{\"AIMode\":\"Online\",\"ChatGPTModel\":\"legacy-model\"}");
        Assert(migratedConfig.ChatGPTModel == "legacy-model", "existing config values survive migration");
        Assert(migratedConfig.ProviderSupportsStreaming && migratedConfig.ProviderSupportsVision && migratedConfig.ProviderSupportsJsonOutput,
            "new capability defaults are applied to legacy config");

        ProviderCapabilityProfile profile = ProviderRequestPolicy.ReadCapabilities(JObject.Parse("{\"SupportsStreaming\":false,\"SupportsVision\":false,\"SupportsJsonOutput\":false,\"ContextWindow\":8192}"), "OpenAI");
        Assert(!profile.SupportsStreaming && !profile.SupportsVision && !profile.SupportsJsonOutput && profile.ContextWindow == 8192,
            "saved capability profile is preserved");
        Assert(ProviderFactory.Create("OpenAI", "https://example.invalid/v1/chat/completions", "key", "model", true) is OpenAIProvider,
            "OpenAI-compatible provider conversion");

        var edits = TextDiffService.CreateEdits("Payment is due in 30 days.", "Payment is due in 15 business days.");
        Assert(ApplyEdits("Payment is due in 30 days.", edits) == "Payment is due in 15 business days.",
            "character diff applies only focused text edits");
        string complexOriginal = "If you already have your text and content. WPS AI can also help you improve it. Whether you need it shorter, I've got you covered.";
        string complexRevised = "If you already have your text and content ready, WPS AI can step in to help. Whether you need more detail or a new format, I've got you covered.";
        Assert(ApplyEdits(complexOriginal, TextDiffService.CreateEdits(complexOriginal, complexRevised)) == complexRevised,
            "multiple precise edits reconstruct the displayed result exactly");
        Assert(TextDiffService.PreserveStructuralSuffix("Clause text\r", "Revised clause") == "Revised clause\r",
            "paragraph marker is preserved during precise replacement");
        var paragraphChanges = DocumentDiffPlanner.CreateChanges(
            new[] { "First", "Second", "Third" },
            new[] { "Inserted", "First", "Second", "Third" });
        Assert(paragraphChanges.Count == 1 && paragraphChanges[0].Type == "新增" && paragraphChanges[0].ParagraphIndex == 1,
            "paragraph insertion does not misclassify all following paragraphs as modified");
        var deletionChanges = DocumentDiffPlanner.CreateChanges(
            new[] { "First", "Remove", "Last" },
            new[] { "First", "Last" });
        Assert(deletionChanges.Count == 1 && deletionChanges[0].Type == "删除" && deletionChanges[0].ParagraphIndex == 2,
            "paragraph deletion keeps the original location without shifting later paragraphs");

        string batchPrompt = PromptTemplateService.BuildReviewSuggestionBatchPrompt(
            new[]
            {
                new KeyValuePair<string, string>("1", "First paragraph"),
                new KeyValuePair<string, string>("2", "Second paragraph")
            },
            "Polish the text",
            "Keep defined terms unchanged");
        Assert(batchPrompt.Contains("\"id\":\"1\"") && batchPrompt.Contains("\"id\":\"2\""),
            "review batch prompt preserves item identifiers");

        string translationPrompt = PromptTemplateService.BuildTranslationWindowPrompt(
            "English", "示例文本", false, "Chinese", "Use the defined legal terms consistently.");
        Assert(translationPrompt.Contains("The source language is Chinese") &&
               translationPrompt.Contains("Terminology and style requirements") &&
               translationPrompt.Contains("defined legal terms"),
            "translation prompt includes source language and terminology scheme");

        Assert(McpToolManager.ParseTransportKind("streamable-http") == McpTransportKind.StreamableHttp, "streamable HTTP transport is parsed");
        Assert(McpToolManager.ParseTransportKind("legacy-sse") == McpTransportKind.Sse, "legacy SSE transport is parsed");
        var environment = McpToolManager.ParseEnvironment(JToken.Parse("{\"TOKEN\":\"value\"}"));
        string token;
        Assert(environment.TryGetValue("TOKEN", out token) && token == "value", "MCP environment configuration is parsed");

        Assert(ToolExecutionPolicy.GetEffects(registry.GetTool("search_text")) == ToolEffect.DocumentRead, "read-only tool policy");
        Assert(ToolExecutionPolicy.GetEffects(registry.GetTool("replace_paragraph")) == ToolEffect.DocumentWrite, "write tool policy");
        Assert(TableInsertRowTool.IsValidAfterRow(1, 3) && TableInsertRowTool.IsValidAfterRow(3, 3) &&
            !TableInsertRowTool.IsValidAfterRow(0, 3) && !TableInsertRowTool.IsValidAfterRow(4, 3),
            "table row insertion rejects invalid indexes instead of silently appending");
        Assert(registry.GetTool("format_paragraph") != null, "formatting tools are compiled and registered");
        Assert(registry.GetTool("page_setup") != null, "page layout tools are compiled and registered");
        Assert(registry.GetTool("read_table") != null, "table tools are compiled and registered");
        Assert(registry.GetTool("execute_command") == null, "external command tools are denied by default");
        Assert(registry.GetTool("run_python_script") == null, "external Python tools are denied by default");
        Assert(registry.GetTool("http_request") == null, "external HTTP tools are denied by default");
        Assert(registry.GetTool("grep") == null, "local file search is denied by default");

        var externalRegistry = new ToolRegistry();
        externalRegistry.DiscoverAndRegister(true);
        Assert(externalRegistry.GetTool("execute_command") != null, "external command tools can be explicitly enabled");
        Assert(externalRegistry.GetTool("run_python_script") != null, "external Python tools can be explicitly enabled");
        Assert(externalRegistry.GetTool("http_request") != null, "external HTTP tools can be explicitly enabled");
        Assert(externalRegistry.GetTool("grep") != null, "local file search can be explicitly enabled");
        Assert((ToolExecutionPolicy.GetEffects(externalRegistry.GetTool("grep")) & ToolEffect.LocalFileRead) != 0,
            "local file search requires local file permission");
        Assert(externalRegistry.GetTool("fetch_url") != null, "safe HTTPS fetch can be explicitly enabled");
        Assert(externalRegistry.GetTool("list_workspace_files") != null, "isolated workspace listing can be explicitly enabled");
        Assert(externalRegistry.GetTool("read_workspace_file") != null, "isolated workspace reads can be explicitly enabled");
        Assert(externalRegistry.GetTool("write_workspace_file") != null, "isolated workspace writes can be explicitly enabled");
        Assert(ToolExecutionPolicy.GetEffects(externalRegistry.GetTool("write_workspace_file")) == ToolEffect.LocalFileWrite,
            "workspace writes require a distinct local file write permission");
        Assert(!ToolExecutionPolicy.RequiresWordCom(ToolExecutionPolicy.GetEffects(externalRegistry.GetTool("write_workspace_file"))),
            "workspace writes never enter the Word COM tracked-changes path");
        Assert(ToolExecutionPolicy.RequiresExternalToolPermission(externalRegistry.GetTool("fetch_url")),
            "safe web fetch still requires explicit external-network permission");
        Assert(externalRegistry.GetTool("run_python") != null && externalRegistry.GetTool("run_node") != null,
            "bounded workspace runtimes can be explicitly enabled");

        var sources = AgentSourceExtractor.Extract("search", "{\"title\":\"Example\",\"url\":\"https://example.com/a\"}");
        Assert(sources.Count == 1 && sources[0].Url == "https://example.com/a", "external sources are extracted and deduplicated");

        string agentPrompt = PromptTemplateService.BuildAgentSystemPrompt("<document></document>", "tools", 1, "skills", "English", "Finish the contract review", "\nOutput preferences:\nConcise");
        Assert(agentPrompt.Contains("Durable Task Goal") && agentPrompt.Contains("Finish the contract review") && agentPrompt.Contains("Output preferences"),
            "Agent goal and memory context are injected");
        Assert(agentPrompt.Contains("You are WordOllama") && !agentPrompt.Contains("Microsoft Word AI Assistant"),
            "English Agent identity is WordOllama rather than Microsoft Word AI");
        string chineseAgentPrompt = PromptTemplateService.BuildAgentSystemPrompt("<document></document>", "tools", 1, "skills", "Chinese");
        Assert(chineseAgentPrompt.Contains("你是 WordOllama") && !chineseAgentPrompt.Contains("微软 Word AI 智能助理"),
            "Chinese Agent identity is WordOllama rather than Microsoft Word AI");
        Assert(PromptTemplateService.BuildPlanDecisionSystemPrompt(false).Contains("你是 WordOllama"),
            "planning prompt keeps the WordOllama identity");

        Console.WriteLine("Core smoke test passed.");
    }

    private static string ApplyEdits(string original, IReadOnlyList<TextEdit> edits)
    {
        string result = original;
        for (int i = edits.Count - 1; i >= 0; i--)
        {
            TextEdit edit = edits[i];
            result = result.Remove(edit.OriginalStart, edit.DeleteLength)
                .Insert(edit.OriginalStart, edit.InsertText ?? "");
        }
        return result;
    }

    private static void Assert(bool condition, string name)
    {
        if (!condition) throw new InvalidOperationException("Failed: " + name);
    }
}
