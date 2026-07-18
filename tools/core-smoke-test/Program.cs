using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Reflection;
using Newtonsoft.Json.Linq;
using WordOllama.lib;
using WordOllama.lib.Providers;

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
        Assert(TextDiffService.PreserveStructuralSuffix("Clause text\r", "Revised clause") == "Revised clause\r",
            "paragraph marker is preserved during precise replacement");

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

        Assert(McpToolManager.ParseTransportKind("streamable-http") == McpTransportKind.StreamableHttp, "streamable HTTP transport is parsed");
        Assert(McpToolManager.ParseTransportKind("legacy-sse") == McpTransportKind.Sse, "legacy SSE transport is parsed");
        var environment = McpToolManager.ParseEnvironment(JToken.Parse("{\"TOKEN\":\"value\"}"));
        string token;
        Assert(environment.TryGetValue("TOKEN", out token) && token == "value", "MCP environment configuration is parsed");

        Assert(ToolExecutionPolicy.GetEffects(registry.GetTool("search_text")) == ToolEffect.DocumentRead, "read-only tool policy");
        Assert(ToolExecutionPolicy.GetEffects(registry.GetTool("replace_paragraph")) == ToolEffect.DocumentWrite, "write tool policy");
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
