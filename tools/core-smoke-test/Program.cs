using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using Newtonsoft.Json.Linq;
using WordOllama.lib;
using WordOllama.lib.Providers;

internal static class Program
{
    private static void Main()
    {
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

        Assert(McpToolManager.ParseTransportKind("streamable-http") == McpTransportKind.StreamableHttp, "streamable HTTP transport is parsed");
        Assert(McpToolManager.ParseTransportKind("legacy-sse") == McpTransportKind.Sse, "legacy SSE transport is parsed");
        var environment = McpToolManager.ParseEnvironment(JToken.Parse("{\"TOKEN\":\"value\"}"));
        string token;
        Assert(environment.TryGetValue("TOKEN", out token) && token == "value", "MCP environment configuration is parsed");

        Assert(ToolExecutionPolicy.GetEffects(registry.GetTool("search_text")) == ToolEffect.DocumentRead, "read-only tool policy");
        Assert(ToolExecutionPolicy.GetEffects(registry.GetTool("replace_paragraph")) == ToolEffect.DocumentWrite, "write tool policy");
        Assert(registry.GetTool("execute_command") == null, "external command tools are denied by default");
        Assert(registry.GetTool("run_python_script") == null, "external Python tools are denied by default");
        Assert(registry.GetTool("http_request") == null, "external HTTP tools are denied by default");

        var externalRegistry = new ToolRegistry();
        externalRegistry.DiscoverAndRegister(true);
        Assert(externalRegistry.GetTool("execute_command") != null, "external command tools can be explicitly enabled");
        Assert(externalRegistry.GetTool("run_python_script") != null, "external Python tools can be explicitly enabled");
        Assert(externalRegistry.GetTool("http_request") != null, "external HTTP tools can be explicitly enabled");

        Console.WriteLine("Core smoke test passed.");
    }

    private static void Assert(bool condition, string name)
    {
        if (!condition) throw new InvalidOperationException("Failed: " + name);
    }
}
