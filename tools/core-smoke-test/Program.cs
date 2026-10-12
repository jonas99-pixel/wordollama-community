using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Reflection;
using System.Linq;
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

        Assert(OpenAIProvider.NormalizeEndpoint("http://localhost:30001/v1") == "http://localhost:30001/v1/chat/completions", "base endpoint normalized");
        Assert(OpenAIProvider.NormalizeEndpoint("https://example.com/custom/chat/completions?api-version=1") == "https://example.com/custom/chat/completions?api-version=1", "custom endpoint and query preserved");
        var probe = new CapabilityProbeResult();
        probe.Steps.Add(new CapabilityProbeStep { Name = "普通聊天", Passed = true });
        probe.Steps.Add(new CapabilityProbeStep { Name = "工具调用", Passed = false });
        Assert(probe.Success, "optional capability failure does not invalidate chat connection");
        var failedProbe = new CapabilityProbeResult();
        failedProbe.Steps.Add(new CapabilityProbeStep { Name = "模型列表", Passed = true });
        Assert(!failedProbe.Success, "model listing alone does not prove chat connection");

        int tokens;
        Assert(ProviderRequestPolicy.TryTokenCount("128k", out tokens) && tokens == 131072, "K context input");
        Assert(ProviderRequestPolicy.TryTokenCount("240000", out tokens) && tokens == 240000, "exact context input");
        Assert(!ProviderRequestPolicy.TryTokenCount("2147483647k", out tokens), "context overflow rejected");
        Assert(!ProviderRequestPolicy.TryTokenCount("-1", out tokens), "negative context rejected");
        var tokenMethod = typeof(OpenAIProvider).GetMethod("GetMaxTokenParamName", BindingFlags.Instance | BindingFlags.NonPublic);
        var explicitTokens = new OpenAIProvider("https://example.com/v1", "", "gpt-5", false, null, false, "max_tokens");
        Assert((string)tokenMethod.Invoke(explicitTokens, null) == "max_tokens", "explicit token parameter overrides model inference");
        var automaticTokens = new OpenAIProvider("https://example.com/v1", "", "gpt-5", false, null, false, "auto");
        Assert((string)tokenMethod.Invoke(automaticTokens, null) == "max_completion_tokens", "automatic reasoning model token parameter");
        var tokenProfile = ProviderRequestPolicy.ReadCapabilities(JObject.Parse("{\"OutputTokenParameter\":\"max_completion_tokens\",\"ContextWindow\":240000}"), "OpenAI");
        Assert(tokenProfile.OutputTokenParameter == "max_completion_tokens" && tokenProfile.ContextWindow == 240000, "model token settings round trip");

        string longValue = new string('x', 300);
        var actionA = new JObject(new JProperty("b", longValue), new JProperty("a", new JObject(new JProperty("y", 2), new JProperty("x", 1))));
        var actionB = new JObject(new JProperty("a", new JObject(new JProperty("x", 1), new JProperty("y", 2))), new JProperty("b", longValue));
        Assert(DocumentAgent.CanonicalActionParameters(actionA) == DocumentAgent.CanonicalActionParameters(actionB), "repeat detection handles long parameters and object key order");
        actionB["b"] = longValue + "changed";
        Assert(DocumentAgent.CanonicalActionParameters(actionA) != DocumentAgent.CanonicalActionParameters(actionB), "repeat detection preserves differences after 100 characters");
        Assert(DocumentAgent.CanonicalActionParameters(new JArray(1, 2)) != DocumentAgent.CanonicalActionParameters(new JArray(2, 1)), "repeat detection preserves array order");

        var nativeMessage = JObject.Parse("{\"role\":\"assistant\",\"content\":null,\"tool_calls\":[{\"id\":\"c1\",\"type\":\"function\",\"function\":{\"name\":\"read\",\"arguments\":\"{}\"}},{\"id\":\"c2\",\"type\":\"function\",\"function\":{\"name\":\"read\",\"arguments\":\"{}\"}}]}");
        var nativeCalls = new JArray(AgentWireProtocol.Call("c1", "read", new JObject()), AgentWireProtocol.Call("c2", "read", new JObject()));
        var nativeEnvelope = AgentWireProtocol.Decode(AgentWireProtocol.Encode("OpenAI", nativeMessage, nativeCalls));
        Assert(((JArray)nativeEnvelope["calls"]).Count == 2, "all native calls preserved");
        var wireHistory = new List<Dictionary<string, string>>();
        wireHistory.Add(new Dictionary<string, string> { { "role", "assistant" }, { "content", "" }, { "native", nativeEnvelope.ToString() } });
        wireHistory.Add(AgentWireProtocol.Result("c1", "read", "first", false));
        wireHistory.Add(AgentWireProtocol.Result("c2", "read", "denied", true));
        var wire = AgentWireProtocol.OpenAiMessages(wireHistory);
        Assert(JToken.DeepEquals(wire[0], nativeMessage), "native assistant payload round trip");
        Assert(wire[1]["tool_call_id"].ToString() == "c1" && wire[2]["tool_call_id"].ToString() == "c2", "tool results correlate with call IDs");
        Assert(!(bool)JObject.Parse(wire[2]["content"].ToString())["ok"], "failed tool result remains structured");
        var schema = JObject.Parse("{\"required\":[\"index\"],\"properties\":{\"index\":{\"type\":\"integer\"}}}");
        Assert(AgentWireProtocol.ValidateArguments(schema, new JObject()) != null, "missing tool parameter rejected");
        Assert(AgentWireProtocol.ValidateArguments(schema, JObject.Parse("{\"index\":\"1\"}")) != null, "wrong tool parameter type rejected");
        Assert(AgentWireProtocol.ValidateArguments(schema, JObject.Parse("{\"index\":1}")) == null, "valid tool parameter accepted");
        Assert(AgentWireProtocol.OllamaMessages(wireHistory)[1]["tool_name"].ToString() == "read", "Ollama tool result name mapped");

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

        Assert(PromptTemplateService.ResolveLanguage("Chinese", "This is Russian: Это русский текст") == "Chinese",
            "explicit Chinese language mode takes precedence");
        Assert(PromptTemplateService.ResolveLanguage("English", "这是中文文本") == "English",
            "explicit English language mode takes precedence");
        Assert(PromptTemplateService.ResolveLanguage("Russian", "Hello world") == "Russian",
            "explicit Russian language mode takes precedence");
        Assert(PromptTemplateService.ResolveLanguage("Auto", "这是中文文本") == "Chinese",
            "Auto detects Chinese text");
        Assert(PromptTemplateService.ResolveLanguage("Auto", "Это русский текст") == "Russian",
            "Auto detects high-confidence Russian text");
        Assert(PromptTemplateService.ResolveLanguage("Auto", "Hello world") == "Auto",
            "Auto preserves the user's language for English text");
        Assert(PromptTemplateService.ResolveLanguage("Auto", "Привет") == "Auto",
            "Auto does not treat short ambiguous Cyrillic text as Russian");

        Assert(!PromptTemplateService.ResolveUseEnglish("Chinese", "这是中文文本"),
            "Chinese uses the Chinese internal prompt path");
        Assert(PromptTemplateService.ResolveUseEnglish("English", "Hello world"),
            "English uses the English internal prompt path");
        Assert(PromptTemplateService.ResolveUseEnglish("Russian", "Это русский текст"),
            "Russian uses the English internal prompt path");
        Assert(PromptTemplateService.ResolveUseEnglish("Auto", "Привет"),
            "Auto uses the English internal prompt path");

        string russianAgentPrompt = PromptTemplateService.BuildAgentSystemPrompt("", "", 0, "", "Russian");
        string autoAgentPrompt = PromptTemplateService.BuildAgentSystemPrompt("", "", 0, "", "Auto");
        Assert(russianAgentPrompt.Contains("You are WordOllama") &&
               !russianAgentPrompt.Contains("Microsoft Word AI Assistant") &&
               russianAgentPrompt.Contains("Always respond to the user in Russian"),
            "Russian Agent prompt preserves the WordOllama identity and Russian response instruction");
        Assert(autoAgentPrompt.Contains("You are WordOllama") &&
               !autoAgentPrompt.Contains("Microsoft Word AI Assistant") &&
               autoAgentPrompt.Contains("Reply in the user's language"),
            "Auto Agent prompt preserves the WordOllama identity and a neutral response instruction");

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

        RunLongDocumentReviewChecks();
        RunReviewSessionChecks();
        RunAgentReviewChecks();
        Console.WriteLine("Core smoke test passed.");
    }

    private static void RunLongDocumentReviewChecks()
    {
        var defaults = AppConfig.Deserialize("{}");
        Assert(defaults.ReviewIssueBatchMaxParagraphs == 80 && defaults.ReviewIssueBatchMaxCharacters == 12000
            && defaults.ReviewSelectionBatchMaxCharacters == 6000 && defaults.ReviewSuggestionBatchMaxItems == 8
            && defaults.ReviewSuggestionBatchMaxCharacters == 6000 && defaults.ReviewMaxConcurrentRequests == 3,
            "old configuration keeps existing limits as per-batch defaults");
        Assert(defaults.OllamaTimeoutSeconds == 600 && defaults.LinterContextTokens == 4096
            && defaults.LinterTimeoutSeconds == 300 && defaults.AgentTaskTimeoutMinutes == 30,
            "old configuration retains timeout defaults");
        var configured = AppConfig.Deserialize("{\"ReviewIssueBatchMaxParagraphs\":25,\"ReviewSelectionBatchMaxCharacters\":1200,\"ReviewOllamaOutputTokens\":4096}");
        Assert(configured.ReviewIssueBatchMaxParagraphs == 25 && configured.ReviewSelectionBatchMaxCharacters == 1200
            && configured.ReviewOllamaOutputTokens == 4096, "review limits are configurable in appconfig.json");

        var issueBudget = new ReviewRequestBudget(32768, 4096, PromptTemplateService.BuildIssueReviewPrompt(""), false);
        var paragraphs = Enumerable.Range(1, 12000).Select(i => new ReviewParagraph
            { Index = i, Text = "Paragraph " + i + ": Check this sentence carefully." }).ToList();
        var batches = ReviewChunkPlanner.CreateIssueBatches(paragraphs, issueBudget, 80, 12000);
        Assert(batches.Count > 1 && batches.Last().LastParagraph == 12000
            && batches.SelectMany(b => b.Paragraphs).Distinct().Count() == 12000,
            "whole-document review reaches the final paragraph beyond the old 80-paragraph cutoff");
        Assert(batches.Sum(b => b.Characters) == paragraphs.Sum(p => p.Text.Length),
            "whole-document batches cover every source character");
        Assert(batches.All(b => b.Paragraphs.Count <= 80 && b.Context.Length <= 12000
            && issueBudget.Fits(PromptTemplateService.BuildIssueReviewPrompt(b.Context), b.Context)),
            "each issue batch respects paragraph, character and prompt budgets");

        string longSelection = string.Concat(Enumerable.Repeat("A sentence in a very long selection. ", 3000));
        var selectionBatches = ReviewChunkPlanner.CreateIssueBatches(
            new[] { new ReviewParagraph { Index = 91, Text = longSelection } }, issueBudget, 80, 6000);
        Assert(selectionBatches.Count > 1 && selectionBatches.Sum(b => b.Characters) == longSelection.Length
            && selectionBatches.All(b => b.Paragraphs.SetEquals(new[] { 91 }) && b.Context.Length <= 6000),
            "selection over 6000 characters and an oversized single paragraph are split without truncation");

        string unicodeText = string.Concat(Enumerable.Repeat("First sentence.\r\n中文😀  Next sentence. ", 900)) + "End.";
        var parts = ReviewChunkPlanner.SplitText(unicodeText, 6000, 512);
        Assert(string.Concat(parts.Select(p => p.Text)) == unicodeText
            && parts.All(p => p.Text.Length <= 6000 && ReviewChunkPlanner.EstimateTokens(p.Text) <= 512),
            "mixed-language long paragraphs are split losslessly within token allowances");
        Assert(parts.All(p => !char.IsHighSurrogate(p.Text[p.Text.Length - 1])
            && !char.IsLowSurrogate(p.Text[0])), "splitting preserves Unicode surrogate pairs");
        Assert(parts.Take(parts.Count - 1).All(p => !p.Text.EndsWith("\r")), "splitting keeps CRLF pairs together");
        string unchangedReview = string.Concat(parts.Select(p => ReviewChunkPlanner.PreserveFragmentWhitespace(p.Text, p.Text.Trim())));
        Assert(unchangedReview == unicodeText, "joining model fragments preserves inter-word spaces and paragraph breaks");
        Assert(ReviewChunkPlanner.PreserveFragmentWhitespace("   ", "") == "   "
            && ReviewChunkPlanner.PreserveFragmentWhitespace("\t Indented text.  ", "Indented text.") == "\t Indented text.  ",
            "whitespace-only fragments are preserved");

        var rewriteBudget = new ReviewRequestBudget(4096, 2048,
            PromptTemplateService.BuildReviewSuggestionPrompt("", "Proofread conservatively.", ""), true);
        foreach (var part in ReviewChunkPlanner.SplitText(unicodeText, 6000, rewriteBudget.TextTokens))
            Assert(rewriteBudget.Fits(PromptTemplateService.BuildReviewSuggestionPrompt(part.Text, "Proofread conservatively.", ""), part.Text),
                "rewriting reserves capacity for prompt, source and returned text");
        bool rejected = false;
        try { new ReviewRequestBudget(256, 2048, "Review prompt", true); }
        catch (InvalidOperationException) { rejected = true; }
        Assert(rejected, "insufficient context fails explicitly instead of sending an oversized request");

        JArray parsed;
        var indexes = new HashSet<int> { 91, 92 };
        Assert(ReviewChunkPlanner.TryReadIssues("[]", indexes, out parsed) && parsed.Count == 0,
            "a valid empty array means review completed with no issues");
        Assert(ReviewChunkPlanner.TryReadIssues("[{\"paragraph\":91,\"title\":\"Typo\"}]", indexes, out parsed),
            "review results retain original document indexes");
        Assert(!ReviewChunkPlanner.TryReadIssues("[{\"paragraph\":1,\"title\":\"Typo\"}]", indexes, out parsed),
            "renumbered or invented paragraph locations fail validation");
        Assert(!ReviewChunkPlanner.TryReadIssues("", indexes, out parsed)
            && !ReviewChunkPlanner.TryReadIssues("[API error] timeout", indexes, out parsed)
            && !ReviewChunkPlanner.TryReadIssues("[{\"paragraph\":91", indexes, out parsed)
            && !ReviewChunkPlanner.TryReadIssues("Unable to review. []", indexes, out parsed)
            && !ReviewChunkPlanner.TryReadIssues("[{}]", indexes, out parsed),
            "empty, API-error, truncated and invalid output are not reported as a clean review");
        var allowedIds = new[] { "1", "2" };
        Assert(ReviewChunkPlanner.ReadSuggestions("[{\"id\":\"1\",\"text\":\"Corrected sentence.\"}]", allowedIds).Count == 1,
            "valid batch suggestions keep their original card identity");
        Assert(ReviewChunkPlanner.ReadSuggestions("Unable to review. []", allowedIds).Count == 0
            && ReviewChunkPlanner.ReadSuggestions("[{\"id\":\"3\",\"text\":\"Wrong card.\"}]", allowedIds).Count == 0
            && ReviewChunkPlanner.ReadSuggestions("[{\"id\":\"1\",\"text\":\"One\"},{\"id\":\"1\",\"text\":\"Two\"}]", allowedIds).Count == 0
            && ReviewChunkPlanner.ReadSuggestions("[{\"id\":\"1\",\"text\":\"[API error] timeout\"}]", allowedIds).Count == 0,
            "unstructured, mismatched, duplicate and error suggestions never become applicable text");
        Assert(ProviderRequestPolicy.IsOutputTruncated(JObject.Parse("{\"done_reason\":\"length\"}"))
            && ProviderRequestPolicy.IsOutputTruncated(JObject.Parse("{\"choices\":[{\"finish_reason\":\"length\"}]}"))
            && ProviderRequestPolicy.IsOutputTruncated(JObject.Parse("{\"stop_reason\":\"max_tokens\"}"))
            && ProviderRequestPolicy.IsOutputTruncated(JObject.Parse("{\"candidates\":[{\"finishReason\":\"MAX_TOKENS\"}]}")),
            "native and online providers reject explicitly truncated output");
        Assert(!ProviderRequestPolicy.IsOutputTruncated(JObject.Parse("{\"choices\":[{\"finish_reason\":\"stop\"}]}"))
            && !ProviderRequestPolicy.IsOutputTruncated(new JObject()), "normal or legacy responses are not marked truncated");
        Console.WriteLine("Long-document review checks passed: 12000 paragraphs, long selection, Unicode, budgets and failure validation.");
    }

    private static void RunAgentReviewChecks()
    {
        var defaults = AppConfig.Deserialize("{}");
        Assert(defaults.ReviewReviewerName == "AI审批" && defaults.ReviewReviewerIdentityMode == "Date", "existing settings receive the default dated AI reviewer");
        var day = new DateTime(2026, 10, 12);
        Assert(ReviewRevisionIdentity.Build(null, "Fixed", day, null) == "AI审批"
            && ReviewRevisionIdentity.Build(" AI审批 ", "Date", day, "abc") == "AI审批-2026-10-12"
            && ReviewRevisionIdentity.Build("AI审批", "Session", day, "0123456789abcdef") == "AI审批-2026-10-12-01234567"
            && ReviewRevisionIdentity.Build("AI", "Session", day, "abc") == "AI-2026-10-12-abc", "reviewer names distinguish fixed, dated and session runs");
        var body = new ReviewStoryAddress();
        var header = new ReviewStoryAddress { Kind = "Header", Section = 2, Variant = 1 };
        var box = new ReviewStoryAddress { Shapes = new List<string> { "group/name", "box:1" } };
        var otherBox = new ReviewStoryAddress { Shapes = new List<string> { "group", "name/box:1" } };
        Assert(body.Key != header.Key && box.Key != otherBox.Key, "story addresses distinguish text areas and escape shape paths");
        var record = new ReviewSessionRecord();
        var edited = new ReviewSessionItem { Index = 1, RangeStart = 10, RangeEnd = 20, Story = header };
        var following = new ReviewSessionItem { Index = 2, RangeStart = 25, RangeEnd = 35, Story = header };
        var main = new ReviewSessionItem { Index = 3, RangeStart = 25, RangeEnd = 35, Story = body };
        var textBox = new ReviewSessionItem { Index = 4, RangeStart = 25, RangeEnd = 35, Story = box };
        record.Items.AddRange(new[] { edited, following, main, textBox });
        record.ShiftRangesAfterEdit(edited.Id, 20, 8, header.Key);
        record.ReplayRangeShifts();
        Assert(following.RangeStart == 33 && main.RangeStart == 25 && textBox.RangeStart == 25
            && edited.RangeStart == 10, "editing a header rebases only that story, exactly once");
        var reloaded = Newtonsoft.Json.JsonConvert.DeserializeObject<ReviewSessionRecord>(Newtonsoft.Json.JsonConvert.SerializeObject(record));
        reloaded.ReplayRangeShifts();
        Assert(reloaded.Items[1].RangeStart == 33 && reloaded.Items[2].RangeStart == 25, "cross-story recovery preserves rebased offsets");
        var spans = new[] { new ProtectedTextSpan { Start = 8, End = 12 } };
        TextDiffService.ValidateProtectedSpans(new[] { new TextEdit { OriginalStart = 2, DeleteLength = 1, InsertText = "a" },
            new TextEdit { OriginalStart = 12, DeleteLength = 0, InsertText = "x" } }, spans);
        ExpectReviewFailure<InvalidOperationException>(delegate { TextDiffService.ValidateProtectedSpans(
            new[] { new TextEdit { OriginalStart = 9, DeleteLength = 1, InsertText = "x" } }, spans); }, "field replacements are rejected before writing");
        ExpectReviewFailure<InvalidOperationException>(delegate { TextDiffService.ValidateProtectedSpans(
            new[] { new TextEdit { OriginalStart = 9, DeleteLength = 0, InsertText = "x" } }, spans); }, "insertions inside fields are rejected");
        _reviewWorkflowCalls.Clear();
        var tools = new List<ReviewWorkflowTool>();
        foreach (string name in new[] { "review_document", "resume_review", "list_review_sessions", "apply_review_revisions", "save_review_document" })
        {
            var constructor = typeof(ReviewWorkflowTool).GetConstructors().Single();
            var callbackType = constructor.GetParameters()[1].ParameterType;
            var parameters = callbackType.GetMethod("Invoke").GetParameters()
                .Select(p => System.Linq.Expressions.Expression.Parameter(p.ParameterType, p.Name)).ToArray();
            var bodyCall = System.Linq.Expressions.Expression.Call(typeof(Program).GetMethod("ReviewTestWorkflow", BindingFlags.Static | BindingFlags.NonPublic),
                System.Linq.Expressions.Expression.Constant(name), parameters[3]);
            var callback = System.Linq.Expressions.Expression.Lambda(callbackType, bodyCall, parameters).Compile();
            var tool = (ReviewWorkflowTool)constructor.Invoke(new object[] { name, callback });
            Assert(tool.IsWriteOperation == (name == "apply_review_revisions")
                && !tool.IsConcurrencySafe && !tool.Description.StartsWith("I18n_", StringComparison.Ordinal), "review tool mode classification and localized description: " + name);
            tools.Add(tool);
        }
        var review = tools[0];
        Assert(AgentWireProtocol.ValidateArguments(review.ParameterSchema, JObject.Parse("{\"kind\":\"suggestions\",\"create_copy\":true}")) == null
            && AgentWireProtocol.ValidateArguments(tools[1].ParameterSchema, new JObject()) != null, "review and resume argument schemas validate");
        var executed = (System.Threading.Tasks.Task<string>)typeof(ReviewWorkflowTool).GetMethod("ExecuteAsync")
            .Invoke(review, new object[] { new JObject(), null, null, System.Threading.CancellationToken.None });
        Assert(executed.GetAwaiter().GetResult().Contains("true")
            && _reviewWorkflowCalls.SequenceEqual(new[] { "review_document" }), "Agent review calls the configured shared workflow");
        string prompt = PromptTemplateService.BuildAgentSystemPrompt("<document>truncated snapshot</document>", "review_document", 5, "", "English");
        Assert(prompt.Contains("apply_review_revisions") && prompt.Contains("Never claim the whole document is complete")
            && prompt.Contains("list_review_sessions"), "Agent prompt routes whole-document review and checks actual completion");
        string instruction = "Check terminology only";
        Func<string, string> buildPrompt = delegate(string context) { return PromptTemplateService.BuildIssueReviewPrompt(context, false, instruction); };
        var budget = new ReviewRequestBudget(8192, 1024, buildPrompt(""), false);
        var paragraphs = new[] { new ReviewParagraph { Index = 51, Text = "Header text needs review", Location = new ReviewParagraphLocation { Story = header, Paragraph = 1 } },
            new ReviewParagraph { Index = 52, Text = "Body text needs review", Location = new ReviewParagraphLocation { Story = body, Paragraph = 2 } } };
        var batches = ReviewChunkPlanner.CreateIssueBatches(paragraphs, budget, 1, 6000, buildPrompt);
        Assert(batches.Count == 2 && batches.All(b => budget.Fits(buildPrompt(b.Context), b.Context))
            && batches[0].Context.Contains(header.Label) && buildPrompt(batches[0].Context).Contains(instruction), "issue batches retain area, original IDs and custom instruction within budget");
        Console.WriteLine("Agent review checks passed: routing, modes, reviewer identities, protected fields and story-local recovery.");
    }

    private static readonly List<string> _reviewWorkflowCalls = new List<string>();
    private static System.Threading.Tasks.Task<string> ReviewTestWorkflow(string name, System.Threading.CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        _reviewWorkflowCalls.Add(name);
        return System.Threading.Tasks.Task.FromResult("{\"ok\":true}");
    }

    private static void RunReviewSessionChecks()
    {
        string temporaryRoot = Path.GetFullPath(Path.Combine(Path.GetTempPath(), "WordOllamaReviewStoreSmoke-" + Guid.NewGuid().ToString("N")));
        string temporaryParent = Path.GetFullPath(Path.GetTempPath()).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        Assert(temporaryRoot.StartsWith(temporaryParent, StringComparison.OrdinalIgnoreCase)
            && Path.GetFileName(temporaryRoot).StartsWith("WordOllamaReviewStoreSmoke-", StringComparison.Ordinal), "review test data stays inside its own temporary directory");
        try
        {
            var store = new ReviewSessionStore(temporaryRoot);
            string document = Path.Combine(temporaryRoot, "original.docx");
            string hash = ReviewSessionStore.DocumentFingerprint("Sample body\r", 12, 2);
            var record = new ReviewSessionRecord { Metadata = new ReviewSessionMetadata
            {
                Kind = "Suggestions", Scope = "Body", DocumentIdentity = document, DocumentName = "original.docx",
                IsSavedDocument = true, ExpectedDocumentHash = hash, State = ReviewRunState.Running,
                ConfigurationFingerprint = "private-configuration-fingerprint",
                ModelSettings = new ReviewModelSettings { Mode = "Ollama", Provider = "Ollama", Model = "model-A", ContextTokens = 32768, OutputTokens = 2048 },
                Instruction = "Private review instruction"
            } };
            var complete = new ReviewSessionItem { Index = 1, RangeStart = 0, RangeEnd = 20, Characters = 20,
                OriginalText = new string('a', 20), SuggestedText = "Completed suggestion", State = ReviewItemState.Succeeded };
            var failed = new ReviewSessionItem { Index = 2, RangeStart = 40, RangeEnd = 50, Characters = 10,
                OriginalText = new string('b', 10), State = ReviewItemState.Failed, Error = "first failed request" };
            failed.BeginAttempt(); failed.FinishAttempt(ReviewItemState.Failed, failed.Error);
            var interrupted = new ReviewSessionItem { Index = 3, RangeStart = 70, RangeEnd = 83, Characters = 13,
                OriginalText = "alpha beta!!!", State = ReviewItemState.Running };
            interrupted.Fragments.Add(new ReviewSessionFragment { Start = 0, Source = "alpha ", Suggestion = "Alpha ", State = ReviewItemState.Succeeded });
            interrupted.Fragments.Add(new ReviewSessionFragment { Start = 6, Source = "beta!!!", State = ReviewItemState.Running });
            interrupted.BeginAttempt(); interrupted.BeginAttempt(6);
            record.Items.AddRange(new[] { complete, failed, interrupted });
            store.Create(record);
            var restored = store.Load(record.Metadata.Id);
            Assert(restored.Items[2].State == ReviewItemState.Pending && restored.Items[2].Fragments[1].State == ReviewItemState.Pending
                && restored.Items[2].Fragments[0].Suggestion == "Alpha " && restored.Items[2].SuggestedText == null,
                "interrupted review retains successful fragments and publishes no incomplete suggestion");
            Assert(restored.Metadata.CompletedItems == 1 && restored.Metadata.FailedItems == 1 && restored.Metadata.PendingItems == 1
                && restored.Metadata.ReviewedCharacters == 26 && restored.Metadata.FailedCharacters == 10 && restored.Metadata.UnreviewedCharacters == 7,
                "recovery rebuilds complete, failed and unreviewed character coverage from work items");
            Assert(restored.Items[2].AttemptLog.All(a => a.State == ReviewItemState.Pending), "in-flight attempt records become retryable after interruption");
            Assert(ReviewSessionStore.CanResume(restored, document.ToUpperInvariant(), hash)
                && !ReviewSessionStore.CanResume(restored, document, hash + "changed")
                && !ReviewSessionStore.CanResume(restored, document + ".copy", hash), "resuming binds to matching file identity and document text");
            restored.Metadata.IsSavedDocument = false;
            Assert(!ReviewSessionStore.CanResume(restored, document, hash), "unsaved document cannot be restored as a stable saved-file session");
            restored.Metadata.IsSavedDocument = true;
            Assert(hash != ReviewSessionStore.DocumentFingerprint("Sample body\r", 12, 3), "document stamp detects changes in paragraph structure");

            var retried = restored.Items[1];
            retried.State = ReviewItemState.Running; retried.BeginAttempt(); store.SaveItem(restored, retried);
            retried.State = ReviewItemState.Succeeded; retried.Error = null; retried.SuggestedText = "Retry succeeded";
            retried.FinishAttempt(ReviewItemState.Succeeded); store.SaveItem(restored, retried);
            store.SaveItem(restored, retried); // Keep a complete encrypted backup for the corruption check.
            var afterRetry = store.Load(restored.Metadata.Id);
            Assert(afterRetry.Metadata.CompletedItems == 2 && afterRetry.Metadata.FailedItems == 0 && afterRetry.Metadata.UnreviewedCharacters == 7
                && afterRetry.Items[1].AttemptLog.Any(a => a.Error == "first failed request"), "successful retry clears failed coverage and preserves previous failure details");
            string directory = Path.Combine(temporaryRoot, restored.Metadata.Id);
            Assert(Directory.GetFiles(directory, "*.dat").All(path => File.ReadAllText(path).StartsWith("dpapi:", StringComparison.Ordinal)
                && !File.ReadAllText(path).Contains("Private review instruction")), "review text and instructions are encrypted on disk");
            string itemFile = Path.Combine(directory, "item-" + retried.Id + ".dat");
            File.WriteAllText(itemFile, "dpapi:invalid-base64");
            var backupRestored = store.Load(restored.Metadata.Id);
            Assert(backupRestored.Items[1].SuggestedText == "Retry succeeded", "damaged primary item recovers from its encrypted atomic-write backup");
            store.SaveItem(backupRestored, backupRestored.Items[1]);

            // The unchanged remaining item is deliberately not rewritten. Its older
            // position is rebased from the metadata journal when loading it again.
            backupRestored.Items[0].UserAction = "replace";
            backupRestored.ShiftRangesAfterEdit(complete.Id, 20, 5);
            store.SaveItem(backupRestored, backupRestored.Items[0]);
            var shifted = store.Load(restored.Metadata.Id);
            Assert(shifted.Items[1].RangeStart == 45 && shifted.Items[2].RangeStart == 75, "pending positions follow edits without rewriting the whole review cache");
            store.SaveItem(shifted, shifted.Items[1]);
            shifted.Items[1].UserAction = "replace";
            shifted.ShiftRangesAfterEdit(retried.Id, 55, -3);
            store.SaveItem(shifted, shifted.Items[1]);
            var shiftedAgain = store.Load(restored.Metadata.Id);
            Assert(shiftedAgain.Items[1].RangeStart == 45 && shiftedAgain.Items[2].RangeStart == 72
                && shiftedAgain.Items[2].PositionRevision == 2, "position replay handles saved and older items without double shifting");
            Assert(!shiftedAgain.Items[0].NeedsProcessing && !shiftedAgain.Items[1].NeedsProcessing && shiftedAgain.Items[2].NeedsProcessing,
                "continuing selects unfinished work while excluding completed and consumed results");

            var another = new ReviewSessionRecord { Metadata = new ReviewSessionMetadata { Kind = "Suggestions", Scope = "Body",
                DocumentIdentity = document, DocumentName = "original.docx", ModelSettings = new ReviewModelSettings { Model = "model-B" } } };
            another.Items.Add(new ReviewSessionItem { Index = 1, State = ReviewItemState.Pending });
            store.Create(another);
            Assert(store.List(document, "Suggestions").Count == 2 && store.List(document, "Issues").Count == 0
                && store.Load(restored.Metadata.Id).Metadata.ModelSettings.Model == "model-A", "history retains separate model sessions and filters by operation");
            ExpectReviewFailure<ArgumentException>(delegate { store.Load("..\\outside"); }, "session identifiers cannot escape the local cache root");
            File.Delete(Path.Combine(temporaryRoot, another.Metadata.Id, "item-" + another.Items[0].Id + ".dat"));
            ExpectReviewFailure<IOException>(delegate { store.Load(another.Metadata.Id); }, "a missing work payload is reported instead of treated as completed");
            var crossStory = new ReviewSessionRecord { Metadata = new ReviewSessionMetadata { SchemaVersion = 2, Kind = "Issues", Scope = "Document",
                DocumentIdentity = document + ".review.docx", OriginalDocumentPath = document, ReviewerIdentity = "AI审批-2026-10-12", ReviewTaskInstruction = "Check terminology only" } };
            crossStory.Items.Add(new ReviewSessionItem { Story = new ReviewStoryAddress { Kind = "Header", Section = 2, Variant = 3 },
                ParagraphLocations = new Dictionary<int, ReviewParagraphLocation> { { 112, new ReviewParagraphLocation { Story = new ReviewStoryAddress { Kind = "Footnote", Note = 5 }, Paragraph = 2 } } } });
            store.Create(crossStory);
            var crossStoryRestored = store.Load(crossStory.Metadata.Id);
            Assert(crossStoryRestored.Metadata.SchemaVersion == 2 && crossStoryRestored.Items[0].Story.Key == crossStory.Items[0].Story.Key
                && crossStoryRestored.Items[0].ParagraphLocations[112].Story.Note == 5
                && crossStoryRestored.Metadata.ReviewerIdentity == "AI审批-2026-10-12"
                && crossStoryRestored.Metadata.ReviewTaskInstruction == "Check terminology only", "encrypted recovery retains stories, reviewer and custom issue instructions");
            Assert(store.List(document, null).Any(m => m.Id == crossStory.Metadata.Id), "original document can discover saved review-copy sessions");
            var legacy = new ReviewSessionRecord { Metadata = new ReviewSessionMetadata { SchemaVersion = 1, Kind = "Issues", Scope = "Body", DocumentIdentity = document } };
            legacy.Items.Add(new ReviewSessionItem { RangeStart = 0, RangeEnd = 20 }); store.Create(legacy);
            Assert(store.Load(legacy.Metadata.Id).Metadata.SchemaVersion == 1 && store.Load(legacy.Metadata.Id).Items[0].Story == null,
                "v2.4 body sessions remain readable with their legacy fingerprint schema");

            var report = new ReviewSessionRecord { Metadata = record.Metadata };
            for (int index = 1; index <= 24; index++)
            {
                var item = new ReviewSessionItem { Index = index, FirstParagraph = index, LastParagraph = index, Characters = 15,
                    OriginalText = "Source with ```` code", State = ReviewItemState.Failed, Error = "Failure number " + index };
                item.BeginAttempt(); item.FinishAttempt(ReviewItemState.Failed, item.Error); report.Items.Add(item);
            }
            report.Items[0].AttemptLog.Add(new ReviewSessionAttempt { State = ReviewItemState.Failed, Error = "Earlier failed attempt\nwith detail" });
            report.Items[0].GeneratedText = "Original model output";
            report.Items[0].SuggestedText = "Human edited suggestion";
            report.Metadata.RunLog.Add(new ReviewRunLogEntry { State = ReviewRunState.Failed, Error = "Previous queue failure" });
            DateTime lastSaved = report.Metadata.UpdatedUtc;
            string markdown = ReviewReportExporter.BuildMarkdown(report);
            string json = ReviewReportExporter.BuildJson(report);
            Assert(Enumerable.Range(1, 24).All(i => markdown.Contains("Failure number " + i)) && markdown.Contains("Earlier failed attempt\nwith detail")
                && markdown.Contains("Source with ```` code") && markdown.Contains("`````"), "report includes every failure, prior attempts and safely fenced source text");
            Assert(markdown.Contains("Original model output") && markdown.Contains("Human edited suggestion") && markdown.Contains("Previous queue failure"),
                "report retains original model output, human edits and earlier queue errors");
            var jsonReport = JObject.Parse(json);
            Assert(jsonReport["Items"].Count() == 24 && (string)jsonReport["Metadata"]["ModelSettings"]["Model"] == "model-A"
                && !json.Contains("ConfigurationFingerprint") && !json.Contains("ApiKey") && report.Metadata.UpdatedUtc == lastSaved,
                "JSON report retains full results and safe settings without changing the session save date");
            ReviewReportExporter.Write(report, Path.Combine(temporaryRoot, "report.md"));
            ReviewReportExporter.Write(report, Path.Combine(temporaryRoot, "report.json"));
            Assert(File.ReadAllText(Path.Combine(temporaryRoot, "report.md")).Contains("Failure number 24"), "local report export writes the complete report");
            Assert(ReviewCopyService.ExtensionFor(13) == ".docm" && ReviewCopyService.ExtensionFor(0) == ".doc"
                && ReviewCopyService.ExtensionFor(24) == ".docx", "copy planning retains native Word file formats");
            File.WriteAllText(document, "Original document sentinel");
            ExpectReviewFailure<InvalidOperationException>(delegate { ReviewCopyService.ValidateTarget(document, document, ".docx"); }, "review copy cannot replace its source file");
            string copyPath = Path.Combine(temporaryRoot, "copy.docx");
            ReviewCopyService.ValidateTarget(document, copyPath, ".docx");
            File.WriteAllText(copyPath, "Existing copy sentinel");
            ExpectReviewFailure<IOException>(delegate { ReviewCopyService.ValidateTarget(document, copyPath, ".docx"); }, "review copy refuses an existing destination");
            ExpectReviewFailure<InvalidOperationException>(delegate { ReviewCopyService.ValidateTarget(document, Path.Combine(temporaryRoot, "copy.docm"), ".docx"); }, "copy rejects a mismatched extension");
            Assert(File.ReadAllText(document) == "Original document sentinel" && File.ReadAllText(copyPath) == "Existing copy sentinel", "copy validation leaves source and existing files intact");
            Console.WriteLine("Review session checks passed: encrypted recovery, fragment retry, positions, complete reports and copy protection.");
        }
        finally
        {
            if (Directory.Exists(temporaryRoot))
            {
                if (!Path.GetFullPath(temporaryRoot).StartsWith(temporaryParent, StringComparison.OrdinalIgnoreCase)
                    || !Path.GetFileName(temporaryRoot).StartsWith("WordOllamaReviewStoreSmoke-", StringComparison.Ordinal))
                    throw new InvalidOperationException("Unsafe review test cleanup path.");
                Directory.Delete(temporaryRoot, true);
            }
        }
    }

    private static void ExpectReviewFailure<T>(Action action, string name) where T : Exception
    {
        bool failed = false;
        try { action(); } catch (T) { failed = true; }
        Assert(failed, name);
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
