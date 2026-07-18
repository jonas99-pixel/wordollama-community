using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json.Linq;

namespace WordOllama.lib
{
    public enum PromptTaskType
    {
        Translate,
        Expand,
        Polish,
        Simplify,
        Summarize,
        ContinueWriting,
        TypoFix,
        LegalRisk,
        FairRewrite,
        Custom
    }

    public enum PromptOutputMode
    {
        InsertBelow,
        InsertBelowWithDiff,
        ReplaceOriginal,
        Comment,
        ReviewPane
    }

    public static class PromptTemplateService
    {
        private const int ReviewPaneCharacterThreshold = 4000;
        private const int ReviewPaneParagraphThreshold = 8;

        public static string BuildPrompt(
            PromptTaskType taskType,
            string selectedText,
            string targetLanguage = null,
            string customInstruction = null,
            string writingProfile = null)
        {
            if (taskType == PromptTaskType.Custom)
                return BuildCustomPrompt(customInstruction, selectedText);

            string taskInstruction = BuildTaskInstruction(taskType, targetLanguage);
            bool commentOutput = taskType == PromptTaskType.LegalRisk;

            var sb = new StringBuilder();
            sb.AppendLine(commentOutput
                ? "You are reviewing selected text inside Microsoft Word."
                : "You are editing selected text inside Microsoft Word.");
            sb.AppendLine();
            sb.AppendLine("Task:");
            sb.AppendLine(taskInstruction);
            sb.AppendLine();
            sb.AppendLine("Rules:");

            if (commentOutput)
            {
                sb.AppendLine("- Return concise Word comment text, not a rewritten document.");
                sb.AppendLine("- Do not overstate risks. If no material issue is found, say so clearly.");
                sb.AppendLine("- Use this format: Risk / Reason / Suggested wording / Severity.");
            }
            else
            {
                sb.AppendLine("- Preserve the author's meaning, voice, paragraph breaks, names, dates, amounts, citations, numbering, and unusual spellings unless the task explicitly requires a change.");
                sb.AppendLine("- Do not add explanations, alternatives, headings, markdown, or code fences.");
                sb.AppendLine("- Return only the final text.");
                sb.AppendLine("- If no change is needed, return the original text unchanged.");
            }

            if (!string.IsNullOrWhiteSpace(writingProfile))
            {
                sb.AppendLine();
                sb.AppendLine("Writing profile / style guide:");
                sb.AppendLine(writingProfile.Trim());
            }

            sb.AppendLine();
            sb.AppendLine("Text:");
            sb.AppendLine("<<<");
            sb.AppendLine(selectedText ?? "");
            sb.AppendLine(">>>");
            return sb.ToString();
        }

        public static string BuildReviewInstruction(
            PromptTaskType taskType,
            string targetLanguage = null,
            string customInstruction = null)
        {
            if (taskType == PromptTaskType.Custom)
                return string.IsNullOrWhiteSpace(customInstruction)
                    ? "Edit the selected text according to the configured custom instruction."
                    : customInstruction.Trim();

            return BuildTaskInstruction(taskType, targetLanguage);
        }

        public static bool ShouldUseReviewPane(PromptTaskType taskType, string selectedText)
        {
            if (taskType == PromptTaskType.LegalRisk)
                return false;

            if (string.IsNullOrWhiteSpace(selectedText))
                return false;

            int paragraphCount = selectedText
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Count(line => !string.IsNullOrWhiteSpace(line));

            return selectedText.Length >= ReviewPaneCharacterThreshold ||
                   paragraphCount >= ReviewPaneParagraphThreshold;
        }

        public static string BuildIssueReviewPrompt(string context, bool concise = false)
        {
            var sb = new StringBuilder();
            sb.AppendLine("You are a Microsoft Word document review assistant.");
            sb.AppendLine("Review the numbered paragraphs below for clear typos, grammar problems, formatting issues, citation/reference errors, inconsistent terminology, material legal/contract risks, and unclear expression.");
            sb.AppendLine();
            sb.AppendLine("Return only a valid JSON array. Do not output markdown, code fences, explanations, or trailing text.");
            sb.AppendLine("Use stable English enum values for category and severity. Write title, description, suggestion, and excerpt in the same language as the source text when possible.");
            if (concise)
                sb.AppendLine("Be conservative and concise. Report only obvious, actionable issues.");
            sb.AppendLine();
            sb.AppendLine("Schema:");
            sb.AppendLine("[");
            sb.AppendLine("  {");
            sb.AppendLine("    \"paragraph\": 1,");
            sb.AppendLine("    \"category\": \"typo|grammar|format|reference|terminology|legal|clarity\",");
            sb.AppendLine("    \"severity\": \"high|medium|low\",");
            sb.AppendLine("    \"title\": \"short issue title\",");
            sb.AppendLine("    \"description\": \"issue explanation\",");
            sb.AppendLine("    \"suggestion\": \"specific fix or recommendation\",");
            sb.AppendLine("    \"excerpt\": \"relevant source excerpt\"");
            sb.AppendLine("  }");
            sb.AppendLine("]");
            sb.AppendLine();
            sb.AppendLine("If no issue is found, return [].");
            sb.AppendLine();
            sb.AppendLine("Text to review:");
            sb.AppendLine(context ?? "");
            return sb.ToString();
        }

        public static string BuildLegalRiskAnalysisPrompt(string selectedText)
        {
            var sb = new StringBuilder();
            sb.AppendLine("You are a senior lawyer specializing in contract and legal risk review.");
            sb.AppendLine("Analyze the following text and provide practical editing advice.");
            sb.AppendLine();
            sb.AppendLine("Output structure:");
            sb.AppendLine("1. Risk points");
            sb.AppendLine("2. Recommended revisions");
            sb.AppendLine("3. Suggested wording");
            sb.AppendLine("4. Additional precautions");
            sb.AppendLine();
            sb.AppendLine("Rules:");
            sb.AppendLine("- Use the same language as the source text when possible.");
            sb.AppendLine("- Be specific and actionable.");
            sb.AppendLine("- Do not overstate risks.");
            sb.AppendLine();
            sb.AppendLine("Text:");
            sb.AppendLine("<<<");
            sb.AppendLine(selectedText ?? "");
            sb.AppendLine(">>>");
            return sb.ToString();
        }

        public static string BuildTranslationWindowPrompt(string targetLanguage, string inputText, bool comparisonMode)
        {
            var sb = new StringBuilder();
            sb.AppendLine("You are a professional translator.");
            sb.AppendLine("Translate the source text into natural, idiomatic " + (string.IsNullOrWhiteSpace(targetLanguage) ? "target language" : targetLanguage.Trim()) + ".");
            sb.AppendLine("Preserve paragraph boundaries, names, dates, amounts, citations, numbering, placeholders, and terminology.");
            sb.AppendLine();
            if (comparisonMode)
            {
                sb.AppendLine("Output mode:");
                sb.AppendLine("- Produce a source/translation comparison.");
                sb.AppendLine("- For multiple source paragraphs, output each original paragraph followed immediately by its translation.");
                sb.AppendLine("- Do not add explanations or extra commentary.");
            }
            else
            {
                sb.AppendLine("Output mode:");
                sb.AppendLine("- Return only the translated text.");
                sb.AppendLine("- Do not add explanations, labels, markdown, or extra commentary.");
            }

            sb.AppendLine();
            sb.AppendLine("Source text:");
            sb.AppendLine("<<<");
            sb.AppendLine(inputText ?? "");
            sb.AppendLine(">>>");
            return sb.ToString();
        }

        public static string BuildFreeWritingPrompt(string theme, string requirement)
        {
            var sb = new StringBuilder();
            sb.AppendLine("You are a professional writing assistant inside Microsoft Word.");
            sb.AppendLine("Create a polished draft according to the user's topic and requirements.");
            sb.AppendLine();
            sb.AppendLine("Rules:");
            sb.AppendLine("- Use natural paragraphs.");
            sb.AppendLine("- Do not use markdown unless the user explicitly asks for it.");
            sb.AppendLine("- Return only the generated content. Do not add greetings, explanations, or meta commentary.");
            sb.AppendLine("- Use the user's requested language when clear; otherwise use the language of the requirement.");
            sb.AppendLine();
            sb.AppendLine("Topic:");
            sb.AppendLine(theme ?? "");
            sb.AppendLine();
            sb.AppendLine("Requirements:");
            sb.AppendLine(requirement ?? "");
            return sb.ToString();
        }

        public static string BuildModificationPrompt(string originalText, string instruction)
        {
            var sb = new StringBuilder();
            sb.AppendLine("You are editing text inside Microsoft Word.");
            sb.AppendLine("Apply the user's instruction while preserving the original meaning, names, dates, amounts, citations, numbering, paragraph breaks, and unusual spellings unless a change is explicitly requested.");
            sb.AppendLine("Return only the complete revised text. Do not add greetings, explanations, alternatives, markdown, or code fences.");
            sb.AppendLine();
            sb.AppendLine("User instruction:");
            sb.AppendLine(instruction ?? "");
            sb.AppendLine();
            sb.AppendLine("Original text:");
            sb.AppendLine("<<<");
            sb.AppendLine(originalText ?? "");
            sb.AppendLine(">>>");
            return sb.ToString();
        }

        public static string BuildMootCourtInvestigationPrompt(string documentType, string documentText, bool english)
        {
            var sb = new StringBuilder();
            if (english)
            {
                sb.AppendLine("You are an experienced senior judge conducting a pre-hearing investigation.");
                sb.AppendLine("Review the supplied pleading and identify focused questions that may reveal legal gaps, factual contradictions, or insufficient evidence.");
                sb.AppendLine();
                sb.AppendLine("Your questions should:");
                sb.AppendLine("- address material facts, legal application, and the evidence chain;");
                sb.AppendLine("- be specific and actionable rather than generic;");
                sb.AppendLine("- help assess the strengths and weaknesses of the case.");
                sb.AppendLine();
                sb.AppendLine("Return a numbered list. After each question, briefly explain its importance or expected effect.");
                sb.AppendLine();
                sb.AppendLine($"Pleading type: {documentType}");
            }
            else
            {
                sb.AppendLine("你是一名经验丰富的中国高级法官，负责在庭审前的调查阶段对案件进行全面审查。");
                sb.AppendLine($"请仔细阅读并分析所提供的{documentType}，提出能够识别法律漏洞、事实矛盾或证据不足的调查问题。");
                sb.AppendLine();
                sb.AppendLine("问题应当触及案件实质，帮助发现事实或澄清法律适用，揭示证据链条缺陷，并且具体、有针对性，避免笼统表述。");
                sb.AppendLine("请以编号列表呈现问题，并在每个问题后简要说明其重要性或预期效果。");
                sb.AppendLine();
                sb.AppendLine($"文书类型：{documentType}");
            }

            sb.AppendLine();
            sb.AppendLine(english ? "Pleading text:" : $"以下为待分析的{documentType}文本内容：");
            sb.AppendLine("<<<");
            sb.AppendLine(documentText ?? "");
            sb.AppendLine(">>>");
            return sb.ToString();
        }

        public static string BuildHtmlAppPrompt(string requirement)
        {
            var sb = new StringBuilder();
            sb.AppendLine("You are a senior full-stack developer.");
            sb.AppendLine("Generate a complete, production-quality HTML5 app for the user's requirement.");
            sb.AppendLine();
            sb.AppendLine("Rules:");
            sb.AppendLine("- Output a single complete HTML document with embedded CSS and JavaScript.");
            sb.AppendLine("- The page must be responsive and usable.");
            sb.AppendLine("- Do not split CSS or JavaScript into separate files.");
            sb.AppendLine("- Return only the HTML code. Do not add explanations outside code comments.");
            sb.AppendLine();
            sb.AppendLine("User requirement:");
            sb.AppendLine(requirement ?? "");
            return sb.ToString();
        }

        public static string BuildTablePrompt(string sourceText, IEnumerable<string> requirements = null)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Analyze the following text and convert it into structured table data.");
            sb.AppendLine("Return only valid JSON. Do not output markdown, code fences, explanations, or trailing text.");
            sb.AppendLine();
            sb.AppendLine("Rules:");
            sb.AppendLine("1. The JSON must contain exactly two top-level fields: \"headers\" and \"rows\".");
            sb.AppendLine("2. \"headers\" must be an array of column names.");
            sb.AppendLine("3. \"rows\" must be an array of row arrays.");
            sb.AppendLine("4. Extract data accurately and completely.");
            sb.AppendLine("5. Use \"N/A\" for missing values.");

            var extraRequirements = (requirements ?? Enumerable.Empty<string>())
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Select(item => item.Trim())
                .ToList();

            for (int i = 0; i < extraRequirements.Count; i++)
                sb.AppendLine($"{6 + i}. {extraRequirements[i]}");

            sb.AppendLine();
            sb.AppendLine("Expected JSON example:");
            sb.AppendLine("{");
            sb.AppendLine("  \"headers\": [\"Column 1\", \"Column 2\", \"Column 3\"],");
            sb.AppendLine("  \"rows\": [");
            sb.AppendLine("    [\"Data 1\", \"Data 2\", \"Data 3\"],");
            sb.AppendLine("    [\"Data 4\", \"Data 5\", \"Data 6\"]");
            sb.AppendLine("  ]");
            sb.AppendLine("}");
            sb.AppendLine();
            sb.AppendLine("Text:");
            sb.AppendLine("<<<");
            sb.AppendLine(sourceText ?? "");
            sb.AppendLine(">>>");
            return sb.ToString();
        }

        public static string BuildPromptOptimizerPrompt(string userPrompt)
        {
            var sb = new StringBuilder();
            sb.AppendLine("You are a prompt optimization specialist.");
            sb.AppendLine("Rewrite the user's raw prompt into a clearer, more professional prompt.");
            sb.AppendLine();
            sb.AppendLine("Optimization guidelines:");
            sb.AppendLine("- Add an appropriate role when useful.");
            sb.AppendLine("- Clarify the task, context, constraints, and expected output.");
            sb.AppendLine("- Add a suitable output format when it improves reliability.");
            sb.AppendLine("- Preserve the user's original intent and language.");
            sb.AppendLine("- Adjust the level of detail to the complexity of the task.");
            sb.AppendLine();
            sb.AppendLine("Output rules:");
            sb.AppendLine("- Return only the optimized prompt.");
            sb.AppendLine("- Do not include analysis, explanations, or alternative versions.");
            sb.AppendLine();
            sb.AppendLine("Raw prompt:");
            sb.AppendLine(userPrompt ?? "");
            return sb.ToString();
        }

        public static string BuildLanguageDetectionPrompt(string text)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Identify the language used in the following text.");
            sb.AppendLine("Return only the language name. Do not add explanations.");
            sb.AppendLine();
            sb.AppendLine("Examples:");
            sb.AppendLine("Input: 你好");
            sb.AppendLine("Output: Chinese");
            sb.AppendLine("Input: Hello");
            sb.AppendLine("Output: English");
            sb.AppendLine();
            sb.AppendLine("Text:");
            sb.AppendLine(text ?? "");
            return sb.ToString();
        }

        public static string BuildRevisionAnalysisPrompt(string reviewerRole, string originalText, string revisedText)
        {
            var sb = new StringBuilder();
            sb.AppendLine("You are an experienced contract lawyer.");
            sb.AppendLine("Review the other party's revisions and provide a practical assessment for " + (string.IsNullOrWhiteSpace(reviewerRole) ? "the user" : reviewerRole.Trim()) + ".");
            sb.AppendLine();
            sb.AppendLine("Output in natural paragraphs with these sections:");
            sb.AppendLine("Revision assessment: choose one of Acceptable / Needs negotiation / Not recommended.");
            sb.AppendLine("Reason: explain why.");
            sb.AppendLine("Suggested revision: provide improved wording or negotiation suggestions.");
            sb.AppendLine();
            sb.AppendLine("Original text:");
            sb.AppendLine("<<<");
            sb.AppendLine(originalText ?? "");
            sb.AppendLine(">>>");
            sb.AppendLine();
            sb.AppendLine("Revised text:");
            sb.AppendLine("<<<");
            sb.AppendLine(revisedText ?? "");
            sb.AppendLine(">>>");
            return sb.ToString();
        }

        public static string BuildAgentPlanningPrompt(string userRequirement)
        {
            var sb = new StringBuilder();
            sb.AppendLine("You are a task planner for a Microsoft Word AI agent.");
            sb.AppendLine("Break the user's request into a small number of independent AI steps.");
            sb.AppendLine();
            sb.AppendLine("Return only a valid JSON array. Do not output markdown, code fences, explanations, or trailing text.");
            sb.AppendLine("Use these exact English field names:");
            sb.AppendLine("- step: step number, starting from 1");
            sb.AppendLine("- description: short user-visible step description");
            sb.AppendLine("- prompt: the complete prompt for that step");
            sb.AppendLine();
            sb.AppendLine("Rules:");
            sb.AppendLine("- Keep the plan concise. Prefer 1-3 steps unless the request is genuinely complex.");
            sb.AppendLine("- Preserve the user's original language in step descriptions and prompts when helpful.");
            sb.AppendLine("- The prompt must be directly executable by a language model.");
            sb.AppendLine();
            sb.AppendLine("Example:");
            sb.AppendLine("[");
            sb.AppendLine("  {\"step\":1,\"description\":\"Analyze the selected text\",\"prompt\":\"Analyze the selected text and identify the main issues.\"}");
            sb.AppendLine("]");
            sb.AppendLine();
            sb.AppendLine("User request:");
            sb.AppendLine(userRequirement ?? "");
            return sb.ToString();
        }

        public static string BuildAgentStepPrompt(string stepPrompt, string previousContext = null)
        {
            if (string.IsNullOrWhiteSpace(previousContext))
                return stepPrompt ?? "";

            var sb = new StringBuilder();
            sb.AppendLine("Previous step results:");
            sb.AppendLine(previousContext);
            sb.AppendLine();
            sb.AppendLine("---");
            sb.AppendLine("Current task:");
            sb.AppendLine(stepPrompt ?? "");
            return sb.ToString();
        }

        public static string BuildAgentSummaryPrompt(string userRequirement, string allResults)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Summarize the following multi-step execution results into one complete, coherent final answer.");
            sb.AppendLine("Use the user's original language when possible.");
            sb.AppendLine();
            sb.AppendLine("Original user request:");
            sb.AppendLine(userRequirement ?? "");
            sb.AppendLine();
            sb.AppendLine("Step results:");
            sb.AppendLine(allResults ?? "");
            sb.AppendLine();
            sb.AppendLine("Final answer:");
            return sb.ToString();
        }

        public static string BuildAgentSystemPrompt(
            string documentSnapshot,
            string toolDescriptions,
            int toolCount,
            string skillCatalog,
            string languageMode)
        {
            string resolvedLanguage = ResolveLanguage(languageMode, null);
            bool useEnglish = string.Equals(resolvedLanguage, "English", StringComparison.OrdinalIgnoreCase);
            string langInstruction = string.Equals(resolvedLanguage, "Chinese", StringComparison.OrdinalIgnoreCase)
                ? "\n- **语言要求**：请始终用中文回复用户。"
                : useEnglish ? "\n- **Language**: Always respond in English." : "";
            bool isFullTextLoaded = !string.IsNullOrEmpty(documentSnapshot) && documentSnapshot.Contains("<document>");
            string readingStrategy = isFullTextLoaded
                ? (useEnglish
                    ? @"## Reading Strategy
- **Full text is already in the snapshot above!** Do NOT call get_doc_overview, read_large_chunk, or read_paragraphs — the entire document is already visible to you.
- **Go straight to analysis**: Review the content directly, identify issues, and call add_comment or replace_exact_text as needed."
                    : @"## 阅读策略
- **文档全文已在上方快照中！** 不要调用 get_doc_overview、read_large_chunk 或 read_paragraphs — 全文已经完整呈现在你面前。
- **直接开始审查**：直接分析已有内容，发现问题后使用 add_comment 或 replace_exact_text。")
                : (useEnglish
                    ? @"## Reading Strategy
- **Read in bulk first**: Start with `get_doc_overview` to understand the document, then use `read_large_chunk(start_paragraph: 1)` to read large blocks (50 paragraphs / 10,000+ words) at once, instead of reading a few paragraphs at a time with `read_paragraphs`.
- **Map long documents first**: For documents over 100 paragraphs (20+ pages), call `build_semantic_map()` first to get an overview map, then use `read_large_chunk` to read specific regions of interest.
- **Minimize round-trips**: Get as much information as possible in each observation. When reading, err on the side of reading more — the tool will auto-truncate long results.
- **Think globally**: Read the full document before making analysis or edits. Build a global terminology list and logic framework, catch cross-paragraph contradictions."
                    : @"## 阅读策略建议
- **优先大块读取**：第一步使用 `get_doc_overview` 了解文档概况，然后直接使用 `read_large_chunk(start_paragraph: 1)` 一次性读入大量内容（50段/10000+字），而不是用 `read_paragraphs` 几段几段地看。
- **超长文档先画地图**：如果文档超过 100 段（约 20+ 页），第一步先调用 `build_semantic_map()` 获取全文语义地图，定位感兴趣的区域，再用 `read_large_chunk` 精读目标区域。
- **减少往返**：尽量在一次 Observation 中获取足够的信息，避免在简单的信息获取上浪费任务轮数。读取时宁可多读也不要少读。
- **全局思维**：读完全文后再做分析和修改。建立全局术语表和逻辑框架，找出跨段落的矛盾。");

            return useEnglish
                ? $@"You are a powerful Microsoft Word AI Assistant. You can read, edit, format, and review the currently open Word document by calling tools.
Follow the ReAct pattern strictly: output thought (reasoning) first, then action. Do not skip the thought step.

## Document Content Snapshot
{documentSnapshot}

## Available Tools (call one at a time)
{toolDescriptions}
{toolCount + 1}. finish(message) — Call when the user's requirement is fully satisfied. message should be a summary for the user. When calling finish, only the message field is needed in params.

## Intent Routing
- First classify the request as conversation, document work, or a multi-step task.
- For greetings, general questions, brainstorming, or other requests that do not need the Word document, answer directly with `finish`; do not call document tools.
- For a narrow read-only request, call only the smallest tool needed. Do not load the whole document just to answer a local question.
- For document edits, use tools only when the user clearly asks to create, change, format, or review document content. Preserve the document until the requested operation is complete.
- If the request is ambiguous or an edit could have materially different results, use `ask_human` before changing the document.

## Extended Skills
{skillCatalog}

If the user's task closely matches an extended skill, **you MUST use read_skill to load it first**, then follow the skill's instructions strictly.

## Important Rules
- When performing write operations (insert, replace, highlight), the system automatically enables Track Changes mode. All modifications will be marked as revisions.
- For a focused wording change, prefer replace_exact_text over replace_paragraph so unchanged text and formatting remain untouched.
- Observe each step's results carefully and adjust subsequent operations accordingly.
- **Do not call the same tool repeatedly**: get_doc_overview only needs to be called once. If the document is empty, use insert_text_at_end directly.
- **Failure retry strategy**: If you encounter [execution exception] or [error], read the error message carefully, modify params, and retry. Do not blindly repeat the same attempt.
- **Eliminate redundancy**: If action is set to finish and done is true, no other tool parameters are needed.
- **Ask proactively**: When key information is missing, there is ambiguity, or user decisions are needed, use the ask_human tool to ask the user. Do not guess.
- **Progress check**: In each thought, briefly summarize what has been accomplished and what remains. Your observation history is preserved in this conversation — review it before acting.
- **Memory**: All previous tool results are available in this conversation history. You do NOT need to re-read content you've already seen.{langInstruction}
- **CRITICAL — insert_text_at_end rule**: NEVER use insert_text_at_end to write your analysis, thinking process, or step-by-step reasoning. It is ONLY for final polished content the user wants in their document. Write your thought process only in the JSON ""thought"" field.

{readingStrategy}

## Output Format
You must and can only output a valid JSON object. Do NOT include Markdown code markers (no ```json ```):
{{
  ""thought"": ""Analyze the current observation and decide which tool to use next"",
  ""action"": ""tool_name"",
  ""params"": {{ ""param_name"": ""param_value"" }},
  ""done"": false
}}
If the task is completed, set action to ""finish"" and done to true."
                : $@"你是一个强大的微软 Word AI 智能助理。你可以通过调用工具来阅读、编辑、排版和审查当前打开的 Word 文档。
请严格遵循 ReAct 模式：先输出 thought（思考推理），再输出 action（行动），不要跳过 thought。

## 当前文档内容快照
{documentSnapshot}

## 可用工具 (每次只能调用一个)
{toolDescriptions}
{toolCount + 1}. finish(message) — 认为用户需求已经完全满足时调用，message 写总结汇报给用户。调用 finish 时，params 中只需 message 字段。

## 意图路由
- 第一步先判断用户是在聊天问答、处理文档，还是要求执行多步骤任务。
- 对于问候、常识问答、头脑风暴以及不需要读取 Word 文档的请求，直接使用 `finish` 回复，不要调用文档工具。
- 对于局部的只读问题，只调用能回答问题的最小工具，不要为了回答一个局部问题加载全文。
- 只有用户明确要求创建、修改、排版或审查文档内容时，才调用写入或审查工具，并在完成请求前保持文档内容可控。
- 当需求有歧义，或修改可能存在多种重要结果时，先使用 `ask_human` 询问，不要自行猜测后改文档。

## 扩展技能库 (Skills)
{skillCatalog}

如果用户的任务与某个扩展技能高度匹配，**务必第一步使用 read_skill 加载它**，然后严格按照技能指令编排工具操作。

## 重要规则
- 执行""写""操作（insert、replace、highlight）时，系统会自动开启修订模式，所有改动会被标记为修订。
- 局部字词修改优先使用 replace_exact_text，不要整段调用 replace_paragraph，以保留未变化的文字和格式。
- 请仔细观察每一步的结果，根据实际情况调整后续操作。
- **不要重复调用同一工具**：get_doc_overview 只需调用一次。如果文档为空，直接用 insert_text_at_end 插入内容。
- **失败重试策略**：如果遇到[执行异常]或[错误]，请认真阅读错误信息，修改 params 后再试，不要盲目重复相同的尝试。
- **消除冗余**：如果 action 设置为 finish，done 设为 true，不需要填写其他工具参数。
- **主动提问**：当文档中缺失关键信息、存在歧义、或需要用户做出决策时，使用 ask_human 工具向用户提问。不要自行猜测。
- **进度核对**：每次思考时简要总结当前已完成什么、还剩什么。你的历史观察结果已经保存在对话历史中 — 行动前请回顾。
- **记忆**：之前所有工具的执行结果都在对话历史中，你不需要重新读取已经看过的段落。{langInstruction}
- **严格禁止 — insert_text_at_end 使用规则**：严禁使用 insert_text_at_end 写入你的分析过程、思考步骤或中间推理。该工具只用于输出最终精炼内容（如润色后的全文、生成的文本）。思考过程只能写在 JSON 的 ""thought"" 字段中。

{readingStrategy}

## 输出格式要求
你必须且只能输出一个合法的 JSON 对象，不要输出 Markdown 代码标记（不要用 ```json ```）：
{{
  ""thought"": ""分析目前的观察结果，决定下一步该用什么工具"",
  ""action"": ""工具名称"",
  ""params"": {{ ""参数名"": ""参数值"" }},
  ""done"": false
}}
如果任务已完成，action 请设为 ""finish""，done 设为 true。";
        }

        public static string BuildReplyLanguagePrefix(string language)
        {
            if (string.Equals(language, "Chinese", StringComparison.OrdinalIgnoreCase))
                return "请始终使用中文回复用户，不要切换为英文。\n";
            if (string.Equals(language, "English", StringComparison.OrdinalIgnoreCase))
                return "Please always reply in English and do not switch to another language.\n";
            return "Please reply in the user's language.\n";
        }

        /// <summary>
        /// Resolve the response language for the current request.
        /// Auto mode follows Chinese input instead of treating every non-explicit mode as English.
        /// </summary>
        public static string ResolveLanguage(string languageMode, string inputText)
        {
            if (string.Equals(languageMode, "Chinese", StringComparison.OrdinalIgnoreCase))
                return "Chinese";
            if (string.Equals(languageMode, "English", StringComparison.OrdinalIgnoreCase))
                return "English";

            return ContainsChineseCharacters(inputText) ? "Chinese" : "English";
        }

        public static bool ResolveUseEnglish(string languageMode, string inputText)
        {
            return string.Equals(ResolveLanguage(languageMode, inputText), "English", StringComparison.OrdinalIgnoreCase);
        }

        private static bool ContainsChineseCharacters(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;

            foreach (char c in text)
            {
                if ((c >= '\u3400' && c <= '\u4DBF') ||
                    (c >= '\u4E00' && c <= '\u9FFF') ||
                    (c >= '\uF900' && c <= '\uFAFF'))
                    return true;
            }

            return false;
        }

        public static string BuildConnectionTestPrompt()
        {
            return "Reply with exactly OK.";
        }

        public static string BuildPlanDecisionSystemPrompt(bool english)
        {
            return english
                ? "You are a Word AI assistant. Before executing the task, output only the plan. Do not call tools."
                : "你是一个 Word AI 助理，请在执行任务前先输出计划。仅输出计划内容，不要调用工具。";
        }

        public static string BuildPlanDecisionUserPrompt(string requirement, bool english)
        {
            return english
                ? $"[System] Before executing the user's task, decide whether planning is needed.\n- If the request is **simple** (greeting, quick Q&A, trivial read, no multi-step editing needed), respond with only **[SIMPLE]** to skip planning.\n- If the request needs **multiple steps or document editing**, output a numbered execution plan prefixed with **[PLAN]**.\n\nThe user's requirement is: {requirement}"
                : $"[系统] 在正式执行任务前，先判断是否需要计划。\n- 如果用户需求**非常简单**（问候、快速问答、纯阅读、无需多步操作），直接回复 **[SIMPLE]** 跳过计划。\n- 如果需要**多步操作或文档编辑**，请输出清晰的执行计划，以编号列表形式输出，开头加 **[PLAN]** 标记。\n\n用户的需求是：{requirement}";
        }

        public static string BuildSelectionHint(int paragraphIndex, string preview, bool english)
        {
            return english
                ? $"\n\n[IMPORTANT] The user currently has text selected at paragraph ~{paragraphIndex}:\n「{preview}」\nPlease prioritize analyzing this selected region first!\n"
                : $"\n\n[重要] 用户当前在第 {paragraphIndex} 段附近选中了文本：\n「{preview}」\n请优先围绕用户的选中区域执行审查指令！\n";
        }

        public static string BuildAgentUserRequirement(string requirement, bool english)
        {
            return english ? $"User's final requirement: {requirement}" : $"用户的最终需求：{requirement}";
        }

        public static string BuildToolObservation(string action, string result, bool english)
        {
            return english
                ? $"[System] Tool {action} result (Observation):\n{result}"
                : $"[系统] 工具 {action} 执行结果（观察）：\n{result}";
        }

        public static string BuildUserAdjustmentMessage(string userDecision, bool english)
        {
            return english
                ? $"[System] The user suggests: {userDecision}. Adjust your approach and retry."
                : $"[系统] 用户建议：{userDecision}。请根据此建议调整策略后重试。";
        }

        public static string BuildRepeatedActionWarning(string action, bool english)
        {
            return english
                ? $"[BLOCKED] You already tried \"{action}\" with the same parameters recently and it did not work. Reflect on why it failed, then try a different tool or parameters."
                : $"[拒绝执行] 你已经在近期尝试过“{action}”且参数完全相同，之前没有效果。请反思为什么前几次尝试失败，尝试不同的工具或参数。";
        }

        public static string BuildConversationCompressionSummary(IEnumerable<string> actionLog, int keepCompleteRounds, bool english)
        {
            string actions = string.Join("\n", actionLog ?? Enumerable.Empty<string>());
            if (!string.IsNullOrWhiteSpace(actions))
            {
                return english
                    ? $"[System] Previous operations (results compressed to save space):\n{actions}\n\nFull records of the most recent {keepCompleteRounds} rounds are preserved below."
                    : $"[系统压缩] 此前完成的操作（详细结果已压缩，保留操作日志）：\n{actions}\n\n最近 {keepCompleteRounds} 轮完整记录见下方。";
            }

            return english ? "[System] Previous conversation rounds were compressed." : "[系统] 旧的对话轮次已压缩。";
        }

        public static string BuildDocumentSnapshotEmpty(bool english)
        {
            return english ? "(empty document)" : "（空文档）";
        }

        public static string BuildDocumentSnapshotFullHeader(int paragraphCount, int characterCount, bool english)
        {
            return english
                ? $"(Full document loaded: {paragraphCount} paragraph(s), approximately {characterCount} characters)"
                : $"（文档共 {paragraphCount} 段，约 {characterCount} 字符，已全文加载）";
        }

        public static string BuildDocumentSnapshotOutlineHeader(int paragraphCount, int characterCount, bool english)
        {
            return english
                ? $"(Document outline: {paragraphCount} paragraph(s), approximately {characterCount} characters)"
                : $"（文档共 {paragraphCount} 段，约 {characterCount} 字符，以下为骨架导航）";
        }

        public static string FormatSnapshotParagraph(int paragraphIndex, string text, bool english)
        {
            return english ? $"[Paragraph {paragraphIndex}] {text}" : $"[段落{paragraphIndex}] {text}";
        }

        public static string FormatSnapshotOutlineEntry(int startParagraph, int endParagraph, string summary, bool english)
        {
            return english ? $"- [Paragraphs {startParagraph}-{endParagraph}] {summary}" : $"- [{startParagraph}-{endParagraph}段] {summary}";
        }

        public static string BuildDocumentSnapshotEmptyParagraph(bool english)
        {
            return english ? "(empty paragraph)" : "（空段落）";
        }

        public static string BuildDocumentSnapshotFullTextNotice(bool english)
        {
            return english
                ? "[System] The snapshot above includes the full document. Do not call read_large_chunk or get_doc_overview; begin the review directly."
                : "[系统提示]：上述快照已包含文档全文！请勿再调用 read_large_chunk 或 get_doc_overview 工具，直接开始审查分析！";
        }

        public static string BuildDocumentSnapshotOutlineNotice(bool english)
        {
            return english
                ? "[System] The snapshot above is an outline only. Use read_large_chunk to retrieve document details."
                : "[系统提示]：上述仅为文档大纲骨架，请使用 read_large_chunk 获取正文细节。";
        }

        public static string BuildAgentApiErrorRepairPrompt(string apiError, bool english)
        {
            return english
                ? (apiError ?? "") + "\nCheck the model configuration, API key, and service endpoint in Settings."
                : (apiError ?? "") + "\n请检查设置中的模型配置、API Key 或服务地址是否正确。";
        }

        public static string BuildAgentJsonRepairPrompt(string rawOutput, bool english)
        {
            return english
                ? $"[System] JSON parsing failed. Original output: {rawOutput}\nReturn strictly valid JSON without Markdown code fences."
                : $"[系统提示] JSON 解析失败。原始输出：{rawOutput}\n请严格按 JSON 格式输出，不要包含 Markdown 代码标记。";
        }

        public static string BuildReviewSuggestionPrompt(string originalText, string instruction, string writingProfile)
        {
            var sb = new StringBuilder();
            sb.AppendLine("You are a careful writing editor working inside Microsoft Word.");
            sb.AppendLine("Return only the revised text. Do not add explanations, markdown fences, alternatives, comments, or headings.");
            sb.AppendLine("Preserve the author's meaning, paragraph structure, proper nouns, character voice, and unusual spellings unless the instruction says otherwise.");
            if (!string.IsNullOrWhiteSpace(writingProfile))
            {
                sb.AppendLine();
                sb.AppendLine("Writing profile / story bible:");
                sb.AppendLine(writingProfile.Trim());
            }
            sb.AppendLine();
            sb.AppendLine("Editing instruction:");
            sb.AppendLine(instruction ?? "");
            sb.AppendLine();
            sb.AppendLine("Text to revise:");
            sb.AppendLine(originalText ?? "");
            return sb.ToString();
        }

        public static string BuildReviewSuggestionBatchPrompt(
            IEnumerable<KeyValuePair<string, string>> items,
            string instruction,
            string writingProfile)
        {
            var payload = new JArray((items ?? Enumerable.Empty<KeyValuePair<string, string>>())
                .Select(item => new JObject
                {
                    ["id"] = item.Key,
                    ["text"] = item.Value ?? ""
                }));

            var sb = new StringBuilder();
            sb.AppendLine("You are a careful writing editor working inside Microsoft Word.");
            sb.AppendLine("Revise every input item independently and return only a valid JSON array.");
            sb.AppendLine("Each output item must contain the unchanged id and the revised text: {\"id\":\"1\",\"text\":\"...\"}.");
            sb.AppendLine("Return exactly one output item for every input item, in the same order. Do not add markdown, explanations, alternatives, or extra fields.");
            sb.AppendLine("Preserve meaning, paragraph structure, proper nouns, numbers, citations, unusual spellings, and character voice unless the instruction says otherwise.");
            if (!string.IsNullOrWhiteSpace(writingProfile))
            {
                sb.AppendLine();
                sb.AppendLine("Writing profile / story bible:");
                sb.AppendLine(writingProfile.Trim());
            }
            sb.AppendLine();
            sb.AppendLine("Editing instruction:");
            sb.AppendLine(instruction ?? "");
            sb.AppendLine();
            sb.AppendLine("Input JSON:");
            sb.AppendLine(payload.ToString(Newtonsoft.Json.Formatting.None));
            return sb.ToString();
        }

        public static string BuildSkillAuthoringPrompt(string apiDocumentation, string skillName, string description, string goal, bool english)
        {
            string language = english ? "English" : "Chinese";
            return $@"
You are creating a local WordOllama SKILL.md workflow in {language}.
The skill name is: {skillName}
The skill description is: {description}
The user's workflow goal is:
{goal}

Use only the exact Agent tool names and JSON parameters from the API catalog below. Prefer Word document read/write tools for ordinary Word work. Do not invent tools or parameters. Do not use raw COM, VBA, PowerShell, Python, shell commands, arbitrary code, hidden network requests, or direct file-system operations. If a workflow would edit the document, state that the user must review or confirm the changes. External network, local process, and MCP tools may only be used when the user explicitly asks for them and the Agent permission dialog allows them.

Return only the Markdown body of the skill. Do not return YAML front matter and do not wrap the answer in a code fence. Include:
1. When to use this skill.
2. A concise step-by-step workflow.
3. Which catalog tools are used and the important parameters.
4. Safety and review notes.

API catalog:
---
{apiDocumentation}
---
";
        }

        public static string FormatAgentStepResult(int stepNumber, string description, string result)
        {
            return $"\n\n[Step {stepNumber}: {description}]\n{result}";
        }

        private static string BuildCustomPrompt(string customInstruction, string selectedText)
        {
            var sb = new StringBuilder();
            sb.AppendLine(string.IsNullOrWhiteSpace(customInstruction)
                ? "Process the following text according to the user's custom instruction."
                : customInstruction.Trim());
            sb.AppendLine();
            sb.AppendLine("Text to process:");
            sb.AppendLine("<<<");
            sb.AppendLine(selectedText ?? "");
            sb.AppendLine(">>>");
            return sb.ToString();
        }

        private static string BuildTaskInstruction(PromptTaskType taskType, string targetLanguage)
        {
            switch (taskType)
            {
                case PromptTaskType.Translate:
                    string language = string.IsNullOrWhiteSpace(targetLanguage) ? "the target language" : targetLanguage.Trim();
                    return "Translate the text into " + language + ". Preserve paragraph breaks, numbering, placeholders, citations, names, dates, amounts, and legal terms.";
                case PromptTaskType.Expand:
                    return "Expand the text with useful detail while preserving the original meaning, style, tone, and language.";
                case PromptTaskType.Polish:
                    return "Polish the text to improve clarity, wording, grammar, structure, and flow while preserving the original meaning and language.";
                case PromptTaskType.Simplify:
                    return "Simplify and condense the text so it is clearer and easier to understand while preserving all important meaning and the original language.";
                case PromptTaskType.Summarize:
                    return "Summarize the text by extracting its core ideas and key points. Preserve important facts, dates, obligations, parties, amounts, and the original language.";
                case PromptTaskType.ContinueWriting:
                    return "Continue writing from the selected text. Match its style, tone, point of view, and language. Do not repeat the selected text.";
                case PromptTaskType.TypoFix:
                    return "Correct only clear typos, grammar mistakes, punctuation errors, and malformed sentences. Do not rewrite the author's style.";
                case PromptTaskType.LegalRisk:
                    return "Review the selected clause or passage for material legal or contractual risks.";
                case PromptTaskType.FairRewrite:
                    return "Rewrite the contract text to be fairer, clearer, and more balanced while preserving its legal purpose and original language.";
                default:
                    return "Edit the selected text according to the requested task.";
            }
        }
    }
}
