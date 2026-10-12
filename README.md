<div align="center">

<h1>📝 WordOllama Community Edition</h1>

<p>Document-aware AI tools for Microsoft Word and compatible WPS Word hosts.</p>

<p>
  <strong>English</strong> · <a href="README.zh-CN.md">简体中文</a>
</p>

<p>
  <img alt="GPL-3.0" src="https://img.shields.io/badge/license-GPL--3.0-blue?style=for-the-badge">
  <img alt=".NET Framework 4.8" src="https://img.shields.io/badge/.NET_Framework-4.8-512BD4?style=for-the-badge&logo=.net&logoColor=white">
  <img alt="Word and WPS" src="https://img.shields.io/badge/host-Word%20%2F%20WPS-2ea44f?style=for-the-badge">
  <img alt="MCP transports" src="https://img.shields.io/badge/MCP-stdio%20%7C%20HTTP%20%7C%20SSE-0f766e?style=for-the-badge">
</p>

<p><strong>Local models · Compatible APIs · Skills · MCP · Review workflows</strong></p>

</div>

WordOllama Community Edition is the clean open-source edition of the main WordOllama project. The main project website is [www.wordollama.com](https://www.wordollama.com).

WordOllama is an open-source AI add-in for Microsoft Word and compatible WPS Word hosts. It brings document-aware chat, editing, translation, legal-document tools, review workflows, Skills, and MCP tools into the document window.

The project is designed for people who want to use a local model or connect an API-compatible provider without moving their normal Word workflow into a separate application.

## Version 2.5.0 — long-document review

Whole-document review now uses consecutive batches with configurable limits, progress and coverage reporting. The built-in Agent can create a native review copy, generate suggestions or findings, resume unfinished local sessions, apply precise tracked edits, and save the copy and reports. Review includes body text and tables, section headers and footers, footnotes, endnotes, and text boxes. AI reviewer names can use a fixed name, a date, or a session suffix. See [the long-document review guide](docs/long-document-review.md) and the original [English update log](WordOllama/Resources/updateLog.en.txt).

This is a source release. To update an existing checkout and build on Windows:

```powershell
git pull --ff-only origin main
powershell -ExecutionPolicy Bypass -File .\tools\build-open-source.ps1 -Configuration Release
powershell -ExecutionPolicy Bypass -File .\tools\verify-release.ps1 -SkipBuild
```

For loading your local build into Word, follow the certificate and deployment guidance in **Build** below. Compilation and automated checks cover the review workflow and session handling; performance and proofreading quality on real, very long Word documents still require testing. Local Ollama review does not require an online document service; restoring NuGet dependencies during a first build may require network access.

## ✨ Highlights

- Agent workflows that can inspect, edit, format, review, and navigate a Word document through a controlled COM tool layer.
- Ollama support, plus OpenAI-compatible and Anthropic-compatible API formats. This can cover vLLM, llama.cpp servers, LM Studio, hosted APIs, and other compatible gateways.
- MCP support for stdio, Streamable HTTP, and SSE transports.
- A separate Review Workspace for paragraph-level suggestions, accept/insert/comment/skip actions, and navigation back to the document.
- Local Skills stored as `SKILL.md` files. The Agent API documentation in the add-in can be used to draft a new Skill.
- Chinese and English UI resources, with automatic language detection available in settings.
- Legal provision lookup through an external public API. The feature is optional and requires network access.

## 🔐 Privacy and permissions

The WordOllama project does not operate telemetry or a document-collection backend. When the user starts an AI operation, the add-in forwards the required document text to the provider configured by the user. Local Ollama requests stay on the configured local endpoint; online providers and remote MCP servers receive data according to their own policies. Do not send confidential material to a provider unless you have permission to do so.

High-risk external Agent tools are disabled by default. These include local commands, Python scripts, arbitrary HTTP requests, restricted local-file search, and MCP tools. Local-file search is limited to the active document directory and Word Agent Skill directories; sensitive credential and key files are skipped. The Agent does not connect configured MCP servers until the user enables external tools for the task or turns on the explicit default-permission setting. Untrusted external operations ask for approval and show a command, URL, path, Skill, or MCP server/tool. A server explicitly marked as trusted is a persistent permission decision and does not show a separate confirmation for every MCP call. Secret-like parameter fields are redacted in the approval message.

Unfinished Agent checkpoints are encrypted for the current Windows user with DPAPI. They remain local and can be cleared from the resume prompt.

The add-in does not require Office.js and does not provide a sandbox around arbitrary external programs. Treat external tool approval as a real system permission decision.

## 🧰 Requirements

For running the add-in:

- Windows with Microsoft Word or a compatible WPS Word installation.
- .NET Framework 4.8.
- Visual Studio Tools for Office Runtime (VSTO Runtime), normally installed with supported Office/Visual Studio setups.
- Microsoft Edge WebView2 Runtime for WebView-based panels. Install the Evergreen Runtime separately if it is not already present.
- A configured model endpoint. Ollama is optional; an OpenAI-compatible or Anthropic-compatible endpoint can be used instead.

For building from source:

- Visual Studio 2022 with the .NET desktop development workload and Office/SharePoint development tools.
- Word or WPS installed for manual host validation.
- NuGet package restore access. The repository intentionally does not commit the `packages` directory or installer output.

## 🛠️ Build

1. Clone the repository and open `WordOllama.sln` in Visual Studio 2022.
2. Restore NuGet packages.
3. Run `powershell -ExecutionPolicy Bypass -File .\tools\build-open-source.ps1 -Configuration Release`.
4. The public source build produces the add-in assembly and dependencies. It intentionally skips VSTO ClickOnce manifest generation. For local F5 debugging, create/select a local test certificate in Visual Studio, set the local project property `OpenSourceBuild=false`, and do not commit the certificate or generated manifest. For end-user distribution, use a separately signed deployment process.

This repository does not contain a private certificate, ClickOnce publishing URL, or installer project. The open-source build intentionally skips ClickOnce manifest generation; signing and deployment are release-owner responsibilities and should be configured outside the public source tree.

## 🔌 Provider setup

### Ollama

Install Ollama, start its local service, pull a model, and enter the endpoint and model name in WordOllama settings. The usual endpoint is `http://localhost:11434`.

### OpenAI-compatible APIs

Select the OpenAI-compatible provider and enter the chat-completions endpoint, API key, and model name. This is also the normal path for vLLM, llama.cpp server mode, LM Studio, and many hosted gateways that expose an OpenAI-compatible API.

### Anthropic-compatible APIs

Select the Anthropic provider and enter its endpoint, API key, and model name. Provider capability detection is used to adapt tool calling, streaming, vision, JSON output, and context-window behavior where supported.

## 🤖 Agent and MCP

The Agent uses Word-specific interfaces rather than exposing raw COM objects to the model. The API documentation tab lists the active tools, parameters, and effects. Use it as the source of truth when authoring a Skill.

MCP servers can be configured in settings. Enable only the individual MCP tools you need, and use the trusted-server option only for servers you control and understand. Untrusted external tools remain subject to the Agent execution policy and per-operation confirmation. Marking an MCP server trusted is an explicit persistent permission decision.

## ✅ Checks

The repository keeps lightweight checks instead of a heavyweight Office UI automation suite:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\verify-localization.ps1
powershell -ExecutionPolicy Bypass -File .\tools\run-core-smoke-test.ps1
powershell -ExecutionPolicy Bypass -File .\tools\run-mcp-smoke-test.ps1
powershell -ExecutionPolicy Bypass -File .\tools\verify-release.ps1 -SkipBuild
```

The legal API contract check is intentionally separate because it needs network access:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\legal-api-smoke-test.ps1
```

Manual validation in both Word and WPS remains necessary for COM loading, ribbon registration, WebView2, task-pane sizing, and host-specific behavior.

## 🗂️ Repository layout

- `WordOllama/`: the VSTO add-in source project.
- `WordOllama/lib/`: Agent runtime, providers, MCP bridge, configuration, and tool policy.
- `WordOllama/lib/Tools/`: Word and external Agent tools.
- `WordOllama/NewUI/`: WPF task panes and windows.
- `WordOllama/Properties/`: localized resources and assembly metadata.
- `tools/`: focused smoke and validation scripts.
- `docs/`: security, contribution, release, and third-party documentation.

## ⚠️ Limitations

- This project targets the desktop Word/WPS COM model. Office.js is not currently used.
- Microsoft Word and WPS differ in loading behavior and available COM features; test the host you intend to support.
- Online API, remote MCP, and legal lookup availability depends on network policy and the service provider.
- The public source tree is not a signed installer. Do not trust an unsigned local build as if it were an official release.

## 📄 License

WordOllama is available under the GNU GPL v3. See [LICENSE](LICENSE) and [third-party notices](docs/THIRD-PARTY-NOTICES.md).
