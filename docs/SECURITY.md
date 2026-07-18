# Security Policy

## Scope

WordOllama can send document text to configured AI providers and can optionally call external Agent tools or MCP servers. Treat provider credentials, document contents, Skill files, and MCP configuration as sensitive.

## Reporting a vulnerability

Please do not publish credentials, private documents, exploit details, or a working malicious MCP configuration in a public issue. Contact the project maintainer privately through the contact channel listed on the project website, and include:

- affected version or commit;
- host application and Windows/.NET environment;
- reproduction steps with sanitized input;
- impact and any suggested mitigation.

Allow reasonable time for triage before public disclosure.

## User safety notes

- External Agent tools are disabled by default.
- Local file search is treated as an external permission. It is limited to the current document directory and WordAgentSkills directories, skips sensitive credential files, and requires approval.
- Review every external operation approval, especially commands, URLs, and MCP calls.
- Do not mark an MCP server trusted unless you control and understand it.
- Legacy SSE message endpoints must remain on the same origin as the configured MCP server so authorization headers are not forwarded elsewhere.
- Agent recovery checkpoints are encrypted for the current Windows user with DPAPI and are removed when a task exits normally or is cancelled.
- Never commit API keys, certificates, document samples containing personal data, or local absolute paths.
