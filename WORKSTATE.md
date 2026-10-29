# WORKSTATE — WordOllama Office 2021

- Result requested: working WordOllama add-in on Windows 11 tablet / Microsoft Word 2021 / local Strata OpenAI-compatible inference.
- Repository/remote: jonas99-pixel/wordollama-community / https://github.com/jonas99-pixel/wordollama-community.git
- Work branch: feature/office2021-install-package; no merge to main.
- Starting main commit: 56a105964d3c195c910736331567e8c275171f01.
- Deployment: user confirmed Visual Studio 2026 Release Any CPU build and VSTO publish succeeded; setup.exe, WordOllama.vsto and Application Files present. User confirmed installation succeeded on Office 2021 tablet and first-run wizard launches. Locally trusted development self-signed certificate used for test installation.
- Strata HTTPS: user verified from tablet PowerShell 5.1 that https://strata-hermes.raettlemail.xyz/v1/models returned HTTP 200 and POST /v1/chat/completions with model swift-1.5-iq3_xxs-amd returned HTTP 200, content OK. Existing API keys unchanged.
- Issue: WordOllama first-run capability probe fetches model list but plain chat/stream/tool tests fail with [OpenAI:transport] An error occurred while sending the request.
- Patch committed: ProviderRequestPolicy.NormalizeError now reports redacted nested exception types, HRESULTs, optional SocketError/WebStatus instead of raw exception text; no request content, endpoint URLs or secrets included in diagnostics. See b24bb2bfb6ad73567e063bd1813564ea2d43a68f.
- Validation: repository content and patch commit verified via GitHub. No Windows build, publish or tablet runtime test of the patched code completed in this environment.
- Local caveat: Windows Visual Studio source checkout may have uncommitted OpenSourceBuild=false and other publish-property edits. Inspect git status before pulling; preserve those changes, do not overwrite.
- Next: fetch branch in Visual Studio 2026; inspect local changes; Release build (OpenSourceBuild=false), publish and install a newer version on tablet; reproduce capability probe and capture redacted error chain. Then fix the verified root cause and complete end-to-end acceptance.
- Stop condition: no changes to Strata, Traefik, keys, main branch, deployment server or certificate trust without explicit further authorization.
