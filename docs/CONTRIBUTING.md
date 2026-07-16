# Contributing to WordOllama

Thank you for helping improve WordOllama.

## Before opening an issue

- Reproduce the problem with the current source when possible.
- Record the host (Microsoft Word or WPS), Windows version, .NET Framework version, provider type, and whether WebView2 is installed.
- Remove document text, API keys, personal paths, and MCP credentials from logs and screenshots.

## Code changes

- Keep Word COM calls on the host/UI thread through the existing execution helpers.
- Keep user-visible strings and Agent prompts in the appropriate `.resx` resource files.
- Preserve the default-deny behavior for local processes, arbitrary HTTP, Python Skills, and MCP.
- Add focused smoke coverage for shared policy or parser changes. Full Office UI automation is not required.
- Avoid committing `bin`, `obj`, `packages`, installer output, certificates, API keys, or personal deployment paths.

## Verification

Run `tools/verify-release.ps1` (or its individual localization/core checks) before submitting a change. Build the VSTO project with Visual Studio 2022 and manually check the affected Word/WPS workflow when the change touches COM or XAML.

## Pull requests

Explain the user-facing behavior, security impact, host coverage, and checks that were run. Small focused pull requests are easier to review.
