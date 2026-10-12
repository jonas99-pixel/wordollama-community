# WordOllama Release Smoke Checklist

## Automated checks

1. Build the Community source output:
   `powershell -ExecutionPolicy Bypass -File tools/build-open-source.ps1 -Configuration Release`
2. Run `tools/verify-localization.ps1`.
3. Run `tools/run-core-smoke-test.ps1`.
4. With Node.js available, run `tools/run-mcp-smoke-test.ps1`.
5. Run `tools/verify-release.ps1 -SkipBuild` to check localization, core review/session/Agent behavior, Office STA dispatch, assembly/project versions and the original embedded TXT update logs. Without `-SkipBuild`, this command first builds Release.
6. Source builds generate the add-in DLL, dependencies and language resources. Local test certificates and VSTO deployment manifests are configured outside the public source tree; they are not required by the Community source verifier.

## Word host

1. Load the local build through Visual Studio or a separately configured VSTO deployment, with .NET Framework 4.8 and WebView2 available.
2. Open Word, enable only the current WordOllama add-in, and verify the ribbon appears.
3. Open Settings and verify the active provider connection test, Agent mode, diagnostics controls, MCP health, and API/Skill tab.
4. Run one read-only Agent task, one tracked-change task, and one Review Workspace task.
5. Verify Word word count, comments, bookmarks, tables, and built-in review controls still work.
6. Close Word and verify no add-in-owned background process or tray icon remains.
7. In Agent Tracked Changes mode, request a whole-document review copy, apply the completed suggestions, and save the copy and reports. Include multiple sections, first/even headers and footers, footnotes/endnotes, linked and grouped text boxes, fields and existing revisions. Check formatting, precise locations, reviewer identities and retention of existing revision authors.
8. Interrupt generation, reopen the same saved copy, and resume the local session. Confirm completed batches are retained and unrelated document changes prevent stale writeback. Real thousand-page document performance and proofreading quality require this hosted validation.

## WPS host

1. Configure the local build on a machine with WPS Office and the required VSTO/Office runtime available.
2. Verify the WPS add-in registration points to the current installed output and that the WPS ribbon loads.
3. Run one document read and one document write operation.

## Local models and deployment

1. With a local Ollama endpoint, verify review works without an online document service; confirm requested context, output and timeout settings.
2. Rebuild from a fresh checkout after restoring NuGet dependencies. Verify the 2.5.0.0 assembly and original English/Chinese TXT resources.
3. For a local VSTO deployment, verify the generated test/deployment manifests refer to the current assembly and that Word loads the current ribbon.
4. Confirm existing encrypted API keys, MCP configuration and review settings survive updating the local build.
