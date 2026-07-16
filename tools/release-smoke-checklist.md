# WordOllama Release Smoke Checklist

## Automated checks

1. Build the current Release output:
   `MSBuild.exe WordOllama.sln /t:Build /p:Configuration=Release`
2. Run `tools/verify-localization.ps1`.
3. Run `tools/run-core-smoke-test.ps1`.
4. Run `tools/verify-release.ps1`.
5. In Advanced Installer, sync files from `WordOllama/bin/Release/`; do not use a stale copied folder.

## Word host

1. Install the package on a clean Windows user profile with .NET Framework 4.8 and WebView2 available.
2. Open Word, enable only the current WordOllama add-in, and verify the ribbon appears.
3. Open Settings and verify the active provider connection test, Agent mode, diagnostics controls, MCP health, and API/Skill tab.
4. Run one read-only Agent task, one tracked-change task, and one Review Workspace task.
5. Verify Word word count, comments, bookmarks, tables, and built-in review controls still work.
6. Close Word and verify no add-in-owned background process or tray icon remains.

## WPS host

1. Install on a machine with WPS Office and the required VSTO/Office runtime available.
2. Verify the WPS add-in registration points to the current installed output and that the WPS ribbon loads.
3. Run one document read and one document write operation.

## Online/offline and upgrade

1. Online package: verify update discovery, certificate behavior, and rollback/repair from Add or Remove Programs.
2. Offline package: verify it does not require the online update certificate and that user-supplied .NET/WebView2 prerequisites are reported clearly.
3. Upgrade from the previous version and confirm the new manifest, assembly, and ribbon are loaded.
4. Confirm existing encrypted API keys and MCP configuration survive the upgrade.
