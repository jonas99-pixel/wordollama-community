# Windows 11 / Word 2021 deployment

## Current state
The upstream community build produces DLL assemblies only; **it is not an installer**. This branch adds staging and installation entrypoints for a **separately signed and published VSTO deployment**, not a finished signed binary. Neither the Windows 11 tablet nor Word has been tested.

## On a Windows build machine

1. Install Visual Studio 2022 with **.NET desktop development** and **Office/SharePoint development** components, .NET Framework 4.8 targeting pack, and restore NuGet packages. Assess dependencies.
2. Open `WordOllama.sln`. Run the source build and `tools/verify-release.ps1`.
3. Use a **local, uncommitted** Visual Studio configuration with `OpenSourceBuild=false`, configure a signing identity under your control, and run the standard VSTO **Publish** workflow. It must emit `WordOllama.vsto` and its `Application Files` directory. The official source tree deliberately excludes signing credentials and manifests.
4. Stage the complete deployment with:

```powershell
powershell -File .\tools\prepare-office2021-package.ps1 -SourceManifest 'C:\path\to\publish\WordOllama.vsto'
```

Do not add private keys to the repository. The staging check only checks that a signature element exists, **not that the certificate is trusted or the signature is valid**.

## On the Windows 11 / Office 2021 tablet

1. Close Word. Ensure official Microsoft VSTO Runtime, .NET Framework 4.8 and WebView2 runtime are installed.
2. Copy the complete deployment folder to the tablet; retain the Application Files tree and never accept an unknown publisher or certificate trust request.
3. From the deployed folder:

```powershell
powershell -File .\install-office2021.ps1 -ManifestPath .\WordOllama.vsto
```

4. Open Word and confirm ribbon loading. Configure the OpenAI-compatible API to the **homelab address**, not 127.0.0.1 (which refers to the tablet). Enable no external command, Python, HTTP, local-file-write or MCP Agent tools for initial testing.
5. Validate a disposable document: text editing, heading formatting, table creation, tracked edits, save, close and reopen.

## Completion criteria
- Word 2021 loads the add-in; local-model inference succeeds; editing and tracked changes are visually confirmed.
- No system-wide certificate import or unintended update service.
- Installation/removal work using Microsoft's VSTO deployment mechanism.

**Not completed:** Signed VSTO publication, installation, Office 2021 E2E test. Do not distribute this as a ready-made installer.
