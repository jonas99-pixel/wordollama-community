<div align="center">

<h1>📝 WordOllama 社区版</h1>

<p>面向 Microsoft Word 和兼容 WPS Word 宿主的文档智能 AI 工具。</p>

<p>
  <a href="README.md">English</a> · <strong>简体中文</strong>
</p>

<p>
  <img alt="GPL-3.0" src="https://img.shields.io/badge/license-GPL--3.0-blue?style=for-the-badge">
  <img alt=".NET Framework 4.8" src="https://img.shields.io/badge/.NET_Framework-4.8-512BD4?style=for-the-badge&logo=.net&logoColor=white">
  <img alt="Word and WPS" src="https://img.shields.io/badge/host-Word%20%2F%20WPS-2ea44f?style=for-the-badge">
  <img alt="MCP transports" src="https://img.shields.io/badge/MCP-stdio%20%7C%20HTTP%20%7C%20SSE-0f766e?style=for-the-badge">
</p>

<p><strong>本地模型 · 兼容 API · Skill · MCP · 文档审阅工作流</strong></p>

</div>

WordOllama 社区版是 WordOllama 主项目的清洁开源版。主项目官网：[www.wordollama.com](https://www.wordollama.com)。

WordOllama 是面向 Microsoft Word 以及兼容 WPS Word 宿主的开源 AI 加载项。它把文档感知对话、编辑、翻译、法律文档工具、审阅工作流、Skill 和 MCP 工具带进文档窗口。

项目适合希望继续在 Word 中工作，同时使用本地模型或兼容 API 的用户。

## ✨ 主要功能

- Agent 可以通过受控的 Word COM 工具层读取、编辑、格式化、审阅和定位文档。
- 支持 Ollama，也支持通用的 OpenAI 兼容和 Anthropic 兼容 API 格式，因此可以连接 vLLM、llama.cpp、LM Studio、在线 API 以及其他兼容网关。
- 支持 stdio、Streamable HTTP 和 SSE 三种 MCP transport。
- 独立的审阅工作台，支持按段落生成建议、接受、插入、批注、跳过，并可跳回文档位置。
- 本地 Skill 以 `SKILL.md` 保存。加载项内的 Agent 接口文档可以用来辅助创建 Skill。
- 中英文资源和自动语言模式。
- 通过外部公益 API 查询法条。该功能需要网络，服务不可用时不会影响其他功能。

## 🔐 隐私和权限

WordOllama 项目方不运营遥测，也没有收集文档的后端。当用户主动开始 AI 操作时，加载项会把所需文档文本发送给用户配置的模型服务。Ollama 请求通常留在配置的本机地址；在线模型和远程 MCP 服务会按照各自的隐私政策接收数据。处理敏感文档前，请确认服务和组织政策允许这样做。

高风险 Agent 外部工具默认关闭，包括本地命令、Python 脚本、任意 HTTP 请求、受限本地文件检索和 MCP 工具。本地文件检索仅限当前文档目录和 Word Agent Skill 目录，并会跳过凭据、私钥等敏感文件。未获得用户授权前，Agent 不会连接配置的 MCP 服务器。用户可以仅为当前任务启用，也可以在设置中主动开启默认权限。不可信外部操作仍会单独询问，并显示命令、URL、路径、Skill 或 MCP 服务/工具；参数中的密钥、令牌、密码等字段会在确认提示中隐藏。把 MCP 服务器标记为可信，是一个持久化的权限决定，此后不会对每次 MCP 调用再次弹窗。

未完成的 Agent 检查点使用 Windows DPAPI 按当前用户加密，数据只保存在本机，并可在恢复提示中清除。

加载项不依赖 Office.js，也没有为任意外部程序提供沙箱。请把外部工具授权当作真正的系统权限决定。

## 🧰 运行环境

- Windows，以及 Microsoft Word 或兼容的 WPS Word。
- .NET Framework 4.8。
- Visual Studio Tools for Office Runtime（VSTO Runtime）。
- WebView 面板需要 Microsoft Edge WebView2 Runtime；如果系统没有，请另外安装 Evergreen Runtime。
- 一个已经配置好的模型地址。Ollama 不是强制依赖，也可以使用 OpenAI 或 Anthropic 兼容 API。

## 🧱 源码构建环境

- Visual Studio 2022，安装 .NET 桌面开发和 Office/SharePoint 开发工具。
- Word 或 WPS，用于手工验证宿主加载。
- 可以恢复 NuGet 包的网络环境。仓库不提交 `packages` 目录，也不提交安装器输出。

## 🛠️ 构建

1. 克隆仓库，在 Visual Studio 2022 中打开 `WordOllama.sln`。
2. 恢复 NuGet 包。
3. 运行 `powershell -ExecutionPolicy Bypass -File .\tools\build-open-source.ps1 -Configuration Release`。
4. 公开源码构建会生成加载项程序集和依赖，但会主动跳过 VSTO ClickOnce 清单。需要在 Visual Studio 中调试时，请创建或选择本地测试证书，将本地项目属性 `OpenSourceBuild=false`，不要提交证书和生成的清单。正式分发请在源码之外完成签名和部署。

本仓库不包含个人证书、ClickOnce 发布地址和安装器工程。开源构建会主动跳过 ClickOnce 清单生成；签名和正式发布应由发布者在源码之外配置。

## 🔌 模型配置

### Ollama

安装并启动 Ollama，准备一个模型，然后在 WordOllama 设置中填写地址和模型名。常用地址是 `http://localhost:11434`。

### OpenAI 兼容 API

选择 OpenAI 兼容模式，填写 chat completions 地址、API Key 和模型名。这也是连接 vLLM、llama.cpp server、LM Studio 以及许多在线网关的常用方式。

### Anthropic 兼容 API

选择 Anthropic 模式，填写地址、API Key 和模型名。加载项会根据服务能力调整工具调用、流式输出、视觉、JSON 输出和上下文窗口行为。

## 🤖 Agent 和 MCP

Agent 使用面向 Word 的接口，不会把原始 COM 对象直接交给模型。接口文档页会列出当前可用工具、参数和副作用，可以以此为依据编写 Skill。

MCP 服务器在设置中配置。只勾选确实需要的工具；只有在服务器由你控制且你理解其行为时，才考虑设为可信服务器。不可信外部工具仍受 Agent 执行模式和每次操作确认约束。把 MCP 服务器标记为可信，是一个明确的持久化权限决定。

## ✅ 检查命令

仓库保留轻量检查，不引入重量级 Office UI 自动化测试：

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\verify-localization.ps1
powershell -ExecutionPolicy Bypass -File .\tools\run-core-smoke-test.ps1
powershell -ExecutionPolicy Bypass -File .\tools\run-mcp-smoke-test.ps1
powershell -ExecutionPolicy Bypass -File .\tools\verify-release.ps1 -SkipBuild
```

法条接口检查需要网络，因此单独运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\legal-api-smoke-test.ps1
```

Word 和 WPS 的 COM 加载、Ribbon、WebView2、任务窗格尺寸以及宿主差异仍需要手工验证。

## 🗂️ 目录结构

- `WordOllama/`：VSTO 加载项源码工程。
- `WordOllama/lib/`：Agent、模型提供方、MCP、配置和权限策略。
- `WordOllama/lib/Tools/`：Word 工具和外部工具。
- `WordOllama/NewUI/`：WPF 任务窗格和窗口。
- `WordOllama/Properties/`：程序集信息和本地化资源。
- `tools/`：轻量 smoke 和验证脚本。
- `docs/`：安全、贡献、发布和第三方说明文档。

## ⚠️ 限制

- 项目面向桌面 Word/WPS COM 模型，目前没有使用 Office.js。
- Word 和 WPS 的加载行为及 COM 能力存在差异，请针对目标宿主测试。
- 在线 API、远程 MCP 和法条服务是否可用取决于网络策略及服务方状态。
- 公开源码不是签名安装包。未签名本地构建不应被当作官方发布版本。

## 📄 许可证

WordOllama 使用 GNU GPL v3，详见 [LICENSE](LICENSE) 和 [第三方声明](docs/THIRD-PARTY-NOTICES.md)。
