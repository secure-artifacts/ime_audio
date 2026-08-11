# VoiceIME — Windows-wide voice input

[![Latest release](https://img.shields.io/github/v/release/secure-artifacts/ime_audio?display_name=tag)](https://github.com/secure-artifacts/ime_audio/releases/latest)
[![Release downloads](https://img.shields.io/github/downloads/secure-artifacts/ime_audio/total)](https://github.com/secure-artifacts/ime_audio/releases)
[![Release build](https://github.com/secure-artifacts/ime_audio/actions/workflows/release.yml/badge.svg)](https://github.com/secure-artifacts/ime_audio/actions/workflows/release.yml)
[![CodeQL](https://github.com/secure-artifacts/ime_audio/actions/workflows/codeql.yml/badge.svg)](https://github.com/secure-artifacts/ime_audio/actions/workflows/codeql.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

VoiceIME is an open-source Windows 10/11 desktop utility that turns speech into text in any application. It combines a native C/Win32 interface with interchangeable cloud and fully offline speech-recognition backends, global hotkeys, multilingual text processing, and a reusable terminology library.

## Why VoiceIME?

- **Works across Windows applications:** start recording with a global hotkey and insert the result at the active text cursor.
- **Cloud or local processing:** use Gemini, Groq, Gladia, an OpenAI-compatible endpoint, or offline Sherpa-onnx models.
- **Privacy choice:** keep audio local with an offline backend, or select the cloud provider that fits your requirements.
- **Accessible and customizable:** configure voice activity detection, continuous dictation, translation, writing style, and domain terminology.
- **Auditable releases:** tagged builds run in GitHub Actions and publish build-provenance attestations for the installer.

Download the [latest release](https://github.com/secure-artifacts/ime_audio/releases/latest), read the [security policy](SECURITY.md), or see [how to contribute](CONTRIBUTING.md).

> [!IMPORTANT]
> Cloud backends require credentials from their respective providers. Never include API keys, private audio, configuration files, or sensitive logs in an issue or pull request.

---

<a id="中文说明"></a>

# Windows 语音输入法 (C + Win32 GUI)

这是一个可后台运行的 Windows 全局语音输入工具，旨在提供极速、高准确率、无缝衔接的桌面语音转文字体验。支持云端（Groq / Gemini / Gladia）与本地（Sherpa-onnx）多种识别后端。

> [!NOTE]
> **v1.0.41 更新亮点**：
> 1. 新增 **统一术语库 (Unified Terms Library)** 支持：可在 `terms.tsv` 中统一维护专有名词、谐音误听纠错、AI 词汇参考与本地模型热词。
> 2. 新增 **GUI 术语编辑器 (Terms Editor)**：在设置界面提供可视化的术语列表编辑器，支持双击编辑、导入/导出 TSV、冲突检查、批量启用/禁用及多行删除。
> 3. **优化 prompt 生成中的缓冲区容量**：重构并扩大了系统 prompt 生成中的缓冲区大小（从 2048 字节升级为 8192 字节），解决长词库规则下的溢出截断问题。
> 4. **增强对 JSON 空响应的兼容性**：改进了 ASR 接口返回空 JSON response 或无 text 字段时的解析逻辑，提升交互稳定性。

## 🌟 核心特性

- **全局热键与悬浮窗**：在任意应用输入框中，通过全局热键随时唤起录音；悬浮按钮智能跟随当前光标（或前台窗口）。
- **多端识别能力**：
  - **Gemini Native Audio (首选)**：利用 Google Gemini 多模态能力，将语音直接交由大模型处理，**一步到位**实现语音转写、润色与翻译，延迟极低。
  - **Groq API**：利用 Groq 极速 API 驱动的 Whisper 模型，实现毫秒级超快识别。
  - **Gladia API**：支持 Gladia 高效的云端语音转文字服务。
  - **Sherpa-onnx (本地)**：完全离线运行，保护隐私，无网络延迟。支持一键自动化安装本地环境与模型（包含 **SenseVoice-Small**、Paraformer、Zipformer 和 FunASR）。
- **AI 智能润色与多语种翻译**：
  - 内置 Gemini 文本处理器。无论使用哪种底层语音识别，皆可将文本结果二次送入 Gemini 进行润色、排版、去除语气词或**多语种翻译**。
  - 支持本地/自定义 OpenAI 兼容模型作为替代 AI 润色引擎。
  - 支持自定义 System Prompt 和 Gemini 3.1 思考模式配置（Thinking Level）。
  - 支持全局提示词模版文件自定义（详见进阶说明）。
- **智能 VAD 与防幻觉 (Anti-Hallucination)**：
  - 增强的底层 VAD（语音活动检测）：自动过滤键盘敲击、鼠标点击等极短孤立底噪，只在真人连续说话时触发录音。
  - 防乱出字：结合特定的 Prompt 防幻觉指令，彻底解决由于底噪引发的大模型“幻觉（如乱出字幕词、无意义乱码）”。
- **灵活的判停与续录逻辑**：
  - 支持**自动静音停录**：说话完毕停顿数秒后自动上屏，无需再次按键。
  - 支持**自动监听（连续模式）**：一次识别上屏后，自动恢复录音状态，实现不间断的语音连续输入。
- **统一术语库**：支持 `terms.tsv` 统一维护专有名词、误听纠错、AI 词汇参考与本地模型热词；旧版 `错词=正词;` 规则仍可继续使用。

## 🛠️ 编译与依赖

- **操作系统**：Windows 10/11
- **构建工具**：CMake 3.20+
- **编译器**：Visual Studio 2022 (MSVC)
- **网络条件**：使用云端服务 (Groq / Gemini / Gladia) 时需确保相应的网络连通性。

```powershell
# 1. 生成构建系统
cmake -S . -B build -G "Visual Studio 17 2022"

# 2. 编译 Release 版本
cmake --build build --config Release
```

编译输出程序：`build\Release\voice_ime.exe`

## 🚀 快速上手

1. 启动 `voice_ime.exe`。
2. 在设置界面配置您的 **热键**（默认为 `Ctrl + Alt + R`）。
3. 选择 **识别后端**：
   - 推荐选择 **Gemini Native Audio**，并在下方填写 `Gemini Key`（也可以填上特定 `Project ID` 或切换 `Gemini 模型`）。
   - 若您想通过先转文字再润色的形式组合不同平台，也可选择 **Groq** 后端 + 勾选 **启用 Gemini 润色/处理/翻译**。
4. **自定义 AI 行为**：
   - 下拉选择翻译目标语言（支持英、日、韩、法等）。
   - 在“Gemini 指令”栏填入您的个性化要求（如：“去除多余的语气词，帮我把这段话改写得更加商务正式”）。
5. 点击 **[保存设置]** 并关闭窗口（程序将隐藏到右下角托盘）。
6. **开始输入**：
   - 将光标放在任意聊天软件或文档中。
   - 按下设置的热键开始说话（界面或悬浮窗提示“录音中”）。
   - 停顿 1.5 秒（默认静音时长），程序会自动停止录音，并瞬间将处理后的文字输出在您的光标处。

## 💻 一键安装本地模型 (Sherpa)

若您希望完全离线使用：
1. 主界面 -> 识别后端选择 **Sherpa (本地)**。
2. 点击界面上的 **[安装本地模型（Sherpa）]** 按钮。
3. 程序会启动脚本自动下载 `sherpa-onnx` 运行包以及 `paraformer-zh` 中文离线大模型。
4. 下载完毕后自动填充路径。点击 **[配置自检]** 显示通过即可离线使用。

## ⚙️ 进阶参数说明

所有配置不仅可以在 GUI 修改，也可通过程序同目录下的 `voice_ime.ini` 批量修改：

- **VAD (声音检测) 阈值设置**：
  - `音量阈值 (voice_threshold)`：默认 1400。若环境嘈杂导致乱触发，可提高至 2500~4000。
  - `静音时长 (silence_timeout_ms)`：默认 1500ms。检测到低于阈值的静音超过此时间即判定断句。
  - `最短录音 (min_record_ms)`：默认 900ms。低于此时长的纯噪音将被直接丢弃。
  - `最长录音 (max_record_ms)`：默认 30000ms。到达后强制结束当前段落的录制并发送识别。
- **统一术语库**：程序目录下的 `terms.tsv` 是推荐的术语维护入口。界面提供 **[打开词库] / [重载词库] / [词库检查]**。
  - 表头格式：`enabled	wrong	correct	aliases	mode	tags	note`
  - `enabled`：`1` 启用，`0` 停用。
  - `wrong`：常见误听、错词或口述别名。
  - `correct`：最终希望输出的标准写法。
  - `aliases`：其他别名，多项可用逗号、顿号、斜杠或分号分隔。
  - `mode`：`all`、`ai`、`replace`、`hotword`，可组合使用。`all` 会同时用于 AI 参考、本地强制替换和本地识别热词。
- **旧版术语纠错格式**：界面中的 `旧词=新词;词语A=词语B;` 仍然有效，会与 `terms.tsv` 自动合并。

### 🎨 提示词模版与优化风格自定义
在程序运行目录的 `prompts/` 文件夹下，您可以直接编辑以下提示词文本文件来深度定制 AI 的行为：
- `system_prompt_openai.txt`：使用本地/自订 OpenAI 兼容引擎时的基础系统提示词。
- `prompt_openai_fewshot.txt`：OpenAI 兼容引擎的 Few-Shot (少样本) 对话示例，用于规范优化格式。
- `system_prompt_gemini.txt`：Gemini 进行二次润色与翻译时的系统提示词。
- `system_prompt_gemini_transcribe.txt`：Gemini Native Audio 语音直输时的核心系统提示词。
- `prompt_style_default.txt` / `prompt_style_business.txt` / `prompt_style_casual.txt` / `prompt_style_concise.txt`：对应界面上“智慧預設/商務正式/日常口語/簡潔扼要”四种优化风格的提示词子模板。

## 📝 日志与排错

程序会在 `.exe` 同级目录下生成 `voice_ime.log` 文件。
若遇到按快捷键无反应、识别不出字或上屏失败的情况，可打开日志文件，里面详细记录了录音设备状态、云端 API 的错误返回值（如 HTTP 400/401/403/500）以及剪贴板注入状态。

## 📦 安全发布流程 (CI/CD)

本项目已经接入了严格的安全自动化发布流程（基于 GitHub Actions 和 OIDC）。当需要发布新版本时，请按照以下步骤操作：

1. **提交代码**：将所有修改提交到 `master` 分支并推送到远程仓库。
2. **打标签触发 CI**：
   ```powershell
   git tag -a v1.0.x -m "版本说明"
   git push origin v1.0.x
   ```
   > 注意：Tag 必须以小写字母 `v` 开头（例如 `v1.0.3`），否则不会触发自动化发布。
3. **等待构建完成**：前往 GitHub 仓库的 **Actions** 页面，等待 Release 工作流运行完毕。
   - 如果构建失败：请先运行 `git push --delete origin v1.0.x` 和 `git tag -d v1.0.x` 删除错误标签，修复代码后重新打 Tag 推送。
   - 如果构建成功：工作流会自动生成并上传 `voice_ime.exe`，并利用 `attest-build-provenance` 为产物附加安全溯源签名 (Attestation)。
4. **审核申请**：构建成功后，在内部安全管理平台提交项目审核，并将生成的 Release 链接发送至管理群组进行审批。

## 已知限制
- 某些具有反作弊或管理员高权限的游戏/应用（如部分网游终端或 UAC 弹窗），可能无法响应热键或拒绝模拟键盘粘贴输入。
- 当前上屏逻辑使用模拟 Unicode 键盘输入；少数高权限或特殊输入框仍可能拒绝模拟输入。

## 🔐 安全与隐私

- 使用 Sherpa-onnx 本地后端时，语音识别可完全在本机完成。
- 云端后端会将音频或文本发送至所选择的服务商，请同时阅读对应服务商的隐私条款。
- API Key 保存在本机的 `voice_ime.ini` 中。请勿提交该文件，也不要在 Issue、日志或截图中暴露密钥和私人音频内容。
- 安装脚本会从 GitHub、Hugging Face 或 NVIDIA 下载第三方运行时和模型；建议在受控环境中检查来源后再运行。
- 安全问题请按照 [SECURITY.md](SECURITY.md) 私下报告，不要先创建包含漏洞细节的公开 Issue。

## 🤝 参与贡献

Bug 报告、功能建议和代码贡献均欢迎。提交前请阅读 [CONTRIBUTING.md](CONTRIBUTING.md)，并使用仓库提供的 Issue 模板。涉及安全的问题请遵循 [安全策略](SECURITY.md)。

## 📄 许可证

本项目采用 [MIT License](LICENSE) 开源。
