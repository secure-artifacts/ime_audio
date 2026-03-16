# Windows 语音输入法 (C + Win32 GUI)

这是一个可后台运行的 Windows 全局语音输入工具，支持云端与本地两种识别后端。

## 当前能力

- 全局热键控制录音（任意输入框可用）
- 设置窗口输入 API Key 与热键
- 设置窗口可直接配置 backend、continuous mode、阈值和 Sherpa 参数
- 启动后自动执行配置自检，界面可手动点击 Self-check
- 一键安装 Sherpa（运行脚本或点击 Install Sherpa）
- 关闭窗口后最小化到托盘继续运行
- 悬浮按钮跟随当前前台输入窗口
- 自动静音停录（静音一段时间自动结束并识别）
- 识别结果自动粘贴（剪贴板 + Ctrl+V）
- 识别日志写入 voice_ime.log
- 后端可切换：groq 或 sherpa

## 依赖

- Windows 10/11
- CMake 3.20+
- Visual Studio 2022 (MSVC)

使用 groq 后端时还需要网络访问 api.groq.com。

## 构建

```powershell
cmake -S . -B build -G "Visual Studio 17 2022"
cmake --build build --config Release
```

输出程序：build\Release\voice_ime.exe

## 快速使用

1. 运行程序后先设置热键，点击 Apply & Save。
2. 使用 groq 时，在界面填写 API Key，并将 backend 选为 groq。
3. 使用 sherpa 时，将 backend 选为 sherpa，并填写 Sherpa EXE / Sherpa Args。
4. 可点击 Self-check 验证配置是否完整。
5. 在任意应用输入框中按热键开始录音。
6. 再按一次热键手动结束，或等待自动静音停录。
7. 转写完成后自动粘贴到目标输入框。
8. 点击主窗口关闭按钮会隐藏到托盘，不会退出。

## 一键安装 Sherpa

方式 1（推荐）：在应用界面点击 Install Sherpa。

方式 2（命令行）：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install_sherpa.ps1 -ConfigureIni
```

该脚本会：

- 下载并解压 sherpa-onnx Windows 运行包（x64 static no-tts）
- 下载中文 paraformer 模型（model.int8.onnx + tokens.txt）
- 自动更新 voice_ime.ini 的 backend/sherpa_exe/sherpa_args

> 初次下载体积较大，视网络情况需要几分钟。

## 配置文件

程序会在 exe 同目录读写 voice_ime.ini。

说明：大部分配置可以在 GUI 中直接修改；INI 仍可用于批量部署或手工微调。

主要字段如下（settings 段）：

- api_key: groq 的 API Key
- hotkey_key: 主键，例如 R、F6、Enter
- hotkey_mods: 修饰键位掩码
- backend: groq 或 sherpa
- sherpa_exe: sherpa 可执行文件路径
- sherpa_args: sherpa 模型参数（不需要写 wav 路径）
- continuous_mode: 1 表示一次识别结束后自动继续录音，0 表示关闭
- sample_rate: 采样率，默认 16000
- voice_threshold: 判定有声的峰值阈值
- silence_timeout_ms: 静音超过该时长自动停录
- min_record_ms: 最短录音时长，未达到不会触发静音停录
- max_record_ms: 最长录音时长，达到后强制停录

hotkey_mods 位掩码：Alt=1, Ctrl=2, Shift=4, Win=8。
例如 Ctrl+Alt 为 3。

## Sherpa 本地后端示例

把 backend 设为 sherpa，并填写 sherpa_exe 与 sherpa_args。

示例：

```ini
[settings]
backend=sherpa
sherpa_exe=C:\\models\\sherpa-onnx\\bin\\sherpa-onnx-offline.exe
sherpa_args=--tokens=C:\\models\\sherpa-onnx\\tokens.txt --encoder=C:\\models\\sherpa-onnx\\encoder.onnx --decoder=C:\\models\\sherpa-onnx\\decoder.onnx --joiner=C:\\models\\sherpa-onnx\\joiner.onnx
```

说明：

- 程序会自动把录音 wav 路径追加到命令末尾。
- 如果路径包含空格，请在 sherpa_args 中自行加引号。
- 如果 sherpa_exe 为空，默认尝试 sherpa-onnx-offline.exe（需在 PATH 中可找到）。

## 自动静音停录逻辑

录音中满足以下任一条件会自动结束并开始识别：

- 已录音时长 >= min_record_ms 且静音时长 >= silence_timeout_ms
- 已录音时长 >= max_record_ms

## 日志

日志文件：voice_ime.log（与 exe 同目录）。

日志包含：

- 启动与配置加载
- 录音开始/结束（手动或自动）
- 转写后端与结果状态
- 粘贴成功/失败
- 失败原因（包含 groq/sherpa 返回文本）

## 实现说明

1. waveIn 采集 PCM 数据并写入临时 WAV。
2. 按 backend 分支调用：
   - groq: 走 HTTP 上传 WAV 转写。
   - sherpa: 启动本地 CLI 执行离线转写。
3. 将识别文本写入剪贴板并模拟 Ctrl+V。
4. 定时器跟踪前台焦点并更新悬浮按钮位置。

## 已知限制

- 某些应用限制前台切换或模拟输入，可能影响自动粘贴。
- 当前会覆盖剪贴板内容（后续可加粘贴后恢复）。
