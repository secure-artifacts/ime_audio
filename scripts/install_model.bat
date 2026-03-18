@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

echo ===================================================
echo   Voice IME - Model Installer
echo ===================================================
echo.

set "MODEL_ID=%~1"
if "%MODEL_ID%"=="" (
    echo [error] No model specified.
    pause
    exit /b 1
)

set "ROOT_DIR=%~dp0.."
set "SHERPA_ROOT=%ROOT_DIR%\third_party\sherpa"
if not exist "%SHERPA_ROOT%" mkdir "%SHERPA_ROOT%"

set "TAG=v1.12.29"
set "HAS_GPU=0"
nvidia-smi >nul 2>&1
if !ERRORLEVEL! equ 0 set "HAS_GPU=1"
if !HAS_GPU! equ 1 (
    echo NVIDIA GPU detected. Using CUDA runtime.
    set "RUNTIME_NAME=sherpa-onnx-%TAG%-win-x64-cuda"
) else (
    echo No NVIDIA GPU detected. Using CPU runtime.
    set "RUNTIME_NAME=sherpa-onnx-%TAG%-win-x64-static-MT-Release-no-tts"
)
set "ARCHIVE_NAME=%RUNTIME_NAME%.tar.bz2"
set "ARCHIVE_PATH=%SHERPA_ROOT%\%ARCHIVE_NAME%"
set "RUNTIME_DIR=%SHERPA_ROOT%\%RUNTIME_NAME%"
if not exist "%RUNTIME_DIR%" (
    echo ===================================================
    echo   Downloading Sherpa-ONNX Runtime...
    echo ===================================================
    if not exist "%ARCHIVE_PATH%" (
        set "URL=https://github.com/k2-fsa/sherpa-onnx/releases/download/%TAG%/%ARCHIVE_NAME%"
        echo Downloading !URL!
        curl -L -o "%ARCHIVE_PATH%" "!URL!"
        if !ERRORLEVEL! neq 0 (
            echo [error] Failed to download runtime.
            pause
            exit /b 1
        )
    )
    echo Extracting %ARCHIVE_NAME% ...
    tar -xjf "%ARCHIVE_PATH%" -C "%SHERPA_ROOT%"
    if !ERRORLEVEL! neq 0 (
        echo [error] Failed to extract runtime.
        pause
        exit /b 1
    )
    echo Runtime extracted successfully.
    echo.
)


:: Mirror URL to avoid huggingface blockage in China
set "HF_MIRROR=https://huggingface.co"

if "%MODEL_ID%"=="paraformer" (
    set "MODEL_DIR=%SHERPA_ROOT%\models\paraformer-zh"
    if not exist "!MODEL_DIR!" mkdir "!MODEL_DIR!"
    echo Downloading Paraformer model...
    if not exist "!MODEL_DIR!\model.int8.onnx" curl -L -o "!MODEL_DIR!\model.int8.onnx" "%HF_MIRROR%/csukuangfj/sherpa-onnx-paraformer-zh-2023-09-14/resolve/main/model.int8.onnx"
    if not exist "!MODEL_DIR!\tokens.txt" curl -L -o "!MODEL_DIR!\tokens.txt" "%HF_MIRROR%/csukuangfj/sherpa-onnx-paraformer-zh-2023-09-14/resolve/main/tokens.txt"
)

if "%MODEL_ID%"=="zipformer" (
    set "MODEL_DIR=%SHERPA_ROOT%\models\zipformer-zh"
    if not exist "!MODEL_DIR!" mkdir "!MODEL_DIR!"
    echo Downloading Zipformer-zh-xlarge model...
    if not exist "!MODEL_DIR!\bpe.model" curl -L -o "!MODEL_DIR!\bpe.model" "%HF_MIRROR%/csukuangfj/sherpa-onnx-streaming-zipformer-zh-xlarge-int8-2025-06-30/resolve/main/bpe.model"
    if not exist "!MODEL_DIR!\decoder.onnx" curl -L -o "!MODEL_DIR!\decoder.onnx" "%HF_MIRROR%/csukuangfj/sherpa-onnx-streaming-zipformer-zh-xlarge-int8-2025-06-30/resolve/main/decoder.onnx"
    if not exist "!MODEL_DIR!\encoder.int8.onnx" curl -L -o "!MODEL_DIR!\encoder.int8.onnx" "%HF_MIRROR%/csukuangfj/sherpa-onnx-streaming-zipformer-zh-xlarge-int8-2025-06-30/resolve/main/encoder.int8.onnx"
    if not exist "!MODEL_DIR!\joiner.int8.onnx" curl -L -o "!MODEL_DIR!\joiner.int8.onnx" "%HF_MIRROR%/csukuangfj/sherpa-onnx-streaming-zipformer-zh-xlarge-int8-2025-06-30/resolve/main/joiner.int8.onnx"
    if not exist "!MODEL_DIR!\tokens.txt" curl -L -o "!MODEL_DIR!\tokens.txt" "%HF_MIRROR%/csukuangfj/sherpa-onnx-streaming-zipformer-zh-xlarge-int8-2025-06-30/resolve/main/tokens.txt"
)

if "%MODEL_ID%"=="funasr" (
    set "MODEL_DIR=%SHERPA_ROOT%\models\funasr"
    if not exist "!MODEL_DIR!" mkdir "!MODEL_DIR!"
    if not exist "!MODEL_DIR!\Qwen3-0.6B" mkdir "!MODEL_DIR!\Qwen3-0.6B"
    echo Downloading FunASR-nano model...
    if not exist "!MODEL_DIR!\embedding.int8.onnx" curl -L -o "!MODEL_DIR!\embedding.int8.onnx" "%HF_MIRROR%/csukuangfj/sherpa-onnx-funasr-nano-int8-2025-12-30/resolve/main/embedding.int8.onnx"
    if not exist "!MODEL_DIR!\encoder_adaptor.int8.onnx" curl -L -o "!MODEL_DIR!\encoder_adaptor.int8.onnx" "%HF_MIRROR%/csukuangfj/sherpa-onnx-funasr-nano-int8-2025-12-30/resolve/main/encoder_adaptor.int8.onnx"
    if not exist "!MODEL_DIR!\llm.int8.onnx" curl -L -o "!MODEL_DIR!\llm.int8.onnx" "%HF_MIRROR%/csukuangfj/sherpa-onnx-funasr-nano-int8-2025-12-30/resolve/main/llm.int8.onnx"
    if not exist "!MODEL_DIR!\tokens.txt" curl -L -o "!MODEL_DIR!\tokens.txt" "%HF_MIRROR%/csukuangfj/sherpa-onnx-funasr-nano-int8-2025-12-30/resolve/main/tokens.txt"
    if not exist "!MODEL_DIR!\Qwen3-0.6B\merges.txt" curl -L -o "!MODEL_DIR!\Qwen3-0.6B\merges.txt" "%HF_MIRROR%/csukuangfj/sherpa-onnx-funasr-nano-int8-2025-12-30/resolve/main/Qwen3-0.6B/merges.txt"
    if not exist "!MODEL_DIR!\Qwen3-0.6B\tokenizer.json" curl -L -o "!MODEL_DIR!\Qwen3-0.6B\tokenizer.json" "%HF_MIRROR%/csukuangfj/sherpa-onnx-funasr-nano-int8-2025-12-30/resolve/main/Qwen3-0.6B/tokenizer.json"
    if not exist "!MODEL_DIR!\Qwen3-0.6B\vocab.json" curl -L -o "!MODEL_DIR!\Qwen3-0.6B\vocab.json" "%HF_MIRROR%/csukuangfj/sherpa-onnx-funasr-nano-int8-2025-12-30/resolve/main/Qwen3-0.6B/vocab.json"
)

echo.
echo ===================================================
echo   Installation Complete!
echo ===================================================
timeout /t 3 >nul
exit /b 0
