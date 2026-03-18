@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

echo ===================================================
echo   Voice IME - Sherpa Onnx Offline Model Installer
echo ===================================================
echo.

set "TAG=v1.12.29"
set "ROOT_DIR=%~dp0.."
set "SHERPA_ROOT=%ROOT_DIR%\third_party\sherpa"
if not exist "%SHERPA_ROOT%" mkdir "%SHERPA_ROOT%"
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
set "MODEL_DIR=%SHERPA_ROOT%\models\paraformer-zh"

if not exist "%SHERPA_ROOT%" mkdir "%SHERPA_ROOT%"
if not exist "%MODEL_DIR%" mkdir "%MODEL_DIR%"

:: 1. Download and extract runtime
if not exist "%RUNTIME_DIR%" (
    if not exist "%ARCHIVE_PATH%" (
        set "URL=https://github.com/k2-fsa/sherpa-onnx/releases/download/%TAG%/%ARCHIVE_NAME%"
        echo [download] !URL!
        curl -L -o "%ARCHIVE_PATH%" "!URL!"
        if %ERRORLEVEL% neq 0 (
            echo [error] Failed to download runtime.
            pause
            exit /b 1
        )
    ) else (
        echo [skip] runtime archive exists: %ARCHIVE_PATH%
    )

    echo [extract] %ARCHIVE_NAME%
    tar -xjf "%ARCHIVE_PATH%" -C "%SHERPA_ROOT%"
    if %ERRORLEVEL% neq 0 (
        echo [error] Failed to extract runtime.
        pause
        exit /b 1
    )
) else (
    echo [skip] runtime dir exists: %RUNTIME_DIR%
)

:: 2. Download models
set "MODEL_ONNX=%MODEL_DIR%\model.int8.onnx"
if not exist "%MODEL_ONNX%" (
    echo [download] https://huggingface.co/csukuangfj/sherpa-onnx-paraformer-zh-2023-09-14/resolve/main/model.int8.onnx
    curl -L -o "%MODEL_ONNX%" "https://huggingface.co/csukuangfj/sherpa-onnx-paraformer-zh-2023-09-14/resolve/main/model.int8.onnx"
) else (
    echo [skip] model file exists: %MODEL_ONNX%
)

set "MODEL_TOKENS=%MODEL_DIR%\tokens.txt"
if not exist "%MODEL_TOKENS%" (
    echo [download] https://huggingface.co/csukuangfj/sherpa-onnx-paraformer-zh-2023-09-14/resolve/main/tokens.txt
    curl -L -o "%MODEL_TOKENS%" "https://huggingface.co/csukuangfj/sherpa-onnx-paraformer-zh-2023-09-14/resolve/main/tokens.txt"
) else (
    echo [skip] tokens file exists: %MODEL_TOKENS%
)

set "EXE_PATH=%RUNTIME_DIR%\bin\sherpa-onnx-offline.exe"
if not exist "%EXE_PATH%" (
    echo [error] Sherpa executable not found: %EXE_PATH%
    pause
    exit /b 1
)

echo.
echo ===================================================
echo   Installation Complete!
echo ===================================================
echo Please go back to the Voice IME settings and click 
echo "配置自检" (Self-Check) to verify.
echo.
pause
exit /b 0