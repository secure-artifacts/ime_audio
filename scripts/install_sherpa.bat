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
if !ERRORLEVEL! equ 0 (
    if not "%CUDA_PATH%"=="" (
        set "HAS_GPU=1"
    ) else (
        echo ===================================================================
        echo NVIDIA GPU detected, but CUDA Toolkit is missing (CUDA_PATH is empty).
        echo.
        echo To run Sherpa-ONNX with full GPU acceleration, CUDA Toolkit is required.
        echo.
        set /p "INSTALL_CUDA=Do you want to automatically download and install CUDA Toolkit 12.6? (Y/N, default Y): "
        if "!INSTALL_CUDA!"=="" set "INSTALL_CUDA=Y"
        if /i "!INSTALL_CUDA!"=="Y" (
            echo Downloading CUDA 12.6.0 Network Installer (approx. 31MB)...
            curl -L -o "%SHERPA_ROOT%\cuda_installer.exe" "https://developer.download.nvidia.com/compute/cuda/12.6.0/network_installers/cuda_12.6.0_windows_network.exe"
            if !ERRORLEVEL! equ 0 (
                echo Installing CUDA Toolkit 12.6.0 silently...
                echo Please approve the UAC (Administrator) prompt if it appears.
                powershell -Command "Start-Process -FilePath '%SHERPA_ROOT%\cuda_installer.exe' -ArgumentList '-s' -Wait -Verb RunAs"
                echo CUDA Toolkit installation finished!
                del "%SHERPA_ROOT%\cuda_installer.exe" >nul 2>&1
                
                :: Refresh environment variables for the current batch session so CUDA_PATH is recognized
                for /f "tokens=2*" %%A in ('reg query "HKLM\System\CurrentControlSet\Control\Session Manager\Environment" /v CUDA_PATH 2^>nul') do set "CUDA_PATH=%%B"
                if not "!CUDA_PATH!"=="" (
                    set "HAS_GPU=1"
                    echo CUDA Environment successfully refreshed: !CUDA_PATH!
                ) else (
                    :: Fallback check
                    if exist "C:\Program Files\NVIDIA GPU Computing Toolkit\CUDA" (
                        set "HAS_GPU=1"
                    ) else (
                        echo CUDA_PATH not detected in registry. We will assume GPU is ready, please restart application after install.
                        set "HAS_GPU=1"
                    )
                )
            ) else (
                echo Failed to download CUDA Toolkit.
            )
        )
        if !HAS_GPU! equ 0 (
            echo Defaulting to stable and high performance CPU runtime.
        )
        echo ===================================================================
    )
)
if !HAS_GPU! equ 1 (
    echo NVIDIA GPU and CUDA Toolkit detected. Using CUDA runtime.
    set "RUNTIME_NAME=sherpa-onnx-%TAG%-win-x64-cuda"
) else (
    echo Using CPU runtime (stable and high performance).
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