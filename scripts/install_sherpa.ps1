param(
    [string]$Tag = "v1.12.29",
    [switch]$ConfigureIni,
    [string]$IniPath = ""
)

$ErrorActionPreference = "Stop"

function Set-IniSetting {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Section,
        [Parameter(Mandatory = $true)][string]$Key,
        [Parameter(Mandatory = $true)][string]$Value
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    if (Test-Path $Path) {
        (Get-Content $Path) | ForEach-Object { [void]$lines.Add($_) }
    }

    if ($lines.Count -gt 0 -and $lines[0].Length -gt 0 -and [int][char]$lines[0][0] -eq 0xFEFF) {
        $lines[0] = $lines[0].Substring(1)
    }

    if ($lines.Count -eq 0) {
        [void]$lines.Add("[$Section]")
    }

    $inSection = $false
    $sectionFound = $false
    $keySet = $false

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i].TrimStart([char]0xFEFF)
        if ($line -match '^\s*\[(.+?)\]\s*$') {
            if ($inSection -and -not $keySet) {
                $lines.Insert($i, "$Key=$Value")
                $keySet = $true
                break
            }

            $inSection = ($matches[1].ToLowerInvariant() -eq $Section.ToLowerInvariant())
            if ($inSection) {
                $sectionFound = $true
            }
            continue
        }

        if ($inSection -and $line -match "^\s*$([Regex]::Escape($Key))\s*=") {
            $lines[$i] = "$Key=$Value"
            $keySet = $true
            break
        }
    }

    if (-not $sectionFound) {
        if ($lines.Count -gt 0 -and $lines[$lines.Count - 1].Trim() -ne "") {
            [void]$lines.Add("")
        }
        [void]$lines.Add("[$Section]")
        [void]$lines.Add("$Key=$Value")
        $keySet = $true
    }

    if (-not $keySet) {
        [void]$lines.Add("$Key=$Value")
    }

    Set-Content -Path $Path -Value $lines -Encoding Unicode
}

function Normalize-SettingsSection {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path $Path)) {
        return
    }

    $src = [System.Collections.Generic.List[string]]::new()
    (Get-Content $Path) | ForEach-Object { [void]$src.Add($_) }

    if ($src.Count -gt 0 -and $src[0].Length -gt 0 -and [int][char]$src[0][0] -eq 0xFEFF) {
        $src[0] = $src[0].Substring(1)
    }

    $dst = [System.Collections.Generic.List[string]]::new()
    $seenSettings = $false
    $skipSection = $false

    foreach ($rawLine in $src) {
        $line = $rawLine.TrimStart([char]0xFEFF)

        if ($line -match '^\s*\[(.+?)\]\s*$') {
            $sectionName = $matches[1].ToLowerInvariant()
            if ($sectionName -eq 'settings') {
                if ($seenSettings) {
                    $skipSection = $true
                    continue
                }

                $seenSettings = $true
                $skipSection = $false
                [void]$dst.Add('[settings]')
                continue
            }

            $skipSection = $false
            [void]$dst.Add($line)
            continue
        }

        if (-not $skipSection) {
            [void]$dst.Add($line)
        }
    }

    Set-Content -Path $Path -Value $dst -Encoding Unicode
}

$root = Split-Path -Parent $PSScriptRoot
$sherpaRoot = Join-Path $root "third_party\sherpa"
New-Item -ItemType Directory -Path $sherpaRoot -Force | Out-Null

$runtimeName = "sherpa-onnx-$Tag-win-x64-static-MT-Release-no-tts"
$archiveName = "$runtimeName.tar.bz2"
$archivePath = Join-Path $sherpaRoot $archiveName
$runtimeDir = Join-Path $sherpaRoot $runtimeName

if (-not (Test-Path $runtimeDir)) {
    if (-not (Test-Path $archivePath)) {
        $runtimeUrl = "https://github.com/k2-fsa/sherpa-onnx/releases/download/$Tag/$archiveName"
        Write-Host "[download] $runtimeUrl"
        Invoke-WebRequest $runtimeUrl -OutFile $archivePath
    } else {
        Write-Host "[skip] runtime archive exists: $archivePath"
    }

    Write-Host "[extract] $archiveName"
    tar -xjf $archivePath -C $sherpaRoot
} else {
    Write-Host "[skip] runtime dir exists: $runtimeDir"
}

$modelDir = Join-Path $sherpaRoot "models\paraformer-zh"
New-Item -ItemType Directory -Path $modelDir -Force | Out-Null

$downloads = @(
    @{
        Url = "https://huggingface.co/csukuangfj/sherpa-onnx-paraformer-zh-2023-09-14/resolve/main/model.int8.onnx"
        Out = (Join-Path $modelDir "model.int8.onnx")
    },
    @{
        Url = "https://huggingface.co/csukuangfj/sherpa-onnx-paraformer-zh-2023-09-14/resolve/main/tokens.txt"
        Out = (Join-Path $modelDir "tokens.txt")
    }
)

foreach ($d in $downloads) {
    if (-not (Test-Path $d.Out)) {
        Write-Host "[download] $($d.Url)"
        Invoke-WebRequest $d.Url -OutFile $d.Out
    } else {
        Write-Host "[skip] model file exists: $($d.Out)"
    }
}

$exePath = Join-Path $runtimeDir "bin\sherpa-onnx-offline.exe"
if (-not (Test-Path $exePath)) {
    throw "Sherpa executable not found: $exePath"
}

$tokensPath = Join-Path $modelDir "tokens.txt"
$paraformerPath = Join-Path $modelDir "model.int8.onnx"
$sherpaArgs = '--tokens="{0}" --paraformer="{1}" --num-threads=2 --decoding-method=greedy_search' -f $tokensPath, $paraformerPath

if ($ConfigureIni) {
    if ([string]::IsNullOrWhiteSpace($IniPath)) {
        $iniPath = Join-Path $root "voice_ime.ini"
    } else {
        $iniPath = [System.IO.Path]::GetFullPath($IniPath)
        $iniDir = Split-Path -Parent $iniPath
        if ($iniDir -and -not (Test-Path $iniDir)) {
            New-Item -ItemType Directory -Path $iniDir -Force | Out-Null
        }
    }

    Normalize-SettingsSection -Path $iniPath
    Set-IniSetting -Path $iniPath -Section "settings" -Key "backend" -Value "sherpa"
    Set-IniSetting -Path $iniPath -Section "settings" -Key "sherpa_exe" -Value $exePath
    Set-IniSetting -Path $iniPath -Section "settings" -Key "sherpa_args" -Value $sherpaArgs
    Set-IniSetting -Path $iniPath -Section "settings" -Key "continuous_mode" -Value "0"
    Write-Host "[done] updated ini: $iniPath"
}

Write-Host "[done] sherpa_exe=$exePath"
Write-Host "[done] sherpa_args=$sherpaArgs"
