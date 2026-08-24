param(
    [switch]$Clean,
    [string]$UvPath = ""
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot
$env:UV_CACHE_DIR = Join-Path $projectRoot ".uv-cache"

$uv = $UvPath
if (-not $uv) {
    $uv = (Get-Command uv -ErrorAction SilentlyContinue).Source
}
if (-not $uv) {
    throw "uv was not found. Install uv or pass -UvPath."
}

& $uv sync --group build

if ($Clean) {
    $buildPath = Join-Path $projectRoot "build"
    $distPath = Join-Path $projectRoot "dist"
    if (Test-Path -LiteralPath $buildPath) {
        Remove-Item -LiteralPath $buildPath -Recurse -Force
    }
    if (Test-Path -LiteralPath $distPath) {
        Remove-Item -LiteralPath $distPath -Recurse -Force
    }
}

& $uv run --group build pyinstaller --noconfirm --clean --onedir --console --name KKS-Audit-Agent .\server.py

$packageRoot = Join-Path $projectRoot "dist\KKS-Audit-Agent"
$packageConfig = Join-Path $packageRoot "config"
New-Item -ItemType Directory -Path $packageConfig -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $projectRoot "config\kks_rules.json") -Destination $packageConfig -Force
$configuredAppConfig = Join-Path $projectRoot "config\app_config.json"
$exampleAppConfig = Join-Path $projectRoot "config\app_config.example.json"
if (Test-Path -LiteralPath $configuredAppConfig) {
    Copy-Item -LiteralPath $configuredAppConfig -Destination (Join-Path $packageConfig "app_config.json") -Force
} else {
    Copy-Item -LiteralPath $exampleAppConfig -Destination (Join-Path $packageConfig "app_config.json") -Force
}
Copy-Item -LiteralPath (Join-Path $projectRoot "README.md") -Destination $packageRoot -Force
$skillPackage = Join-Path $packageRoot "kks-audit"
if (Test-Path -LiteralPath $skillPackage) {
    Remove-Item -LiteralPath $skillPackage -Recurse -Force
}
Copy-Item -LiteralPath (Join-Path $projectRoot "kks-audit") -Destination $skillPackage -Recurse -Force

$launcher = @'
@echo off
cd /d "%~dp0"
start "KKS Audit Agent" "%~dp0KKS-Audit-Agent.exe"
timeout /t 2 /nobreak >nul
start "" "http://127.0.0.1:8080/"
'@
Set-Content -LiteralPath (Join-Path $packageRoot "start-kks-audit.cmd") -Value $launcher -Encoding ascii

Write-Host "Build complete: $packageRoot\KKS-Audit-Agent.exe"
Write-Host "Double-click start-kks-audit.cmd to start the service and open the browser."
