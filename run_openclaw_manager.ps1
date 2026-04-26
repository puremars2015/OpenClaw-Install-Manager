Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# 若未以管理員身份執行，重新以管理員身份啟動
$currentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $scriptPath = $MyInvocation.MyCommand.Path
    Start-Process pwsh -ArgumentList "-NoLogo -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs
    exit
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$appPath = Join-Path $scriptRoot 'openclaw_manager.py'

$pythonLauncher = Get-Command py -ErrorAction SilentlyContinue
if ($pythonLauncher) {
    & $pythonLauncher.Source -3 $appPath
    exit $LASTEXITCODE
}

$python = Get-Command python -ErrorAction SilentlyContinue
if ($python) {
    & $python.Source $appPath
    exit $LASTEXITCODE
}

throw 'Python was not found. Install Python 3 before running run_openclaw_manager.ps1.'