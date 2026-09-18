Write-Host ""
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host " ROBO HORAS - PORTABLE v0.9" -ForegroundColor Cyan
Write-Host " TESTE MULTIDESTINATARIO | WHATSAPP WEB" -ForegroundColor Cyan
Write-Host " SEM INSTALACAO | SEM PYTHON | SEM ACTIONS" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""
$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root
. (Join-Path $Root "src\bootstrap.ps1")
. (Join-Path $Root "src\excel.ps1")
. (Join-Path $Root "src\normalize_input.ps1")
. (Join-Path $Root "src\multi_send.ps1")
. (Join-Path $Root "src\main.ps1")
