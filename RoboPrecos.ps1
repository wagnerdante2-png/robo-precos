Write-Host ""
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host " ROBO PRECOS - PROTOTIPO PDA v0.1" -ForegroundColor Cyan
Write-Host " LEITURA DIRETA DOS TOTALIZADORES DE AUDITORIA" -ForegroundColor Cyan
Write-Host " SEM INSTALACAO | SEM SELENIUM | SEM ACTIONS" -ForegroundColor Cyan
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host ""

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

. (Join-Path $Root "src\bootstrap.ps1")
. (Join-Path $Root "src\cdp.ps1")
. (Join-Path $Root "src\pda.ps1")

$socket = $null

try {
    $config = Get-RoboPrecosConfig

    $store = [string]$config.test.store
    $startDate = [string]$config.test.startDate
    $endDate = [string]$config.test.endDate

    if ([string]::IsNullOrWhiteSpace($store)) {
        $store = Read-Host "Loja para teste (ex.: 5 ou ML05)"
    }
    if ([string]::IsNullOrWhiteSpace($startDate)) {
        $startDate = Read-Host "Data inicial (dd/mm/aaaa)"
    }
    if ([string]::IsNullOrWhiteSpace($endDate)) {
        $endDate = Read-Host "Data final (dd/mm/aaaa)"
    }

    if ($startDate -notmatch '^\d{2}/\d{2}/\d{4}$') {
        throw "Data inicial invalida. Use dd/mm/aaaa."
    }
    if ($endDate -notmatch '^\d{2}/\d{2}/\d{4}$') {
        throw "Data final invalida. Use dd/mm/aaaa."
    }

    $socket = Start-RoboPrecosBrowser -Config $config
    Ensure-PdaAuditPage -Socket $socket -Config $config

    Write-Host ""
    Write-Host ("Consultando {0} de {1} a {2}..." -f $store, $startDate, $endDate) -ForegroundColor Yellow

    $result = Invoke-PdaAuditQuery -Socket $socket -Config $config -Store $store -StartDate $startDate -EndDate $endDate

    Write-Host ""
    Write-Host "RESULTADO VALIDADO" -ForegroundColor Green
    Write-Host ("Loja          : {0}" -f $result.Loja)
    Write-Host ("OK            : {0}" -f $result.Ok)
    Write-Host ("Divergente    : {0}" -f $result.Divergente)
    Write-Host ("Sem etiqueta  : {0}" -f $result.SemEtiqueta)
    Write-Host ("Total auditado: {0}" -f $result.Total)
    Write-Host ""

    $testOutput = Join-Path $OutputPath ("teste_pda_{0}_{1}.csv" -f $result.Loja, (Get-Date -Format "yyyyMMdd_HHmmss"))
    @($result) | Export-Csv -LiteralPath $testOutput -NoTypeInformation -Encoding UTF8
    Write-RoboLog ("Teste concluido e salvo em " + $testOutput)

    Write-Host ("Arquivo de teste: {0}" -f $testOutput) -ForegroundColor Cyan
    Write-Host ""
    Read-Host "Pressione ENTER para fechar"
}
catch {
    Write-RoboLog $_.Exception.Message "ERRO"
    Write-Host ""
    Write-Host ("[ERRO] " + $_.Exception.Message) -ForegroundColor Red
    Write-Host ""
    Read-Host "Pressione ENTER para fechar"
    exit 1
}
finally {
    if ($socket) {
        Close-CdpPage -Socket $socket
    }
}
