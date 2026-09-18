Write-Host ""
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host " ROBO PRECOS - COLETA PDA v0.2" -ForegroundColor Cyan
Write-Host " AUDITORIA DE PRECOS | LOJA UNICA OU REDE INTEIRA" -ForegroundColor Cyan
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
. (Join-Path $Root "src\network.ps1")
. (Join-Path $Root "src\control_workbook.ps1")

$socket = $null

try {
    $config = Get-RoboPrecosConfig

    Write-Host "Modo de execucao:" -ForegroundColor Cyan
    Write-Host "  1 - Testar uma unica loja"
    Write-Host "  2 - Coletar rede e preencher planilha de controle"
    $mode = Read-Host "Escolha [2]"
    if ([string]::IsNullOrWhiteSpace($mode)) { $mode = "2" }

    if ($mode -notin @("1","2")) {
        throw "Modo invalido. Use 1 ou 2."
    }

    $store = ""
    if ($mode -eq "1") {
        $store = Read-Host "Loja para teste (ex.: 5 ou ML05)"
    }

    $startDate = Read-Host "Data inicial (dd/mm/aaaa)"
    $endDate = Read-Host "Data final (dd/mm/aaaa)"

    if ($startDate -notmatch '^\d{2}/\d{2}/\d{4}$') {
        throw "Data inicial invalida. Use dd/mm/aaaa."
    }
    if ($endDate -notmatch '^\d{2}/\d{2}/\d{4}$') {
        throw "Data final invalida. Use dd/mm/aaaa."
    }

    [void](ConvertTo-RoboPrecosDateKey $startDate)
    [void](ConvertTo-RoboPrecosDateKey $endDate)

    # Mesma regra para modo 1 e modo 2: credencial precisa existir ANTES de abrir/coletar.
    Write-RoboLog "Validando credencial PDA antes de iniciar navegador/coleta."
    $null = Get-PdaCredential -Config $config

    $socket = Start-RoboPrecosBrowser -Config $config
    Ensure-PdaAuditPage -Socket $socket -Config $config

    if ($mode -eq "1") {
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
    }
    else {
        $summary = Invoke-RoboPrecosNetworkCollection -Socket $socket -Config $config -StartDate $startDate -EndDate $endDate -RetryPerStore 3

        if ([int]$summary.ErrorCount -gt 0) {
            Write-Host ""
            Write-Host ("A coleta do PDA teve {0} centro(s) com erro. A planilha sera validada pela whitelist antes de qualquer gravacao." -f $summary.ErrorCount) -ForegroundColor Yellow
        }
        else {
            Write-Host ""
            Write-Host "Todas as lojas retornadas pelo PDA foram coletadas e validadas." -ForegroundColor Green
        }

        $controlResult = Invoke-RoboPrecosControlWorkbook -Rows @($summary.Rows) -StartDate $startDate -EndDate $endDate

        Write-Host ("Planilha de controle: " + $controlResult.WorkbookPath) -ForegroundColor Cyan
    }

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
