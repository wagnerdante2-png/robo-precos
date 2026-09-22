Write-Host ""
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host " ROBO PRECOS - PDA + POWER BI v0.4.8" -ForegroundColor Cyan
Write-Host " AUDITORIA DE PRECOS + DESCONTOS PRECO ERRADO" -ForegroundColor Cyan
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
. (Join-Path $Root "src\bi.ps1")
. (Join-Path $Root "src\discount_workbook.ps1")

$pdaSocket = $null

try {
    $config = Get-RoboPrecosConfig

    Write-Host "Modo de execucao:" -ForegroundColor Cyan
    Write-Host "  1 - Testar uma unica loja no PDA"
    Write-Host "  2 - Fluxo completo: PDA + planilha + Power BI + descontos"
    Write-Host "  3 - Testar somente a leitura do Power BI (NAO grava planilha)"
    $mode = Read-Host "Escolha [2]"
    if ([string]::IsNullOrWhiteSpace($mode)) { $mode = "2" }

    if ($mode -notin @("1","2","3")) {
        throw "Modo invalido. Use 1, 2 ou 3."
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

    if ($mode -eq "3") {
        Write-Host ""
        Write-Host "MODO DE TESTE POWER BI - nenhuma celula da planilha sera alterada." -ForegroundColor Yellow

        $discountResult = Invoke-RoboPrecosBiDiscountCollection -Config $config -StartDate $startDate -EndDate $endDate

        Write-Host ""
        Write-Host "AMOSTRA COLETADA DO POWER BI" -ForegroundColor Cyan
        @($discountResult.Records | Sort-Object Empresa | Select-Object Loja,Empresa,QuantidadeCupons,Desconto,Fonte) | Format-Table -AutoSize
        Write-Host ("Total de lojas com valores validos: {0}" -f $discountResult.RecordCount) -ForegroundColor Green

        if ($discountResult.Integrity) {
            Write-Host ("Soma de cupons coletada        : {0}" -f $discountResult.Integrity.SumQuantity) -ForegroundColor Cyan
            Write-Host ("Soma de desconto coletada      : R$ {0}" -f ([double]$discountResult.Integrity.SumDiscount).ToString("N2",[Globalization.CultureInfo]::GetCultureInfo("pt-BR"))) -ForegroundColor Cyan

            if ([string]$discountResult.Mode -eq "CURRENT") {
                Write-Host ("Total do visual Power BI       : {0} cupons / R$ {1}" -f $discountResult.Integrity.VisualQuantity, ([double]$discountResult.Integrity.VisualDiscount).ToString("N2",[Globalization.CultureInfo]::GetCultureInfo("pt-BR"))) -ForegroundColor Green
                Write-Host "Reconciliacao                   : OK" -ForegroundColor Green

                if (@($discountResult.Integrity.MissingCompanies).Count -gt 0) {
                    Write-Host ("IDs de empresa nao retornados   : " + (@($discountResult.Integrity.MissingCompanies) -join ", ")) -ForegroundColor Yellow
                }
            }
        }
        Write-Host ""
        Read-Host "Pressione ENTER para fechar"
        return
    }

    # PDA: mesma logica validada anteriormente.
    Write-RoboLog "Validando credencial PDA antes de iniciar navegador/coleta."
    $null = Get-PdaCredential -Config $config

    $pdaSocket = Start-RoboPrecosBrowser -Config $config
    Ensure-PdaAuditPage -Socket $pdaSocket -Config $config

    if ($mode -eq "1") {
        Write-Host ""
        Write-Host ("Consultando {0} de {1} a {2}..." -f $store, $startDate, $endDate) -ForegroundColor Yellow

        $result = Invoke-PdaAuditQuery -Socket $pdaSocket -Config $config -Store $store -StartDate $startDate -EndDate $endDate

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
        # 1) COLETA PDA
        $summary = Invoke-RoboPrecosNetworkCollection -Socket $pdaSocket -Config $config -StartDate $startDate -EndDate $endDate -RetryPerStore 3

        if ([int]$summary.ErrorCount -gt 0) {
            Write-Host ""
            Write-Host ("A coleta do PDA teve {0} centro(s) com erro. A planilha sera validada pela whitelist antes de qualquer gravacao." -f $summary.ErrorCount) -ForegroundColor Yellow
        }
        else {
            Write-Host ""
            Write-Host "Todas as lojas retornadas pelo PDA foram coletadas e validadas." -ForegroundColor Green
        }

        # 2) GRAVACAO AUDITORIA
        $controlResult = Invoke-RoboPrecosControlWorkbook -Rows @($summary.Rows) -StartDate $startDate -EndDate $endDate
        Write-Host ("Planilha de controle: " + $controlResult.WorkbookPath) -ForegroundColor Cyan

        # O PDA ja terminou. Fecha apenas o canal CDP antes de iniciar o Chrome dedicado do BI.
        if ($pdaSocket) {
            Close-CdpPage -Socket $pdaSocket
            $pdaSocket = $null
        }

        # 3) POWER BI -> PRECO ERRADO
        try {
            $discountResult = Invoke-RoboPrecosBiDiscountCollection -Config $config -StartDate $startDate -EndDate $endDate

            # 4) GRAVACAO SELETIVA DOS DESCONTOS
            $discountWrite = Invoke-RoboPrecosDiscountWorkbook -DiscountResult $discountResult

            Write-Host ""
            Write-Host "FLUXO COMPLETO CONCLUIDO" -ForegroundColor Green
            Write-Host ("Auditoria - lojas     : {0}" -f $controlResult.StoreCount)
            Write-Host ("Descontos - coletadas : {0}" -f $discountResult.RecordCount)
            Write-Host ("Descontos - gravadas  : {0}" -f $discountWrite.WrittenStoreCount)
            Write-Host ("Arquivo final         : {0}" -f $discountWrite.WorkbookPath)
        }
        catch {
            Write-RoboLog ("Auditoria ja foi gravada, mas o modulo de descontos falhou: " + $_.Exception.Message) "ERRO"
            Write-Host ""
            Write-Host "A AUDITORIA DO PDA FOI PRESERVADA." -ForegroundColor Yellow
            Write-Host "O Power BI/descontos falhou antes de uma conclusao valida. Nenhum dado ausente foi convertido em zero." -ForegroundColor Yellow
            throw
        }
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
    if ($pdaSocket) {
        Close-CdpPage -Socket $pdaSocket
    }
}
