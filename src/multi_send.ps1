$script:RoboSingleRecipientInvokeExcel = ${function:Invoke-RoboExcel}

function Invoke-RoboExcel {
    param(
        [string]$ExcelPath,
        $Config,
        [hashtable]$StoreMap,
        [string]$ChromePath,
        [string]$ProfilePath
    )

    $testMode = $false
    if ($Config.whatsapp.PSObject.Properties.Name -contains "testMode") {
        $testMode = [bool]$Config.whatsapp.testMode
    }

    $selectedRecipients = @()
    if ($testMode -and ($Config.whatsapp.PSObject.Properties.Name -contains "selectedTestRecipients")) {
        $selectedRecipients = @($Config.whatsapp.selectedTestRecipients)
    }

    if (-not $testMode -or $selectedRecipients.Count -le 1) {
        & $script:RoboSingleRecipientInvokeExcel $ExcelPath $Config $StoreMap $ChromePath $ProfilePath
        return
    }

    $continueOnError = $true
    $retryCount = 1
    $retryDelaySeconds = 3
    $delayBetweenMessages = 4

    if ($Config.whatsapp.PSObject.Properties.Name -contains "continueOnSendError") {
        $continueOnError = [bool]$Config.whatsapp.continueOnSendError
    }
    if ($Config.whatsapp.PSObject.Properties.Name -contains "sendRetryCount") {
        $retryCount = [int]$Config.whatsapp.sendRetryCount
    }
    if ($Config.whatsapp.PSObject.Properties.Name -contains "sendRetryDelaySeconds") {
        $retryDelaySeconds = [int]$Config.whatsapp.sendRetryDelaySeconds
    }
    if ($Config.whatsapp.PSObject.Properties.Name -contains "delayBetweenMessagesSeconds") {
        $delayBetweenMessages = [int]$Config.whatsapp.delayBetweenMessagesSeconds
    }

    $originalPhone = ""
    if ($Config.whatsapp.PSObject.Properties.Name -contains "testPhone") {
        $originalPhone = [string]$Config.whatsapp.testPhone
    }

    $results = New-Object System.Collections.ArrayList
    $total = $selectedRecipients.Count
    $position = 0

    Write-RoboLog ("Inicio do lote de teste com {0} destinatarios." -f $total)

    foreach ($recipient in $selectedRecipients) {
        $position++
        $name = [string]$recipient.name
        if ([string]::IsNullOrWhiteSpace($name)) {
            $name = "Teste $position"
        }
        $phone = ConvertTo-RoboPhone ([string]$recipient.phone)

        if ($phone.Length -lt 10) {
            $message = "Telefone invalido para $name."
            Write-RoboLog $message "ERRO"
            [void]$results.Add([PSCustomObject]@{
                ordem = $position
                nome = $name
                telefone = $phone
                status = "FALHOU"
                tentativas = 0
                erro = $message
                horario = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            })
            if (-not $continueOnError) { throw $message }
            continue
        }

        $Config.whatsapp.testPhone = $phone
        $success = $false
        $lastError = ""
        $attempts = 0

        for ($attempt = 1; $attempt -le ($retryCount + 1); $attempt++) {
            $attempts = $attempt
            try {
                Write-RoboLog ("[{0}/{1}] Enviando para {2} | {3} | tentativa {4}" -f $position, $total, $name, $phone, $attempt)
                & $script:RoboSingleRecipientInvokeExcel $ExcelPath $Config $StoreMap $ChromePath $ProfilePath
                $success = $true
                Write-RoboLog ("[{0}/{1}] {2} | {3} | ENVIADO" -f $position, $total, $name, $phone)
                break
            }
            catch {
                $lastError = $_.Exception.Message
                Write-RoboLog ("[{0}/{1}] {2} | {3} | FALHOU tentativa {4}: {5}" -f $position, $total, $name, $phone, $attempt, $lastError) "ERRO"
                if ($attempt -le $retryCount) {
                    Start-Sleep -Seconds $retryDelaySeconds
                }
            }
        }

        $status = if ($success) { "ENVIADO" } else { "FALHOU" }
        [void]$results.Add([PSCustomObject]@{
            ordem = $position
            nome = $name
            telefone = $phone
            status = $status
            tentativas = $attempts
            erro = $lastError
            horario = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        })

        if (-not $success -and -not $continueOnError) {
            throw "Falha no envio para $name ($phone): $lastError"
        }

        if ($position -lt $total -and $delayBetweenMessages -gt 0) {
            Start-Sleep -Seconds $delayBetweenMessages
        }
    }

    $Config.whatsapp.testPhone = $originalPhone

    $summaryPath = Join-Path $LogDirectory ("dispatch_summary_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
    $results | Export-Csv -LiteralPath $summaryPath -NoTypeInformation -Encoding UTF8

    $sent = @($results | Where-Object { $_.status -eq "ENVIADO" }).Count
    $failed = @($results | Where-Object { $_.status -eq "FALHOU" }).Count

    Write-Host ""
    Write-Host "RESUMO DO LOTE DE TESTE" -ForegroundColor Cyan
    Write-Host ("Previstos: {0}" -f $total)
    Write-Host ("Enviados : {0}" -f $sent) -ForegroundColor Green
    Write-Host ("Falhas   : {0}" -f $failed) -ForegroundColor $(if ($failed -gt 0) { "Yellow" } else { "Green" })
    Write-Host ("Resumo CSV: {0}" -f $summaryPath)

    Write-RoboLog ("Fim do lote de teste. Previstos={0}; Enviados={1}; Falhas={2}; Resumo={3}" -f $total, $sent, $failed, $summaryPath)
}
