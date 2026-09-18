function New-RoboMessage {
    param(
        [string]$Store,
        $Worksheet,
        [int]$RowNumber,
        [array]$Headers,
        [hashtable]$HeaderIndex,
        $Config,
        [int]$RecordCount
    )

    $lines = New-Object "System.Collections.Generic.List[string]"
    $title = ([string]$Config.message.title).Replace("{loja}", $Store)

    $lines.Add($title)
    $lines.Add("")

    if ([bool]$Config.message.includeRecordCount -and $RecordCount -gt 1) {
        $lines.Add(("Registros no relatorio: {0}" -f $RecordCount))
        $lines.Add("")
    }

    $fields = @($Config.message.fields)

    if ($fields.Count -eq 0) {
        $fields = @()
        $maxFields = [int]$Config.message.maxAutoFields

        foreach ($header in $Headers) {
            if ($header -ne [string]$Config.excel.storeColumn) {
                $fields += $header
            }

            if ($fields.Count -ge $maxFields) {
                break
            }
        }
    }

    foreach ($field in $fields) {
        $fieldName = [string]$field

        if ($HeaderIndex.ContainsKey($fieldName)) {
            $columnNumber = [int]$HeaderIndex[$fieldName]
            $value = [string]$Worksheet.Cells.Item($RowNumber, $columnNumber).Text

            if ([string]::IsNullOrWhiteSpace($value)) {
                $value = "-"
            }

            $lines.Add(("{0}: {1}" -f $fieldName, $value))
        }
    }

    $footer = [string]$Config.message.footer
    if (-not [string]::IsNullOrWhiteSpace($footer)) {
        $lines.Add("")
        $lines.Add($footer.Trim())
    }

    return [string]::Join([Environment]::NewLine, $lines)
}

function Invoke-RoboWhatsAppSendButton {
    param(
        [string[]]$ButtonNames = @("Enviar", "Send"),
        [int]$TimeoutSeconds = 25
    )

    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $windowCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Window
    )
    $buttonCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Button
    )

    while ((Get-Date) -lt $deadline) {
        $root = [System.Windows.Automation.AutomationElement]::RootElement
        $windows = $root.FindAll([System.Windows.Automation.TreeScope]::Children, $windowCondition)

        foreach ($window in $windows) {
            $windowName = [string]$window.Current.Name
            if ($windowName -notmatch "WhatsApp") {
                continue
            }

            $buttons = $window.FindAll([System.Windows.Automation.TreeScope]::Descendants, $buttonCondition)
            foreach ($button in $buttons) {
                $name = ([string]$button.Current.Name).Trim()
                if ([string]::IsNullOrWhiteSpace($name)) {
                    continue
                }

                $matches = $false
                foreach ($expected in $ButtonNames) {
                    if ($name.Equals([string]$expected, [StringComparison]::OrdinalIgnoreCase)) {
                        $matches = $true
                        break
                    }
                }

                if (-not $matches) {
                    continue
                }

                try {
                    $patternObject = $null
                    if ($button.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$patternObject)) {
                        ([System.Windows.Automation.InvokePattern]$patternObject).Invoke()
                        Write-RoboLog ("Botao do WhatsApp acionado via UI Automation: " + $name)
                        return $true
                    }
                }
                catch {
                    Write-RoboLog ("Falha ao acionar botao '$name': " + $_.Exception.Message) "AVISO"
                }
            }
        }

        Start-Sleep -Milliseconds 500
    }

    return $false
}

function Send-RoboWhatsApp {
    param(
        [string]$ChromePath,
        [string]$ProfilePath,
        [string]$Phone,
        [string]$Message,
        [int]$WaitSeconds,
        [bool]$UseExistingChromeSession = $true,
        [bool]$UseUiAutomationSend = $true,
        [int]$SendButtonTimeoutSeconds = 25,
        [string[]]$SendButtonNames = @("Enviar", "Send")
    )

    $encoded = [Uri]::EscapeDataString($Message)
    $url = "https://web.whatsapp.com/send?phone=$Phone&text=$encoded"

    if ($UseExistingChromeSession) {
        $arguments = @("--start-maximized", $url)
    }
    else {
        $arguments = @("--user-data-dir=$ProfilePath", "--start-maximized", $url)
    }

    Start-Process -FilePath $ChromePath -ArgumentList $arguments | Out-Null
    Start-Sleep -Seconds $WaitSeconds

    if ($UseUiAutomationSend) {
        $clicked = Invoke-RoboWhatsAppSendButton -ButtonNames $SendButtonNames -TimeoutSeconds $SendButtonTimeoutSeconds
        if ($clicked) {
            Start-Sleep -Seconds 2
            return
        }

        throw "O WhatsApp abriu a conversa, mas o botao Enviar nao foi localizado automaticamente."
    }

    $shell = New-Object -ComObject WScript.Shell
    $activated = $false

    for ($attempt = 0; $attempt -lt 10; $attempt++) {
        if ($shell.AppActivate("WhatsApp")) {
            $activated = $true
            break
        }
        Start-Sleep -Seconds 1
    }

    if (-not $activated) {
        throw "Janela do WhatsApp Web nao encontrada. Confirme que o Chrome abriu a conversa."
    }

    Start-Sleep -Milliseconds 700
    $shell.SendKeys("{ENTER}")
}

function Invoke-RoboExcel {
    param(
        [string]$ExcelPath,
        $Config,
        [hashtable]$StoreMap,
        [string]$ChromePath,
        [string]$ProfilePath
    )

    $excelApp = $null
    $workbook = $null
    $worksheet = $null

    try {
        $excelApp = New-Object -ComObject Excel.Application
        $excelApp.Visible = $false
        $excelApp.DisplayAlerts = $false

        $workbook = $excelApp.Workbooks.Open($ExcelPath, 0, $true)

        $sheetName = [string]$Config.excel.worksheet
        if ([string]::IsNullOrWhiteSpace($sheetName)) {
            $worksheet = $workbook.Worksheets.Item(1)
        }
        else {
            $worksheet = $workbook.Worksheets.Item($sheetName)
        }

        $usedRange = $worksheet.UsedRange
        $rowCount = [int]$usedRange.Rows.Count
        $columnCount = [int]$usedRange.Columns.Count

        if ($rowCount -lt 2) {
            throw "O Excel exportado nao possui linhas de dados."
        }

        $headers = @()
        $headerIndex = @{}

        for ($column = 1; $column -le $columnCount; $column++) {
            $header = ([string]$worksheet.Cells.Item(1, $column).Text).Trim()

            if ([string]::IsNullOrWhiteSpace($header)) {
                $header = "Coluna$column"
            }

            $headers += $header
            $headerIndex[$header] = $column
        }

        $storeColumnName = [string]$Config.excel.storeColumn
        if (-not $headerIndex.ContainsKey($storeColumnName)) {
            throw "Coluna '$storeColumnName' nao encontrada. Colunas detectadas: $([string]::Join(', ', $headers))"
        }

        $storeColumnNumber = [int]$headerIndex[$storeColumnName]
        $groups = @{}

        for ($row = 2; $row -le $rowCount; $row++) {
            $store = ConvertTo-RoboStore $worksheet.Cells.Item($row, $storeColumnNumber).Value2

            if ([string]::IsNullOrWhiteSpace($store)) {
                continue
            }

            if (-not $groups.ContainsKey($store)) {
                $groups[$store] = New-Object System.Collections.ArrayList
            }

            [void]$groups[$store].Add($row)
        }

        Write-RoboLog ("Lojas encontradas no Excel: " + $groups.Count)
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

        $testMode = $false
        $testStore = "ML01"
        $testPhone = ""
        $useExistingChromeSession = $true
        $useUiAutomationSend = $true
        $sendButtonTimeoutSeconds = 25
        $sendButtonNames = @("Enviar", "Send")

        if ($Config.whatsapp.PSObject.Properties.Name -contains "testMode") {
            $testMode = [bool]$Config.whatsapp.testMode
        }
        if ($Config.whatsapp.PSObject.Properties.Name -contains "testStore") {
            $testStore = ConvertTo-RoboStore ([string]$Config.whatsapp.testStore)
        }
        if ($Config.whatsapp.PSObject.Properties.Name -contains "testPhone") {
            $testPhone = ConvertTo-RoboPhone ([string]$Config.whatsapp.testPhone)
        }
        if ($Config.whatsapp.PSObject.Properties.Name -contains "useExistingChromeSession") {
            $useExistingChromeSession = [bool]$Config.whatsapp.useExistingChromeSession
        }
        if ($Config.whatsapp.PSObject.Properties.Name -contains "useUiAutomationSend") {
            $useUiAutomationSend = [bool]$Config.whatsapp.useUiAutomationSend
        }
        if ($Config.whatsapp.PSObject.Properties.Name -contains "sendButtonTimeoutSeconds") {
            $sendButtonTimeoutSeconds = [int]$Config.whatsapp.sendButtonTimeoutSeconds
        }
        if ($Config.whatsapp.PSObject.Properties.Name -contains "sendButtonNames") {
            $sendButtonNames = @($Config.whatsapp.sendButtonNames)
        }

        $messagesSent = 0

        foreach ($store in ($groups.Keys | Sort-Object)) {
            if ($testMode -and $store -ne $testStore) {
                continue
            }

            $sourceRows = $groups[$store]
            $storeFile = Join-Path $StoreOutputDirectory ("{0}_{1}.xlsx" -f $store, $timestamp)

            $outWorkbook = $null
            $outWorksheet = $null

            try {
                $outWorkbook = $excelApp.Workbooks.Add()
                $outWorksheet = $outWorkbook.Worksheets.Item(1)

                for ($column = 1; $column -le $columnCount; $column++) {
                    $outWorksheet.Cells.Item(1, $column).Value2 = $headers[$column - 1]
                }

                $destinationRow = 2

                foreach ($sourceRow in $sourceRows) {
                    for ($column = 1; $column -le $columnCount; $column++) {
                        $outWorksheet.Cells.Item($destinationRow, $column).Value2 = $worksheet.Cells.Item($sourceRow, $column).Value2
                    }
                    $destinationRow++
                }

                $outWorkbook.SaveAs($storeFile, 51)
            }
            finally {
                if ($outWorkbook) {
                    try { $outWorkbook.Close($false) } catch {}
                }
                if ($outWorksheet) {
                    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($outWorksheet) } catch {}
                }
                if ($outWorkbook) {
                    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($outWorkbook) } catch {}
                }
            }

            $phone = ""
            if ($testMode) {
                $phone = $testPhone
            }
            elseif ($StoreMap.ContainsKey($store)) {
                $phone = [string]$StoreMap[$store]
            }

            if ([string]::IsNullOrWhiteSpace($phone)) {
                Write-RoboLog "$store sem telefone valido; arquivo separado criado." "AVISO"
                continue
            }

            $message = New-RoboMessage $store $worksheet ([int]$sourceRows[0]) $headers $headerIndex $Config $sourceRows.Count

            $previewFile = Join-Path $PreviewDirectory ("{0}_{1}.txt" -f $store, $timestamp)
            $nl = [Environment]::NewLine
            $previewText = "Loja: $store" + $nl + "Telefone: $phone" + $nl + $nl + $message
            Set-Content -LiteralPath $previewFile -Value $previewText -Encoding UTF8

            if ([bool]$Config.whatsapp.dryRun) {
                Write-Host ""
                Write-Host ("PREVIA {0} -> {1}" -f $store, $phone) -ForegroundColor Yellow
                Write-Host $message
                Write-RoboLog "$store preparado em DRY-RUN."
            }
            else {
                Send-RoboWhatsApp $ChromePath $ProfilePath $phone $message ([int]$Config.whatsapp.waitSeconds) $useExistingChromeSession $useUiAutomationSend $sendButtonTimeoutSeconds $sendButtonNames
                Write-RoboLog "$store enviado."
                $messagesSent++

                if ($testMode -and $messagesSent -ge 1) {
                    Write-RoboLog "Modo de teste: limite de uma mensagem atingido."
                    break
                }

                Start-Sleep -Seconds ([int]$Config.whatsapp.delayBetweenMessagesSeconds)
            }
        }

        if ($testMode -and -not $groups.ContainsKey($testStore)) {
            throw "A loja de teste '$testStore' nao foi encontrada no Excel."
        }
    }
    finally {
        if ($workbook) {
            try { $workbook.Close($false) } catch {}
        }

        if ($excelApp) {
            try { $excelApp.Quit() } catch {}
        }

        if ($worksheet) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) } catch {}
        }
        if ($workbook) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } catch {}
        }
        if ($excelApp) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp) } catch {}
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
