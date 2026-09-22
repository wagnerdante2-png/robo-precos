function ConvertTo-RoboPrecosWorkbookStore {
    param($Value)

    if ($null -eq $Value) { return "" }
    $text = ([string]$Value).Trim().ToUpperInvariant()

    if ($text -match 'ML\s*0*(\d+)') {
        return ("ML{0:D2}" -f [int]$Matches[1])
    }

    if ($text -match '^\d+([.,]0+)?$') {
        $numberText = $text -replace ',', '.'
        return ("ML{0:D2}" -f [int][double]::Parse($numberText, [Globalization.CultureInfo]::InvariantCulture))
    }

    return ""
}

function Get-RoboPrecosDiscountStoreMap {
    param(
        $Worksheet,
        [int]$StartRow,
        [int]$EndRow
    )

    $map = @{}

    for ($row = $StartRow; $row -le $EndRow; $row++) {
        $store = ConvertTo-RoboPrecosWorkbookStore $Worksheet.Cells.Item($row, 2).Value2
        if ([string]::IsNullOrWhiteSpace($store)) { continue }

        if ($map.ContainsKey($store)) {
            throw ("Loja duplicada no bloco DESCONTOS B{0}:B{1}: {2}" -f $StartRow, $EndRow, $store)
        }

        $map[$store] = $row
    }

    return $map
}

function Merge-RoboPrecosDiscountStoreMaps {
    param(
        [hashtable]$Primary,
        [hashtable]$Expansion
    )

    $result = @{}

    foreach ($key in $Primary.Keys) {
        $result[$key] = [int]$Primary[$key]
    }

    foreach ($key in $Expansion.Keys) {
        if ($result.ContainsKey($key)) {
            throw ("Loja duplicada entre bloco principal e expansao de DESCONTOS: " + $key)
        }
        $result[$key] = [int]$Expansion[$key]
    }

    return $result
}

function Get-RoboPrecosMonthColumnAtRow {
    param(
        $Worksheet,
        [int]$HeaderRow,
        [datetime]$MonthDate
    )

    $xlToLeft = -4159
    $lastColumn = [int]$Worksheet.Cells.Item($HeaderRow, $Worksheet.Columns.Count).End($xlToLeft).Column
    $targetSerial = [double]$MonthDate.ToOADate()

    for ($column = 3; $column -le $lastColumn; $column++) {
        $value = $Worksheet.Cells.Item($HeaderRow, $column).Value2
        if ($null -eq $value) { continue }

        $numeric = 0.0
        if ([double]::TryParse(
            ([string]$value),
            [Globalization.NumberStyles]::Any,
            [Globalization.CultureInfo]::InvariantCulture,
            [ref]$numeric
        )) {
            if ([Math]::Abs($numeric - $targetSerial) -lt 0.01) {
                return $column
            }
        }
    }

    throw ("Mes " + $MonthDate.ToString("MM/yyyy") + " nao encontrado na linha " + $HeaderRow + " da aba DESCONTOS.")
}

function Invoke-RoboPrecosDiscountWorkbook {
    param(
        [Parameter(Mandatory = $true)]$DiscountResult
    )

    $records = @($DiscountResult.Records)
    if ($records.Count -eq 0) {
        throw "Nenhum registro de desconto valido foi fornecido para gravacao."
    }

    $workbookPath = Get-RoboPrecosControlWorkbookPath
    Assert-RoboPrecosWorkbookUnlocked -Path $workbookPath

    $excel = $null
    $workbook = $null
    $sheet = $null
    $saved = $false
    $backupPath = ""

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($workbookPath, 0, $false)

        try {
            $sheet = $workbook.Worksheets.Item("DESCONTOS")
        }
        catch {
            throw "A aba obrigatoria 'DESCONTOS' nao foi encontrada na planilha."
        }

        try {
            $excel.Calculation = -4105
            $excel.CalculateFullRebuild()
        }
        catch {
            try { $workbook.Calculate() } catch {}
        }

        # Valor de descontos: bloco historico principal e bloco de expansao.
        $valuePrimary = Get-RoboPrecosDiscountStoreMap -Worksheet $sheet -StartRow 3 -EndRow 62
        $valueExpansion = Get-RoboPrecosDiscountStoreMap -Worksheet $sheet -StartRow 141 -EndRow 180
        $valueMap = Merge-RoboPrecosDiscountStoreMaps -Primary $valuePrimary -Expansion $valueExpansion

        # Quantidade de cupons: bloco principal e expansao.
        $qtyPrimary = Get-RoboPrecosDiscountStoreMap -Worksheet $sheet -StartRow 66 -EndRow 125
        $qtyExpansion = Get-RoboPrecosDiscountStoreMap -Worksheet $sheet -StartRow 183 -EndRow 222
        $qtyMap = Merge-RoboPrecosDiscountStoreMaps -Primary $qtyPrimary -Expansion $qtyExpansion

        $monthDate = [datetime]$DiscountResult.MonthDate
        $valueColumn = Get-RoboPrecosMonthColumnAtRow -Worksheet $sheet -HeaderRow 2 -MonthDate $monthDate
        $qtyColumn = Get-RoboPrecosMonthColumnAtRow -Worksheet $sheet -HeaderRow 65 -MonthDate $monthDate

        if ($valueColumn -ne $qtyColumn) {
            throw ("As colunas de valor e quantidade nao coincidem para " + $monthDate.ToString("MM/yyyy") + ".")
        }

        $backupPath = New-RoboPrecosWorkbookBackup -Path $workbookPath -Workbook $workbook

        $writtenStores = 0
        $ignoredStores = New-Object System.Collections.Generic.List[string]
        $writtenCells = 0

        foreach ($record in $records) {
            $store = ConvertTo-RoboPrecosWorkbookStore $record.Loja
            if ([string]::IsNullOrWhiteSpace($store)) {
                continue
            }

            # Loja ausente da planilha: ignora e registra. Nunca cria/realoca linha aqui.
            if (-not $valueMap.ContainsKey($store) -or -not $qtyMap.ContainsKey($store)) {
                $ignoredStores.Add($store)
                Write-RoboLog ("Desconto ignorado: loja nao encontrada nos blocos da planilha: " + $store) "AVISO"
                continue
            }

            $quantity = $record.QuantidadeCupons
            $discount = $record.Desconto

            # Ausencia de valor nao vira zero. Zero explicito e valido.
            if ($null -eq $quantity -or $null -eq $discount) {
                Write-RoboLog ("Desconto ignorado por campo vazio: " + $store) "AVISO"
                continue
            }

            $valueRow = [int]$valueMap[$store]
            $qtyRow = [int]$qtyMap[$store]

            $sheet.Cells.Item($valueRow, $valueColumn).Value2 = [double]$discount
            $sheet.Cells.Item($qtyRow, $qtyColumn).Value2 = [int]$quantity

            $writtenStores++
            $writtenCells += 2
        }

        if ($writtenStores -eq 0) {
            throw "Nenhuma loja coletada no Power BI correspondeu a uma loja valida na planilha. Nenhuma celula foi gravada."
        }

        $workbook.Save()
        $saved = $true

        Write-RoboLog ("Descontos gravados: " + $writtenStores + " lojas | " + $writtenCells + " celulas | Mes " + $monthDate.ToString("MM/yyyy"))

        Write-Host ""
        Write-Host "DESCONTOS GRAVADOS NA PLANILHA" -ForegroundColor Green
        Write-Host ("Mes                : {0}" -f $monthDate.ToString("MM/yyyy"))
        Write-Host ("Lojas coletadas BI : {0}" -f $records.Count)
        Write-Host ("Lojas gravadas     : {0}" -f $writtenStores)
        Write-Host ("Lojas ignoradas    : {0}" -f $ignoredStores.Count)
        Write-Host ("Celulas gravadas   : {0}" -f $writtenCells)
        Write-Host ("Backup              : {0}" -f $backupPath)

        if ($ignoredStores.Count -gt 0) {
            Write-Host ("Ignoradas           : " + (($ignoredStores | Sort-Object -Unique) -join ", ")) -ForegroundColor Yellow
        }
        Write-Host ""

        return [PSCustomObject]@{
            WorkbookPath = $workbookPath
            BackupPath = $backupPath
            Month = $monthDate.ToString("MM/yyyy")
            CollectedCount = $records.Count
            WrittenStoreCount = $writtenStores
            IgnoredStoreCount = $ignoredStores.Count
            IgnoredStores = @($ignoredStores)
            CellsWritten = $writtenCells
        }
    }
    finally {
        if ($workbook) {
            try { $workbook.Close($false) } catch {}
        }
        if ($excel) {
            try { $excel.Quit() } catch {}
        }
        if ($sheet) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) } catch {}
        }
        if ($workbook) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } catch {}
        }
        if ($excel) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {}
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()

        if (-not $saved -and -not [string]::IsNullOrWhiteSpace($backupPath)) {
            Write-RoboLog ("Gravacao dos descontos nao foi concluida. Backup preservado em: " + $backupPath) "AVISO"
        }
    }
}
