function ConvertTo-RoboNormalizedExcel {
    param(
        [Parameter(Mandatory = $true)][string]$ExcelPath,
        [Parameter(Mandatory = $true)]$Config
    )

    $excel = $null
    $workbook = $null
    $worksheet = $null
    $outWorkbook = $null
    $outWorksheet = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($ExcelPath, 0, $true)
        $configuredSheet = [string]$Config.excel.worksheet

        if ([string]::IsNullOrWhiteSpace($configuredSheet)) {
            $worksheet = $workbook.Worksheets.Item(1)
        }
        else {
            $worksheet = $workbook.Worksheets.Item($configuredSheet)
        }

        $used = $worksheet.UsedRange
        $rowCount = [int]$used.Rows.Count
        $columnCount = [int]$used.Columns.Count
        $wanted = ([string]$Config.excel.storeColumn).Trim()
        $aliases = @($wanted, "Loja", "NOME", "Nome")

        $headerRow = 0
        $storeColumn = 0

        $scanRows = [Math]::Min(20, $rowCount)
        for ($row = 1; $row -le $scanRows -and $headerRow -eq 0; $row++) {
            for ($column = 1; $column -le $columnCount; $column++) {
                $text = ([string]$worksheet.Cells.Item($row, $column).Text).Trim()
                foreach ($alias in $aliases) {
                    if (-not [string]::IsNullOrWhiteSpace($alias) -and $text.Equals($alias, [StringComparison]::OrdinalIgnoreCase)) {
                        $headerRow = $row
                        $storeColumn = $column
                        break
                    }
                }
                if ($headerRow -gt 0) { break }
            }
        }

        if ($headerRow -eq 0) {
            throw "Nao foi possivel localizar a coluna de loja nas primeiras 20 linhas do Excel."
        }

        $normalizedPath = Join-Path $OutputPath ("_entrada_normalizada_{0}.xlsx" -f (Get-Date -Format "yyyyMMdd_HHmmssfff"))
        $outWorkbook = $excel.Workbooks.Add()
        $outWorksheet = $outWorkbook.Worksheets.Item(1)

        for ($column = 1; $column -le $columnCount; $column++) {
            $header = ([string]$worksheet.Cells.Item($headerRow, $column).Text).Trim()
            if ($column -eq $storeColumn) {
                $header = $wanted
            }
            if ([string]::IsNullOrWhiteSpace($header)) {
                $header = "Coluna$column"
            }
            $outWorksheet.Cells.Item(1, $column).Value2 = $header
        }

        $destinationRow = 2
        for ($sourceRow = $headerRow + 1; $sourceRow -le $rowCount; $sourceRow++) {
            $rawStore = ([string]$worksheet.Cells.Item($sourceRow, $storeColumn).Text).Trim()
            if ([string]::IsNullOrWhiteSpace($rawStore)) {
                continue
            }

            $normalizedStore = $rawStore.ToUpperInvariant()
            if ($normalizedStore -match '^ML[\s\.\-_]*0*(\d+)') {
                $normalizedStore = "ML{0:D2}" -f [int]$Matches[1]
            }
            elseif ($normalizedStore -match '^\d+([.,]0+)?$') {
                $numberText = $normalizedStore -replace ',', '.'
                $normalizedStore = "ML{0:D2}" -f [int][double]::Parse($numberText, [Globalization.CultureInfo]::InvariantCulture)
            }

            for ($column = 1; $column -le $columnCount; $column++) {
                if ($column -eq $storeColumn) {
                    $outWorksheet.Cells.Item($destinationRow, $column).Value2 = $normalizedStore
                }
                else {
                    $outWorksheet.Cells.Item($destinationRow, $column).Value2 = $worksheet.Cells.Item($sourceRow, $column).Value2
                }
            }
            $destinationRow++
        }

        $outWorkbook.SaveAs($normalizedPath, 51)
        return $normalizedPath
    }
    finally {
        if ($outWorkbook) { try { $outWorkbook.Close($false) } catch {} }
        if ($workbook) { try { $workbook.Close($false) } catch {} }
        if ($excel) { try { $excel.Quit() } catch {} }
        if ($outWorksheet) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($outWorksheet) } catch {} }
        if ($outWorkbook) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($outWorkbook) } catch {} }
        if ($worksheet) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) } catch {} }
        if ($workbook) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } catch {} }
        if ($excel) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {} }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

$script:RoboBaseInvokeExcel = ${function:Invoke-RoboExcel}

function Invoke-RoboExcel {
    param(
        [string]$ExcelPath,
        $Config,
        [hashtable]$StoreMap,
        [string]$ChromePath,
        [string]$ProfilePath
    )

    $normalizedPath = ConvertTo-RoboNormalizedExcel -ExcelPath $ExcelPath -Config $Config
    Write-RoboLog ("Entrada normalizada para processamento: " + $normalizedPath)

    try {
        & $script:RoboBaseInvokeExcel $normalizedPath $Config $StoreMap $ChromePath $ProfilePath
    }
    finally {
        if ($normalizedPath -and (Test-Path -LiteralPath $normalizedPath)) {
            Remove-Item -LiteralPath $normalizedPath -Force -ErrorAction SilentlyContinue
        }
    }
}
