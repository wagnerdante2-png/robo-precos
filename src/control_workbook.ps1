function Get-RoboPrecosControlWorkbookPath {
    $downloads = Join-Path $env:USERPROFILE "Downloads"
    if (-not (Test-Path -LiteralPath $downloads)) {
        throw ("Pasta Downloads nao encontrada: " + $downloads)
    }

    $cedilla = [char]0x00E7
    $expectedName = "Controle de Auditoria de Pre" + $cedilla + "os.xlsx"
    $expectedPath = Join-Path $downloads $expectedName

    if (Test-Path -LiteralPath $expectedPath) {
        return $expectedPath
    }

    $matches = @(
        Get-ChildItem -LiteralPath $downloads -File -Filter "*.xlsx" -ErrorAction SilentlyContinue |
        Where-Object {
            -not $_.Name.StartsWith("~$") -and
            $_.BaseName -like ("Controle de Auditoria de Pre" + $cedilla + "os*") -and
            $_.DirectoryName -eq $downloads
        }
    )

    if ($matches.Count -eq 1) {
        Write-RoboLog ("Planilha localizada por nome compativel: " + $matches[0].FullName) "AVISO"
        return $matches[0].FullName
    }

    if ($matches.Count -gt 1) {
        $names = @($matches | ForEach-Object { $_.Name }) -join ", "
        throw ("Mais de uma planilha de controle foi encontrada em Downloads. Deixe apenas o arquivo correto ou use o nome exato '" + $expectedName + "'. Encontrados: " + $names)
    }

    throw ("Planilha de controle nao encontrada. Deixe o arquivo em: " + $expectedPath)
}

function Assert-RoboPrecosWorkbookUnlocked {
    param([Parameter(Mandatory = $true)][string]$Path)

    $stream = $null
    try {
        $stream = [System.IO.File]::Open(
            $Path,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::ReadWrite,
            [System.IO.FileShare]::None
        )
    }
    catch {
        throw ("A planilha de controle esta aberta ou bloqueada. Feche o Excel antes de executar o robo. Arquivo: " + $Path)
    }
    finally {
        if ($stream) {
            try { $stream.Dispose() } catch {}
        }
    }
}

function New-RoboPrecosWorkbookBackup {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        $Workbook
    )

    $downloads = Split-Path -Parent $Path
    $backupDirectory = Join-Path $downloads "RoboPrecos_Backups"
    Ensure-RoboDirectory $backupDirectory

    $base = [IO.Path]::GetFileNameWithoutExtension($Path)
    $extension = [IO.Path]::GetExtension($Path)
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupPath = Join-Path $backupDirectory ($base + "_" + $timestamp + $extension)

    if ($Workbook) {
        $Workbook.SaveCopyAs($backupPath)
    }
    else {
        Copy-Item -LiteralPath $Path -Destination $backupPath -Force
    }

    Write-RoboLog ("Backup da planilha criado: " + $backupPath)

    return $backupPath
}

function Get-RoboPrecosExcelSheetNameConfig {
    $circumflexE = [char]0x00CA
    $cedillaUpper = [char]0x00C7

    return @(
        [PSCustomObject]@{
            Sheet = "ETIQUETAS"
            Field = "Total"
            Label = "TOTAL"
        },
        [PSCustomObject]@{
            Sheet = ("DIVERG" + $circumflexE + "NCIAS")
            Field = "Divergente"
            Label = "DIVERGENTE"
        },
        [PSCustomObject]@{
            Sheet = ("SEM PRE" + $cedillaUpper + "O")
            Field = "SemEtiqueta"
            Label = "SEM ETIQUETA"
        }
    )
}

function Get-RoboPrecosSheetStoreMap {
    param($Worksheet)

    $xlUp = -4162
    $lastRow = [int]$Worksheet.Cells.Item($Worksheet.Rows.Count, 2).End($xlUp).Row

    if ($lastRow -lt 3) {
        throw ("A aba '" + [string]$Worksheet.Name + "' nao possui lojas na coluna B.")
    }

    $map = @{}

    for ($row = 3; $row -le $lastRow; $row++) {
        $raw = $Worksheet.Cells.Item($row, 2).Value2

        if ($null -eq $raw -or [string]::IsNullOrWhiteSpace([string]$raw)) {
            continue
        }

        $store = ConvertTo-RoboStore $raw
        if ($store -notmatch '^ML\d+$') {
            throw ("Valor de loja invalido na aba '" + [string]$Worksheet.Name + "', celula B" + $row + ": " + [string]$raw)
        }

        if ($map.ContainsKey($store)) {
            throw ("Loja duplicada na aba '" + [string]$Worksheet.Name + "': " + $store)
        }

        $map[$store] = $row
    }

    if ($map.Count -eq 0) {
        throw ("Nenhuma loja valida encontrada na coluna B da aba '" + [string]$Worksheet.Name + "'.")
    }

    return $map
}

function Compare-RoboPrecosStoreMaps {
    param(
        [hashtable]$ReferenceMap,
        [hashtable]$OtherMap,
        [string]$ReferenceName,
        [string]$OtherName
    )

    $reference = @($ReferenceMap.Keys | Sort-Object)
    $other = @($OtherMap.Keys | Sort-Object)

    $missing = @($reference | Where-Object { -not $OtherMap.ContainsKey($_) })
    $extra = @($other | Where-Object { -not $ReferenceMap.ContainsKey($_) })

    if ($missing.Count -gt 0 -or $extra.Count -gt 0) {
        $parts = @()
        if ($missing.Count -gt 0) {
            $parts += ("faltando em " + $OtherName + ": " + ($missing -join ", "))
        }
        if ($extra.Count -gt 0) {
            $parts += ("extras em " + $OtherName + ": " + ($extra -join ", "))
        }

        throw ("As whitelists da coluna B nao sao identicas entre '" + $ReferenceName + "' e '" + $OtherName + "'. " + ($parts -join " | "))
    }
}

function Get-RoboPrecosMonthColumn {
    param(
        $Worksheet,
        [datetime]$MonthDate
    )

    $xlToLeft = -4159
    $lastColumn = [int]$Worksheet.Cells.Item(2, $Worksheet.Columns.Count).End($xlToLeft).Column
    $targetSerial = [double]$MonthDate.ToOADate()

    for ($column = 3; $column -le $lastColumn; $column++) {
        $value = $Worksheet.Cells.Item(2, $column).Value2
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

    throw ("Mes " + $MonthDate.ToString("MM/yyyy") + " nao encontrado na linha 2 da aba '" + [string]$Worksheet.Name + "'. O robo nao cria novas colunas de mes automaticamente.")
}

function Get-RoboPrecosRowMap {
    param([array]$Rows)

    $map = @{}

    foreach ($row in @($Rows)) {
        $store = ConvertTo-RoboStore $row.Loja
        if ([string]::IsNullOrWhiteSpace($store)) { continue }

        if ([string]$row.Status -ne "OK") {
            continue
        }

        $map[$store] = $row
    }

    return $map
}

function Invoke-RoboPrecosControlWorkbook {
    param(
        [array]$Rows,
        [Parameter(Mandatory = $true)][string]$StartDate,
        [Parameter(Mandatory = $true)][string]$EndDate
    )

    $culture = [Globalization.CultureInfo]::InvariantCulture
    $start = [datetime]::ParseExact($StartDate, "dd/MM/yyyy", $culture)
    $end = [datetime]::ParseExact($EndDate, "dd/MM/yyyy", $culture)

    if ($start.Year -ne $end.Year -or $start.Month -ne $end.Month) {
        throw "A planilha de controle e mensal. Data inicial e data final precisam pertencer ao mesmo mes."
    }

    $monthDate = Get-Date -Year $start.Year -Month $start.Month -Day 1 -Hour 0 -Minute 0 -Second 0
    $workbookPath = Get-RoboPrecosControlWorkbookPath
    Assert-RoboPrecosWorkbookUnlocked -Path $workbookPath

    $excel = $null
    $workbook = $null
    $sheetObjects = @()
    $saved = $false
    $backupPath = ""

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($workbookPath, 0, $false)

        $configs = @(Get-RoboPrecosExcelSheetNameConfig)
        $sheetMaps = @{}
        $monthColumns = @{}

        foreach ($config in $configs) {
            try {
                $sheet = $workbook.Worksheets.Item([string]$config.Sheet)
            }
            catch {
                throw ("A aba obrigatoria '" + [string]$config.Sheet + "' nao foi encontrada na planilha.")
            }

            $sheetObjects += $sheet
            $sheetMaps[[string]$config.Sheet] = Get-RoboPrecosSheetStoreMap -Worksheet $sheet
            $monthColumns[[string]$config.Sheet] = Get-RoboPrecosMonthColumn -Worksheet $sheet -MonthDate $monthDate
        }

        $referenceName = [string]$configs[0].Sheet
        $referenceMap = $sheetMaps[$referenceName]

        foreach ($config in @($configs | Select-Object -Skip 1)) {
            Compare-RoboPrecosStoreMaps -ReferenceMap $referenceMap -OtherMap $sheetMaps[[string]$config.Sheet] -ReferenceName $referenceName -OtherName ([string]$config.Sheet)
        }

        $collectedMap = Get-RoboPrecosRowMap -Rows $Rows

        $missing = @($referenceMap.Keys | Where-Object { -not $collectedMap.ContainsKey($_) } | Sort-Object)
        if ($missing.Count -gt 0) {
            throw ("Existem lojas na planilha sem coleta valida no PDA: " + ($missing -join ", ") + ". Nenhum dado foi gravado.")
        }

        $ignored = @($collectedMap.Keys | Where-Object { -not $referenceMap.ContainsKey($_) } | Sort-Object)

        Write-Host ""
        Write-Host "VALIDACAO DA PLANILHA DE CONTROLE" -ForegroundColor Cyan
        Write-Host ("Arquivo             : {0}" -f $workbookPath)
        Write-Host ("Mes                  : {0}" -f $monthDate.ToString("MM/yyyy"))
        Write-Host ("Lojas na whitelist   : {0}" -f $referenceMap.Count)
        Write-Host ("Centros coletados PDA: {0}" -f $collectedMap.Count)
        Write-Host ("Centros ignorados    : {0}" -f $ignored.Count)

        if ($ignored.Count -gt 0) {
            Write-Host ("Ignorados            : " + ($ignored -join ", ")) -ForegroundColor DarkGray
            Write-RoboLog ("Centros coletados no PDA e ignorados por nao existirem na coluna B da planilha: " + ($ignored -join ", "))
        }

        $backupPath = New-RoboPrecosWorkbookBackup -Path $workbookPath -Workbook $workbook

        $written = 0

        foreach ($config in $configs) {
            $sheetName = [string]$config.Sheet
            $field = [string]$config.Field
            $sheet = $workbook.Worksheets.Item($sheetName)
            $rowMap = $sheetMaps[$sheetName]
            $column = [int]$monthColumns[$sheetName]

            foreach ($store in ($referenceMap.Keys | Sort-Object)) {
                $rowNumber = [int]$rowMap[$store]
                $record = $collectedMap[$store]
                $property = $record.PSObject.Properties[$field]

                if (-not $property) {
                    throw ("Campo '" + $field + "' ausente no resultado coletado de " + $store + ".")
                }

                $value = [int]$property.Value
                $sheet.Cells.Item($rowNumber, $column).Value2 = $value
                $written++
            }
        }

        $workbook.Save()
        $saved = $true

        Write-RoboLog ("Planilha de controle atualizada com sucesso: " + $workbookPath)
        Write-RoboLog ("Celulas gravadas: " + $written + " | Lojas: " + $referenceMap.Count + " | Mes: " + $monthDate.ToString("MM/yyyy"))

        Write-Host ""
        Write-Host "PLANILHA ATUALIZADA COM SUCESSO" -ForegroundColor Green
        Write-Host ("Lojas preenchidas : {0}" -f $referenceMap.Count)
        Write-Host ("Celulas gravadas   : {0}" -f $written)
        Write-Host ("Backup criado      : {0}" -f $backupPath)
        Write-Host ""

        return [PSCustomObject]@{
            WorkbookPath = $workbookPath
            BackupPath = $backupPath
            Month = $monthDate.ToString("MM/yyyy")
            StoreCount = $referenceMap.Count
            IgnoredCount = $ignored.Count
            IgnoredStores = $ignored
            CellsWritten = $written
        }
    }
    finally {
        if ($workbook) {
            try { $workbook.Close($false) } catch {}
        }

        if ($excel) {
            try { $excel.Quit() } catch {}
        }

        foreach ($sheet in $sheetObjects) {
            if ($sheet) {
                try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) } catch {}
            }
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
            Write-RoboLog ("A gravacao nao foi concluida. A copia de seguranca permanece em: " + $backupPath) "AVISO"
        }
    }
}
