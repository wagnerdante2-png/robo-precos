$ConfigPath = Join-Path $Root "config.json"
$ConfigExamplePath = Join-Path $Root "config.example.json"
$StorePath = Join-Path $Root "data\lojas.csv"
$StoreExamplePath = Join-Path $Root "data\lojas.exemplo.csv"
$OutputPath = Join-Path $Root "output"
$LogDirectory = Join-Path $OutputPath "logs"
$PreviewDirectory = Join-Path $OutputPath "previews"
$StoreOutputDirectory = Join-Path $OutputPath "lojas"

function Ensure-RoboDirectory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

Ensure-RoboDirectory $OutputPath
Ensure-RoboDirectory $LogDirectory
Ensure-RoboDirectory $PreviewDirectory
Ensure-RoboDirectory $StoreOutputDirectory
Ensure-RoboDirectory (Join-Path $Root "data")

$CurrentLogFile = Join-Path $LogDirectory ("run_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

function Write-RoboLog {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $line = "{0} | {1} | {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -LiteralPath $CurrentLogFile -Value $line -Encoding UTF8
    Write-Host $line
}

function ConvertTo-RoboStore {
    param($Value)

    if ($null -eq $Value) { return "" }

    $text = ([string]$Value).Trim().ToUpperInvariant()
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }

    if ($text -match '^\d+([.,]0+)?$') {
        $numberText = $text -replace ',', '.'
        $number = [int][double]::Parse($numberText, [Globalization.CultureInfo]::InvariantCulture)
        return ("ML{0:D2}" -f $number)
    }

    if ($text -match '^ML\s*0*(\d+)$') {
        return ("ML{0:D2}" -f [int]$Matches[1])
    }

    return ($text -replace '\s+', '')
}

function ConvertTo-RoboPhone {
    param($Value)
    if ($null -eq $Value) { return "" }
    return (([string]$Value) -replace '\D', '')
}

function Get-RoboChrome {
    $programFilesX86 = [Environment]::GetFolderPath("ProgramFilesX86")
    $candidates = @(
        (Join-Path $env:ProgramFiles "Google\Chrome\Application\chrome.exe"),
        (Join-Path $programFilesX86 "Google\Chrome\Application\chrome.exe"),
        (Join-Path $env:LOCALAPPDATA "Google\Chrome\Application\chrome.exe")
    )

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }

    return $null
}

function Get-RoboStoreMap {
    if (-not (Test-Path -LiteralPath $StorePath)) {
        throw "Cadastro de lojas ausente: $StorePath"
    }

    $map = @{}
    foreach ($row in (Import-Csv -LiteralPath $StorePath)) {
        $active = $true

        if ($row.PSObject.Properties.Name -contains "ativo") {
            $activeText = ([string]$row.ativo).Trim().ToLowerInvariant()
            $active = @("1", "true", "sim", "s", "ativo") -contains $activeText
        }

        if (-not $active) { continue }

        $store = ConvertTo-RoboStore $row.loja
        $phone = ConvertTo-RoboPhone $row.telefone

        if ($store -and $phone) {
            $map[$store] = $phone
        }
    }

    if ($map.Count -eq 0) {
        throw "Nenhuma loja ativa com telefone valido foi encontrada em $StorePath."
    }

    return $map
}

function Get-RoboDownloadSnapshot {
    $directory = Join-Path $env:USERPROFILE "Downloads"
    if (-not (Test-Path -LiteralPath $directory)) {
        $directory = $env:USERPROFILE
    }

    $files = @{}
    Get-ChildItem -LiteralPath $directory -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -in @(".xlsx", ".xls") -and -not $_.Name.StartsWith("~$") } |
        ForEach-Object { $files[$_.FullName] = $_.LastWriteTimeUtc.Ticks }

    return [PSCustomObject]@{
        Directory = $directory
        Files = $files
    }
}

function Wait-RoboExcel {
    param(
        [string]$Directory,
        [hashtable]$Before,
        [int]$TimeoutSeconds
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)

    Write-Host ""
    Write-Host "BI aberto. Exporte o relatorio para Excel." -ForegroundColor Cyan
    Write-Host "O robo detectara automaticamente o novo arquivo em:"
    Write-Host "  $Directory"
    Write-Host ""

    while ((Get-Date) -lt $deadline) {
        $files = Get-ChildItem -LiteralPath $Directory -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -in @(".xlsx", ".xls") -and -not $_.Name.StartsWith("~$") } |
            Sort-Object LastWriteTimeUtc -Descending

        foreach ($file in $files) {
            $isNew = -not $Before.ContainsKey($file.FullName)
            $wasChanged = $false

            if (-not $isNew) {
                $wasChanged = $Before[$file.FullName] -ne $file.LastWriteTimeUtc.Ticks
            }

            if ($isNew -or $wasChanged) {
                Start-Sleep -Seconds 2
                Write-RoboLog ("Excel detectado: " + $file.FullName)
                return $file.FullName
            }
        }

        Start-Sleep -Seconds 1
    }

    throw "Tempo esgotado aguardando a exportacao do Excel."
}
