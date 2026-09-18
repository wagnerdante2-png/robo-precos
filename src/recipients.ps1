$RecipientDefaultPath = Join-Path $Root "data\destinatarios.csv"
$RecipientExamplePath = Join-Path $Root "data\destinatarios.exemplo.csv"

function Get-RoboRecipientPath {
    param($Config)

    $configured = ""
    if ($Config -and $Config.whatsapp -and ($Config.whatsapp.PSObject.Properties.Name -contains "recipientListFile")) {
        $configured = [string]$Config.whatsapp.recipientListFile
    }

    if ([string]::IsNullOrWhiteSpace($configured)) {
        return $RecipientDefaultPath
    }

    if ([IO.Path]::IsPathRooted($configured)) {
        return $configured
    }

    return (Join-Path $Root $configured)
}

function Test-RoboRecipientActive {
    param($Value)
    if ($null -eq $Value) { return $true }
    $text = ([string]$Value).Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($text)) { return $true }
    return @("1", "true", "sim", "s", "ativo", "yes", "y") -contains $text
}

function Get-RoboCsvDelimiter {
    param([string]$Path)

    $firstLine = Get-Content -LiteralPath $Path -TotalCount 1 -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($firstLine)) { return "," }

    $lower = $firstLine.ToLowerInvariant()
    if ($lower -match 'loja\s*;\s*nome\s*;\s*telefone') { return ";" }
    if ($lower -match 'loja\s*,\s*nome\s*,\s*telefone') { return "," }

    $semicolonCount = ([regex]::Matches($firstLine, ';')).Count
    $commaCount = ([regex]::Matches($firstLine, ',')).Count
    if ($semicolonCount -gt $commaCount) { return ";" }
    return ","
}

function Import-RoboRecipientCsv {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Lista de destinatarios nao encontrada: $Path"
    }

    $delimiter = Get-RoboCsvDelimiter -Path $Path
    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($raw)) { return @() }

    try {
        $rows = @($raw | ConvertFrom-Csv -Delimiter $delimiter)
    }
    catch {
        throw "Nao foi possivel ler a lista de destinatarios. Verifique o CSV. Erro: $($_.Exception.Message)"
    }

    if ($rows.Count -eq 0) { return @() }

    $props = @($rows[0].PSObject.Properties.Name)
    $required = @("loja", "nome", "telefone", "ativo")
    $missing = @()
    foreach ($requiredName in $required) {
        if (-not ($props -contains $requiredName)) { $missing += $requiredName }
    }

    if ($missing.Count -gt 0) {
        throw ("Cabecalho invalido em destinatarios.csv. Esperado: loja,nome,telefone,ativo,mensagem. Detectado: {0}. Separador detectado: '{1}'." -f ([string]::Join(", ", $props)), $delimiter)
    }

    return @($rows)
}

function Export-RoboRecipientCsv {
    param([array]$Rows, [string]$Path)
    $Rows | Select-Object loja,nome,telefone,ativo,mensagem |
        Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

function Get-RoboRecipients {
    param($Config)

    $path = Get-RoboRecipientPath $Config
    if (-not (Test-Path -LiteralPath $path)) {
        if (Test-Path -LiteralPath $RecipientExamplePath) {
            Copy-Item -LiteralPath $RecipientExamplePath -Destination $path -Force
        }
        else {
            throw "Lista de destinatarios nao encontrada: $path"
        }
    }

    $result = New-Object System.Collections.ArrayList
    $seen = @{}
    $rows = @(Import-RoboRecipientCsv -Path $path)

    foreach ($row in $rows) {
        if (-not (Test-RoboRecipientActive $row.ativo)) { continue }

        $store = ConvertTo-RoboStore $row.loja
        $name = ([string]$row.nome).Trim()
        $phone = ConvertTo-RoboPhone $row.telefone
        $customMessage = ""
        if ($row.PSObject.Properties.Name -contains "mensagem") {
            $customMessage = [string]$row.mensagem
        }

        if ([string]::IsNullOrWhiteSpace($store)) { continue }
        if ([string]::IsNullOrWhiteSpace($name)) { $name = $store }

        $key = ("{0}|{1}|{2}" -f $store, $name.ToUpperInvariant(), $phone)
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true

        [void]$result.Add([PSCustomObject]@{
            loja = $store
            nome = $name
            telefone = $phone
            ativo = $true
            mensagem = $customMessage
        })
    }

    return @($result)
}

function Get-RoboRecipientsForStore {
    param([string]$Store, $Config)
    $normalizedStore = ConvertTo-RoboStore $Store
    return @(Get-RoboRecipients $Config | Where-Object { $_.loja -eq $normalizedStore })
}

function Set-RoboRecipientPhone {
    param([string]$Store, [string]$Name, [string]$Phone, $Config)

    $path = Get-RoboRecipientPath $Config
    $normalizedStore = ConvertTo-RoboStore $Store
    $normalizedPhone = ConvertTo-RoboPhone $Phone
    if ($normalizedPhone.Length -lt 10) {
        throw "Telefone invalido para $Name."
    }

    $rows = @(Import-RoboRecipientCsv -Path $path)
    $changed = $false

    foreach ($row in $rows) {
        $rowStore = ConvertTo-RoboStore $row.loja
        $rowName = ([string]$row.nome).Trim()

        if (-not $changed -and $rowStore -eq $normalizedStore -and $rowName -eq $Name) {
            $row.telefone = $normalizedPhone
            $changed = $true
        }

        if (-not ($row.PSObject.Properties.Name -contains "mensagem")) {
            $row | Add-Member -NotePropertyName "mensagem" -NotePropertyValue "" -Force
        }
    }

    if (-not $changed) {
        $rows += [PSCustomObject]@{
            loja = $normalizedStore
            nome = $Name
            telefone = $normalizedPhone
            ativo = "sim"
            mensagem = ""
        }
    }

    Export-RoboRecipientCsv -Rows $rows -Path $path
    Write-RoboLog ("Lista de destinatarios atualizada: {0} | {1}" -f $Name, $normalizedPhone)
}
