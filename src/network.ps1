function Get-PdaStoreList {
    param([System.Net.WebSockets.ClientWebSocket]$Socket)

    $expression = @'
(() => {
  const norm = s => (s || '').replace(/\s+/g,' ').trim().toLowerCase();
  const selects = [...document.querySelectorAll('select')];
  const center = selects.find(s => [...s.options].some(o =>
    norm(o.textContent).includes('centerlar comercio de utilidades')
  ));

  if (!center) {
    return { ok:false, reason:'center-select-not-found', stores:[] };
  }

  const stores = [];

  for (const option of [...center.options]) {
    const text = (option.textContent || '').replace(/[\r\n\t|]+/g,' ').trim();
    const m = text.match(/^(\d+)-(\d+)\s*-\s*(.+)$/);
    if (!m) continue;

    const a = parseInt(m[1], 10);
    const b = parseInt(m[2], 10);
    if (!Number.isFinite(a) || a <= 0 || a !== b) continue;

    stores.push({
      numero: a,
      loja: 'ML' + String(a).padStart(2,'0'),
      textoPda: text,
      value: option.value
    });
  }

  stores.sort((x,y) => x.numero - y.numero);

  return {
    ok:true,
    count:stores.length,
    stores
  };
})()
'@

    $result = Invoke-CdpJsonExpression -Socket $Socket -Expression $expression

    if ($null -eq $result) {
        throw "O PDA nao retornou a lista de lojas."
    }

    if (-not [bool]$result.ok) {
        throw ("Nao foi possivel listar as lojas do PDA: " + [string]$result.reason)
    }

    $stores = @()

    foreach ($row in @($result.stores)) {
        $number = [int]$row.numero
        if ($number -le 0) { continue }

        $stores += [PSCustomObject]@{
            Numero = $number
            Loja = [string]$row.loja
            TextoPda = [string]$row.textoPda
            ValuePda = [string]$row.value
        }
    }

    if ($stores.Count -eq 0) {
        throw "Nenhuma loja valida foi identificada no dropdown Centro do PDA."
    }

    Write-RoboLog ("Dropdown Centro lido com sucesso: {0} lojas identificadas." -f $stores.Count)

    return @($stores | Sort-Object Numero)
}

function ConvertTo-RoboPrecosDateKey {
    param([Parameter(Mandatory = $true)][string]$DateText)

    $culture = [Globalization.CultureInfo]::InvariantCulture
    $date = [datetime]::ParseExact($DateText, "dd/MM/yyyy", $culture)
    return $date.ToString("yyyyMMdd")
}

function Get-RoboPrecosCollectionPaths {
    param(
        [string]$StartDate,
        [string]$EndDate
    )

    $startKey = ConvertTo-RoboPrecosDateKey $StartDate
    $endKey = ConvertTo-RoboPrecosDateKey $EndDate
    $key = $startKey + "_" + $endKey

    $checkpointDirectory = Join-Path $OutputPath "checkpoints"
    Ensure-RoboDirectory $checkpointDirectory

    return [PSCustomObject]@{
        Key = $key
        Checkpoint = Join-Path $checkpointDirectory ("auditoria_" + $key + ".csv")
        Consolidated = Join-Path $OutputPath ("auditoria_rede_" + $key + ".csv")
    }
}

function New-RoboPrecosRecordMap {
    param(
        [string]$CheckpointPath,
        [string]$StartDate,
        [string]$EndDate
    )

    $map = @{}

    if (-not (Test-Path -LiteralPath $CheckpointPath)) {
        return $map
    }

    try {
        foreach ($row in (Import-Csv -LiteralPath $CheckpointPath)) {
            if ([string]$row.DataInicio -ne $StartDate -or [string]$row.DataFim -ne $EndDate) {
                continue
            }

            $store = ConvertTo-RoboStore $row.Loja
            if ([string]::IsNullOrWhiteSpace($store)) { continue }

            $map[$store] = $row
        }

        Write-RoboLog ("Checkpoint carregado: {0} registros em {1}" -f $map.Count, $CheckpointPath)
    }
    catch {
        Write-RoboLog ("Nao foi possivel ler o checkpoint existente. Um novo sera criado. Detalhe: " + $_.Exception.Message) "AVISO"
        $map = @{}
    }

    return $map
}

function Save-RoboPrecosRecordMap {
    param(
        [hashtable]$Records,
        [string]$Path
    )

    $directory = Split-Path -Parent $Path
    Ensure-RoboDirectory $directory

    $temp = $Path + ".tmp"

    $ordered = @($Records.Values | Sort-Object {
        $n = 999999
        if ([string]$_.Loja -match '(\d+)') { $n = [int]$Matches[1] }
        $n
    })

    $ordered | Export-Csv -LiteralPath $temp -NoTypeInformation -Encoding UTF8
    Move-Item -LiteralPath $temp -Destination $Path -Force
}

function Test-RoboPrecosCompletedRecord {
    param($Record)

    if ($null -eq $Record) { return $false }
    if ([string]$Record.Status -ne "OK") { return $false }

    try {
        $ok = [int]$Record.Ok
        $div = [int]$Record.Divergente
        $sem = [int]$Record.SemEtiqueta
        $total = [int]$Record.Total
        return (($ok + $div + $sem) -eq $total)
    }
    catch {
        return $false
    }
}

function Invoke-RoboPrecosNetworkCollection {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        $Config,
        [string]$StartDate,
        [string]$EndDate,
        [int]$RetryPerStore = 3
    )

    $stores = @(Get-PdaStoreList -Socket $Socket)

    if ($stores.Count -lt 2) {
        throw ("Modo rede identificou apenas {0} loja(s). Coleta interrompida para evitar falso sucesso." -f $stores.Count)
    }

    Write-RoboLog ("Modo rede confirmado com {0} lojas. Primeira={1}; Ultima={2}" -f $stores.Count, $stores[0].Loja, $stores[$stores.Count - 1].Loja)

    $paths = Get-RoboPrecosCollectionPaths -StartDate $StartDate -EndDate $EndDate
    $records = New-RoboPrecosRecordMap -CheckpointPath $paths.Checkpoint -StartDate $StartDate -EndDate $EndDate

    Write-Host ""
    Write-Host ("LOJAS IDENTIFICADAS NO PDA: {0}" -f $stores.Count) -ForegroundColor Cyan
    Write-Host ("Periodo: {0} a {1}" -f $StartDate, $EndDate)
    Write-Host ("Checkpoint: {0}" -f $paths.Checkpoint)
    Write-Host ""

    $alreadyDone = 0
    foreach ($store in $stores) {
        if ($records.ContainsKey($store.Loja) -and (Test-RoboPrecosCompletedRecord $records[$store.Loja])) {
            $alreadyDone++
        }
    }

    if ($alreadyDone -gt 0) {
        Write-RoboLog ("Retomada detectada: {0} loja(s) ja concluidas serao preservadas." -f $alreadyDone)
    }

    $position = 0
    foreach ($store in $stores) {
        $position++
        $storeCode = [string]$store.Loja

        if ($records.ContainsKey($storeCode) -and (Test-RoboPrecosCompletedRecord $records[$storeCode])) {
            Write-Host ("[{0}/{1}] {2} - ja concluida, pulando." -f $position, $stores.Count, $storeCode) -ForegroundColor DarkGray
            continue
        }

        Write-Host ""
        Write-Host ("[{0}/{1}] Consultando {2}..." -f $position, $stores.Count, $storeCode) -ForegroundColor Yellow

        $success = $false
        $lastError = ""

        for ($attempt = 1; $attempt -le $RetryPerStore; $attempt++) {
            try {
                if ($attempt -gt 1) {
                    Write-RoboLog ("Nova tentativa {0}/{1} para {2}" -f $attempt, $RetryPerStore, $storeCode) "AVISO"
                    Ensure-PdaAuditPage -Socket $Socket -Config $Config
                }

                $result = Invoke-PdaAuditQuery -Socket $Socket -Config $Config -Store $store.Numero -StartDate $StartDate -EndDate $EndDate

                $records[$storeCode] = [PSCustomObject]@{
                    Loja = $result.Loja
                    NumeroPda = $store.Numero
                    CentroPda = $store.TextoPda
                    DataInicio = $StartDate
                    DataFim = $EndDate
                    Ok = $result.Ok
                    Divergente = $result.Divergente
                    SemEtiqueta = $result.SemEtiqueta
                    Total = $result.Total
                    Status = "OK"
                    Tentativas = $attempt
                    ColetadoEm = $result.ColetadoEm
                    Erro = ""
                }

                Save-RoboPrecosRecordMap -Records $records -Path $paths.Checkpoint
                $success = $true

                Write-Host ("  OK={0} | Divergente={1} | Sem etiqueta={2} | Total={3}" -f $result.Ok, $result.Divergente, $result.SemEtiqueta, $result.Total) -ForegroundColor Green
                break
            }
            catch {
                $lastError = $_.Exception.Message
                Write-RoboLog ("Falha em {0} tentativa {1}/{2}: {3}" -f $storeCode, $attempt, $RetryPerStore, $lastError) "AVISO"

                if ($attempt -lt $RetryPerStore) {
                    Start-Sleep -Seconds 2
                }
            }
        }

        if (-not $success) {
            $records[$storeCode] = [PSCustomObject]@{
                Loja = $storeCode
                NumeroPda = $store.Numero
                CentroPda = $store.TextoPda
                DataInicio = $StartDate
                DataFim = $EndDate
                Ok = ""
                Divergente = ""
                SemEtiqueta = ""
                Total = ""
                Status = "ERRO"
                Tentativas = $RetryPerStore
                ColetadoEm = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                Erro = $lastError
            }

            Save-RoboPrecosRecordMap -Records $records -Path $paths.Checkpoint
            Write-Host ("  ERRO apos {0} tentativa(s). Registrado no checkpoint." -f $RetryPerStore) -ForegroundColor Red
        }
    }

    Save-RoboPrecosRecordMap -Records $records -Path $paths.Checkpoint

    $allRows = @($records.Values | Sort-Object {
        $n = 999999
        if ([string]$_.Loja -match '(\d+)') { $n = [int]$Matches[1] }
        $n
    })

    $finalTemp = $paths.Consolidated + ".tmp"
    $allRows | Export-Csv -LiteralPath $finalTemp -NoTypeInformation -Encoding UTF8
    Move-Item -LiteralPath $finalTemp -Destination $paths.Consolidated -Force

    $okRows = @($allRows | Where-Object { [string]$_.Status -eq "OK" })
    $errorRows = @($allRows | Where-Object { [string]$_.Status -eq "ERRO" })

    Write-Host ""
    Write-Host "COLETA DA REDE CONCLUIDA" -ForegroundColor Cyan
    Write-Host ("Lojas identificadas : {0}" -f $stores.Count)
    Write-Host ("Lojas com sucesso   : {0}" -f $okRows.Count) -ForegroundColor Green
    Write-Host ("Lojas com erro      : {0}" -f $errorRows.Count)
    Write-Host ("Arquivo consolidado : {0}" -f $paths.Consolidated)
    Write-Host ("Checkpoint          : {0}" -f $paths.Checkpoint)
    Write-Host ""

    if ($errorRows.Count -gt 0) {
        Write-Host "As lojas com erro nao serao consideradas prontas para preencher a planilha." -ForegroundColor Yellow
        Write-Host "Ao executar novamente o mesmo periodo, o robo preservara as lojas OK e tentara apenas as pendentes." -ForegroundColor Yellow
    }

    return [PSCustomObject]@{
        Stores = $stores
        Rows = $allRows
        SuccessCount = $okRows.Count
        ErrorCount = $errorRows.Count
        CheckpointPath = $paths.Checkpoint
        ConsolidatedPath = $paths.Consolidated
    }
}
