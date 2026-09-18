$script:CdpCommandId = 0

function Wait-CdpEndpoint {
    param(
        [int]$Port,
        [int]$TimeoutSeconds = 30
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        try {
            $version = Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/json/version" -f $Port) -UseBasicParsing -TimeoutSec 2
            if ($version) { return }
        }
        catch {}
        Start-Sleep -Milliseconds 500
    } while ((Get-Date) -lt $deadline)

    throw "Chrome DevTools nao respondeu na porta $Port."
}

function Get-CdpPageTarget {
    param(
        [int]$Port,
        [string]$UrlContains = ""
    )

    $targets = @(Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/json" -f $Port) -UseBasicParsing -TimeoutSec 5)
    $pages = @($targets | Where-Object { $_.type -eq "page" })

    if (-not [string]::IsNullOrWhiteSpace($UrlContains)) {
        $match = $pages | Where-Object { ([string]$_.url) -like ("*" + $UrlContains + "*") } | Select-Object -First 1
        if ($match) { return $match }
    }

    $target = $pages | Select-Object -First 1
    if (-not $target) {
        throw "Nenhuma aba do Chrome disponivel para automacao."
    }
    return $target
}

function New-CdpPageTarget {
    param(
        [int]$Port,
        [Parameter(Mandatory = $true)][string]$Url
    )

    $encoded = [Uri]::EscapeDataString($Url)
    $endpoint = "http://127.0.0.1:{0}/json/new?{1}" -f $Port, $encoded

    try {
        $target = Invoke-RestMethod -Method Put -Uri $endpoint -UseBasicParsing -TimeoutSec 10
    }
    catch {
        throw ("Nao foi possivel criar uma aba dedicada do PDA no Chrome DevTools. Detalhe: " + $_.Exception.Message)
    }

    if (-not $target -or [string]::IsNullOrWhiteSpace([string]$target.webSocketDebuggerUrl)) {
        throw "Chrome DevTools criou a aba, mas nao retornou webSocketDebuggerUrl."
    }

    return $target
}

function Connect-CdpPage {
    param(
        [int]$Port,
        [string]$UrlContains = "",
        [int]$TimeoutSeconds = 30
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $target = $null
    $lastPages = @()

    do {
        $targets = @(Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/json" -f $Port) -UseBasicParsing -TimeoutSec 5)
        $pages = @($targets | Where-Object { $_.type -eq "page" })
        $lastPages = $pages

        if (-not [string]::IsNullOrWhiteSpace($UrlContains)) {
            $target = @($pages | Where-Object {
                ([string]$_.url) -like ("*" + $UrlContains + "*")
            }) | Select-Object -First 1

            if ($target) { break }

            # O Chrome pode publicar extensoes/background antes da aba PDA.
            # Quando um dominio foi solicitado, nunca caimos no primeiro target.
            Start-Sleep -Milliseconds 300
            continue
        }

        $target = @($pages | Where-Object {
            -not ([string]$_.url).StartsWith("chrome-extension://")
        }) | Select-Object -First 1

        if (-not $target) {
            $target = $pages | Select-Object -First 1
        }

        if ($target) { break }
        Start-Sleep -Milliseconds 300
    } while ((Get-Date) -lt $deadline)

    if (-not $target) {
        $seen = @($lastPages | ForEach-Object { [string]$_.url }) -join " | "
        if (-not [string]::IsNullOrWhiteSpace($UrlContains)) {
            throw ("Nenhuma aba do Chrome correspondente a '" + $UrlContains + "' apareceu dentro do tempo limite. Targets vistos: " + $seen)
        }
        throw ("Nenhuma aba do Chrome disponivel para automacao. Targets vistos: " + $seen)
    }

    if ($target -is [System.Array]) {
        $target = @($target)[0]
    }

    $wsValues = @($target.webSocketDebuggerUrl)
    if ($wsValues.Count -eq 0 -or [string]::IsNullOrWhiteSpace([string]$wsValues[0])) {
        throw "A aba do Chrome nao expos webSocketDebuggerUrl."
    }

    $wsUrl = [string]$wsValues[0]
    $wsUrl = $wsUrl -replace 'ws://localhost:', 'ws://127.0.0.1:'
    $wsUrl = $wsUrl -replace 'ws://\[::1\]:', 'ws://127.0.0.1:'

    Write-RoboLog ("Target Chrome selecionado: " + [string]$target.url)
    Write-Host ("Conectando ao Chrome local: " + $wsUrl) -ForegroundColor DarkGray

    $socket = New-Object System.Net.WebSockets.ClientWebSocket

    try {
        $socket.Options.Proxy = New-Object System.Net.WebProxy
    }
    catch {
        try { $socket.Options.Proxy = $null } catch {}
    }

    $uri = New-Object System.Uri($wsUrl)

    try {
        [void]$socket.ConnectAsync($uri, [Threading.CancellationToken]::None).GetAwaiter().GetResult()
    }
    catch {
        $detail = $_.Exception.Message
        $inner = $_.Exception.InnerException
        while ($inner) {
            $detail += " -> " + $inner.Message
            $inner = $inner.InnerException
        }
        try { $socket.Dispose() } catch {}
        throw ("Falha na conexao local com o Chrome DevTools. " +
               "O navegador abriu, mas o PowerShell nao conseguiu conectar ao WebSocket local. " +
               "URL: " + $wsUrl + " | Detalhe: " + $detail)
    }

    return $socket
}

function Receive-CdpMessage {
    param([System.Net.WebSockets.ClientWebSocket]$Socket)

    $buffer = New-Object byte[] 65536
    $stream = New-Object System.IO.MemoryStream

    try {
        do {
            $segment = [System.ArraySegment[byte]]::new($buffer)
            $result = $Socket.ReceiveAsync($segment, [Threading.CancellationToken]::None).GetAwaiter().GetResult()

            if ($result.MessageType -eq [System.Net.WebSockets.WebSocketMessageType]::Close) {
                throw "Conexao CDP encerrada pelo Chrome."
            }

            if ($result.Count -gt 0) {
                $stream.Write($buffer, 0, $result.Count)
            }
        } while (-not $result.EndOfMessage)

        return [Text.Encoding]::UTF8.GetString($stream.ToArray())
    }
    finally {
        $stream.Dispose()
    }
}

function Invoke-CdpCommand {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [Parameter(Mandatory = $true)][string]$Method,
        [hashtable]$Params = @{}
    )

    $script:CdpCommandId++
    $id = $script:CdpCommandId
    $payload = @{
        id = $id
        method = $Method
        params = $Params
    } | ConvertTo-Json -Depth 20 -Compress

    $bytes = [Text.Encoding]::UTF8.GetBytes($payload)
    $segment = [System.ArraySegment[byte]]::new($bytes)
    [void]$Socket.SendAsync(
        $segment,
        [System.Net.WebSockets.WebSocketMessageType]::Text,
        $true,
        [Threading.CancellationToken]::None
    ).GetAwaiter().GetResult()

    while ($true) {
        $raw = Receive-CdpMessage -Socket $Socket
        $message = $raw | ConvertFrom-Json

        if (($message.PSObject.Properties.Name -contains "id") -and ([int]$message.id -eq $id)) {
            if ($message.PSObject.Properties.Name -contains "error") {
                throw ("CDP {0}: {1}" -f $Method, ($message.error | ConvertTo-Json -Compress))
            }
            return $message.result
        }
    }
}

function Invoke-CdpExpression {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [Parameter(Mandatory = $true)][string]$Expression
    )

    $response = Invoke-CdpCommand -Socket $Socket -Method "Runtime.evaluate" -Params @{
        expression = $Expression
        returnByValue = $true
        awaitPromise = $true
        userGesture = $true
    }

    if ($null -eq $response) {
        return $null
    }

    $exceptionProp = $response.PSObject.Properties["exceptionDetails"]
    if ($exceptionProp -and $exceptionProp.Value) {
        throw ("Erro JavaScript no navegador: " + ($exceptionProp.Value | ConvertTo-Json -Depth 12 -Compress))
    }

    $resultProp = $response.PSObject.Properties["result"]
    if (-not $resultProp) {
        throw ("Resposta inesperada do Runtime.evaluate: " + ($response | ConvertTo-Json -Depth 12 -Compress))
    }

    $remote = $resultProp.Value
    if ($null -eq $remote) {
        return $null
    }

    $valueProp = $remote.PSObject.Properties["value"]
    if ($valueProp) {
        return $valueProp.Value
    }

    $unserializableProp = $remote.PSObject.Properties["unserializableValue"]
    if ($unserializableProp) {
        return $unserializableProp.Value
    }

    $descriptionProp = $remote.PSObject.Properties["description"]
    if ($descriptionProp -and -not [string]::IsNullOrWhiteSpace([string]$descriptionProp.Value)) {
        return [string]$descriptionProp.Value
    }

    return $null
}

function Invoke-CdpJsonExpression {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [Parameter(Mandatory = $true)][string]$Expression
    )

    $wrapped = "JSON.stringify(" + $Expression + ")"
    $json = Invoke-CdpExpression -Socket $Socket -Expression $wrapped

    if ([string]::IsNullOrWhiteSpace([string]$json)) {
        return $null
    }

    try {
        return ([string]$json | ConvertFrom-Json)
    }
    catch {
        throw ("Chrome retornou JSON invalido ao ler a pagina: " + [string]$json)
    }
}

function Navigate-Cdp {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [Parameter(Mandatory = $true)][string]$Url,
        [int]$TimeoutSeconds = 45
    )

    [void](Invoke-CdpCommand -Socket $Socket -Method "Page.enable")
    [void](Invoke-CdpCommand -Socket $Socket -Method "Runtime.enable")

    $result = Invoke-CdpCommand -Socket $Socket -Method "Page.navigate" -Params @{ url = $Url }
    if ($result -and ($result.PSObject.Properties.Name -contains "errorText") -and -not [string]::IsNullOrWhiteSpace([string]$result.errorText)) {
        throw ("Chrome nao conseguiu navegar para " + $Url + ": " + [string]$result.errorText)
    }

    # O PDA pode nunca chegar a readyState=complete por manter requisicoes pendentes.
    # A validacao da tela e feita depois, por elementos reais da pagina.
    Start-Sleep -Milliseconds 1200
}

function Close-CdpPage {
    param([System.Net.WebSockets.ClientWebSocket]$Socket)

    if ($Socket) {
        try {
            if ($Socket.State -eq [System.Net.WebSockets.WebSocketState]::Open) {
                [void]$Socket.CloseAsync(
                    [System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure,
                    "fim",
                    [Threading.CancellationToken]::None
                ).GetAwaiter().GetResult()
            }
        }
        catch {}
        try { $Socket.Dispose() } catch {}
    }
}
