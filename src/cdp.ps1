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
            if ($version) { return $true }
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

function Connect-CdpPage {
    param(
        [int]$Port,
        [string]$UrlContains = ""
    )

    $target = Get-CdpPageTarget -Port $Port -UrlContains $UrlContains
    if (-not $target.webSocketDebuggerUrl) {
        throw "A aba do Chrome nao expos webSocketDebuggerUrl."
    }

    $socket = New-Object System.Net.WebSockets.ClientWebSocket
    $uri = New-Object System.Uri([string]$target.webSocketDebuggerUrl)
    $socket.ConnectAsync($uri, [Threading.CancellationToken]::None).GetAwaiter().GetResult()
    return $socket
}

function Receive-CdpMessage {
    param([System.Net.WebSockets.ClientWebSocket]$Socket)

    $buffer = New-Object byte[] 65536
    $stream = New-Object System.IO.MemoryStream

    try {
        do {
            $segment = New-Object System.ArraySegment[byte] -ArgumentList (, $buffer)
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
    $segment = New-Object System.ArraySegment[byte] -ArgumentList (, $bytes)
    $Socket.SendAsync(
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

    $result = Invoke-CdpCommand -Socket $Socket -Method "Runtime.evaluate" -Params @{
        expression = $Expression
        returnByValue = $true
        awaitPromise = $true
        userGesture = $true
    }

    if ($result.exceptionDetails) {
        throw ("Erro JavaScript no navegador: " + ($result.exceptionDetails | ConvertTo-Json -Depth 8 -Compress))
    }

    if ($result.result -and ($result.result.PSObject.Properties.Name -contains "value")) {
        return $result.result.value
    }

    return $null
}

function Navigate-Cdp {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [Parameter(Mandatory = $true)][string]$Url,
        [int]$TimeoutSeconds = 45
    )

    [void](Invoke-CdpCommand -Socket $Socket -Method "Page.enable")
    [void](Invoke-CdpCommand -Socket $Socket -Method "Runtime.enable")
    [void](Invoke-CdpCommand -Socket $Socket -Method "Page.navigate" -Params @{ url = $Url })

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        Start-Sleep -Milliseconds 300
        try {
            $state = Invoke-CdpExpression -Socket $Socket -Expression "document.readyState"
            if ($state -eq "complete") { return }
        }
        catch {}
    } while ((Get-Date) -lt $deadline)

    throw "Timeout aguardando carregamento da pagina: $Url"
}

function Close-CdpPage {
    param([System.Net.WebSockets.ClientWebSocket]$Socket)

    if ($Socket) {
        try {
            if ($Socket.State -eq [System.Net.WebSockets.WebSocketState]::Open) {
                $Socket.CloseAsync(
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
