$script:RoboPrecosBiDefaultConfig = [ordered]@{
    loginUrl = "https://app.powerbi.com/"
    summaryUrl = "https://app.powerbi.com/groups/cb155eaa-6b0a-4190-ae3f-5447f2fb3b58/reports/13902a52-6ab0-4c33-b73c-5352e1c490df/5e1f3f49a475efe362c0?experience=power-bi"
    historicalUrl = "https://app.powerbi.com/groups/cb155eaa-6b0a-4190-ae3f-5447f2fb3b58/reports/13902a52-6ab0-4c33-b73c-5352e1c490df/d42a8bc5428005c25ae7?experience=power-bi"
    debugPort = 9224
    profileDirectory = "output\\chrome_bi"
    credentialFile = "data\\bi_credential.json"
    pageLoadTimeoutSeconds = 120
    loginTimeoutSeconds = 180
}

function Get-RoboPrecosBiConfig {
    param($Config)

    $result = [ordered]@{}
    foreach ($key in $script:RoboPrecosBiDefaultConfig.Keys) {
        $result[$key] = $script:RoboPrecosBiDefaultConfig[$key]
    }

    if ($Config -and ($Config.PSObject.Properties.Name -contains "bi") -and $Config.bi) {
        foreach ($key in $script:RoboPrecosBiDefaultConfig.Keys) {
            if ($Config.bi.PSObject.Properties.Name -contains $key) {
                $value = $Config.bi.PSObject.Properties[$key].Value
                if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
                    $result[$key] = $value
                }
            }
        }
    }

    return [PSCustomObject]$result
}

function ConvertTo-RoboPrecosNormalizedText {
    param($Value)

    if ($null -eq $Value) { return "" }

    $text = ([string]$Value).Trim().ToUpperInvariant().Normalize([Text.NormalizationForm]::FormD)
    $builder = New-Object Text.StringBuilder

    foreach ($char in $text.ToCharArray()) {
        $category = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($char)
        if ($category -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$builder.Append($char)
        }
    }

    return (($builder.ToString().Normalize([Text.NormalizationForm]::FormC)) -replace '\s+', ' ').Trim()
}

function ConvertFrom-RoboPrecosBiDecimal {
    param($Value)

    if ($null -eq $Value) { return $null }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }

    $text = $text -replace '[Rr]\$', ''
    $text = $text -replace '\s', ''

    $culture = [Globalization.CultureInfo]::GetCultureInfo("pt-BR")
    $number = 0.0

    if ([double]::TryParse($text, [Globalization.NumberStyles]::Number, $culture, [ref]$number)) {
        return [double]$number
    }

    $invariant = $text -replace '\.', ''
    $invariant = $invariant -replace ',', '.'

    if ([double]::TryParse($invariant, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return [double]$number
    }

    return $null
}

function ConvertFrom-RoboPrecosBiInteger {
    param($Value)

    $number = ConvertFrom-RoboPrecosBiDecimal $Value
    if ($null -eq $number) { return $null }

    return [int][Math]::Round([double]$number, 0, [MidpointRounding]::AwayFromZero)
}

function Set-RoboPrecosBiCredential {
    param($Config)

    $bi = Get-RoboPrecosBiConfig -Config $Config
    $credentialPath = Resolve-RoboPrecosPath ([string]$bi.credentialFile)
    $credentialDirectory = Split-Path -Parent $credentialPath
    Ensure-RoboDirectory $credentialDirectory

    Write-Host ""
    Write-Host "CREDENCIAL POWER BI - armazenamento local protegido pelo Windows" -ForegroundColor Cyan
    Write-Host "Ela sera solicitada pelo navegador somente se a sessao do BI nao estiver autenticada." -ForegroundColor DarkGray
    $username = Read-Host "Usuario / e-mail do Power BI"
    $securePassword = Read-Host "Senha do Power BI" -AsSecureString

    if ([string]::IsNullOrWhiteSpace($username)) {
        throw "Usuario do Power BI nao informado."
    }

    $payload = [PSCustomObject]@{
        username = $username.Trim()
        password = ($securePassword | ConvertFrom-SecureString)
        createdAt = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    }

    $payload | ConvertTo-Json | Set-Content -LiteralPath $credentialPath -Encoding UTF8
    Write-RoboLog ("Credencial Power BI protegida criada em " + $credentialPath)
}

function Get-RoboPrecosBiCredential {
    param($Config)

    $bi = Get-RoboPrecosBiConfig -Config $Config
    $credentialPath = Resolve-RoboPrecosPath ([string]$bi.credentialFile)

    if (-not (Test-Path -LiteralPath $credentialPath)) {
        Set-RoboPrecosBiCredential -Config $Config
    }

    $payload = Get-Content -LiteralPath $credentialPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $secure = ConvertTo-SecureString ([string]$payload.password)

    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try {
        $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }

    return [PSCustomObject]@{
        Username = [string]$payload.username
        Password = $plain
    }
}

function Connect-RoboPrecosBiTarget {
    param(
        [int]$Port,
        [Parameter(Mandatory = $true)][string]$Url
    )

    $target = New-CdpPageTarget -Port $Port -Url $Url
    $wsUrl = [string]$target.webSocketDebuggerUrl
    $wsUrl = $wsUrl -replace 'ws://localhost:', 'ws://127.0.0.1:'
    $wsUrl = $wsUrl -replace 'ws://\[::1\]:', 'ws://127.0.0.1:'

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
        try { $socket.Dispose() } catch {}
        throw ("Falha conectando a aba Power BI pelo Chrome DevTools: " + $_.Exception.Message)
    }

    $probe = Invoke-CdpExpression -Socket $socket -Expression "'ROBO_BI_CDP_OK'"
    if ([string]$probe -ne "ROBO_BI_CDP_OK") {
        try { $socket.Dispose() } catch {}
        throw "Canal CDP do Power BI nao respondeu ao teste Runtime.evaluate."
    }

    return $socket
}

function Start-RoboPrecosBiBrowser {
    param($Config)

    $bi = Get-RoboPrecosBiConfig -Config $Config
    $chromePath = Get-RoboChrome
    if (-not $chromePath) {
        throw "Google Chrome nao encontrado."
    }

    $port = [int]$bi.debugPort
    $profilePath = Resolve-RoboPrecosPath ([string]$bi.profileDirectory)
    Ensure-RoboDirectory $profilePath

    $endpointReady = $false
    try {
        $null = Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/json/version" -f $port) -UseBasicParsing -TimeoutSec 1
        $endpointReady = $true
    }
    catch {}

    if (-not $endpointReady) {
        Write-RoboLog ("Abrindo Chrome dedicado ao Power BI na porta CDP " + $port)
        $arguments = @(
            "--remote-debugging-port=$port",
            "--remote-debugging-address=127.0.0.1",
            "--remote-allow-origins=*",
            ("--user-data-dir=" + '"' + $profilePath + '"'),
            "--no-first-run",
            "--no-default-browser-check",
            "--new-window",
            "--start-maximized",
            "about:blank"
        )
        Start-Process -FilePath $chromePath -ArgumentList $arguments | Out-Null
    }
    else {
        Write-RoboLog "Chrome dedicado ao Power BI ja esta em execucao."
    }

    Wait-CdpEndpoint -Port $port -TimeoutSeconds 30
    return (Connect-RoboPrecosBiTarget -Port $port -Url ([string]$bi.loginUrl))
}

function Get-RoboPrecosBiPageState {
    param([System.Net.WebSockets.ClientWebSocket]$Socket)

    $expression = @'
(() => {
  const norm = s => (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ').trim().toLowerCase();
  const visible = e => !!(e && (e.offsetWidth || e.offsetHeight || e.getClientRects().length));
  const body = document.body ? document.body.innerText : '';
  const text = norm(body);
  const host = (location.hostname || '').toLowerCase();
  const href = location.href || '';
  const path = (location.pathname || '').toLowerCase();

  const inputs = [...document.querySelectorAll('input')].filter(visible);
  const password = inputs.find(e => (e.type || '').toLowerCase() === 'password' || (e.name || '').toLowerCase() === 'passwd');
  const email = inputs.find(e => {
    const t = (e.type || '').toLowerCase();
    const n = (e.name || '').toLowerCase();
    const p = norm(e.getAttribute('placeholder') || '');
    return t === 'email' || n === 'loginfmt' || p.includes('email') || p.includes('e-mail');
  });
  const genericText = inputs.find(e => {
    const t = (e.type || '').toLowerCase();
    return t === 'text' || t === '' || t === 'email';
  });

  const microsoft = host.includes('login.microsoftonline.com') || host.includes('login.live.com');
  const powerbi = host.includes('app.powerbi.com');
  const singleSignOn = powerbi && path.includes('/singlesignon');

  let kind = 'OTHER';

  if (singleSignOn || (powerbi && (
      text.includes('insira seu endereco de email corporativo') ||
      text.includes('inserir endereco de email') ||
      text.includes('email corporativo')
  ))) {
    kind = 'POWERBI_EMAIL';
  }
  else if (microsoft && (
      !!password ||
      text.includes('insira a senha') ||
      text.includes('digite a senha') ||
      text.includes('enter password') ||
      text.includes('password')
  )) {
    kind = 'MICROSOFT_PASSWORD';
  }
  else if (microsoft && (
      text.includes('continuar conectado') ||
      text.includes('permanecer conectado') ||
      text.includes('manter conectado') ||
      text.includes('stay signed in')
  )) {
    kind = 'MICROSOFT_STAY';
  }
  else if (microsoft && (
      !!email ||
      !!genericText ||
      text.includes('insira seu email') ||
      text.includes('insira seu e-mail') ||
      text.includes('enter email') ||
      text.includes('entrar em sua conta') ||
      text.includes('sign in')
  )) {
    kind = 'MICROSOFT_EMAIL';
  }
  else if (powerbi && !singleSignOn) {
    kind = 'AUTHENTICATED';
  }

  return {
    kind,
    host,
    href,
    path,
    title: document.title || '',
    ready: document.readyState || '',
    body: body.slice(0,5000),
    hasPassword: !!password,
    hasEmail: !!email,
    inputCount: inputs.length
  };
})()
'@

    return Invoke-CdpJsonExpression -Socket $Socket -Expression $expression
}

function Get-RoboPrecosBiFlatDomNodes {
    param([System.Net.WebSockets.ClientWebSocket]$Socket)

    [void](Invoke-CdpCommand -Socket $Socket -Method "DOM.enable")
    [void](Invoke-CdpCommand -Socket $Socket -Method "Runtime.enable")

    $result = Invoke-CdpCommand -Socket $Socket -Method "DOM.getFlattenedDocument" -Params @{
        depth = -1
        pierce = $true
    }

    if (-not $result -or -not ($result.PSObject.Properties.Name -contains "nodes")) {
        return @()
    }

    return @($result.nodes)
}

function Resolve-RoboPrecosBiDomObject {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [Parameter(Mandatory = $true)]$Node
    )

    $backendProperty = $Node.PSObject.Properties["backendNodeId"]
    if (-not $backendProperty) {
        return $null
    }

    try {
        $resolved = Invoke-CdpCommand -Socket $Socket -Method "DOM.resolveNode" -Params @{
            backendNodeId = [int]$backendProperty.Value
        }

        if ($resolved -and ($resolved.PSObject.Properties.Name -contains "object")) {
            return $resolved.object
        }
    }
    catch {}

    return $null
}

function Invoke-RoboPrecosBiObjectFunction {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [Parameter(Mandatory = $true)]$RemoteObject,
        [Parameter(Mandatory = $true)][string]$FunctionDeclaration,
        [array]$Arguments = @()
    )

    $objectIdProperty = $RemoteObject.PSObject.Properties["objectId"]
    if (-not $objectIdProperty -or [string]::IsNullOrWhiteSpace([string]$objectIdProperty.Value)) {
        return $null
    }

    $params = @{
        objectId = [string]$objectIdProperty.Value
        functionDeclaration = $FunctionDeclaration
        returnByValue = $true
        awaitPromise = $true
        userGesture = $true
    }

    if ($Arguments.Count -gt 0) {
        $params.arguments = @($Arguments | ForEach-Object { @{ value = $_ } })
    }

    $response = Invoke-CdpCommand -Socket $Socket -Method "Runtime.callFunctionOn" -Params $params
    if (-not $response) { return $null }

    $exception = $response.PSObject.Properties["exceptionDetails"]
    if ($exception -and $exception.Value) {
        throw ("Erro JavaScript no controle Power BI: " + ($exception.Value | ConvertTo-Json -Depth 10 -Compress))
    }

    $resultProperty = $response.PSObject.Properties["result"]
    if (-not $resultProperty -or -not $resultProperty.Value) {
        return $null
    }

    $remote = $resultProperty.Value
    $valueProperty = $remote.PSObject.Properties["value"]
    if ($valueProperty) {
        return $valueProperty.Value
    }

    return $null
}

function Get-RoboPrecosBiNodeInfo {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [Parameter(Mandatory = $true)]$Node
    )

    $object = Resolve-RoboPrecosBiDomObject -Socket $Socket -Node $Node
    if (-not $object) { return $null }

    $json = Invoke-RoboPrecosBiObjectFunction -Socket $Socket -RemoteObject $object -FunctionDeclaration @'
function() {
  try {
    const r = this.getBoundingClientRect ? this.getBoundingClientRect() : {width:0,height:0};
    const s = window.getComputedStyle ? getComputedStyle(this) : null;
    const visible = !!(
      r &&
      r.width > 1 &&
      r.height > 1 &&
      (!s || (s.display !== 'none' && s.visibility !== 'hidden' && Number(s.opacity || 1) > 0))
    );
    return JSON.stringify({
      visible,
      tag: (this.tagName || '').toLowerCase(),
      type: (this.type || '').toLowerCase(),
      name: this.name || '',
      id: this.id || '',
      placeholder: this.placeholder || '',
      aria: this.getAttribute ? (this.getAttribute('aria-label') || '') : '',
      title: this.title || '',
      value: this.value == null ? '' : String(this.value),
      text: (this.innerText || this.textContent || '').replace(/\s+/g,' ').trim(),
      disabled: !!this.disabled,
      readOnly: !!this.readOnly
    });
  } catch (e) {
    return JSON.stringify({visible:false,error:String(e)});
  }
}
'@

    if ([string]::IsNullOrWhiteSpace([string]$json)) { return $null }

    try {
        $info = ([string]$json | ConvertFrom-Json)
        $info | Add-Member -NotePropertyName RemoteObject -NotePropertyValue $object -Force
        $info | Add-Member -NotePropertyName SourceNode -NotePropertyValue $Node -Force
        return $info
    }
    catch {
        return $null
    }
}

function Find-RoboPrecosBiVisibleInput {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [ValidateSet("EMAIL","PASSWORD")][string]$Kind
    )

    $nodes = @(Get-RoboPrecosBiFlatDomNodes -Socket $Socket)
    $best = $null
    $bestScore = -1

    foreach ($node in $nodes) {
        $nodeNameProperty = $node.PSObject.Properties["nodeName"]
        if (-not $nodeNameProperty -or ([string]$nodeNameProperty.Value).ToUpperInvariant() -ne "INPUT") {
            continue
        }

        $info = Get-RoboPrecosBiNodeInfo -Socket $Socket -Node $node
        if (-not $info -or -not [bool]$info.visible -or [bool]$info.disabled -or [bool]$info.readOnly) {
            continue
        }

        $type = ConvertTo-RoboPrecosNormalizedText ([string]$info.type)
        $name = ConvertTo-RoboPrecosNormalizedText ([string]$info.name)
        $id = ConvertTo-RoboPrecosNormalizedText ([string]$info.id)
        $placeholder = ConvertTo-RoboPrecosNormalizedText ([string]$info.placeholder)
        $aria = ConvertTo-RoboPrecosNormalizedText ([string]$info.aria)
        $score = 0

        if ($Kind -eq "EMAIL") {
            if ($type -eq "EMAIL") { $score += 200 }
            if ($type -eq "TEXT" -or [string]::IsNullOrWhiteSpace($type)) { $score += 25 }
            if ($name.Contains("EMAIL") -or $name.Contains("LOGIN")) { $score += 120 }
            if ($id.Contains("EMAIL") -or $id.Contains("LOGIN")) { $score += 100 }
            if ($placeholder.Contains("EMAIL")) { $score += 180 }
            if ($aria.Contains("EMAIL")) { $score += 180 }
            if ($type -eq "PASSWORD" -or $type -eq "HIDDEN") { $score = -1000 }
        }
        else {
            if ($type -eq "PASSWORD") { $score += 300 }
            if ($name.Contains("PASSWD") -or $name.Contains("PASSWORD") -or $name.Contains("SENHA")) { $score += 150 }
            if ($id.Contains("PASSWD") -or $id.Contains("PASSWORD") -or $id.Contains("SENHA")) { $score += 120 }
            if ($placeholder.Contains("PASSWORD") -or $placeholder.Contains("SENHA")) { $score += 180 }
            if ($aria.Contains("PASSWORD") -or $aria.Contains("SENHA")) { $score += 180 }
        }

        if ($score -gt $bestScore) {
            $bestScore = $score
            $best = $info
        }
    }

    if ($bestScore -gt 0) { return $best }
    return $null
}

function Wait-RoboPrecosBiVisibleInput {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [ValidateSet("EMAIL","PASSWORD")][string]$Kind,
        [int]$TimeoutSeconds = 20
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)

    do {
        $info = Find-RoboPrecosBiVisibleInput -Socket $Socket -Kind $Kind
        if ($info) { return $info }
        Start-Sleep -Milliseconds 300
    } while ((Get-Date) -lt $deadline)

    return $null
}

function Find-RoboPrecosBiVisibleButton {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [Parameter(Mandatory = $true)][string[]]$Labels
    )

    $nodes = @(Get-RoboPrecosBiFlatDomNodes -Socket $Socket)
    $normalizedLabels = @($Labels | ForEach-Object { ConvertTo-RoboPrecosNormalizedText $_ })
    $best = $null
    $bestScore = -1

    foreach ($node in $nodes) {
        $nodeNameProperty = $node.PSObject.Properties["nodeName"]
        if (-not $nodeNameProperty) { continue }

        $nodeName = ([string]$nodeNameProperty.Value).ToUpperInvariant()
        if ($nodeName -notin @("BUTTON","INPUT")) { continue }

        $info = Get-RoboPrecosBiNodeInfo -Socket $Socket -Node $node
        if (-not $info -or -not [bool]$info.visible -or [bool]$info.disabled) {
            continue
        }

        if ($nodeName -eq "INPUT" -and ([string]$info.type) -notin @("submit","button")) {
            continue
        }

        $search = ConvertTo-RoboPrecosNormalizedText (
            ([string]$info.text) + " " +
            ([string]$info.value) + " " +
            ([string]$info.aria) + " " +
            ([string]$info.title) + " " +
            ([string]$info.name) + " " +
            ([string]$info.id)
        )

        $score = 0
        foreach ($label in $normalizedLabels) {
            if ([string]::IsNullOrWhiteSpace($label)) { continue }
            if ($search -eq $label) { $score = [Math]::Max($score, 300) }
            elseif ($search.Contains($label)) { $score = [Math]::Max($score, 150) }
        }

        if ($score -gt $bestScore) {
            $bestScore = $score
            $best = $info
        }
    }

    if ($bestScore -gt 0) { return $best }
    return $null
}

function Set-RoboPrecosBiInputValue {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [Parameter(Mandatory = $true)]$InputInfo,
        [Parameter(Mandatory = $true)][string]$Text
    )

    $result = Invoke-RoboPrecosBiObjectFunction -Socket $Socket -RemoteObject $InputInfo.RemoteObject -Arguments @($Text) -FunctionDeclaration @'
function(value) {
  try {
    let proto = this;
    let desc = null;
    while (proto && !desc) {
      proto = Object.getPrototypeOf(proto);
      if (proto) desc = Object.getOwnPropertyDescriptor(proto, 'value');
    }

    if (desc && desc.set) {
      desc.set.call(this, value);
    } else {
      this.value = value;
    }

    this.dispatchEvent(new Event('input', {bubbles:true, composed:true}));
    this.dispatchEvent(new Event('change', {bubbles:true, composed:true}));
    this.dispatchEvent(new KeyboardEvent('keyup', {bubbles:true, composed:true, key:'Unidentified'}));

    return String(this.value == null ? '' : this.value);
  } catch (e) {
    return 'ERROR:' + String(e);
  }
}
'@

    if ([string]$result -like "ERROR:*") {
        throw ("Falha preenchendo controle Power BI: " + [string]$result)
    }

    return [string]$result
}

function Click-RoboPrecosBiButton {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [Parameter(Mandatory = $true)]$ButtonInfo
    )

    $result = Invoke-RoboPrecosBiObjectFunction -Socket $Socket -RemoteObject $ButtonInfo.RemoteObject -FunctionDeclaration @'
function() {
  try {
    this.click();
    return true;
  } catch (e) {
    return false;
  }
}
'@

    return [bool]$result
}

function Wait-RoboPrecosBiVisibleButton {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [Parameter(Mandatory = $true)][string[]]$Labels,
        [int]$TimeoutSeconds = 10
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $button = Find-RoboPrecosBiVisibleButton -Socket $Socket -Labels $Labels
        if ($button) { return $button }
        Start-Sleep -Milliseconds 250
    } while ((Get-Date) -lt $deadline)

    return $null
}

function Invoke-RoboPrecosBiLoginStep {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [Parameter(Mandatory = $true)]$State,
        [Parameter(Mandatory = $true)]$Credential
    )

    $kind = [string]$State.kind

    if ($kind -eq "POWERBI_EMAIL") {
        $email = Wait-RoboPrecosBiVisibleInput -Socket $Socket -Kind "EMAIL" -TimeoutSeconds 20
        if (-not $email) {
            return "ERROR:POWERBI_VISIBLE_EMAIL_NOT_FOUND"
        }

        $actual = Set-RoboPrecosBiInputValue -Socket $Socket -InputInfo $email -Text ([string]$Credential.Username)
        if ($actual -ne [string]$Credential.Username) {
            return "ERROR:POWERBI_EMAIL_VALUE_NOT_APPLIED"
        }

        Start-Sleep -Milliseconds 300
        $send = Wait-RoboPrecosBiVisibleButton -Socket $Socket -Labels @("ENVIAR","SUBMIT","CONTINUAR","CONTINUE") -TimeoutSeconds 5
        if (-not $send -or -not (Click-RoboPrecosBiButton -Socket $Socket -ButtonInfo $send)) {
            return "ERROR:POWERBI_SEND_BUTTON_NOT_CLICKED"
        }

        return "POWERBI_EMAIL_SUBMITTED_DIRECT"
    }

    if ($kind -eq "MICROSOFT_EMAIL") {
        $email = Wait-RoboPrecosBiVisibleInput -Socket $Socket -Kind "EMAIL" -TimeoutSeconds 8
        if ($email) {
            $actual = Set-RoboPrecosBiInputValue -Socket $Socket -InputInfo $email -Text ([string]$Credential.Username)
            if ($actual -ne [string]$Credential.Username) {
                return "ERROR:MICROSOFT_EMAIL_VALUE_NOT_APPLIED"
            }

            Start-Sleep -Milliseconds 300
            $next = Wait-RoboPrecosBiVisibleButton -Socket $Socket -Labels @("AVANCAR","PROXIMO","NEXT","ENTRAR","SIGN IN") -TimeoutSeconds 5
            if ($next -and (Click-RoboPrecosBiButton -Socket $Socket -ButtonInfo $next)) {
                return "MICROSOFT_EMAIL_SUBMITTED_DIRECT"
            }
        }

        return "WAITING:MICROSOFT_EMAIL_NO_VISIBLE_INPUT"
    }

    if ($kind -eq "MICROSOFT_PASSWORD") {
        $password = Wait-RoboPrecosBiVisibleInput -Socket $Socket -Kind "PASSWORD" -TimeoutSeconds 20
        if (-not $password) {
            return "ERROR:MICROSOFT_VISIBLE_PASSWORD_NOT_FOUND"
        }

        $actualPassword = Set-RoboPrecosBiInputValue -Socket $Socket -InputInfo $password -Text ([string]$Credential.Password)
        if ([string]::IsNullOrWhiteSpace($actualPassword)) {
            return "ERROR:MICROSOFT_PASSWORD_VALUE_NOT_APPLIED"
        }

        Start-Sleep -Milliseconds 300
        $enter = Wait-RoboPrecosBiVisibleButton -Socket $Socket -Labels @("ENTRAR","SIGN IN","CONTINUAR","CONTINUE") -TimeoutSeconds 5
        if (-not $enter -or -not (Click-RoboPrecosBiButton -Socket $Socket -ButtonInfo $enter)) {
            return "ERROR:MICROSOFT_ENTER_BUTTON_NOT_CLICKED"
        }

        return "MICROSOFT_PASSWORD_SUBMITTED_DIRECT"
    }

    if ($kind -eq "MICROSOFT_STAY") {
        $yes = Wait-RoboPrecosBiVisibleButton -Socket $Socket -Labels @("SIM","YES") -TimeoutSeconds 10
        if (-not $yes -or -not (Click-RoboPrecosBiButton -Socket $Socket -ButtonInfo $yes)) {
            return "ERROR:MICROSOFT_STAY_YES_NOT_CLICKED"
        }

        return "MICROSOFT_STAY_CONFIRMED_DIRECT"
    }

    return ("WAITING:" + $kind)
}

function Invoke-RoboPrecosBiLogin {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        $Config,
        $Credential
    )

    $bi = Get-RoboPrecosBiConfig -Config $Config
    $deadline = (Get-Date).AddSeconds([int]$bi.loginTimeoutSeconds)
    $lastKind = ""
    $manualNoticeShown = $false

    while ((Get-Date) -lt $deadline) {
        $state = Get-RoboPrecosBiPageState -Socket $Socket
        $kind = if ($state) { [string]$state.kind } else { "EMPTY" }

        if ($kind -ne $lastKind) {
            Write-RoboLog ("Estado login Power BI: " + $kind + " | " + [string]$state.href)
            $lastKind = $kind
        }

        if ($kind -eq "AUTHENTICATED") {
            Write-RoboLog ("Sessao Power BI autenticada de fato. URL: " + [string]$state.href)
            return
        }

        if ($kind -in @("POWERBI_EMAIL","MICROSOFT_EMAIL","MICROSOFT_PASSWORD","MICROSOFT_STAY")) {
            $action = Invoke-RoboPrecosBiLoginStep -Socket $Socket -State $state -Credential $Credential

            if ($action.StartsWith("ERROR:")) {
                throw ("Falha no estado " + $kind + ": " + $action)
            }

            Write-RoboLog ("Power BI login: " + $action)
            Start-Sleep -Milliseconds 900
            continue
        }

        if (-not $manualNoticeShown) {
            Write-Host ""
            Write-Host "Autenticacao Power BI em andamento." -ForegroundColor Yellow
            Write-Host "O robo preenche automaticamente e-mail, senha e confirmacao de permanencia." -ForegroundColor Yellow
            Write-Host "Se surgir MFA ou aprovacao externa, conclua apenas essa etapa manualmente." -ForegroundColor DarkGray
            $manualNoticeShown = $true
        }

        Start-Sleep -Milliseconds 700
    }

    $last = Get-RoboPrecosBiPageState -Socket $Socket
    throw ("Timeout autenticando no Power BI. Ultimo estado: " + [string]$last.kind + " | URL: " + [string]$last.href)
}

function Get-RoboPrecosBiAccessibilityText {
    param([System.Net.WebSockets.ClientWebSocket]$Socket)

    try {
        [void](Invoke-CdpCommand -Socket $Socket -Method "Accessibility.enable")
        $tree = Invoke-CdpCommand -Socket $Socket -Method "Accessibility.getFullAXTree"

        if (-not $tree -or -not ($tree.PSObject.Properties.Name -contains "nodes")) {
            return ""
        }

        $parts = New-Object System.Collections.Generic.List[string]

        foreach ($node in @($tree.nodes)) {
            $ignoredProp = $node.PSObject.Properties["ignored"]
            if ($ignoredProp -and [bool]$ignoredProp.Value) {
                continue
            }

            foreach ($propertyName in @("name", "value", "description")) {
                $property = $node.PSObject.Properties[$propertyName]
                if (-not $property -or -not $property.Value) {
                    continue
                }

                $valueProperty = $property.Value.PSObject.Properties["value"]
                if (-not $valueProperty) {
                    continue
                }

                $text = ([string]$valueProperty.Value).Trim()
                if (-not [string]::IsNullOrWhiteSpace($text)) {
                    $parts.Add($text)
                }
            }
        }

        return ($parts -join [Environment]::NewLine)
    }
    catch {
        Write-RoboLog ("Fallback de acessibilidade do Power BI indisponivel: " + $_.Exception.Message) "AVISO"
        return ""
    }
}

function Wait-RoboPrecosBiText {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [Parameter(Mandatory = $true)][string[]]$RequiredTexts,
        [int]$TimeoutSeconds = 120
    )

    $required = @($RequiredTexts | ForEach-Object { ConvertTo-RoboPrecosNormalizedText $_ })
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $lastCombined = ""
    $attempt = 0

    while ((Get-Date) -lt $deadline) {
        $attempt++

        $body = ""
        try {
            $body = [string](Invoke-CdpExpression -Socket $Socket -Expression "(document.body && document.body.innerText) ? document.body.innerText : ''")
        }
        catch {}

        $axText = ""
        if (($attempt -eq 1) -or ($attempt % 3 -eq 0)) {
            $axText = Get-RoboPrecosBiAccessibilityText -Socket $Socket
        }

        $lastCombined = ConvertTo-RoboPrecosNormalizedText ($body + [Environment]::NewLine + $axText)

        $all = $true
        foreach ($item in $required) {
            if (-not $lastCombined.Contains($item)) {
                $all = $false
                break
            }
        }

        if ($all) {
            $source = if (-not [string]::IsNullOrWhiteSpace($axText)) { "DOM/AX" } else { "DOM" }
            Write-RoboLog ("Elementos Power BI reconhecidos via " + $source + ": " + ($RequiredTexts -join ", "))
            return
        }

        Start-Sleep -Milliseconds 700
    }

    $diagnostic = ($lastCombined -replace '\s+', ' ')
    if ($diagnostic.Length -gt 500) {
        $diagnostic = $diagnostic.Substring(0, 500)
    }

    throw ("Power BI nao apresentou os elementos esperados: " + ($RequiredTexts -join ", ") + ". Texto detectado: " + $diagnostic)
}

function Open-RoboPrecosBiPage {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        $Config,
        [Parameter(Mandatory = $true)][string]$Url,
        [string[]]$RequiredTexts
    )

    $bi = Get-RoboPrecosBiConfig -Config $Config
    $credential = Get-RoboPrecosBiCredential -Config $Config

    # REGRA: nunca navegar direto ao relatorio antes de concluir TODA a autenticacao.
    $state = Get-RoboPrecosBiPageState -Socket $Socket
    if (-not $state -or [string]$state.kind -ne "AUTHENTICATED") {
        $loginUrl = [string]$bi.loginUrl
        if ([string]::IsNullOrWhiteSpace($loginUrl)) {
            $loginUrl = "https://app.powerbi.com/"
        }

        if (-not $state -or -not ([string]$state.href).StartsWith($loginUrl, [StringComparison]::OrdinalIgnoreCase)) {
            Navigate-Cdp -Socket $Socket -Url $loginUrl -TimeoutSeconds ([int]$bi.pageLoadTimeoutSeconds)
        }

        Invoke-RoboPrecosBiLogin -Socket $Socket -Config $Config -Credential $credential
    }

    # Somente depois do estado AUTHENTICATED usa o endereco do relatorio.
    Write-RoboLog ("Autenticacao Power BI concluida. Abrindo relatorio autorizado: " + $Url)
    Navigate-Cdp -Socket $Socket -Url $Url -TimeoutSeconds ([int]$bi.pageLoadTimeoutSeconds)

    # O link do relatorio nao pode retornar para singleSignOn/login.
    $deadline = (Get-Date).AddSeconds([int]$bi.pageLoadTimeoutSeconds)
    do {
        $reportState = Get-RoboPrecosBiPageState -Socket $Socket

        if ($reportState -and [string]$reportState.kind -eq "AUTHENTICATED" -and
            ([string]$reportState.href).Contains("/reports/")) {
            break
        }

        if ($reportState -and [string]$reportState.kind -in @("POWERBI_EMAIL","MICROSOFT_EMAIL","MICROSOFT_PASSWORD","MICROSOFT_STAY")) {
            throw ("A sessao Power BI voltou ao login depois de abrir o relatorio. Estado: " + [string]$reportState.kind)
        }

        Start-Sleep -Milliseconds 500
    } while ((Get-Date) -lt $deadline)

    if (-not $reportState -or -not ([string]$reportState.href).Contains("/reports/")) {
        throw ("Power BI nao chegou ao relatorio apos autenticacao. URL atual: " + [string]$reportState.href)
    }

    if ($RequiredTexts -and $RequiredTexts.Count -gt 0) {
        Wait-RoboPrecosBiText -Socket $Socket -RequiredTexts $RequiredTexts -TimeoutSeconds ([int]$bi.pageLoadTimeoutSeconds)
    }
}

function Clear-RoboPrecosBiEmpresaSlicer {
    param([System.Net.WebSockets.ClientWebSocket]$Socket)

    $result = Invoke-CdpExpression -Socket $Socket -Expression @'
(async () => {
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const norm = s => (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ').trim().toLowerCase();
  const visible = e => !!(e && (e.offsetWidth || e.offsetHeight || e.getClientRects().length));
  const labels = [...document.querySelectorAll('*')].filter(e =>
    visible(e) && norm(e.innerText || e.textContent) === 'empresa'
  );

  for (const label of labels) {
    let node = label;
    for (let level=0; level<8 && node; level++, node=node.parentElement) {
      const clear = [...node.querySelectorAll('button,[role="button"]')].find(b => {
        const t = norm((b.getAttribute('aria-label') || '') + ' ' + (b.getAttribute('title') || '') + ' ' + (b.innerText || ''));
        return t.includes('limpar') || t.includes('clear selection') || t.includes('clear filter');
      });
      if (clear && visible(clear)) {
        clear.click();
        await sleep(1200);
        return 'CLEARED_BUTTON';
      }

      const combo = node.querySelector('[role="combobox"]');
      if (combo && visible(combo)) {
        const current = norm(combo.innerText || combo.textContent);
        if (current === 'todos' || current === 'all') return 'ALREADY_ALL';

        combo.click();
        await sleep(500);
        const options = [...document.querySelectorAll('[role="option"],[role="menuitem"],[role="listbox"] *')].filter(visible);
        const all = options.find(o => {
          const t = norm(o.innerText || o.textContent);
          return t === 'todos' || t === 'all' || t === 'selecionar tudo' || t === 'select all';
        });
        if (all) {
          all.click();
          await sleep(1200);
          return 'CLEARED_OPTION';
        }
      }
    }
  }

  return 'NOT_FOUND';
})()
'@

    Write-RoboLog ("Slicer Empresa Power BI: " + [string]$result)
    return [string]$result
}

function Get-RoboPrecosBiAxPropertyValue {
    param(
        $Node,
        [Parameter(Mandatory = $true)][string]$PropertyName
    )

    if (-not $Node) { return "" }

    $property = $Node.PSObject.Properties[$PropertyName]
    if (-not $property -or -not $property.Value) { return "" }

    $valueProperty = $property.Value.PSObject.Properties["value"]
    if (-not $valueProperty) { return "" }

    return [string]$valueProperty.Value
}

function Get-RoboPrecosBiAccessibilityRows {
    param([System.Net.WebSockets.ClientWebSocket]$Socket)

    [void](Invoke-CdpCommand -Socket $Socket -Method "Accessibility.enable")
    $tree = Invoke-CdpCommand -Socket $Socket -Method "Accessibility.getFullAXTree"

    if (-not $tree -or -not ($tree.PSObject.Properties.Name -contains "nodes")) {
        return @()
    }

    $nodes = @($tree.nodes)
    $byId = @{}

    foreach ($node in $nodes) {
        $idProperty = $node.PSObject.Properties["nodeId"]
        if ($idProperty) {
            $byId[[string]$idProperty.Value] = $node
        }
    }

    $rows = New-Object System.Collections.ArrayList
    $cellRoles = @("gridcell", "cell", "columnheader", "rowheader")

    foreach ($rowNode in $nodes) {
        $role = Get-RoboPrecosBiAxPropertyValue -Node $rowNode -PropertyName "role"
        if ($role -ne "row") {
            continue
        }

        $childIdsProperty = $rowNode.PSObject.Properties["childIds"]
        if (-not $childIdsProperty) {
            continue
        }

        $values = New-Object System.Collections.Generic.List[string]

        foreach ($rootChildId in @($childIdsProperty.Value)) {
            $queue = New-Object System.Collections.Generic.Queue[string]
            $queue.Enqueue([string]$rootChildId)

            while ($queue.Count -gt 0) {
                $nodeId = $queue.Dequeue()
                if (-not $byId.ContainsKey($nodeId)) {
                    continue
                }

                $node = $byId[$nodeId]
                $nodeRole = Get-RoboPrecosBiAxPropertyValue -Node $node -PropertyName "role"

                if ($cellRoles -contains $nodeRole) {
                    $value = Get-RoboPrecosBiAxPropertyValue -Node $node -PropertyName "name"
                    if ([string]::IsNullOrWhiteSpace($value)) {
                        $value = Get-RoboPrecosBiAxPropertyValue -Node $node -PropertyName "value"
                    }

                    # Mantem celula vazia para preservar a posicao das colunas.
                    $values.Add([string]$value)
                    continue
                }

                $children = $node.PSObject.Properties["childIds"]
                if ($children) {
                    foreach ($childId in @($children.Value)) {
                        $queue.Enqueue([string]$childId)
                    }
                }
            }
        }

        if ($values.Count -gt 1) {
            [void]$rows.Add([PSCustomObject]@{
                Cells = @($values.ToArray())
            })
        }
    }

    return @($rows)
}

function Get-RoboPrecosBiGridRows {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [string]$TitleContains = "",
        [string[]]$RequiredHeaders = @()
    )

    $titleJson = ($TitleContains | ConvertTo-Json -Compress)
    $headersJson = ($RequiredHeaders | ConvertTo-Json -Compress)

    $snapshotExpression = @"
(() => {
  const title = $titleJson;
  const required = $headersJson;
  const norm = s => (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ').trim().toUpperCase();
  const titleN = norm(title);
  const reqN = (required || []).map(norm);

  const findGrid = () => {
    const grids = [...document.querySelectorAll('[role="grid"],[role="table"]')];
    const candidates = grids.map(grid => {
      let node = grid;
      let context = '';

      for (let i=0;i<12 && node;i++,node=node.parentElement) {
        const text = norm(node.innerText || node.textContent);
        if (text.length > context.length && text.length < 220000) context = text;
      }

      const own = norm(grid.innerText || grid.textContent);
      let score = reqN.filter(h => own.includes(h) || context.includes(h)).length * 15;
      if (titleN && context.includes(titleN)) score += 150;
      return {grid,score};
    }).sort((a,b)=>b.score-a.score);

    return candidates.length ? candidates[0] : null;
  };

  const selected = findGrid();
  if (!selected || selected.score < Math.max(15,reqN.length*8)) {
    return JSON.stringify({ok:false,message:'GRID_NOT_FOUND',rows:[],companies:[]});
  }

  const grid = selected.grid;

  let visual = grid.parentElement || grid;
  let node = grid;
  const initialRect = grid.getBoundingClientRect();

  for (let i=0;i<14 && node;i++,node=node.parentElement) {
    const text = norm(node.innerText || node.textContent);
    if (!titleN || !text.includes(titleN)) continue;

    const r = node.getBoundingClientRect();
    const ratio = Math.max(1,(r.width*r.height)/Math.max(1,initialRect.width*initialRect.height));
    if (r.width >= initialRect.width*0.85 && r.height >= initialRect.height*0.85 && ratio < 12) {
      visual = node;
      break;
    }
  }

  try {
    visual.scrollIntoView({block:'center',inline:'center',behavior:'instant'});
  } catch {}

  const rows = [];
  const dataCells = [];

  for (const row of [...grid.querySelectorAll('[role="row"]')]) {
    let cells = [...row.querySelectorAll('[role="columnheader"],[role="gridcell"],[role="rowheader"]')];
    let values = cells.map(c => (c.innerText || c.textContent || '').replace(/\s+/g,' ').trim());

    if (!values.length) {
      values = (row.innerText || '').split(/\r?\n|\t/).map(s=>s.trim());
    }

    if (!values.length) continue;
    rows.push(values);

    const company = String(values[0] || '').trim();
    if (/^\d+$/.test(company)) {
      let focusCell = cells.find(c => {
        const r = c.getBoundingClientRect();
        return r.width > 4 && r.height > 4;
      });

      if (!focusCell && cells.length) focusCell = cells[0];

      if (focusCell) {
        const r = focusCell.getBoundingClientRect();
        if (r.width > 4 && r.height > 4) {
          dataCells.push({
            company,
            x:Math.max(3,Math.min(innerWidth-3,r.left + Math.min(r.width*0.5,40))),
            y:Math.max(3,Math.min(innerHeight-3,r.top + r.height*0.5))
          });
        }
      }
    }
  }

  const companies = rows.map(v=>String(v[0]||'').trim()).filter(v=>/^\d+$/.test(v));
  const first = dataCells.length ? dataCells[0] : null;
  const last = dataCells.length ? dataCells[dataCells.length-1] : null;

  return JSON.stringify({
    ok:true,
    rows,
    companies,
    signature:companies.join(','),
    firstCompany:companies.length?companies[0]:'',
    lastCompany:companies.length?companies[companies.length-1]:'',
    firstFocus:first,
    lastFocus:last,
    activeRole:document.activeElement ? (document.activeElement.getAttribute('role') || '') : '',
    activeText:document.activeElement ? ((document.activeElement.innerText || document.activeElement.textContent || '').replace(/\s+/g,' ').trim().slice(0,120)) : ''
  });
})()
"@

    function Get-RoboPrecosBiGridSnapshot {
        $json = [string](Invoke-CdpExpression -Socket $Socket -Expression $snapshotExpression)

        if ([string]::IsNullOrWhiteSpace($json)) {
            return $null
        }

        try {
            return ($json | ConvertFrom-Json)
        }
        catch {
            return $null
        }
    }

    function Add-RoboPrecosBiGridRows {
        param(
            [Parameter(Mandatory = $true)]$Snapshot,
            [Parameter(Mandatory = $true)][hashtable]$Store
        )

        foreach ($row in @($Snapshot.rows)) {
            $values = @($row)
            if ($values.Count -eq 0) { continue }

            $key = ($values | ForEach-Object { [string]$_ }) -join ([char]31)
            $Store[$key] = $values
        }
    }

    function Click-RoboPrecosBiGridPoint {
        param(
            [double]$X,
            [double]$Y
        )

        [void](Invoke-CdpCommand -Socket $Socket -Method "Input.dispatchMouseEvent" -Params @{
            type = "mouseMoved"
            x = $X
            y = $Y
            button = "none"
        })

        [void](Invoke-CdpCommand -Socket $Socket -Method "Input.dispatchMouseEvent" -Params @{
            type = "mousePressed"
            x = $X
            y = $Y
            button = "left"
            buttons = 1
            clickCount = 1
        })

        [void](Invoke-CdpCommand -Socket $Socket -Method "Input.dispatchMouseEvent" -Params @{
            type = "mouseReleased"
            x = $X
            y = $Y
            button = "left"
            buttons = 0
            clickCount = 1
        })
    }

    function Send-RoboPrecosBiGridKey {
        param(
            [Parameter(Mandatory = $true)][string]$Key,
            [Parameter(Mandatory = $true)][string]$Code,
            [Parameter(Mandatory = $true)][int]$VirtualKeyCode,
            [int]$Modifiers = 0
        )

        [void](Invoke-CdpCommand -Socket $Socket -Method "Input.dispatchKeyEvent" -Params @{
            type = "rawKeyDown"
            key = $Key
            code = $Code
            windowsVirtualKeyCode = $VirtualKeyCode
            nativeVirtualKeyCode = $VirtualKeyCode
            modifiers = $Modifiers
        })

        Start-Sleep -Milliseconds 40

        [void](Invoke-CdpCommand -Socket $Socket -Method "Input.dispatchKeyEvent" -Params @{
            type = "keyUp"
            key = $Key
            code = $Code
            windowsVirtualKeyCode = $VirtualKeyCode
            nativeVirtualKeyCode = $VirtualKeyCode
            modifiers = $Modifiers
        })
    }

    $deadline = (Get-Date).AddSeconds(45)
    $store = @{}
    $snapshot = $null

    do {
        $snapshot = Get-RoboPrecosBiGridSnapshot
        if ($snapshot -and [bool]$snapshot.ok -and @($snapshot.companies).Count -gt 0) {
            break
        }

        Start-Sleep -Milliseconds 600
    } while ((Get-Date) -lt $deadline)

    if (-not $snapshot -or -not [bool]$snapshot.ok -or @($snapshot.companies).Count -eq 0) {
        throw "Power BI nao materializou as linhas da tabela esperada."
    }

    Add-RoboPrecosBiGridRows -Snapshot $snapshot -Store $store

    Write-RoboLog (
        "Grid Power BI confirmado. Janela inicial=" +
        [string]$snapshot.firstCompany + "..." + [string]$snapshot.lastCompany +
        " | empresas visiveis=" + @($snapshot.companies).Count
    )

    # PASSO A: foca uma celula real da tabela.
    if ($snapshot.lastFocus) {
        Click-RoboPrecosBiGridPoint -X ([double]$snapshot.lastFocus.x) -Y ([double]$snapshot.lastFocus.y)
        Start-Sleep -Milliseconds 250
    }

    # PASSO B: tenta ir ao inicio logico do grid com Ctrl+Home.
    Send-RoboPrecosBiGridKey -Key "Home" -Code "Home" -VirtualKeyCode 36 -Modifiers 2
    Start-Sleep -Milliseconds 500

    $afterHome = Get-RoboPrecosBiGridSnapshot
    if ($afterHome -and [bool]$afterHome.ok) {
        Add-RoboPrecosBiGridRows -Snapshot $afterHome -Store $store
        $snapshot = $afterHome
        Write-RoboLog (
            "Grid Ctrl+Home -> janela=" +
            [string]$snapshot.firstCompany + "..." + [string]$snapshot.lastCompany
        )
    }

    # PASSO C: percorre para frente com PageDown.
    $stable = 0
    $lastSignature = ""

    for ($i=0; $i -lt 12; $i++) {
        if ($snapshot.lastFocus) {
            Click-RoboPrecosBiGridPoint -X ([double]$snapshot.lastFocus.x) -Y ([double]$snapshot.lastFocus.y)
            Start-Sleep -Milliseconds 100
        }

        Send-RoboPrecosBiGridKey -Key "PageDown" -Code "PageDown" -VirtualKeyCode 34
        Start-Sleep -Milliseconds 500

        $next = Get-RoboPrecosBiGridSnapshot
        if (-not $next -or -not [bool]$next.ok) { continue }

        Add-RoboPrecosBiGridRows -Snapshot $next -Store $store

        $sig = [string]$next.signature
        if ($sig -eq $lastSignature) { $stable++ } else { $stable = 0 }
        $lastSignature = $sig
        $snapshot = $next

        Write-RoboLog (
            "Grid PageDown " + $i +
            " -> janela=" + [string]$snapshot.firstCompany +
            "..." + [string]$snapshot.lastCompany
        )

        if ($stable -ge 2) { break }
    }

    # PASSO D: forca a outra extremidade com Ctrl+End.
    if ($snapshot.lastFocus) {
        Click-RoboPrecosBiGridPoint -X ([double]$snapshot.lastFocus.x) -Y ([double]$snapshot.lastFocus.y)
        Start-Sleep -Milliseconds 100
    }

    Send-RoboPrecosBiGridKey -Key "End" -Code "End" -VirtualKeyCode 35 -Modifiers 2
    Start-Sleep -Milliseconds 600

    $afterEnd = Get-RoboPrecosBiGridSnapshot
    if ($afterEnd -and [bool]$afterEnd.ok) {
        Add-RoboPrecosBiGridRows -Snapshot $afterEnd -Store $store
        $snapshot = $afterEnd
        Write-RoboLog (
            "Grid Ctrl+End -> janela=" +
            [string]$snapshot.firstCompany + "..." + [string]$snapshot.lastCompany
        )
    }

    # PASSO E: passada reversa com PageUp.
    $stable = 0
    $lastSignature = ""

    for ($i=0; $i -lt 12; $i++) {
        if ($snapshot.firstFocus) {
            Click-RoboPrecosBiGridPoint -X ([double]$snapshot.firstFocus.x) -Y ([double]$snapshot.firstFocus.y)
            Start-Sleep -Milliseconds 100
        }

        Send-RoboPrecosBiGridKey -Key "PageUp" -Code "PageUp" -VirtualKeyCode 33
        Start-Sleep -Milliseconds 500

        $next = Get-RoboPrecosBiGridSnapshot
        if (-not $next -or -not [bool]$next.ok) { continue }

        Add-RoboPrecosBiGridRows -Snapshot $next -Store $store

        $sig = [string]$next.signature
        if ($sig -eq $lastSignature) { $stable++ } else { $stable = 0 }
        $lastSignature = $sig
        $snapshot = $next

        Write-RoboLog (
            "Grid PageUp " + $i +
            " -> janela=" + [string]$snapshot.firstCompany +
            "..." + [string]$snapshot.lastCompany
        )

        if ($stable -ge 2) { break }
    }

    $finalRows = @($store.Values)
    $companies = @(
        $finalRows |
        ForEach-Object {
            $v = @($_)
            if ($v.Count -gt 0 -and ([string]$v[0]).Trim() -match '^\d+$') {
                ([string]$v[0]).Trim()
            }
        } |
        Sort-Object {[int]$_} -Unique
    )

    Write-RoboLog (
        "Varredura por teclado concluida: " +
        $companies.Count + " empresa(s) | IDs=" + ($companies -join ",")
    )

    if ($companies.Count -eq 0) {
        throw "A varredura por teclado nao materializou nenhuma empresa."
    }

    return $finalRows
}

function Get-RoboPrecosDiscountMode {
    param(
        [Parameter(Mandatory = $true)][string]$StartDate,
        [Parameter(Mandatory = $true)][string]$EndDate
    )

    $culture = [Globalization.CultureInfo]::InvariantCulture
    $start = [datetime]::ParseExact($StartDate, "dd/MM/yyyy", $culture)
    $end = [datetime]::ParseExact($EndDate, "dd/MM/yyyy", $culture)

    if ($start.Year -ne $end.Year -or $start.Month -ne $end.Month) {
        throw "O fluxo de descontos e mensal. Data inicial e final precisam estar no mesmo mes."
    }

    $now = Get-Date
    $targetKey = ($start.Year * 100) + $start.Month
    $currentKey = ($now.Year * 100) + $now.Month
    $monthDate = [datetime]::new($start.Year, $start.Month, 1, 0, 0, 0)

    if ($targetKey -gt $currentKey) {
        throw ("Periodo futuro nao pode ser consultado no Power BI: " + $monthDate.ToString("MM/yyyy"))
    }

    if ($targetKey -eq $currentKey) {
        return [PSCustomObject]@{ Mode="CURRENT"; MonthDate=$monthDate }
    }

    return [PSCustomObject]@{ Mode="HISTORICAL"; MonthDate=$monthDate }
}

function ConvertFrom-RoboPrecosBiCurrentRows {
    param([array]$Rows)

    $records = @{}

    foreach ($row in @($Rows)) {
        $cells = if ($row -and ($row.PSObject.Properties.Name -contains "Cells")) { @($row.Cells) } else { @($row) }
        if ($cells.Count -lt 3) { continue }

        $companyText = ([string]$cells[0]).Trim()
        if ($companyText -notmatch '^\d+([.,]0+)?$') { continue }

        # No visual "DESCONTO POR MOTIVO", PRECO ERRADO e o ultimo par:
        # CUPONS + DESCONTO. Celulas vazias anteriores permanecem como colunas.
        if ($cells.Count -lt 9) { continue }

        $quantity = ConvertFrom-RoboPrecosBiInteger $cells[$cells.Count - 2]
        $discount = ConvertFrom-RoboPrecosBiDecimal $cells[$cells.Count - 1]

        # Ausencia de dado NAO significa zero. Somente grava quando os dois
        # campos existem; zero explicito e aceito como valor valido.
        if ($null -eq $quantity -or $null -eq $discount) { continue }

        $store = ConvertTo-RoboStore $companyText
        $company = [int][double]$companyText.Replace(",", ".")

        if ($records.ContainsKey($store)) {
            $existing = $records[$store]

            $sameQuantity = ([int]$existing.QuantidadeCupons -eq [int]$quantity)
            $sameDiscount = ([Math]::Abs([double]$existing.Desconto - [double]$discount) -lt 0.005)

            if (-not $sameQuantity -or -not $sameDiscount) {
                throw (
                    "Power BI reciclou a empresa " + $company +
                    " com valores conflitantes. Primeiro: " +
                    [int]$existing.QuantidadeCupons + " / R$ " +
                    ([double]$existing.Desconto).ToString("N2",[Globalization.CultureInfo]::GetCultureInfo("pt-BR")) +
                    " | Novo: " + [int]$quantity + " / R$ " +
                    ([double]$discount).ToString("N2",[Globalization.CultureInfo]::GetCultureInfo("pt-BR"))
                )
            }

            continue
        }

        $records[$store] = [PSCustomObject]@{
            Loja = $store
            Empresa = $company
            QuantidadeCupons = [int]$quantity
            Desconto = [double]$discount
            Motivo = "PRECO ERRADO"
            Fonte = "RESUMO"
        }
    }

    return @($records.Values | Sort-Object Empresa)
}

function Get-RoboPrecosBiCurrentVisualTotal {
    param([array]$Rows)

    foreach ($row in @($Rows)) {
        $cells = if ($row -and ($row.PSObject.Properties.Name -contains "Cells")) { @($row.Cells) } else { @($row) }
        if ($cells.Count -lt 3) { continue }

        $label = ConvertTo-RoboPrecosNormalizedText ([string]$cells[0])
        if ($label -ne "TOTAL") { continue }

        if ($cells.Count -lt 9) { continue }

        $quantity = ConvertFrom-RoboPrecosBiInteger $cells[$cells.Count - 2]
        $discount = ConvertFrom-RoboPrecosBiDecimal $cells[$cells.Count - 1]

        if ($null -eq $quantity -or $null -eq $discount) {
            continue
        }

        return [PSCustomObject]@{
            QuantidadeCupons = [int]$quantity
            Desconto = [double]$discount
        }
    }

    return $null
}

function Test-RoboPrecosBiCurrentIntegrity {
    param(
        [Parameter(Mandatory = $true)][array]$Records,
        [Parameter(Mandatory = $true)][array]$Rows
    )

    if ($Records.Count -eq 0) {
        throw "Integridade Power BI: nenhuma loja valida foi coletada."
    }

    $visualTotal = Get-RoboPrecosBiCurrentVisualTotal -Rows $Rows
    if (-not $visualTotal) {
        throw "Integridade Power BI: o rodape TOTAL do visual PRECO ERRADO nao foi capturado. A coleta nao sera considerada completa."
    }

    $sumQuantity = 0
    $sumDiscount = 0.0
    $companies = New-Object System.Collections.Generic.List[int]

    foreach ($record in $Records) {
        $sumQuantity += [int]$record.QuantidadeCupons
        $sumDiscount += [double]$record.Desconto
        $companies.Add([int]$record.Empresa)
    }

    $sumDiscount = [Math]::Round($sumDiscount, 2, [MidpointRounding]::AwayFromZero)
    $visualDiscount = [Math]::Round([double]$visualTotal.Desconto, 2, [MidpointRounding]::AwayFromZero)

    $quantityMatches = ($sumQuantity -eq [int]$visualTotal.QuantidadeCupons)
    $discountMatches = ([Math]::Abs($sumDiscount - $visualDiscount) -lt 0.005)

    $uniqueCompanies = @($companies | Sort-Object -Unique)
    $maxCompany = if ($uniqueCompanies.Count -gt 0) { [int]($uniqueCompanies | Measure-Object -Maximum).Maximum } else { 0 }
    $missingCompanies = @()

    if ($maxCompany -gt 0) {
        $set = @{}
        foreach ($company in $uniqueCompanies) { $set[[int]$company] = $true }
        for ($i = 1; $i -le $maxCompany; $i++) {
            if (-not $set.ContainsKey($i)) {
                $missingCompanies += $i
            }
        }
    }

    if (-not $quantityMatches -or -not $discountMatches) {
        throw (
            "Integridade Power BI falhou. Soma capturada: " +
            $sumQuantity + " cupons / R$ " + $sumDiscount.ToString("N2",[Globalization.CultureInfo]::GetCultureInfo("pt-BR")) +
            " | Total do visual: " + [int]$visualTotal.QuantidadeCupons + " cupons / R$ " +
            $visualDiscount.ToString("N2",[Globalization.CultureInfo]::GetCultureInfo("pt-BR"))
        )
    }

    Write-RoboLog (
        "Integridade Power BI OK: " + $Records.Count + " lojas | " +
        $sumQuantity + " cupons | R$ " +
        $sumDiscount.ToString("N2",[Globalization.CultureInfo]::GetCultureInfo("pt-BR")) +
        " | Total do visual reconciliado."
    )

    return [PSCustomObject]@{
        Passed = $true
        RecordCount = $Records.Count
        SumQuantity = $sumQuantity
        SumDiscount = $sumDiscount
        VisualQuantity = [int]$visualTotal.QuantidadeCupons
        VisualDiscount = $visualDiscount
        MaxCompany = $maxCompany
        MissingCompanies = @($missingCompanies)
    }
}

function Get-RoboPrecosMonthNumber {
    param([string]$MonthName)

    $name = ConvertTo-RoboPrecosNormalizedText $MonthName
    $map = @{
        "JANEIRO"=1; "FEVEREIRO"=2; "MARCO"=3; "ABRIL"=4;
        "MAIO"=5; "JUNHO"=6; "JULHO"=7; "AGOSTO"=8;
        "SETEMBRO"=9; "OUTUBRO"=10; "NOVEMBRO"=11; "DEZEMBRO"=12
    }

    if ($map.ContainsKey($name)) { return [int]$map[$name] }
    return 0
}

function ConvertFrom-RoboPrecosBiHistoricalRows {
    param(
        [array]$Rows,
        [datetime]$MonthDate
    )

    $records = @{}

    foreach ($row in @($Rows)) {
        $cells = if ($row -and ($row.PSObject.Properties.Name -contains "Cells")) { @($row.Cells) } else { @($row) }
        if ($cells.Count -lt 7) { continue }

        $year = 0
        if (-not [int]::TryParse(([string]$cells[0]).Trim(), [ref]$year)) { continue }

        $month = Get-RoboPrecosMonthNumber ([string]$cells[1])
        if ($month -le 0) { continue }

        $companyText = ([string]$cells[2]).Trim()
        if ($companyText -notmatch '^\d+([.,]0+)?$') { continue }

        $type = ConvertTo-RoboPrecosNormalizedText ([string]$cells[3])
        if ($type -ne "PRECO ERRADO") { continue }

        if ($year -ne $MonthDate.Year -or $month -ne $MonthDate.Month) { continue }

        # Historico: Ano | Mes | Empresa | TIPO | Valor Total | Desconto | Quantidade Cupons
        # Valor Total e deliberadamente ignorado.
        $discount = ConvertFrom-RoboPrecosBiDecimal $cells[5]
        $quantity = ConvertFrom-RoboPrecosBiInteger $cells[6]

        if ($null -eq $quantity -or $null -eq $discount) { continue }

        $store = ConvertTo-RoboStore $companyText
        $records[$store] = [PSCustomObject]@{
            Loja = $store
            Empresa = [int][double]$companyText.Replace(",", ".")
            QuantidadeCupons = [int]$quantity
            Desconto = [double]$discount
            Motivo = "PRECO ERRADO"
            Fonte = "DESCONTOS MES ANTERIOR"
        }
    }

    return @($records.Values | Sort-Object Loja)
}

function Save-RoboPrecosBiDiscountSnapshot {
    param(
        [array]$Records,
        [datetime]$MonthDate,
        [string]$Mode
    )

    $directory = Join-Path $OutputPath "descontos"
    Ensure-RoboDirectory $directory
    $path = Join-Path $directory ("preco_errado_{0}_{1}_{2}.csv" -f $MonthDate.ToString("yyyyMM"), $Mode.ToLowerInvariant(), (Get-Date -Format "yyyyMMdd_HHmmss"))

    @($Records) | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
    Write-RoboLog ("Snapshot descontos Power BI salvo em " + $path)
    return $path
}

function Invoke-RoboPrecosBiDiscountCollection {
    param(
        $Config,
        [Parameter(Mandatory = $true)][string]$StartDate,
        [Parameter(Mandatory = $true)][string]$EndDate
    )

    $decision = Get-RoboPrecosDiscountMode -StartDate $StartDate -EndDate $EndDate
    $bi = Get-RoboPrecosBiConfig -Config $Config

    # Pede/cria a credencial antes de abrir o fluxo, assim como o PDA.
    $null = Get-RoboPrecosBiCredential -Config $Config

    $socket = $null
    try {
        $socket = Start-RoboPrecosBiBrowser -Config $Config

        if ([string]$decision.Mode -eq "CURRENT") {
            Write-Host ""
            Write-Host ("DESCONTOS: mes corrente {0} -> usando RESUMO / DESCONTO POR MOTIVO" -f $decision.MonthDate.ToString("MM/yyyy")) -ForegroundColor Cyan

            Open-RoboPrecosBiPage -Socket $socket -Config $Config -Url ([string]$bi.summaryUrl) -RequiredTexts @("DESCONTO POR MOTIVO", "PRECO ERRADO")
            $rows = @(Get-RoboPrecosBiGridRows -Socket $socket -TitleContains "DESCONTO POR MOTIVO" -RequiredHeaders @("EMPRESA", "PRECO ERRADO", "CUPONS", "DESCONTO"))
            $records = @(ConvertFrom-RoboPrecosBiCurrentRows -Rows $rows)
            $integrity = Test-RoboPrecosBiCurrentIntegrity -Records $records -Rows $rows
        }
        else {
            Write-Host ""
            Write-Host ("DESCONTOS: mes fechado {0} -> usando DESCONTOS MES ANTERIOR" -f $decision.MonthDate.ToString("MM/yyyy")) -ForegroundColor Cyan

            Open-RoboPrecosBiPage -Socket $socket -Config $Config -Url ([string]$bi.historicalUrl) -RequiredTexts @("Quantidade Cupons", "Valor Total", "Desconto")
            $slicerState = Clear-RoboPrecosBiEmpresaSlicer -Socket $socket
            if ([string]$slicerState -eq "NOT_FOUND") {
                throw "Nao foi possivel confirmar o filtro Empresa=Todos no historico do Power BI. Nenhum desconto sera gravado para evitar leitura parcial por filtro persistente."
            }
            Start-Sleep -Seconds 2

            $rows = @(Get-RoboPrecosBiGridRows -Socket $socket -RequiredHeaders @("ANO", "MES", "EMPRESA", "TIPO", "VALOR TOTAL", "DESCONTO", "QUANTIDADE CUPONS"))
            $records = @(ConvertFrom-RoboPrecosBiHistoricalRows -Rows $rows -MonthDate $decision.MonthDate)

            $histQuantity = 0
            $histDiscount = 0.0
            foreach ($record in $records) {
                $histQuantity += [int]$record.QuantidadeCupons
                $histDiscount += [double]$record.Desconto
            }
            $integrity = [PSCustomObject]@{
                Passed = $true
                RecordCount = $records.Count
                SumQuantity = $histQuantity
                SumDiscount = [Math]::Round($histDiscount, 2, [MidpointRounding]::AwayFromZero)
                VisualQuantity = $null
                VisualDiscount = $null
                MaxCompany = 0
                MissingCompanies = @()
            }
        }

        if ($records.Count -eq 0) {
            throw ("Power BI carregou, mas nenhum registro valido de PRECO ERRADO foi obtido para " + $decision.MonthDate.ToString("MM/yyyy") + ". Nenhuma celula de desconto sera alterada.")
        }

        $snapshot = Save-RoboPrecosBiDiscountSnapshot -Records $records -MonthDate $decision.MonthDate -Mode ([string]$decision.Mode)

        Write-Host ""
        Write-Host "DESCONTOS POWER BI COLETADOS" -ForegroundColor Green
        Write-Host ("Fonte              : {0}" -f $decision.Mode)
        Write-Host ("Mes                : {0}" -f $decision.MonthDate.ToString("MM/yyyy"))
        Write-Host ("Lojas com valores  : {0}" -f $records.Count)
        Write-Host ("Cupons coletados   : {0}" -f $integrity.SumQuantity)
        Write-Host ("Desconto coletado  : R$ {0}" -f ([double]$integrity.SumDiscount).ToString("N2",[Globalization.CultureInfo]::GetCultureInfo("pt-BR")))

        if ([string]$decision.Mode -eq "CURRENT") {
            Write-Host ("Total visual BI    : {0} cupons / R$ {1}" -f $integrity.VisualQuantity, ([double]$integrity.VisualDiscount).ToString("N2",[Globalization.CultureInfo]::GetCultureInfo("pt-BR"))) -ForegroundColor Green
            Write-Host "Integridade         : OK - soma das lojas = Total do BI" -ForegroundColor Green

            if (@($integrity.MissingCompanies).Count -gt 0) {
                Write-Host ("IDs nao retornados  : " + (@($integrity.MissingCompanies) -join ", ")) -ForegroundColor Yellow
            }
        }

        Write-Host ("Snapshot de controle: {0}" -f $snapshot)
        Write-Host ""

        return [PSCustomObject]@{
            Mode = [string]$decision.Mode
            MonthDate = $decision.MonthDate
            Records = $records
            RecordCount = $records.Count
            SnapshotPath = $snapshot
            Integrity = $integrity
        }
    }
    finally {
        if ($socket) {
            Close-CdpPage -Socket $socket
        }
    }
}
