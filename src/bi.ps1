$script:RoboPrecosBiDefaultConfig = [ordered]@{
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
    return (Connect-RoboPrecosBiTarget -Port $port -Url ([string]$bi.summaryUrl))
}

function Get-RoboPrecosBiPageState {
    param([System.Net.WebSockets.ClientWebSocket]$Socket)

    $expression = @'
(() => {
  const norm = s => (s || '').replace(/\s+/g,' ').trim().toLowerCase();
  const body = document.body ? document.body.innerText : '';
  const host = location.hostname || '';
  const href = location.href || '';
  const hasPassword = !!document.querySelector('input[type="password"], input[name="passwd"]');
  const hasEmail = !!document.querySelector('input[type="email"], input[name="loginfmt"]');
  const loginHost = host.includes('login.microsoftonline.com') || host.includes('login.live.com');
  const powerBi = host.includes('app.powerbi.com');
  let kind = 'OTHER';
  if (loginHost || hasPassword || hasEmail) kind = 'LOGIN';
  if (powerBi) kind = 'POWERBI';
  return {
    kind,
    host,
    href,
    title: document.title || '',
    ready: document.readyState || '',
    body: body.slice(0,5000),
    hasPassword,
    hasEmail,
    text: norm(body).slice(0,5000)
  };
})()
'@

    return Invoke-CdpJsonExpression -Socket $Socket -Expression $expression
}

function Invoke-RoboPrecosBiLogin {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        $Config,
        $Credential
    )

    $bi = Get-RoboPrecosBiConfig -Config $Config
    $deadline = (Get-Date).AddSeconds([int]$bi.loginTimeoutSeconds)
    $manualNoticeShown = $false

    while ((Get-Date) -lt $deadline) {
        $state = Get-RoboPrecosBiPageState -Socket $Socket

        if ($state -and [string]$state.kind -eq "POWERBI") {
            Write-RoboLog "Sessao Power BI autenticada."
            return
        }

        $usernameJson = ([string]$Credential.Username | ConvertTo-Json -Compress)
        $passwordJson = ([string]$Credential.Password | ConvertTo-Json -Compress)

        $action = Invoke-CdpExpression -Socket $Socket -Expression @"
(async () => {
  const user = $usernameJson;
  const pass = $passwordJson;
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const visible = e => !!(e && (e.offsetWidth || e.offsetHeight || e.getClientRects().length));
  const click = e => { if (!e) return false; e.click(); return true; };

  const email = document.querySelector('input[type="email"], input[name="loginfmt"]');
  if (email && visible(email)) {
    email.focus();
    email.value = user;
    email.dispatchEvent(new Event('input',{bubbles:true}));
    email.dispatchEvent(new Event('change',{bubbles:true}));
    await sleep(150);
    const next = document.querySelector('#idSIButton9, input[type="submit"], button[type="submit"]');
    click(next);
    return 'EMAIL_SUBMITTED';
  }

  const password = document.querySelector('input[type="password"], input[name="passwd"]');
  if (password && visible(password)) {
    password.focus();
    password.value = pass;
    password.dispatchEvent(new Event('input',{bubbles:true}));
    password.dispatchEvent(new Event('change',{bubbles:true}));
    await sleep(150);
    const submit = document.querySelector('#idSIButton9, input[type="submit"], button[type="submit"]');
    click(submit);
    return 'PASSWORD_SUBMITTED';
  }

  const buttons = [...document.querySelectorAll('button,input[type="button"],input[type="submit"],a')].filter(visible);
  const yes = buttons.find(e => {
    const t = (e.innerText || e.value || e.textContent || '').replace(/\s+/g,' ').trim().toLowerCase();
    return t === 'sim' || t === 'yes' || t === 'continuar' || t === 'continue';
  });
  if (yes) {
    click(yes);
    return 'CONFIRM_SUBMITTED';
  }

  const userChoice = [...document.querySelectorAll('*')].find(e => {
    const t = (e.innerText || e.textContent || '').trim().toLowerCase();
    return t === user.toLowerCase();
  });
  if (userChoice && visible(userChoice)) {
    click(userChoice);
    return 'ACCOUNT_SELECTED';
  }

  return 'WAITING';
})()
"@

        if ($action -and [string]$action -ne "WAITING") {
            Write-RoboLog ("Power BI login: " + [string]$action)
            Start-Sleep -Seconds 2
            continue
        }

        if (-not $manualNoticeShown) {
            Write-Host ""
            Write-Host "Aguardando autenticacao do Power BI..." -ForegroundColor Yellow
            Write-Host "Se a Microsoft solicitar MFA, aprovacao ou escolha de conta, conclua essa etapa na janela aberta." -ForegroundColor Yellow
            Write-Host "O robo retomara automaticamente apos a autenticacao." -ForegroundColor DarkGray
            $manualNoticeShown = $true
        }

        Start-Sleep -Seconds 2
    }

    throw "Timeout aguardando autenticacao no Power BI."
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

    Navigate-Cdp -Socket $Socket -Url $Url -TimeoutSeconds ([int]$bi.pageLoadTimeoutSeconds)
    Invoke-RoboPrecosBiLogin -Socket $Socket -Config $Config -Credential $credential

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

    $expression = @"
(async () => {
  const title = $titleJson;
  const required = $headersJson;
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const norm = s => (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ').trim().toUpperCase();
  const titleN = norm(title);
  const reqN = (required || []).map(norm);

  const allGrids = [...document.querySelectorAll('[role="grid"],[role="table"]')];
  const candidates = allGrids.map(grid => {
    let node = grid;
    let context = '';
    for (let i=0; i<7 && node; i++, node=node.parentElement) {
      const text = norm(node.innerText || node.textContent);
      if (text.length > context.length && text.length < 120000) context = text;
    }
    const own = norm(grid.innerText || grid.textContent);
    let score = reqN.filter(h => own.includes(h) || context.includes(h)).length * 10;
    if (titleN && context.includes(titleN)) score += 100;
    return {grid, score, context};
  }).sort((a,b) => b.score-a.score);

  if (!candidates.length || candidates[0].score < Math.max(10, reqN.length * 5)) {
    return { ok:false, message:'GRID_NOT_FOUND', grids:allGrids.length, rows:[] };
  }

  const grid = candidates[0].grid;
  let visual = grid;
  if (titleN) {
    let node = grid;
    for (let i=0; i<7 && node; i++, node=node.parentElement) {
      if (norm(node.innerText || node.textContent).includes(titleN)) {
        visual = node;
      }
    }
  }

  const scrollCandidates = [visual, grid, ...visual.querySelectorAll('*')].filter(e => {
    try {
      const style = getComputedStyle(e);
      return e.scrollHeight > e.clientHeight + 20 &&
        e.clientHeight > 50 &&
        (style.overflowY === 'auto' || style.overflowY === 'scroll' || e.getAttribute('role') === 'grid');
    } catch { return false; }
  }).sort((a,b) => (b.scrollHeight-b.clientHeight)-(a.scrollHeight-a.clientHeight));

  const scroller = scrollCandidates[0] || grid;
  const found = new Map();

  const capture = () => {
    const rows = [...grid.querySelectorAll('[role="row"]')];
    for (const row of rows) {
      let cells = [...row.querySelectorAll('[role="columnheader"],[role="gridcell"],[role="rowheader"]')];
      let values = cells.map(c => (c.innerText || c.textContent || '').replace(/\s+/g,' ').trim());

      if (!values.length) {
        values = (row.innerText || '').split(/\r?\n|\t/).map(s => s.trim()).filter(Boolean);
      }

      if (!values.length) continue;
      const key = values.join('\u001f');
      found.set(key, values);
    }
  };

  const oldTop = scroller.scrollTop || 0;
  scroller.scrollTop = 0;
  await sleep(300);
  capture();

  let last = -1;
  let guard = 0;
  while (guard++ < 250) {
    const max = Math.max(0, scroller.scrollHeight - scroller.clientHeight);
    const now = scroller.scrollTop || 0;
    if (now >= max - 2) break;

    const step = Math.max(120, Math.floor(scroller.clientHeight * 0.80));
    scroller.scrollTop = Math.min(max, now + step);
    await sleep(180);
    capture();

    const moved = scroller.scrollTop || 0;
    if (moved === last || moved === now) break;
    last = moved;
  }

  capture();
  scroller.scrollTop = oldTop;

  return {
    ok:true,
    score:candidates[0].score,
    scrollHeight:scroller.scrollHeight,
    clientHeight:scroller.clientHeight,
    rows:[...found.values()]
  };
})()
"@

    $jsonExpression = "(async () => JSON.stringify(await (" + $expression + ")))()"
    $json = [string](Invoke-CdpExpression -Socket $Socket -Expression $jsonExpression)

    if ([string]::IsNullOrWhiteSpace($json)) {
        throw "Power BI nao devolveu dados da tabela."
    }

    $result = $json | ConvertFrom-Json

    if ($result -and [bool]$result.ok) {
        return @($result.rows)
    }

    $axRows = @(Get-RoboPrecosBiAccessibilityRows -Socket $Socket)
    if ($axRows.Count -gt 0) {
        Write-RoboLog ("Tabela Power BI lida pelo fallback AX: " + $axRows.Count + " linha(s) acessiveis.")
        return @($axRows)
    }

    throw ("Nao foi possivel localizar/ler a tabela esperada no Power BI via DOM ou acessibilidade. Detalhe DOM: " + ($result | ConvertTo-Json -Compress -Depth 5))
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

    $target = Get-Date -Year $start.Year -Month $start.Month -Day 1 -Hour 0 -Minute 0 -Second 0
    $current = Get-Date -Year (Get-Date).Year -Month (Get-Date).Month -Day 1 -Hour 0 -Minute 0 -Second 0

    if ($target -gt $current) {
        throw ("Periodo futuro nao pode ser consultado no Power BI: " + $target.ToString("MM/yyyy"))
    }

    if ($target -eq $current) {
        return [PSCustomObject]@{ Mode="CURRENT"; MonthDate=$target }
    }

    return [PSCustomObject]@{ Mode="HISTORICAL"; MonthDate=$target }
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
        $records[$store] = [PSCustomObject]@{
            Loja = $store
            Empresa = [int][double]$companyText.Replace(",", ".")
            QuantidadeCupons = [int]$quantity
            Desconto = [double]$discount
            Motivo = "PRECO ERRADO"
            Fonte = "RESUMO"
        }
    }

    return @($records.Values | Sort-Object Loja)
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
        Write-Host ("Snapshot de controle: {0}" -f $snapshot)
        Write-Host ""

        return [PSCustomObject]@{
            Mode = [string]$decision.Mode
            MonthDate = $decision.MonthDate
            Records = $records
            RecordCount = $records.Count
            SnapshotPath = $snapshot
        }
    }
    finally {
        if ($socket) {
            Close-CdpPage -Socket $socket
        }
    }
}
