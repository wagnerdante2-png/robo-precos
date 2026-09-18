$script:RoboPrecosConfigPath = Join-Path $Root "config.precos.json"
$script:RoboPrecosConfigExamplePath = Join-Path $Root "config.precos.example.json"

function Get-RoboPrecosConfig {
    if (-not (Test-Path -LiteralPath $script:RoboPrecosConfigPath)) {
        if (-not (Test-Path -LiteralPath $script:RoboPrecosConfigExamplePath)) {
            throw "Arquivo de configuracao exemplo ausente."
        }
        Copy-Item -LiteralPath $script:RoboPrecosConfigExamplePath -Destination $script:RoboPrecosConfigPath -Force
        Write-RoboLog "config.precos.json criado a partir do exemplo."
    }

    return (Get-Content -LiteralPath $script:RoboPrecosConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json)
}

function Resolve-RoboPrecosPath {
    param([string]$Path)

    if ([IO.Path]::IsPathRooted($Path)) { return $Path }
    return (Join-Path $Root $Path)
}

function Set-PdaCredential {
    param($Config)

    $credentialPath = Resolve-RoboPrecosPath ([string]$Config.pda.credentialFile)
    $credentialDirectory = Split-Path -Parent $credentialPath
    if (-not (Test-Path -LiteralPath $credentialDirectory)) {
        New-Item -ItemType Directory -Path $credentialDirectory -Force | Out-Null
    }

    Write-Host ""
    Write-Host "CREDENCIAL PDA - armazenamento local protegido pelo Windows" -ForegroundColor Cyan
    $username = Read-Host "Usuario PDA"
    $securePassword = Read-Host "Senha PDA" -AsSecureString

    if ([string]::IsNullOrWhiteSpace($username)) {
        throw "Usuario PDA nao informado."
    }

    $payload = [PSCustomObject]@{
        username = $username.Trim()
        password = ($securePassword | ConvertFrom-SecureString)
        createdAt = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    }

    $payload | ConvertTo-Json | Set-Content -LiteralPath $credentialPath -Encoding UTF8
    Write-RoboLog ("Credencial PDA protegida criada em " + $credentialPath)
}

function Get-PdaCredential {
    param($Config)

    $credentialPath = Resolve-RoboPrecosPath ([string]$Config.pda.credentialFile)
    if (-not (Test-Path -LiteralPath $credentialPath)) {
        Set-PdaCredential -Config $Config
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

function Start-RoboPrecosBrowser {
    param($Config)

    $chromePath = Get-RoboChrome
    if (-not $chromePath) {
        throw "Google Chrome nao encontrado."
    }

    $port = [int]$Config.pda.debugPort
    $profilePath = Resolve-RoboPrecosPath ([string]$Config.pda.profileDirectory)
    if (-not (Test-Path -LiteralPath $profilePath)) {
        New-Item -ItemType Directory -Path $profilePath -Force | Out-Null
    }

    $baseUrl = [string]$Config.pda.baseUrl

    $endpointReady = $false
    try {
        $null = Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/json/version" -f $port) -UseBasicParsing -TimeoutSec 1
        $endpointReady = $true
    }
    catch {}

    if (-not $endpointReady) {
        Write-RoboLog ("Abrindo Chrome controlado localmente na porta CDP " + $port)
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
        Write-RoboLog "Chrome controlado ja esta em execucao."
    }

    Wait-CdpEndpoint -Port $port -TimeoutSeconds 30

    # Cria UMA aba dedicada para o PDA e conecta exatamente nela.
    # Nao usa mais o primeiro target do Chrome e nunca conecta em extensoes/background.
    $target = New-CdpPageTarget -Port $port -Url $baseUrl
    Write-RoboLog ("Aba PDA dedicada criada: " + [string]$target.url)

    $wsUrl = [string]$target.webSocketDebuggerUrl
    $wsUrl = $wsUrl -replace 'ws://localhost:', 'ws://127.0.0.1:'
    $wsUrl = $wsUrl -replace 'ws://\[::1\]:', 'ws://127.0.0.1:'

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
        throw ("Falha na conexao com a aba PDA dedicada. Detalhe: " + $detail)
    }

    $probe = Invoke-CdpExpression -Socket $socket -Expression "'ROBO_CDP_OK'"
    if ([string]$probe -ne "ROBO_CDP_OK") {
        throw ("Canal Chrome DevTools conectado, mas Runtime.evaluate nao devolveu o valor esperado. Recebido: " + [string]$probe)
    }
    Write-RoboLog "Canal Chrome DevTools validado (Runtime.evaluate OK)."

    return $socket
}

function Get-PdaPageState {
    param([System.Net.WebSockets.ClientWebSocket]$Socket)

    $expression = @'
(() => {
  try {
    const norm = s => (s || '').replace(/\s+/g,' ').trim().toLowerCase();
    const bodyText = document.body ? document.body.innerText : '';
    const text = norm(bodyText);
    const hasPassword = document.querySelectorAll('input[type="password"]').length > 0;
    const hasLoginWord = text.includes('login') || text.includes('usuario') || text.includes('usuário');
    const hasAuditSelect = [...document.querySelectorAll('select')].some(s =>
      [...s.options].some(o => norm(o.textContent).includes('centerlar comercio de utilidades'))
    );
    const hasSearch = [...document.querySelectorAll('button,input[type="button"],input[type="submit"],a')].some(e =>
      norm(e.innerText || e.value || e.textContent) === 'pesquisar'
    );

    let kind = 'OTHER';
    if (hasPassword || hasLoginWord) kind = 'LOGIN';
    if (hasAuditSelect && hasSearch) kind = 'AUDIT';

    const safe = v => String(v == null ? '' : v).replace(/[|\r\n]+/g,' ').slice(0,180);
    return [
      kind,
      safe(location.href),
      safe(document.readyState),
      document.querySelectorAll('input').length,
      document.querySelectorAll('select').length,
      document.querySelectorAll('iframe').length,
      safe(document.title),
      safe(bodyText)
    ].join('|');
  } catch (e) {
    return 'ERROR|' + String(e && e.message ? e.message : e);
  }
})()
'@

    $raw = [string](Invoke-CdpExpression -Socket $Socket -Expression $expression)
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return [PSCustomObject]@{ kind="EMPTY"; url=""; ready=""; inputCount=0; selectCount=0; iframeCount=0; title=""; snippet=""; login=$false; audit=$false }
    }

    $parts = $raw.Split('|', 8)
    if ($parts[0] -eq "ERROR") {
        throw ("Erro lendo pagina PDA: " + (($parts | Select-Object -Skip 1) -join "|"))
    }

    return [PSCustomObject]@{
        kind = $parts[0]
        url = if ($parts.Count -gt 1) { $parts[1] } else { "" }
        ready = if ($parts.Count -gt 2) { $parts[2] } else { "" }
        inputCount = if ($parts.Count -gt 3) { [int]$parts[3] } else { 0 }
        selectCount = if ($parts.Count -gt 4) { [int]$parts[4] } else { 0 }
        iframeCount = if ($parts.Count -gt 5) { [int]$parts[5] } else { 0 }
        title = if ($parts.Count -gt 6) { $parts[6] } else { "" }
        snippet = if ($parts.Count -gt 7) { $parts[7] } else { "" }
        login = ($parts[0] -eq "LOGIN")
        audit = ($parts[0] -eq "AUDIT")
    }
}

function Wait-PdaRecognizedPage {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [int]$TimeoutSeconds = 30
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $lastState = $null
    $lastError = ""

    do {
        try {
            $lastState = Get-PdaPageState -Socket $Socket
            if ($lastState -and ([bool]$lastState.login -or [bool]$lastState.audit)) {
                return $lastState
            }
        }
        catch {
            $lastError = $_.Exception.Message
        }
        Start-Sleep -Milliseconds 500
    } while ((Get-Date) -lt $deadline)

    if ($lastState) {
        $diag = "kind={0} | URL={1} | ready={2} | inputs={3} | selects={4} | iframes={5} | titulo={6} | trecho={7}" -f $lastState.kind,$lastState.url,$lastState.ready,$lastState.inputCount,$lastState.selectCount,$lastState.iframeCount,$lastState.title,$lastState.snippet
        throw ("PDA nao apresentou tela reconhecida. " + $diag)
    }

    if (-not [string]::IsNullOrWhiteSpace($lastError)) {
        throw ("Nao foi possivel ler o DOM da pagina PDA via Chrome DevTools. Ultimo erro: " + $lastError)
    }

    throw "PDA nao apresentou uma tela reconhecivel dentro do tempo limite."
}

function Invoke-PdaLogin {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        $Config
    )

    $credential = Get-PdaCredential -Config $Config
    $userJs = ($credential.Username | ConvertTo-Json -Compress)
    $passwordJs = ($credential.Password | ConvertTo-Json -Compress)

    $expression = @"
(() => {
  try {
    const norm = s => (s || '').replace(/\s+/g,' ').trim().toLowerCase();
    const visible = e => !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
    const pwd = [...document.querySelectorAll('input[type="password"]')].find(visible);
    if (!pwd) return 'ERROR:password-not-found';

    const inputs = [...document.querySelectorAll('input')].filter(visible);
    const user = inputs.find(e => e !== pwd && ['text','email',''].includes((e.type || '').toLowerCase()));
    if (!user) return 'ERROR:username-not-found';

    const setValue = (el, value) => {
      const proto = el instanceof HTMLInputElement ? HTMLInputElement.prototype : HTMLElement.prototype;
      const desc = Object.getOwnPropertyDescriptor(proto, 'value');
      if (desc && desc.set) desc.set.call(el, value); else el.value = value;
      el.dispatchEvent(new Event('input', {bubbles:true}));
      el.dispatchEvent(new Event('change', {bubbles:true}));
    };

    setValue(user, $userJs);
    setValue(pwd, $passwordJs);

    const controls = [...document.querySelectorAll('button,input[type="submit"],input[type="button"],a')].filter(visible);
    const enter = controls.find(e => norm(e.innerText || e.value || e.textContent) === 'entrar');
    if (!enter) return 'ERROR:enter-not-found';
    enter.click();
    return 'OK';
  } catch (e) {
    return 'ERROR:' + String(e && e.message ? e.message : e);
  }
})()
"@

    $result = [string](Invoke-CdpExpression -Socket $Socket -Expression $expression)
    if ($result -ne "OK") {
        throw ("Nao foi possivel acionar o login PDA: " + $result)
    }

    Write-RoboLog "Login PDA enviado. Aguardando autenticacao."
    $deadline = (Get-Date).AddSeconds([int]$Config.pda.pageLoadTimeoutSeconds)
    do {
        Start-Sleep -Milliseconds 500
        $state = Get-PdaPageState -Socket $Socket
        if (-not [bool]$state.login) {
            Write-RoboLog ("Login PDA concluido. URL atual: " + [string]$state.url)
            return
        }
    } while ((Get-Date) -lt $deadline)

    throw "Timeout aguardando autenticacao no PDA. Confira usuario e senha."
}

function Ensure-PdaAuditPage {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        $Config
    )

    $baseUrl = ([string]$Config.pda.baseUrl).TrimEnd('/')
    $auditUrl = $baseUrl + [string]$Config.pda.auditPath

    # 1) Primeiro reconhece a pagina atual. Se for login, autentica ANTES
    # de tentar abrir diretamente a Auditoria.
    $state = $null
    try {
        $state = Wait-PdaRecognizedPage -Socket $Socket -TimeoutSeconds 15
    }
    catch {
        Write-RoboLog ("Pagina inicial ainda nao reconhecida. Abrindo pagina base do PDA. Detalhe: " + $_.Exception.Message) "AVISO"
        Navigate-Cdp -Socket $Socket -Url $baseUrl -TimeoutSeconds ([int]$Config.pda.pageLoadTimeoutSeconds)
        $state = Wait-PdaRecognizedPage -Socket $Socket -TimeoutSeconds ([int]$Config.pda.pageLoadTimeoutSeconds)
    }

    if ([bool]$state.login) {
        Write-RoboLog "Tela de login detectada. Autenticando antes de abrir a Auditoria."
        Invoke-PdaLogin -Socket $Socket -Config $Config
        Start-Sleep -Milliseconds 800
    }

    # 2) Com a sessao autenticada (ou ja existente), abre a tela de Auditoria.
    Write-RoboLog ("Abrindo tela de Auditoria de Preco: " + $auditUrl)
    Navigate-Cdp -Socket $Socket -Url $auditUrl -TimeoutSeconds ([int]$Config.pda.pageLoadTimeoutSeconds)
    $state = Wait-PdaRecognizedPage -Socket $Socket -TimeoutSeconds ([int]$Config.pda.pageLoadTimeoutSeconds)

    # 3) Se a sessao expirou no meio do caminho, refaz login uma vez e retorna.
    if ([bool]$state.login) {
        Write-RoboLog "Sessao PDA expirou ao abrir Auditoria. Refazendo login." "AVISO"
        Invoke-PdaLogin -Socket $Socket -Config $Config
        Navigate-Cdp -Socket $Socket -Url $auditUrl -TimeoutSeconds ([int]$Config.pda.pageLoadTimeoutSeconds)
        $state = Wait-PdaRecognizedPage -Socket $Socket -TimeoutSeconds ([int]$Config.pda.pageLoadTimeoutSeconds)
    }

    if (-not [bool]$state.audit) {
        throw ("Tela de Auditoria de Preco nao reconhecida apos autenticacao. URL: " + [string]$state.url)
    }

    Write-RoboLog "Tela de Auditoria de Preco pronta."
}

function Get-PdaTotals {
    param([System.Net.WebSockets.ClientWebSocket]$Socket)

    $expression = @'
(() => {
  const norm = s => (s || '').replace(/\s+/g,' ').trim().toLowerCase();
  const lines = (document.body ? document.body.innerText : '').split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  const labels = ['ok','divergente','sem etiqueta','total'];

  const numberAfter = label => {
    const target = norm(label);
    for (let i=0; i<lines.length; i++) {
      if (norm(lines[i]) !== target) continue;
      for (let j=i+1; j<Math.min(lines.length, i+5); j++) {
        const n = norm(lines[j]);
        if (labels.includes(n)) break;
        const cleaned = lines[j].replace(/[.\s]/g,'');
        if (/^\d+$/.test(cleaned)) return parseInt(cleaned,10);
        if (lines[j] === '...') return null;
      }
    }
    return null;
  };

  let busy = false;
  try {
    if (window.Sys && Sys.WebForms && Sys.WebForms.PageRequestManager) {
      busy = Sys.WebForms.PageRequestManager.getInstance().get_isInAsyncPostBack();
    }
  } catch(e) {}

  return {
    url: location.href,
    login: !!document.querySelector('input[type="password"]'),
    ready: document.readyState,
    busy,
    ok: numberAfter('ok'),
    divergente: numberAfter('divergente'),
    semEtiqueta: numberAfter('sem etiqueta'),
    total: numberAfter('total')
  };
})()
'@
    return (Invoke-CdpJsonExpression -Socket $Socket -Expression $expression)
}

function Set-PdaAuditFilters {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [string]$Store,
        [string]$StartDate,
        [string]$EndDate
    )

    $storeNumber = 0
    if ($Store -match '(\d+)') {
        $storeNumber = [int]$Matches[1]
    }
    if ($storeNumber -le 0) {
        throw "Loja invalida: $Store"
    }

    $storeJs = ($storeNumber | ConvertTo-Json -Compress)
    $startJs = ($StartDate | ConvertTo-Json -Compress)
    $endJs = ($EndDate | ConvertTo-Json -Compress)

    $expression = @"
(() => {
  const norm = s => (s || '').replace(/\s+/g,' ').trim().toLowerCase();
  const visible = e => !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
  const setValue = (el, value) => {
    const proto = el instanceof HTMLInputElement ? HTMLInputElement.prototype :
                  el instanceof HTMLSelectElement ? HTMLSelectElement.prototype : HTMLElement.prototype;
    const desc = Object.getOwnPropertyDescriptor(proto, 'value');
    if (desc && desc.set) desc.set.call(el, value); else el.value = value;
    el.dispatchEvent(new Event('input', {bubbles:true}));
    el.dispatchEvent(new Event('change', {bubbles:true}));
  };

  const storeNumber = $storeJs;
  const selects = [...document.querySelectorAll('select')].filter(visible);
  const center = selects.find(s => [...s.options].some(o => norm(o.textContent).includes('centerlar comercio de utilidades')));
  if (!center) return {ok:false, reason:'center-select-not-found'};

  const prefix = storeNumber + '-' + storeNumber;
  const option = [...center.options].find(o => norm(o.textContent).startsWith(norm(prefix + ' -'))) ||
                 [...center.options].find(o => norm(o.textContent).startsWith(norm(prefix)));
  if (!option) return {ok:false, reason:'store-not-found', prefix};
  setValue(center, option.value);

  const dateInputs = [...document.querySelectorAll('input')].filter(e =>
    visible(e) && ((e.type || '').toLowerCase() === 'text' || !(e.type)) &&
    (/^\d{2}\/\d{2}\/\d{4}$/.test((e.value || '').trim()) || /data/i.test((e.id || '') + ' ' + (e.name || '') + ' ' + (e.placeholder || '')))
  );

  if (dateInputs.length < 2) return {ok:false, reason:'date-inputs-not-found', count:dateInputs.length};
  setValue(dateInputs[0], $startJs);
  setValue(dateInputs[1], $endJs);

  const reason = selects.find(s => [...s.options].some(o => norm(o.textContent) === 'todos'));
  if (reason) {
    const all = [...reason.options].find(o => norm(o.textContent) === 'todos');
    if (all) setValue(reason, all.value);
  }

  return {
    ok:true,
    storeText: option.textContent.trim(),
    startDate: dateInputs[0].value,
    endDate: dateInputs[1].value,
    reason: reason ? reason.options[reason.selectedIndex].textContent.trim() : ''
  };
})()
"@

    $result = Invoke-CdpJsonExpression -Socket $Socket -Expression $expression
    if (-not $result.ok) {
        throw ("Falha ao preencher filtros PDA: " + ($result | ConvertTo-Json -Compress))
    }

    Write-RoboLog ("Filtros PDA: {0} | {1} a {2} | Motivo={3}" -f $result.storeText, $result.startDate, $result.endDate, $result.reason)
    return $result
}

function Invoke-PdaSearch {
    param([System.Net.WebSockets.ClientWebSocket]$Socket)

    $expression = @'
(() => {
  const norm = s => (s || '').replace(/\s+/g,' ').trim().toLowerCase();
  const visible = e => !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
  const controls = [...document.querySelectorAll('button,input[type="button"],input[type="submit"],a')].filter(visible);
  const search = controls.find(e => norm(e.innerText || e.value || e.textContent) === 'pesquisar');
  if (!search) return {ok:false, reason:'search-not-found'};
  window.__roboPrecosSearchAt = Date.now();
  search.click();
  return {ok:true, at:window.__roboPrecosSearchAt};
})()
'@

    $result = Invoke-CdpJsonExpression -Socket $Socket -Expression $expression
    if (-not $result.ok) {
        throw ("Botao Pesquisar nao localizado: " + [string]$result.reason)
    }
}

function Wait-PdaTotals {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        $Config
    )

    $deadline = (Get-Date).AddSeconds([int]$Config.pda.queryTimeoutSeconds)
    Start-Sleep -Milliseconds 1200

    do {
        $state = Get-PdaTotals -Socket $Socket

        if ([bool]$state.login) {
            throw "SESSION_EXPIRED"
        }

        $hasAll = ($null -ne $state.ok) -and ($null -ne $state.divergente) -and ($null -ne $state.semEtiqueta) -and ($null -ne $state.total)
        if ($hasAll -and -not [bool]$state.busy) {
            $sum = [int]$state.ok + [int]$state.divergente + [int]$state.semEtiqueta
            if ($sum -eq [int]$state.total) {
                return $state
            }
        }

        Start-Sleep -Milliseconds 650
    } while ((Get-Date) -lt $deadline)

    throw "Timeout aguardando totalizadores validos da auditoria."
}

function Invoke-PdaAuditQuery {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        $Config,
        [string]$Store,
        [string]$StartDate,
        [string]$EndDate
    )

    $attempt = 0
    while ($attempt -lt 2) {
        $attempt++
        try {
            $pageState = Get-PdaPageState -Socket $Socket
            if ([bool]$pageState.login -or -not [bool]$pageState.audit) {
                Ensure-PdaAuditPage -Socket $Socket -Config $Config
            }

            [void](Set-PdaAuditFilters -Socket $Socket -Store $Store -StartDate $StartDate -EndDate $EndDate)
            Invoke-PdaSearch -Socket $Socket
            $totals = Wait-PdaTotals -Socket $Socket -Config $Config

            Write-RoboLog ("Resultado {0}: OK={1}; Divergente={2}; SemEtiqueta={3}; Total={4}" -f $Store, $totals.ok, $totals.divergente, $totals.semEtiqueta, $totals.total)

            return [PSCustomObject]@{
                Loja = ConvertTo-RoboStore $Store
                DataInicio = $StartDate
                DataFim = $EndDate
                Ok = [int]$totals.ok
                Divergente = [int]$totals.divergente
                SemEtiqueta = [int]$totals.semEtiqueta
                Total = [int]$totals.total
                ColetadoEm = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            }
        }
        catch {
            if ($_.Exception.Message -eq "SESSION_EXPIRED" -and $attempt -lt 2) {
                Write-RoboLog "Sessao expirou durante a consulta. Refazendo login e repetindo a loja." "AVISO"
                Ensure-PdaAuditPage -Socket $Socket -Config $Config
                continue
            }
            throw
        }
    }
}
