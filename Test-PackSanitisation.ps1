<#
    Test-PackSanitisation.ps1 - Brief Builder Knowledge Pack sanitisation gate.

    Scans every .md file in -PackPath (top level only; add -Recurse to go deeper) and reports
    each hit as  file:line:rule:snippet . Snippets for possible secrets are masked.

    Rules (enforce BRIEF_FOR_TEAM.md, "Sanitisation rules (hard)"):
      IPV4, UNC_PATH, DRIVE_PATH, PORT (host:port, :port), URL / HOSTNAME (any host not
      listed as "host:" in the allow file), EMAIL, PHONE_AU, GUID, INI_SECTION ([UPPER_SNAKE]),
      VAULT_REF (.ini files, the vault file name), SECRET (password, pwd=, secret, client
      secret, api key, token, bearer, connection-string parts), CODE_FENCE, SQL, DMS_TABLE
      (XX_XXXX style), DBO (dbo.), EXCLUDED (any phrase in the exclusion file, case-insensitive,
      whole phrase).

    Allow file: one entry per line, # = comment.
      host: example.com.au   -> URL/HOSTNAME hits on exactly that host (or www.<host>) pass.
      any other line         -> an exact phrase known to be safe. A hit that lies wholly
                                inside an occurrence of the phrase on that line is ignored.
    Exclusion file: one phrase per line, # = comment. It holds names, so it is gitignored.

    Exit: 0 = clean, 1 = hits, 2 = could not run (missing pack / exclusion file).
    -PassThru also emits a result object (Clean, HitCount, Files, Hits).

    ASCII only. Windows PowerShell 5.1. Riley, 14/11/2026.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $PackPath,
    [string] $ExclusionFile,
    [string] $AllowFile,
    [string] $ReportPath,
    [switch] $Recurse,
    [switch] $PassThru
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $ExclusionFile) { $ExclusionFile = Join-Path $here 'Sanitisation-Exclusions.txt' }
if (-not $AllowFile)     { $AllowFile     = Join-Path $here 'Sanitisation-Allow.txt' }

function Read-ListFile {
    param([string]$Path)
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($l in [IO.File]::ReadAllLines($Path, [Text.Encoding]::UTF8)) {
        $t = $l.Trim()
        if ($t -and -not $t.StartsWith('#')) { $out.Add($t) }
    }
    return ,$out
}

if (-not (Test-Path -LiteralPath $PackPath -PathType Container)) { Write-Host "Pack folder not found: $PackPath"; exit 2 }
if (-not (Test-Path -LiteralPath $ExclusionFile)) {
    Write-Host "Exclusion file not found: $ExclusionFile (it is gitignored; it lives only in the OneDrive Scripts folder). Not scanning."
    exit 2
}

$exclusions   = Read-ListFile $ExclusionFile
$allowPhrases = New-Object System.Collections.Generic.List[string]
$allowHosts   = New-Object System.Collections.Generic.List[string]
if (Test-Path -LiteralPath $AllowFile) {
    foreach ($e in (Read-ListFile $AllowFile)) {
        if ($e -match '^(?i)host:\s*(.+)$') { $allowHosts.Add($Matches[1].Trim().ToLowerInvariant()) } else { $allowPhrases.Add($e) }
    }
}

$IC = [Text.RegularExpressions.RegexOptions]::IgnoreCase
$CS = [Text.RegularExpressions.RegexOptions]::None
# Rule table. Mask = how the snippet is shown: none | value (mask the match and what follows) | partial.
$rules = @(
    @{ Id='IPV4';        Mask='none';    Opt=$CS; Rx='(?<![\d.])(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|1?\d?\d)(?![\d.]*\d)' }
    @{ Id='UNC_PATH';    Mask='none';    Opt=$CS; Rx='\\\\[A-Za-z0-9._$-]+\\' }
    @{ Id='DRIVE_PATH';  Mask='none';    Opt=$CS; Rx='(?<![A-Za-z0-9])[A-Za-z]:[\\/](?![/\\])' }
    @{ Id='PORT';        Mask='none';    Opt=$CS; Rx='\b(?=[A-Za-z0-9.-]*[A-Za-z])[A-Za-z0-9][A-Za-z0-9.-]*:\d{2,5}\b|(?<![\w:]):\d{4,5}\b' }
    @{ Id='URL';         Mask='none';    Opt=$IC; Rx='\b(?:https?|ftp|file)://[^\s<>()\[\]"''`]+|\bwww\.[a-z0-9-]+(?:\.[a-z0-9-]+)+' }
    @{ Id='HOSTNAME';    Mask='none';    Opt=$IC; Rx='\b(?:[a-z0-9-]+\.)+(?:com|net|org|au|cloud|io|ai|local|lan|corp|internal|app|dev|biz|info|gov)\b' }
    @{ Id='EMAIL';       Mask='none';    Opt=$IC; Rx='\b[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}\b' }
    @{ Id='PHONE_AU';    Mask='none';    Opt=$CS; Rx='(?<!\d)(?:\+?61[\s-]?|0)4\d{2}[\s-]?\d{3}[\s-]?\d{3}(?!\d)|(?<!\d)(?:\+?61[\s-]?|0)[2378][\s-]?\d{4}[\s-]?\d{4}(?!\d)|\(0[2378]\)\s?\d{4}[\s-]?\d{4}(?!\d)|(?<!\d)1[38]00[\s-]?\d{3}[\s-]?\d{3}(?!\d)|(?<!\d)13\s\d{2}\s\d{2}(?!\d)|(?<!\d)5\d{3}[\s-]\d{4}(?!\d)' }
    @{ Id='GUID';        Mask='partial'; Opt=$IC; Rx='\b[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\b' }
    @{ Id='INI_SECTION'; Mask='none';    Opt=$CS; Rx='\[[A-Z][A-Z0-9]*(?:_[A-Z0-9]+)+\](?!\()|^\s*\[[A-Z][A-Z0-9_]{2,}\]\s*$' }
    @{ Id='VAULT_REF';   Mask='none';    Opt=$IC; Rx='\b[\w-]+\.ini\b|\bdwyers_apps\b|\\Secure\\' }
    @{ Id='SECRET';      Mask='value';   Opt=$IC; Rx='\bpass(?:word|wd)\b|\bpwd\s*=|\bclient[_ -]?secret\b|\bsecrets?\b|\bapi[ _-]?keys?\b|\btokens?\b|\bbearer\b|\b(?:user\s?id|uid)\s*=|\bserver\s*=[^;\s]*;|\bdata source\s*=|\binitial catalog\s*=|\bdatabase\s*=[^;\s]*;|\btrusted_connection\b|\bconnection\s?string\b|\bconn_?str\b|\bprivate key\b|\bsk-ant-[\w-]+|\bAKIA[0-9A-Z]{16}\b' }
    @{ Id='CODE_FENCE';  Mask='none';    Opt=$CS; Rx='^\s*(?:```|~~~)' }
    @{ Id='SQL';         Mask='none';    Opt=$CS; Rx='\bSELECT\b.{0,200}?\bFROM\b|\bINSERT\s+INTO\b|\bUPDATE\s+[\w.\[\]]+\s+SET\b|\bDELETE\s+FROM\b|\bMERGE\s+(?:INTO\s+)?[\w.\[\]]+\s+(?:AS|USING)\b|\b(?:CREATE|ALTER|DROP)\s+(?:TABLE|VIEW|PROCEDURE|INDEX)\b|\bJOIN\s+[\w.\[\]]+\s+(?:\w+\s+)?ON\b|(?i:\bselect\s+(?:\*|top\s+\d+|distinct)\s)' }
    @{ Id='DMS_TABLE';   Mask='none';    Opt=$CS; Rx='\b[A-Z]{2}_[A-Z0-9]{4,}\b' }
    @{ Id='DBO';         Mask='none';    Opt=$IC; Rx='\bdbo\.' }
)
foreach ($r in $rules) { $r.Re = New-Object Text.RegularExpressions.Regex($r.Rx, $r.Opt) }

# Exclusion phrases: case-insensitive, whole phrase, any run of whitespace between words.
$exRules = foreach ($p in $exclusions) {
    $body = ([regex]::Escape($p)) -replace '(\\ )+', '\s+'
    @{ Phrase = $p; Re = New-Object Text.RegularExpressions.Regex("(?<![\p{L}\p{N}])$body(?![\p{L}\p{N}])", $IC) }
}
$allowRes = foreach ($p in $allowPhrases) { New-Object Text.RegularExpressions.Regex([regex]::Escape($p), $IC) }

function Get-Snippet {
    param([string]$Line, [int]$Index, [int]$Length, [string]$Mask)
    $s = [Math]::Max(0, $Index - 30)
    $pre = $Line.Substring($s, $Index - $s)
    $hit = $Line.Substring($Index, $Length)
    $endPost = [Math]::Min($Line.Length, $Index + $Length + 30)
    $post = $Line.Substring($Index + $Length, $endPost - ($Index + $Length))
    switch ($Mask) {
        'value'   { $hit = $hit.Substring(0, [Math]::Min(3, $hit.Length)) + '****'; $post = ' [rest of line masked]' }
        'partial' { $hit = $hit.Substring(0, 4) + '****-****' }
    }
    $snip = ($pre + '>>' + $hit + '<<' + $post) -replace '\s+', ' '
    return $snip.Trim()
}

function Get-UrlHost {
    param([string]$Text)
    $h = $Text -replace '^(?i)(?:https?|ftp|file)://', ''
    $h = ($h -split '[/:?#]')[0]
    return $h.Trim('.').ToLowerInvariant()
}

$files = @(Get-ChildItem -LiteralPath $PackPath -Filter '*.md' -File -Recurse:$Recurse | Sort-Object FullName)
$hits  = New-Object System.Collections.Generic.List[object]

foreach ($f in $files) {
    $rel = if ($Recurse) { $f.FullName.Substring((Resolve-Path -LiteralPath $PackPath).Path.TrimEnd('\').Length + 1) } else { $f.Name }
    $lines = [IO.File]::ReadAllLines($f.FullName, [Text.Encoding]::UTF8)
    for ($i = 0; $i -lt $lines.Length; $i++) {
        $line = $lines[$i]
        if (-not $line.Trim()) { continue }
        # Spans of allowed phrases on this line.
        $spans = @(foreach ($ar in $allowRes) { foreach ($m in $ar.Matches($line)) { ,@($m.Index, ($m.Index + $m.Length)) } })
        $add = {
            param($ruleId, $m, $mask)
            foreach ($sp in $spans) { if ($m.Index -ge $sp[0] -and ($m.Index + $m.Length) -le $sp[1]) { return } }
            $hits.Add([pscustomobject]@{
                File = $rel; Line = $i + 1; Rule = $ruleId
                Snippet = (Get-Snippet -Line $line -Index $m.Index -Length $m.Length -Mask $mask)
            })
        }
        foreach ($r in $rules) {
            foreach ($m in $r.Re.Matches($line)) {
                if ($r.Id -in @('URL', 'HOSTNAME')) {
                    if ($allowHosts.Contains((Get-UrlHost $m.Value)) -or $allowHosts.Contains(((Get-UrlHost $m.Value) -replace '^www\.', ''))) { continue }
                    # A hostname inside an email address is reported once, as EMAIL.
                    if ($r.Id -eq 'HOSTNAME' -and $m.Index -gt 0 -and $line[$m.Index - 1] -eq '@') { continue }
                }
                & $add $r.Id $m $r.Mask
            }
        }
        foreach ($x in $exRules) { foreach ($m in $x.Re.Matches($line)) { & $add 'EXCLUDED' $m 'none' } }
    }
}

$report = New-Object System.Text.StringBuilder
[void]$report.AppendLine("Sanitisation scan - $((Get-Date).ToString('dd/MM/yyyy HH:mm'))")
[void]$report.AppendLine("Pack: $PackPath")
[void]$report.AppendLine("Files scanned: $($files.Count); exclusion phrases: $($exclusions.Count); allow phrases: $($allowPhrases.Count); allowed hosts: $($allowHosts.Count)")
[void]$report.AppendLine("Result: $(if ($hits.Count) { "FAIL - $($hits.Count) hit(s)" } else { 'CLEAN - 0 hits' })")
[void]$report.AppendLine('')
foreach ($h in $hits) { [void]$report.AppendLine(('{0}:{1}:{2}:{3}' -f $h.File, $h.Line, $h.Rule, $h.Snippet)) }
$text = $report.ToString()
Write-Host $text
if ($ReportPath) {
    $dir = Split-Path -Parent $ReportPath
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    [IO.File]::WriteAllText($ReportPath, $text, (New-Object Text.UTF8Encoding($true)))   # BOM: PS 5.1 Get-Content reads it right
}
if ($PassThru) {
    [pscustomobject]@{ Clean = ($hits.Count -eq 0); HitCount = $hits.Count; Files = $files.Count; Hits = $hits.ToArray() }
}
if ($hits.Count) { exit 1 } else { exit 0 }
