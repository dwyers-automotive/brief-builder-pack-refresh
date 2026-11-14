<#
    Invoke-PackRefresh.ps1 - the deterministic half of the monthly Brief Builder Knowledge Pack refresh.

    The judgement half (regenerating the 8 .md files) happens first, in a Claude session:
    "AVA, refresh Brief Builder pack". This script then checks and announces the result.
    It stops at the FIRST failure.

      (a) all 8 expected files are present in the Knowledge Pack folder
      (b) 00_README and 04_Project_Index carry the snapshot date (today, dd/mm/yyyy)
      (c) sanitisation passes (Test-PackSanitisation.ps1)
      (d) diff against the newest dated folder in Knowledge Pack\_Archive -> change_summary.md
      (e) unless -NoSend     : email Ant from ava@ (draft -> GET-verify -> /send) with the
                               change summary, the sanitisation result and the upload checklist
      (f) unless -NoHeartbeat: SUCCESS heartbeat for BriefBuilder_PackRefresh in DMG.dbo.IntegrationControl

    Log + outputs: Output\_Refresh Work\<yyyy-MM-dd>\ (refresh.log, sanitisation_report.txt, change_summary.md)
    Ends by printing True or False (exit 0 / 1).

    Runs on Ant's laptop (needs the network for e and f). Paths are built from $env:USERPROFILE.
    Secrets: Graph from the vault [M365_GRAPH_AVA] over UNC; SQL from [SQL_FMT] read ON dmg-web inside
    Invoke-Command, so the SQL credential never leaves the server. Nothing secret is printed or logged.
    -SnapshotDate dd/MM/yyyy: only when this runs on a later day than the regeneration (max 7 days back).

    ASCII only. Windows PowerShell 5.1. Riley, 14/11/2026.
#>
[CmdletBinding()]
param(
    [switch] $NoSend,
    [switch] $NoHeartbeat,
    [string] $SnapshotDate,
    [string] $PackPath,
    [string] $ArchiveRoot,
    [string] $PreviousPath,
    [string] $WorkRoot,
    [string] $IniPath  = '\\dmg-web\c$\Secure\dwyers_apps.ini',
    [string] $AntEmail = 'adwyer@dwyers.com.au',
    [string] $FromUser = 'ava@dwyers.com.au',
    [string] $SendVia  = 'adwyer@dwyers.com.au',   # mailbox the Graph call posts through (see M365-Graph-Connector.ps1)
    [string] $SqlHost  = 'dmg-web'
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Process  = 'BriefBuilder_PackRefresh'
$Base     = Join-Path $env:USERPROFILE 'OneDrive\PKA_Claude\Projects\Brief Builder'
$Scripts  = Join-Path $Base 'Scripts'
if (-not $PackPath)    { $PackPath    = Join-Path $Base 'Output\Knowledge Pack' }
if (-not $ArchiveRoot) { $ArchiveRoot = Join-Path $Base 'Output\Knowledge Pack\_Archive' }
if (-not $WorkRoot)    { $WorkRoot    = Join-Path $Base 'Output\_Refresh Work' }
$Expected = @('00_README.md', '01_Company_Background.md', '02_Tech_Stack.md', '03_Team_Roster.md',
              '04_Project_Index.md', '05_DMS_and_Integrations.md', '06_Lessons_and_Pitfalls.md', '07_Example_Briefs.md')
$Utf8  = New-Object Text.UTF8Encoding($false)
$G     = 'https://graph.microsoft.com/v1.0'
$Dash  = [string][char]0x2014

$today = (Get-Date).Date
if ($SnapshotDate) {
    $snap = [datetime]::ParseExact($SnapshotDate, 'dd/MM/yyyy', [Globalization.CultureInfo]::InvariantCulture)
} else { $snap = $today }
$snapText = $snap.ToString('dd/MM/yyyy', [Globalization.CultureInfo]::InvariantCulture)
$WorkDir  = Join-Path $WorkRoot $snap.ToString('yyyy-MM-dd')
New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
$LogFile  = Join-Path $WorkDir 'refresh.log'

function Log {
    param([string]$Msg, [string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Msg
    Write-Host $line
    [IO.File]::AppendAllText($LogFile, $line + "`r`n", $Utf8)
}
function Enc { param($s) [System.Net.WebUtility]::HtmlEncode([string]$s) }

function Get-VaultSection {
    param([string]$Path, [string]$Section)
    $kv = @{}; $in = $false
    foreach ($l in [IO.File]::ReadAllLines($Path)) {
        $t = $l.Trim()
        if ($t -match '^\[(.+)\]$') { $in = ($Matches[1] -eq $Section); continue }
        if ($in -and $t -match '^([^;#=]+?)\s*=\s*(.*)$') { $kv[$Matches[1].Trim()] = $Matches[2].Trim() }
    }
    return $kv
}

function Get-ErrorBody {
    param($Err)
    try {
        $resp = $Err.Exception.Response
        if ($resp) { $sr = New-Object IO.StreamReader($resp.GetResponseStream()); return $sr.ReadToEnd() }
    } catch {}
    return ''
}

function Invoke-Graph {
    # UTF-8 safe in both directions (PS 5.1 mangles UTF-8 on Invoke-RestMethod bodies and responses).
    param([string]$Uri, [string]$Token, [string]$Method = 'GET', $Body)
    $p = @{ Uri = $Uri; Method = $Method; Headers = @{ Authorization = "Bearer $Token" }; UseBasicParsing = $true }
    if ($null -ne $Body) { $p.ContentType = 'application/json; charset=utf-8'; $p.Body = $Utf8.GetBytes(($Body | ConvertTo-Json -Depth 10)) }
    try { $wr = Invoke-WebRequest @p }
    catch { throw ("Graph {0} {1} failed: {2} {3}" -f $Method, ($Uri -replace '\?.*$', ''), $_.Exception.Message, (Get-ErrorBody $_)) }
    $txt = $Utf8.GetString($wr.RawContentStream.ToArray())
    if ($txt) { return ($txt | ConvertFrom-Json) } else { return $null }
}

function Get-GraphToken {
    if (-not (Test-Path -LiteralPath $IniPath)) { throw "Vault not reachable: $IniPath (are you on the internal network?)" }
    $c = Get-VaultSection -Path $IniPath -Section 'M365_GRAPH_AVA'
    foreach ($k in 'TenantId', 'ClientId', 'ClientSecret') { if (-not $c.ContainsKey($k) -or -not $c[$k]) { throw "Vault [M365_GRAPH_AVA] is missing $k" } }
    $body = @{ client_id = $c.ClientId; client_secret = $c.ClientSecret; scope = 'https://graph.microsoft.com/.default'; grant_type = 'client_credentials' }
    try { $r = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$($c.TenantId)/oauth2/v2.0/token" -ContentType 'application/x-www-form-urlencoded' -Body $body }
    catch { throw "Graph token request failed: $($_.Exception.Message)" }
    finally { $c = $null; $body = $null }
    return $r.access_token
}

function ConvertTo-SummaryHtml {
    param($d)
    $sb = New-Object System.Text.StringBuilder
    $h3 = 'style="font-size:15px;margin:16px 0 4px 0;color:#1F3864;"'
    $ul = 'style="margin:4px 0 8px 18px;padding:0;"'
    function li($s) { [void]$sb.AppendLine("<li>$s</li>") }
    if (-not $d.HasChanges) { [void]$sb.AppendLine('<p><b>No changes</b> - this version is identical to the previous one.</p>') }
    [void]$sb.AppendLine("<h3 $h3>04 Project Index ($($d.ProjectRowsOld) &rarr; $($d.ProjectRowsNew) rows)</h3><ul $ul>")
    li ("<b>Added ($($d.ProjectsAdded.Count)):</b> " + $(if ($d.ProjectsAdded.Count) { (@($d.ProjectsAdded | ForEach-Object { (Enc $_.Project) + ' <i>(' + (Enc $_.Status) + ')</i>' }) -join '; ') } else { 'none' }))
    li ("<b>Removed ($($d.ProjectsRemoved.Count)):</b> " + $(if ($d.ProjectsRemoved.Count) { (@($d.ProjectsRemoved | ForEach-Object { Enc $_.Project }) -join '; ') } else { 'none' }))
    li ("<b>Status changed ($($d.StatusChanged.Count)):</b> " + $(if ($d.StatusChanged.Count) { (@($d.StatusChanged | ForEach-Object { (Enc $_.Project) + ': ' + (Enc $_.Old) + ' &rarr; ' + (Enc $_.New) }) -join '; ') } else { 'none' }))
    [void]$sb.AppendLine('</ul>')
    [void]$sb.AppendLine("<h3 $h3>03 Team Roster</h3><ul $ul>")
    li ('<b>Added:</b> ' + $(if ($d.RosterAdded.Count) { Enc ($d.RosterAdded -join ', ') } else { 'none' }))
    li ('<b>Removed:</b> ' + $(if ($d.RosterRemoved.Count) { Enc ($d.RosterRemoved -join ', ') } else { 'none' }))
    [void]$sb.AppendLine('</ul>')
    [void]$sb.AppendLine("<h3 $h3>06 Lessons &amp; Pitfalls</h3><ul $ul>")
    li ("<b>Added ($($d.LessonsAdded.Count)):</b> " + $(if ($d.LessonsAdded.Count) { (@($d.LessonsAdded | ForEach-Object { Enc ($_ -replace '\*\*', '') }) -join '<br>') } else { 'none' }))
    li ("<b>Removed ($($d.LessonsRemoved.Count)):</b> " + $(if ($d.LessonsRemoved.Count) { (@($d.LessonsRemoved | ForEach-Object { Enc ($_ -replace '\*\*', '') }) -join '<br>') } else { 'none' }))
    [void]$sb.AppendLine('</ul>')
    [void]$sb.AppendLine("<h3 $h3>All files</h3><table cellpadding=""4"" style=""border-collapse:collapse;font-size:13px;""><tr style=""background:#EEF2F7;""><th align=""left"">File</th><th>Changed</th><th>Lines (old &rarr; new)</th></tr>")
    foreach ($f in $d.Files) { [void]$sb.AppendLine("<tr><td>$(Enc $f.File)</td><td align=""center"">$(if ($f.Changed) { 'yes' } else { 'no' })</td><td align=""center"">$($f.OldLines) &rarr; $($f.NewLines)</td></tr>") }
    [void]$sb.AppendLine('</table>')
    return $sb.ToString()
}

$step = 'start'
try {
    Log "=== $Process started (snapshot $snapText; send=$(-not $NoSend); heartbeat=$(-not $NoHeartbeat)) ==="
    Log "Pack: $PackPath"
    if ($SnapshotDate) {
        if ($snap -gt $today -or $snap -lt $today.AddDays(-7)) { throw "SnapshotDate $snapText must be between 7 days ago and today." }
        Log "SnapshotDate override in use: $snapText (run date $($today.ToString('dd/MM/yyyy')))" 'WARN'
    }

    # (a) files present
    $step = '(a) files present'
    if (-not (Test-Path -LiteralPath $PackPath -PathType Container)) { throw "Pack folder not found: $PackPath" }
    $missing = @($Expected | Where-Object { -not (Test-Path -LiteralPath (Join-Path $PackPath $_) -PathType Leaf) })
    if ($missing.Count) { throw "Missing file(s): $($missing -join ', ')" }
    $extra = @(Get-ChildItem -LiteralPath $PackPath -Filter '*.md' -File | Where-Object { $_.Name -notin $Expected } | ForEach-Object Name)
    if ($extra.Count) { Log "Extra .md file(s) in the pack folder (not part of the 8; do not upload): $($extra -join ', ')" 'WARN' }
    Log "(a) OK - all 8 files present"

    # (b) snapshot date in 00 and 04
    $step = '(b) snapshot date'
    foreach ($name in '00_README.md', '04_Project_Index.md') {
        $head = @([IO.File]::ReadAllLines((Join-Path $PackPath $name), [Text.Encoding]::UTF8) | Select-Object -First 15)
        $found = $null
        foreach ($l in $head) { if ($l -match '(?i)(generated|snapshot)\W{0,6}(?:as at\s+)?(\d{2}/\d{2}/\d{4})') { $found = $Matches[2]; break } }
        if (-not $found) { throw "$name has no 'Generated:' / 'snapshot' date in its first 15 lines" }
        if ($found -ne $snapText) { throw "$name is dated $found, expected $snapText. Regenerate the pack (or set its date) first." }
        Log "(b) OK - $name dated $found"
    }

    # (c) sanitisation
    $step = '(c) sanitisation'
    $scanReport = Join-Path $WorkDir 'sanitisation_report.txt'
    $scan = & (Join-Path $Scripts 'Test-PackSanitisation.ps1') -PackPath $PackPath -ReportPath $scanReport -PassThru
    $scanExit = $LASTEXITCODE
    if ($scanExit -ne 0 -or -not $scan -or -not $scan.Clean) {
        $n = if ($scan) { $scan.HitCount } else { '?' }
        throw "Sanitisation FAILED (exit $scanExit, $n hit(s)). See $scanReport"
    }
    Log "(c) OK - sanitisation clean: $($scan.Files) files, 0 hits"

    # (d) diff vs previous version
    $step = '(d) change summary'
    if (-not $PreviousPath) {
        $packFull = (Resolve-Path -LiteralPath $PackPath).Path.TrimEnd('\')
        $cands = @(Get-ChildItem -LiteralPath $ArchiveRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            if ($_.Name -match '(\d{4}-\d{2}-\d{2})' -and $_.FullName.TrimEnd('\') -ne $packFull) {
                $d = [datetime]::ParseExact($Matches[1], 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
                if ($d -lt $snap) { [pscustomobject]@{ Path = $_.FullName; Date = $d; Written = $_.LastWriteTime } }
            } })
        if (-not $cands.Count) { throw "No previous version (a yyyy-MM-dd folder dated before $snapText) in $ArchiveRoot" }
        $PreviousPath = ($cands | Sort-Object Date, Written -Descending | Select-Object -First 1).Path
    }
    Log "Previous version: $PreviousPath"
    $summaryPath = Join-Path $WorkDir 'change_summary.md'
    $diff = & (Join-Path $Scripts 'Compare-PackVersions.ps1') -OldPath $PreviousPath -NewPath $PackPath -OutPath $summaryPath
    if (-not $diff) { throw 'Compare-PackVersions.ps1 returned nothing' }
    Log ("(d) OK - {0}; 04 rows {1} -> {2}; +{3} / -{4} projects; {5} status change(s); roster +{6} / -{7}; lessons +{8} / -{9}" -f `
        $summaryPath, $diff.ProjectRowsOld, $diff.ProjectRowsNew, $diff.ProjectsAdded.Count, $diff.ProjectsRemoved.Count, $diff.StatusChanged.Count,
        $diff.RosterAdded.Count, $diff.RosterRemoved.Count, $diff.LessonsAdded.Count, $diff.LessonsRemoved.Count)

    # (e) email Ant
    $step = '(e) email'
    if ($NoSend) { Log '(e) skipped (-NoSend)' }
    else {
        $subject = "Brief Builder pack refreshed $snapText $Dash upload to claude.ai"
        $fileList = ($Expected | ForEach-Object { "<li>$(Enc $_)</li>" }) -join ''
        $html = @"
<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#222;">
<p>The Brief Builder Knowledge Pack has been regenerated (snapshot <b>$snapText</b>) and passed its checks.</p>
<p><b>Sanitisation:</b> PASSED $Dash $($scan.Files) files scanned, 0 hits.</p>
<h2 style="font-size:17px;color:#1F3864;margin:18px 0 6px 0;">Upload checklist</h2>
<ol style="margin:4px 0 8px 18px;padding:0;">
<li>Open claude.ai and go to the <b>Brief Builder</b> Project.</li>
<li>Open <b>Project knowledge</b> and <b>delete the 8 old files</b>.</li>
<li>Upload the 8 files from <b>OneDrive\PKA_Claude\Projects\Brief Builder\Output\Knowledge Pack\</b> (not the _Archive folder):<ul style="margin:4px 0 4px 18px;padding:0;">$fileList</ul></li>
<li><b>Start a test brief</b> (any one-line idea) and check it reflects the new snapshot.</li>
</ol>
<p style="color:#666;font-size:12px;">The heartbeat proves the pack was generated, not uploaded. Uploading is the step only you can do.</p>
<h2 style="font-size:17px;color:#1F3864;margin:18px 0 6px 0;">What changed since $(Enc (Split-Path -Leaf $PreviousPath))</h2>
$(ConvertTo-SummaryHtml $diff)
<p style="color:#666;font-size:12px;margin-top:18px;">Sent by Invoke-PackRefresh.ps1 (Brief Builder). Log and full summary: Projects\Brief Builder\Output\_Refresh Work\$($snap.ToString('yyyy-MM-dd'))\</p>
</div>
"@
        $token = Get-GraphToken
        $msg = @{
            subject      = $subject
            body         = @{ contentType = 'HTML'; content = $html }
            from         = @{ emailAddress = @{ address = $FromUser } }
            toRecipients = @(@{ emailAddress = @{ address = $AntEmail } })
        }
        $draft = Invoke-Graph -Uri "$G/users/$SendVia/messages" -Token $token -Method Post -Body $msg
        if (-not $draft -or -not $draft.id) { throw 'Draft was not created (no id returned)' }
        Log "Draft created; verifying before send"
        $back = Invoke-Graph -Uri "$G/users/$SendVia/messages/$($draft.id)?`$select=subject,body,toRecipients,from" -Token $token
        $to   = @($back.toRecipients | ForEach-Object { $_.emailAddress.address })
        $text = [string]$back.body.content
        $problems = @()
        if ([string]::IsNullOrWhiteSpace($back.subject) -or $back.subject -ne $subject) { $problems += "subject mismatch ('$($back.subject)')" }
        if ([string]::IsNullOrWhiteSpace($text)) { $problems += 'body empty' }
        foreach ($m in 'Upload checklist', 'Project knowledge', $snapText, 'Sanitisation:') { if ($text.IndexOf($m) -lt 0) { $problems += "body lacks '$m'" } }
        if ($text -match ([string][char]0x00E2 + [string][char]0x20AC)) { $problems += 'body has mojibake' }
        if ($to -notcontains $AntEmail) { $problems += "recipient is '$($to -join ',')'" }
        if ($back.from -and $back.from.emailAddress.address -ne $FromUser) { $problems += "from is '$($back.from.emailAddress.address)'" }
        if ($problems.Count) {
            throw "Draft failed verification, NOT sent (draft left in $SendVia Drafts): $($problems -join '; ')"
        }
        Invoke-Graph -Uri "$G/users/$SendVia/messages/$($draft.id)/send" -Token $token -Method Post | Out-Null
        Log "(e) SENT to $AntEmail from $FromUser : $subject"
        # Best effort: confirm the copy in Sent Items has a body (a successful /send is not proof of content).
        $seen = $false
        for ($t = 0; $t -lt 6 -and -not $seen; $t++) {
            Start-Sleep -Seconds 5
            try {
                $sent = Invoke-Graph -Uri "$G/users/$SendVia/mailFolders/sentitems/messages?`$top=10&`$orderby=sentDateTime desc&`$select=subject,body" -Token $token
                $hit = @($sent.value | Where-Object { $_.subject -eq $subject -and ([string]$_.body.content).IndexOf('Upload checklist') -ge 0 })
                if ($hit.Count) { $seen = $true }
            } catch { Log "Sent Items check: $($_.Exception.Message)" 'WARN' }
        }
        if ($seen) { Log '(e) Sent Items copy verified (subject + body)' } else { Log '(e) Sent copy not seen within 30 s - check Sent Items by hand' 'WARN' }
    }

    # (f) heartbeat
    $step = '(f) heartbeat'
    if ($NoHeartbeat) { Log '(f) skipped (-NoHeartbeat)' }
    else {
        $hbMsg = ("Pack {0}: clean scan, {1} rows in 04, +{2}/-{3} projects, {4} status changes; email {5}" -f $snapText, $diff.ProjectRowsNew,
                  $diff.ProjectsAdded.Count, $diff.ProjectsRemoved.Count, $diff.StatusChanged.Count, $(if ($NoSend) { 'skipped' } else { 'sent' }))
        $rows = Invoke-Command -ComputerName $SqlHost -ErrorAction Stop -ArgumentList $Process, $hbMsg -ScriptBlock {
            param($p, $m)
            $ErrorActionPreference = 'Stop'
            $sec = $null; $kv = @{}
            foreach ($l in [IO.File]::ReadAllLines('C:\Secure\dwyers_apps.ini')) {
                if ($l -match '^\s*\[(.+)\]') { $sec = $Matches[1]; continue }
                if ($sec -eq 'SQL_FMT' -and $l -match '^\s*(Server|UID|PWD)\s*=\s*(.*)$') { $kv[$Matches[1]] = $Matches[2].Trim() }
            }
            if ($kv.Count -ne 3) { throw '[SQL_FMT] Server/UID/PWD not found in the server vault' }
            $cs = "Server=$($kv['Server']);Database=DMG;User Id=$($kv['UID']);Password=$($kv['PWD']);TrustServerCertificate=True;Connection Timeout=30"
            $kv = $null
            $c = New-Object System.Data.SqlClient.SqlConnection $cs
            $cs = $null
            $c.Open()
            try {
                $cmd = $c.CreateCommand()
                $cmd.CommandText = "UPDATE dbo.IntegrationControl
                                    SET LastAttemptedRun=GETDATE(), LastSuccessfulRun=GETDATE(), LastStatus='SUCCESS',
                                        LastMessage=@m, ModifiedAt=GETDATE()
                                    WHERE ProcessName=@p"
                [void]$cmd.Parameters.AddWithValue('@m', $m.Substring(0, [Math]::Min(400, $m.Length)))
                [void]$cmd.Parameters.AddWithValue('@p', $p)
                return $cmd.ExecuteNonQuery()
            } finally { $c.Close() }
        }
        if ([int]$rows -ne 1) { throw "Heartbeat updated $rows row(s); expected 1. Is the row registered? Run Register-BriefBuilderHeartbeat.sql." }
        Log "(f) OK - SUCCESS heartbeat written for $Process on $SqlHost"
    }

    Log "=== $Process completed OK ==="
    Write-Host 'True'
    exit 0
}
catch {
    Log "FAILED at $step : $($_.Exception.Message)" 'ERROR'
    Log "=== $Process stopped ===" 'ERROR'
    Write-Host 'False'
    exit 1
}
