<#
    Invoke-PackRefresh.ps1 - the deterministic half of the WEEKLY Knowledge Pack refresh
    (shared by the claude.ai Projects "Brief Builder" and "Ideas Capture").

    The judgement half (regenerating the 8 .md files) happens first, in a Claude session:
    "AVA, refresh Brief Builder pack". The pack folder is a git repo (private
    dwyers-automotive/ava-knowledge-pack, branch main); the claude.ai Projects pick it up via
    GitHub "Sync now". This script stops at the FIRST failure.

      (a) pack repo is a git work tree on main, and all 8 expected files are present
      (b) 00_README and 04_Project_Index carry the snapshot date (today, dd/mm/yyyy)
      (c) sanitisation passes (Test-PackSanitisation.ps1) - this is the PUSH GATE
      (d) diff of the working tree against the repo's HEAD (git show HEAD:<file> into a temp folder)
          -> change_summary.md. If local main is ahead of origin/main (an earlier push failed), the base is
          origin/main and the stranded commit(s) are pushed in (e).
      (e) unless -NoPush     : git add 0*.md, commit "Knowledge Pack refresh <dd/mm/yyyy>", git push origin main,
                               then verify with ls-remote. No changes vs HEAD -> no commit, no push.
      (f) unless -NoSend     : email Ant from ava@ (draft -> GET-verify -> /send): change summary, scan result,
                               sync checklist. No changes -> a short "no changes" note.
      (g) unless -NoHeartbeat: SUCCESS heartbeat for BriefBuilder_PackRefresh in DMG.dbo.IntegrationControl
                               (written whether or not anything changed).

    Log + outputs: <Brief Builder>\Output\_Refresh Work\<yyyy-MM-dd>\ (refresh.log, sanitisation_report.txt,
    change_summary.md). Ends by printing True or False (exit 0 / 1).

    Runs on Ant's laptop (needs the network for e, f and g). Paths are built from $env:USERPROFILE.
    Secrets: Graph from the vault [M365_GRAPH_AVA] over UNC; SQL from [SQL_FMT] read ON dmg-web inside
    Invoke-Command, so the SQL credential never leaves the server. Nothing secret is printed or logged.
    Git push uses the laptop's Git Credential Manager (it may pop a GitHub sign-in).
    -SnapshotDate dd/MM/yyyy: only when this runs on a later day than the regeneration (max 7 days back).

    ASCII only. Windows PowerShell 5.1. Riley, 14/11/2026 (weekly + git, same day).
#>
[CmdletBinding()]
param(
    [switch] $NoPush,
    [switch] $NoSend,
    [switch] $NoHeartbeat,
    [string] $SnapshotDate,
    [string] $PackPath,
    [string] $WorkRoot,
    [string] $Branch   = 'main',
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
if (-not $PackPath) { $PackPath = Join-Path $Base 'Output\Knowledge Pack' }
if (-not $WorkRoot) { $WorkRoot = Join-Path $Base 'Output\_Refresh Work' }
$Expected = @('00_README.md', '01_Company_Background.md', '02_Tech_Stack.md', '03_Team_Roster.md',
              '04_Project_Index.md', '05_DMS_and_Integrations.md', '06_Lessons_and_Pitfalls.md', '07_Example_Briefs.md')
$CoAuthor = 'Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>'
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

function Invoke-Git {
    # Runs git through System.Diagnostics.Process: byte-exact stdout (git show), stderr kept apart, and no
    # PS 5.1 NativeCommandError noise. -ReadOnly sets GIT_OPTIONAL_LOCKS=0 so status never rewrites the index.
    param([Parameter(Mandatory)][string[]]$GitArgs, [switch]$ReadOnly, [switch]$AllowFail)
    $q = foreach ($a in $GitArgs) { if ($a -eq '' -or $a -match '[\s"]') { '"' + ($a -replace '"', '\"') + '"' } else { $a } }
    $psi = New-Object Diagnostics.ProcessStartInfo 'git'
    $psi.Arguments = ($q -join ' ')
    $psi.UseShellExecute = $false; $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true
    if ($ReadOnly) { $psi.EnvironmentVariables['GIT_OPTIONAL_LOCKS'] = '0' }
    try { $p = [Diagnostics.Process]::Start($psi) } catch { throw "git is not available on this PC: $($_.Exception.Message)" }
    $errTask = $p.StandardError.ReadToEndAsync()
    $ms = New-Object IO.MemoryStream
    $p.StandardOutput.BaseStream.CopyTo($ms)
    $p.WaitForExit()
    $r = [pscustomobject]@{ Exit = $p.ExitCode; Bytes = $ms.ToArray(); Out = $Utf8.GetString($ms.ToArray()).Trim(); Err = $errTask.Result.Trim() }
    if ($r.Exit -ne 0 -and -not $AllowFail) { throw ("git {0} failed (exit {1}): {2}" -f ($GitArgs -join ' '), $r.Exit, $r.Err) }
    return $r
}

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

function New-RefreshEmail {
    # Returns @{ Subject; Html; Markers }. Kept separate from the send so it can be tested offline.
    param($Diff, [bool]$Changed, [string]$SnapText, [string]$DateFolder, [int]$ScanFiles, [string]$BaseLabel, [string]$CommitInfo)
    $p    = 'style="margin:6px 0;"'
    $foot = "<p style=""color:#666;font-size:12px;margin-top:18px;"">Sent by Invoke-PackRefresh.ps1 (Brief Builder). Log and full summary: Projects\Brief Builder\Output\_Refresh Work\$DateFolder\</p>"
    if (-not $Changed) {
        $subject = "Brief Builder pack checked $SnapText $Dash no changes"
        $html = @"
<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#222;">
<p $p>Weekly Knowledge Pack refresh, snapshot <b>$SnapText</b>: <b>no changes</b> since the last commit ($(Enc $BaseLabel)).</p>
<p $p><b>Sanitisation:</b> PASSED $Dash $ScanFiles files scanned, 0 hits. Nothing was committed or pushed, so there is nothing to sync in claude.ai this week.</p>
$foot
</div>
"@
        return @{ Subject = $subject; Html = $html; Markers = @('no changes', $SnapText, 'Sanitisation:') }
    }
    $subject = "Brief Builder pack refreshed $SnapText $Dash sync in claude.ai"
    $html = @"
<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#222;">
<p $p>The shared Knowledge Pack has been regenerated (snapshot <b>$SnapText</b>), passed its checks and was pushed to GitHub$(if ($CommitInfo) { " ($(Enc $CommitInfo))" }).</p>
<p $p><b>Sanitisation:</b> PASSED $Dash $ScanFiles files scanned, 0 hits.</p>
<h2 style="font-size:17px;color:#1F3864;margin:18px 0 6px 0;">Sync checklist</h2>
<ol style="margin:4px 0 8px 18px;padding:0;">
<li>In claude.ai, open the <b>Brief Builder</b> Project &rarr; <b>Project knowledge</b> &rarr; <b>GitHub</b> &rarr; <b>Sync now</b>.</li>
<li>Do the same in the <b>Ideas Capture</b> Project.</li>
<li><b>Start a test brief</b> (Brief Builder) and a <b>test card</b> (Ideas Capture) and check they reflect the new snapshot.</li>
</ol>
<p style="color:#666;font-size:12px;">Fallback if GitHub sync is unavailable: delete the 8 old files in Project knowledge and upload the 8 from OneDrive\PKA_Claude\Projects\Brief Builder\Output\Knowledge Pack\ (00&ndash;07 only). The heartbeat proves the pack was generated and pushed, not synced.</p>
<h2 style="font-size:17px;color:#1F3864;margin:18px 0 6px 0;">What changed since $(Enc $BaseLabel)</h2>
$(ConvertTo-SummaryHtml $Diff)
$foot
</div>
"@
    return @{ Subject = $subject; Html = $html; Markers = @('Sync checklist', 'Sync now', 'Ideas Capture', $SnapText, 'Sanitisation:') }
}

$step = 'start'
$headTemp = $null
try {
    Log "=== $Process started (snapshot $snapText; push=$(-not $NoPush); send=$(-not $NoSend); heartbeat=$(-not $NoHeartbeat)) ==="
    Log "Pack: $PackPath"
    if ($SnapshotDate) {
        if ($snap -gt $today -or $snap -lt $today.AddDays(-7)) { throw "SnapshotDate $snapText must be between 7 days ago and today." }
        Log "SnapshotDate override in use: $snapText (run date $($today.ToString('dd/MM/yyyy')))" 'WARN'
    }

    # (a) repo + files present
    $step = '(a) repo and files'
    if (-not (Test-Path -LiteralPath $PackPath -PathType Container)) { throw "Pack folder not found: $PackPath" }
    if ((Invoke-Git -ReadOnly -AllowFail -GitArgs @('-C', $PackPath, 'rev-parse', '--is-inside-work-tree')).Out -ne 'true') { throw "$PackPath is not a git work tree" }
    $cur = (Invoke-Git -ReadOnly -GitArgs @('-C', $PackPath, 'rev-parse', '--abbrev-ref', 'HEAD')).Out
    if ($cur -ne $Branch) { throw "Pack repo is on branch '$cur', expected '$Branch'" }
    $headSha = (Invoke-Git -ReadOnly -GitArgs @('-C', $PackPath, 'rev-parse', '--short', 'HEAD')).Out
    $missing = @($Expected | Where-Object { -not (Test-Path -LiteralPath (Join-Path $PackPath $_) -PathType Leaf) })
    if ($missing.Count) { throw "Missing file(s): $($missing -join ', ')" }
    $extra = @(Get-ChildItem -LiteralPath $PackPath -Filter '0*.md' -File | Where-Object { $_.Name -notin $Expected } | ForEach-Object Name)
    if ($extra.Count) { Log "Extra 0*.md file(s) in the pack - git add 0*.md WILL include them: $($extra -join ', ')" 'WARN' }
    Log "(a) OK - git repo on $Branch at $headSha; all 8 files present"

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

    # (c) sanitisation = push gate
    $step = '(c) sanitisation'
    $scanReport = Join-Path $WorkDir 'sanitisation_report.txt'
    $scan = & (Join-Path $Scripts 'Test-PackSanitisation.ps1') -PackPath $PackPath -ReportPath $scanReport -PassThru
    $scanExit = $LASTEXITCODE
    if ($scanExit -ne 0 -or -not $scan -or -not $scan.Clean) {
        $n = if ($scan) { $scan.HitCount } else { '?' }
        throw "Sanitisation FAILED (exit $scanExit, $n hit(s)) - nothing committed or pushed. See $scanReport"
    }
    Log "(c) OK - sanitisation clean: $($scan.Files) files, 0 hits"

    # (d) diff vs HEAD
    $step = '(d) change summary'
    $headTemp = Join-Path ([IO.Path]::GetTempPath()) ("BriefBuilderPack_HEAD_{0}" -f (Get-Date -Format 'yyyyMMddHHmmss'))
    New-Item -ItemType Directory -Path $headTemp -Force | Out-Null
    # Normally the base is HEAD. If local main is AHEAD of origin/main (an earlier run committed but its push
    # failed), the base is origin/main instead, so this run summarises AND pushes the stranded commit(s).
    $aheadOut = (Invoke-Git -ReadOnly -AllowFail -GitArgs @('-C', $PackPath, 'rev-list', '--count', "origin/$Branch..HEAD")).Out
    $ahead = if ($aheadOut -match '^\d+$') { [int]$aheadOut } else { Log "No origin/$Branch tracking ref - diffing against HEAD" 'WARN'; 0 }
    $baseRef = if ($ahead -gt 0) { "origin/$Branch" } else { 'HEAD' }
    $baseSha = (Invoke-Git -ReadOnly -GitArgs @('-C', $PackPath, 'rev-parse', '--short', $baseRef)).Out
    if ($ahead -gt 0) { Log "Local $Branch is $ahead commit(s) ahead of origin/$Branch (earlier push failed?) - diffing against origin/$Branch and pushing them" 'WARN' }
    $inBase = @((Invoke-Git -ReadOnly -GitArgs @('-C', $PackPath, 'ls-tree', '--name-only', $baseRef)).Out -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -like '0*.md' })
    foreach ($f in $inBase) {
        $r = Invoke-Git -ReadOnly -GitArgs @('-C', $PackPath, 'show', "${baseRef}:$f")
        [IO.File]::WriteAllBytes((Join-Path $headTemp $f), $r.Bytes)
    }
    $porcelain = (Invoke-Git -ReadOnly -GitArgs @('-C', $PackPath, 'status', '--porcelain', '--', '0*.md')).Out
    $dirty   = [bool]$porcelain
    $changed = $dirty -or ($ahead -gt 0)
    $summaryPath = Join-Path $WorkDir 'change_summary.md'
    $baseLabel = "$baseRef $baseSha"
    $diff = & (Join-Path $Scripts 'Compare-PackVersions.ps1') -OldPath $headTemp -NewPath $PackPath -OutPath $summaryPath -OldLabel $baseLabel -NewLabel 'working tree'
    if (-not $diff) { throw 'Compare-PackVersions.ps1 returned nothing' }
    Log ("(d) OK - vs {0}: git changes={1}; 04 rows {2} -> {3}; +{4} / -{5} projects; {6} status change(s); roster +{7} / -{8}; lessons +{9} / -{10}" -f `
        $baseLabel, $changed, $diff.ProjectRowsOld, $diff.ProjectRowsNew, $diff.ProjectsAdded.Count, $diff.ProjectsRemoved.Count, $diff.StatusChanged.Count,
        $diff.RosterAdded.Count, $diff.RosterRemoved.Count, $diff.LessonsAdded.Count, $diff.LessonsRemoved.Count)
    if ($changed) { Log ("    git status: " + (($porcelain -split "`n" | ForEach-Object { $_.Trim() }) -join '; ')) }

    # (e) commit + push
    $step = '(e) commit and push'
    $commitInfo = ''
    if ($NoPush) { Log '(e) skipped (-NoPush)' }
    elseif (-not $changed) {
        Log "(e) no changes vs $baseLabel - nothing to commit or push"
    }
    else {
        if ($dirty) {
            Invoke-Git -GitArgs @('-C', $PackPath, 'add', '--', '0*.md') | Out-Null
            $msgFile = Join-Path $WorkDir 'commit_message.txt'
            [IO.File]::WriteAllText($msgFile, "Knowledge Pack refresh $snapText`n`n$CoAuthor`n", $Utf8)
            Invoke-Git -GitArgs @('-C', $PackPath, 'commit', '-F', $msgFile) | Out-Null
            Remove-Item -LiteralPath $msgFile -Force
            Log "    committed $((Invoke-Git -ReadOnly -GitArgs @('-C', $PackPath, 'rev-parse', '--short', 'HEAD')).Out): Knowledge Pack refresh $snapText"
        } else { Log "    working tree clean; pushing $ahead stranded commit(s)" }
        $newSha  = (Invoke-Git -ReadOnly -GitArgs @('-C', $PackPath, 'rev-parse', 'HEAD')).Out
        $push = Invoke-Git -AllowFail -GitArgs @('-C', $PackPath, 'push', 'origin', $Branch)
        if ($push.Exit -ne 0) { throw "git push failed (exit $($push.Exit)) - the commit is local only: $($push.Err)" }
        $remote = ((Invoke-Git -GitArgs @('-C', $PackPath, 'ls-remote', 'origin', "refs/heads/$Branch")).Out -split '\s+')[0]
        if ($remote -ne $newSha) { throw "Push reported success but origin/$Branch is '$remote', not $newSha" }
        $commitInfo = "commit $($newSha.Substring(0, 7)) on $Branch"
        Log "(e) OK - pushed; origin/$Branch = $($newSha.Substring(0, 7)) (verified with ls-remote)"
    }

    # (f) email Ant
    $step = '(f) email'
    if ($NoSend) { Log '(f) skipped (-NoSend)' }
    else {
        $mail = New-RefreshEmail -Diff $diff -Changed $changed -SnapText $snapText -DateFolder $snap.ToString('yyyy-MM-dd') `
                                 -ScanFiles $scan.Files -BaseLabel $baseLabel -CommitInfo $commitInfo
        $token = Get-GraphToken
        $msg = @{
            subject      = $mail.Subject
            body         = @{ contentType = 'HTML'; content = $mail.Html }
            from         = @{ emailAddress = @{ address = $FromUser } }
            toRecipients = @(@{ emailAddress = @{ address = $AntEmail } })
        }
        $draft = Invoke-Graph -Uri "$G/users/$SendVia/messages" -Token $token -Method Post -Body $msg
        if (-not $draft -or -not $draft.id) { throw 'Draft was not created (no id returned)' }
        Log 'Draft created; verifying before send'
        $back = Invoke-Graph -Uri "$G/users/$SendVia/messages/$($draft.id)?`$select=subject,body,toRecipients,from" -Token $token
        $to   = @($back.toRecipients | ForEach-Object { $_.emailAddress.address })
        $text = [string]$back.body.content
        $problems = @()
        if ([string]::IsNullOrWhiteSpace($back.subject) -or $back.subject -ne $mail.Subject) { $problems += "subject mismatch ('$($back.subject)')" }
        if ([string]::IsNullOrWhiteSpace($text)) { $problems += 'body empty' }
        foreach ($m in $mail.Markers) { if ($text.IndexOf($m) -lt 0) { $problems += "body lacks '$m'" } }
        if ($text -match ([string][char]0x00E2 + [string][char]0x20AC)) { $problems += 'body has mojibake' }
        if ($to -notcontains $AntEmail) { $problems += "recipient is '$($to -join ',')'" }
        if ($back.from -and $back.from.emailAddress.address -ne $FromUser) { $problems += "from is '$($back.from.emailAddress.address)'" }
        if ($problems.Count) { throw "Draft failed verification, NOT sent (draft left in $SendVia Drafts): $($problems -join '; ')" }
        Invoke-Graph -Uri "$G/users/$SendVia/messages/$($draft.id)/send" -Token $token -Method Post | Out-Null
        Log "(f) SENT to $AntEmail from $FromUser : $($mail.Subject)"
        # Best effort: confirm the Sent Items copy has a body (a successful /send is not proof of content).
        $seen = $false
        for ($t = 0; $t -lt 6 -and -not $seen; $t++) {
            Start-Sleep -Seconds 5
            try {
                $sent = Invoke-Graph -Uri "$G/users/$SendVia/mailFolders/sentitems/messages?`$top=10&`$orderby=sentDateTime desc&`$select=subject,body" -Token $token
                $hit = @($sent.value | Where-Object { $_.subject -eq $mail.Subject -and ([string]$_.body.content).IndexOf($mail.Markers[0]) -ge 0 })
                if ($hit.Count) { $seen = $true }
            } catch { Log "Sent Items check: $($_.Exception.Message)" 'WARN' }
        }
        if ($seen) { Log '(f) Sent Items copy verified (subject + body)' } else { Log '(f) Sent copy not seen within 30 s - check Sent Items by hand' 'WARN' }
    }

    # (g) heartbeat - written whether or not anything changed
    $step = '(g) heartbeat'
    if ($NoHeartbeat) { Log '(g) skipped (-NoHeartbeat)' }
    else {
        $hbMsg = ("Pack {0}: clean scan; {1}; email {2}" -f $snapText,
                  $(if (-not $changed) { "no changes vs $baseLabel" } elseif ($commitInfo) { "pushed $commitInfo, +$($diff.ProjectsAdded.Count)/-$($diff.ProjectsRemoved.Count) projects, $($diff.StatusChanged.Count) status changes" } else { 'changes not pushed (-NoPush)' }),
                  $(if ($NoSend) { 'skipped' } else { 'sent' }))
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
        Log "(g) OK - SUCCESS heartbeat written for $Process on $SqlHost"
    }

    Log "=== $Process completed OK ==="
    $ok = $true
}
catch {
    Log "FAILED at $step : $($_.Exception.Message)" 'ERROR'
    Log "=== $Process stopped ===" 'ERROR'
    $ok = $false
}
finally {
    if ($headTemp -and (Test-Path -LiteralPath $headTemp)) { Remove-Item -LiteralPath $headTemp -Recurse -Force -ErrorAction SilentlyContinue }
}
if ($ok) { Write-Host 'True'; exit 0 } else { Write-Host 'False'; exit 1 }
