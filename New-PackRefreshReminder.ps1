<#
    New-PackRefreshReminder.ps1 - ONE recurring To Do task in Ant's default list for the WEEKLY
    Knowledge Pack refresh (Brief Builder + Ideas Capture): every Monday, starting Monday 23/11/2026, 09:00.

    Idempotent: if any task in the default list already has this exact title, it does nothing.
    If the old MONTHLY reminder (first title below) exists, it is reported but left unchanged.
    App-only Graph (AVA app, vault [M365_GRAPH_AVA]); the same /users/{id}/todo/lists/{id}/tasks
    pattern Generate-MorningBrief.ps1 already uses to create tasks (Tasks.ReadWrite.All).
    The title deliberately does NOT start with "[d MMM]" - the Morning Brief skips those.

    -Preview: reads the lists and reports what it would do; creates nothing.
    Ends by printing True/False (exit 0/1). ASCII only. Windows PowerShell 5.1. Riley, 14/11/2026.
#>
[CmdletBinding()]
param(
    [switch] $Preview,
    [string] $IniPath   = '\\dmg-web\c$\Secure\dwyers_apps.ini',
    [string] $AntEmail  = 'adwyer@dwyers.com.au',
    [string] $StartDate = '2026-11-23'
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Title    = "Brief Builder + Ideas Capture: weekly pack refresh (say 'AVA, refresh Brief Builder pack')"
$OldTitle = "Brief Builder: monthly pack refresh (say 'AVA, refresh Brief Builder pack')"
$Body  = @"
Weekly Knowledge Pack refresh (shared by the Brief Builder and Ideas Capture Projects in claude.ai).
1. In a Claude session say: "AVA, refresh Brief Builder pack". AVA regenerates the 8 files, dated today.
2. On your laptop, in PowerShell, run:
   powershell -NoProfile -ExecutionPolicy Bypass -File "`$env:USERPROFILE\OneDrive\PKA_Claude\Projects\Brief Builder\Scripts\Invoke-PackRefresh.ps1"
   It checks the files, date and sanitisation, commits and pushes the pack to GitHub, emails you the change summary from ava@ and writes the heartbeat.
3. Follow the email: in claude.ai open Brief Builder and Ideas Capture > Project knowledge > GitHub > Sync now, then start a test brief/card.
Runbook: Projects\Brief Builder\Scripts\README.md
"@
$G    = 'https://graph.microsoft.com/v1.0'
$Utf8 = New-Object Text.UTF8Encoding($false)

function Get-ErrorBody { param($Err) try { $r = $Err.Exception.Response; if ($r) { return (New-Object IO.StreamReader($r.GetResponseStream())).ReadToEnd() } } catch {}; '' }
function Invoke-Graph {
    param([string]$Uri, [string]$Token, [string]$Method = 'GET', $Body)
    $p = @{ Uri = $Uri; Method = $Method; Headers = @{ Authorization = "Bearer $Token" }; UseBasicParsing = $true }
    if ($null -ne $Body) { $p.ContentType = 'application/json; charset=utf-8'; $p.Body = $Utf8.GetBytes(($Body | ConvertTo-Json -Depth 10)) }
    try { $wr = Invoke-WebRequest @p }
    catch { throw ("Graph {0} {1} failed: {2} {3}" -f $Method, ($Uri -replace '\?.*$', ''), $_.Exception.Message, (Get-ErrorBody $_)) }
    $txt = $Utf8.GetString($wr.RawContentStream.ToArray())
    if ($txt) { $txt | ConvertFrom-Json } else { $null }
}
function Get-GraphToken {
    if (-not (Test-Path -LiteralPath $IniPath)) { throw "Vault not reachable: $IniPath (are you on the internal network?)" }
    $c = @{}; $in = $false
    foreach ($l in [IO.File]::ReadAllLines($IniPath)) {
        $t = $l.Trim()
        if ($t -match '^\[(.+)\]$') { $in = ($Matches[1] -eq 'M365_GRAPH_AVA'); continue }
        if ($in -and $t -match '^([^;#=]+?)\s*=\s*(.*)$') { $c[$Matches[1].Trim()] = $Matches[2].Trim() }
    }
    foreach ($k in 'TenantId', 'ClientId', 'ClientSecret') { if (-not $c.ContainsKey($k) -or -not $c[$k]) { throw "Vault [M365_GRAPH_AVA] is missing $k" } }
    $b = @{ client_id = $c.ClientId; client_secret = $c.ClientSecret; scope = 'https://graph.microsoft.com/.default'; grant_type = 'client_credentials' }
    try { (Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$($c.TenantId)/oauth2/v2.0/token" -ContentType 'application/x-www-form-urlencoded' -Body $b).access_token }
    catch { throw "Graph token request failed: $($_.Exception.Message)" }
}

try {
    $start = [datetime]::ParseExact($StartDate, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
    if ($start.DayOfWeek -ne 'Monday') { throw "StartDate $StartDate is not a Monday" }

    $token = Get-GraphToken
    $lists = Invoke-Graph -Uri "$G/users/$AntEmail/todo/lists" -Token $token
    $list  = @($lists.value | Where-Object { $_.wellknownListName -eq 'defaultList' }) | Select-Object -First 1
    if (-not $list) { throw "No default To Do list found for $AntEmail" }
    Write-Host "Default list: '$($list.displayName)'"

    # Idempotency: page through every task (any status). No $select: Graph returns 400
    # (RequestBroker--ParseUri) for $select on To Do tasks (seen 14/11/2026).
    $u = "$G/users/$AntEmail/todo/lists/$($list.id)/tasks?`$top=100"
    $existing = $null; $old = $null; $n = 0
    while ($u) {
        $page = Invoke-Graph -Uri $u -Token $token
        $n += @($page.value).Count
        if (-not $existing) { $existing = @($page.value | Where-Object { $_.title -eq $Title }) | Select-Object -First 1 }
        if (-not $old)      { $old      = @($page.value | Where-Object { $_.title -eq $OldTitle }) | Select-Object -First 1 }
        $u = if ($page.PSObject.Properties['@odata.nextLink']) { $page.'@odata.nextLink' } else { $null }
    }
    Write-Host "Checked $n task(s)."
    if ($old) { Write-Host "NOTE - the old MONTHLY reminder exists (status $($old.status)). Left unchanged; delete it in To Do if no longer wanted." }
    if ($existing) {
        Write-Host "SKIP - the weekly reminder already exists (status $($existing.status)). Nothing created."
        Write-Host 'True'; exit 0
    }

    $due = $start.ToString('yyyy-MM-dd') + 'T09:00:00'
    $task = @{
        title            = $Title
        importance       = 'normal'
        body             = @{ contentType = 'text'; content = $Body }
        dueDateTime      = @{ dateTime = $due; timeZone = 'AUS Eastern Standard Time' }
        isReminderOn     = $true
        reminderDateTime = @{ dateTime = $due; timeZone = 'AUS Eastern Standard Time' }
        recurrence       = @{
            pattern = @{ type = 'weekly'; interval = 1; daysOfWeek = @('monday'); firstDayOfWeek = 'monday' }
            range   = @{ type = 'noEnd'; startDate = $start.ToString('yyyy-MM-dd') }
        }
    }
    if ($Preview) {
        Write-Host "PREVIEW - would create: '$Title', every Monday from $($start.ToString('dd/MM/yyyy')) 09:00. Nothing created."
        Write-Host 'True'; exit 0
    }
    $new = Invoke-Graph -Uri "$G/users/$AntEmail/todo/lists/$($list.id)/tasks" -Token $token -Method Post -Body $task
    # Read it back: prove title and recurrence actually landed.
    $back = Invoke-Graph -Uri "$G/users/$AntEmail/todo/lists/$($list.id)/tasks/$($new.id)" -Token $token
    $pat = if ($back.PSObject.Properties['recurrence'] -and $back.recurrence) { $back.recurrence.pattern } else { $null }
    if ($back.title -ne $Title) { throw "Task created but title reads back as '$($back.title)'" }
    if (-not $pat -or $pat.type -ne 'weekly' -or @($pat.daysOfWeek) -notcontains 'monday') {
        throw ("Task created but recurrence reads back as '{0}'. Check it in To Do (task kept)." -f $(if ($pat) { "$($pat.type)/$(@($pat.daysOfWeek) -join ',')" } else { 'none' }))
    }
    Write-Host ("CREATED - '{0}' in '{1}': every Monday from {2}, due {3}." -f $back.title, $list.displayName, $start.ToString('dd/MM/yyyy'), $back.dueDateTime.dateTime)
    Write-Host 'True'; exit 0
}
catch {
    Write-Host "FAILED: $($_.Exception.Message)"
    Write-Host 'False'; exit 1
}
