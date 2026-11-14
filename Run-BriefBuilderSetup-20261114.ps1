<#
    Run-BriefBuilderSetup-20261114.ps1 - the ONE script Ant runs on his laptop to switch on the
    WEEKLY Knowledge Pack refresh (Brief Builder + Ideas Capture; pack repo dwyers-automotive/ava-knowledge-pack).
    Stops at the first failure.

      0. Pre-flight : vault reachable over UNC; git present; the pack folder is a git repo on main with an
                      origin remote; all scripts + the exclusion file present.
      1. Dry check  : Invoke-PackRefresh.ps1 -NoSend -NoHeartbeat -NoPush. Proves the pack passes (files, date,
                      sanitisation, diff vs HEAD) BEFORE anything is registered, committed or pushed, so a half-done
                      setup never leaves an enabled heartbeat row that alerts NEVER.
      2. SQL        : Register-BriefBuilderHeartbeat.sql via the existing SQL admin runner
                      (Invoke-GrantFmtIntegration.ps1: runs on dmg-web as the vault login; no password typed).
                      Preview first, then apply.
      3. Reminder   : New-PackRefreshReminder.ps1 (recurring To Do, every Monday from 23/11/2026).
      4. Refresh    : Invoke-PackRefresh.ps1 with push + send + heartbeat (first real run).
    Ends by printing True/False.

    Run (one line, in PowerShell on the laptop):
      powershell -NoProfile -ExecutionPolicy Bypass -File "$env:USERPROFILE\OneDrive\PKA_Claude\Projects\Brief Builder\Scripts\Run-BriefBuilderSetup-20261114.ps1"
    If the pack was regenerated on an earlier day, add: -SnapshotDate dd/mm/yyyy (the date in 00/04).

    ASCII only. Windows PowerShell 5.1. Riley, 14/11/2026.
#>
[CmdletBinding()]
param(
    [string] $SnapshotDate
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$b      = Join-Path $env:USERPROFILE 'OneDrive\PKA_Claude\Projects\Brief Builder\Scripts'
$pack   = Join-Path $env:USERPROFILE 'OneDrive\PKA_Claude\Projects\Brief Builder\Output\Knowledge Pack'
$vault  = '\\dmg-web\c$\Secure\dwyers_apps.ini'
$runner = Join-Path $env:USERPROFILE 'OneDrive\PKA_Claude\Projects\FMT Fuel Management Integration\Scripts\Invoke-GrantFmtIntegration.ps1'
$sql    = Join-Path $b 'Register-BriefBuilderHeartbeat.sql'
$refreshArgs = @{}
if ($SnapshotDate) { $refreshArgs.SnapshotDate = $SnapshotDate }

function Step { param([string]$s) Write-Host ''; Write-Host "==== $s ====" -ForegroundColor Cyan }
function Stop-Setup {
    param([string]$why)
    Write-Host ''
    Write-Host "STOPPED: $why" -ForegroundColor Red
    Write-Host 'Nothing after this step was run.'
    Write-Host 'False'
    exit 1
}
function Git-Out {
    # Read-only git query; returns trimmed stdout or '' (no PS 5.1 NativeCommandError noise).
    param([string[]]$GitArgs)
    $psi = New-Object Diagnostics.ProcessStartInfo 'git'
    $psi.Arguments = (($GitArgs | ForEach-Object { if ($_ -match '\s') { '"' + $_ + '"' } else { $_ } }) -join ' ')
    $psi.UseShellExecute = $false; $psi.CreateNoWindow = $true; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true
    $psi.EnvironmentVariables['GIT_OPTIONAL_LOCKS'] = '0'
    $p = [Diagnostics.Process]::Start($psi); $o = $p.StandardOutput.ReadToEnd(); [void]$p.StandardError.ReadToEnd(); $p.WaitForExit()
    if ($p.ExitCode -ne 0) { return '' } else { return $o.Trim() }
}

try {
    Step '0. Pre-flight'
    if (-not (Test-Path -LiteralPath $vault)) {
        Stop-Setup "The vault ($vault) is not reachable. Connect to the internal network (office or VPN) and run this again."
    }
    Write-Host 'Vault reachable.'
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) { Stop-Setup 'git is not installed or not on PATH.' }
    if ((Git-Out @('-C', $pack, 'rev-parse', '--is-inside-work-tree')) -ne 'true') { Stop-Setup "The pack folder is not a git repo: $pack" }
    $br = Git-Out @('-C', $pack, 'rev-parse', '--abbrev-ref', 'HEAD')
    if ($br -ne 'main') { Stop-Setup "The pack repo is on branch '$br', expected 'main'." }
    $origin = Git-Out @('-C', $pack, 'remote', 'get-url', 'origin')
    if (-not $origin) { Stop-Setup 'The pack repo has no origin remote.' }
    Write-Host "Pack repo on main; origin = $origin"
    foreach ($f in 'Test-PackSanitisation.ps1', 'Compare-PackVersions.ps1', 'Invoke-PackRefresh.ps1', 'New-PackRefreshReminder.ps1',
                   'Register-BriefBuilderHeartbeat.sql', 'Sanitisation-Exclusions.txt', 'Sanitisation-Allow.txt') {
        if (-not (Test-Path -LiteralPath (Join-Path $b $f))) { Stop-Setup "Missing $b\$f (has OneDrive finished syncing?)" }
    }
    if (-not (Test-Path -LiteralPath $runner)) { Stop-Setup "SQL admin runner not found: $runner" }
    Write-Host 'All files present.'

    Step '1. Dry check of the pack (no push, no email, no heartbeat)'
    & (Join-Path $b 'Invoke-PackRefresh.ps1') -NoSend -NoHeartbeat -NoPush @refreshArgs
    if ($LASTEXITCODE -ne 0) { Stop-Setup 'The pack did not pass its checks (see the log lines above). Nothing was registered, committed or pushed. Ask AVA to fix the pack, then re-run.' }

    Step '2. Register the heartbeat row (SQL admin runner on dmg-web)'
    $pre = & $runner -SqlFile $sql -Preview 2>&1 | Out-String
    Write-Host $pre
    if ($pre -notmatch 'sysadmin=1' -or $pre -notmatch 'PREVIEW:') { Stop-Setup 'SQL preview did not confirm a sysadmin connection. Nothing was run.' }
    $run = & $runner -SqlFile $sql 2>&1 | Out-String
    Write-Host $run
    if ($run -notmatch 'DONE:' -or $run -notmatch 'BriefBuilder_PackRefresh') { Stop-Setup 'SQL registration did not complete (no DONE / no read-back row).' }
    Write-Host '(Ignore the runner''s "Now run Verify-Fmt..." line - that is for the FMT project.)'

    Step '3. Create the weekly To Do reminder'
    & (Join-Path $b 'New-PackRefreshReminder.ps1')
    if ($LASTEXITCODE -ne 0) { Stop-Setup 'The To Do reminder was not created.' }

    Step '4. Refresh: commit + push the pack, email Ant, write the heartbeat'
    & (Join-Path $b 'Invoke-PackRefresh.ps1') @refreshArgs
    if ($LASTEXITCODE -ne 0) { Stop-Setup 'The refresh run failed (see the log lines above).' }

    Write-Host ''
    Write-Host 'Setup complete. Check your inbox for the "Brief Builder pack" email and follow its sync checklist.' -ForegroundColor Green
    Write-Host 'True'
    exit 0
}
catch {
    Stop-Setup $_.Exception.Message
}
