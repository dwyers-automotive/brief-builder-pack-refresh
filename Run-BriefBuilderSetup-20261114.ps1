<#
    Run-BriefBuilderSetup-20261114.ps1 - the ONE script Ant runs on his laptop to switch on the
    Brief Builder monthly pack refresh (option A, approved 14/11/2026). Stops at the first failure.

      0. Pre-flight : vault reachable over UNC, all scripts + the exclusion file present.
      1. Dry check  : Invoke-PackRefresh.ps1 -NoSend -NoHeartbeat. Proves the regenerated pack passes
                      (files, date, sanitisation, diff) BEFORE anything is registered, so a half-done
                      setup never leaves an enabled heartbeat row that alerts NEVER.
      2. SQL        : Register-BriefBuilderHeartbeat.sql via the existing SQL admin runner
                      (Invoke-GrantFmtIntegration.ps1: runs on dmg-web as the vault login; no password typed).
                      Preview first, then apply.
      3. Reminder   : New-PackRefreshReminder.ps1 (recurring To Do, first Monday, from 07/12/2026).
      4. Refresh    : Invoke-PackRefresh.ps1 with send + heartbeat (first real run = v2 baseline).
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

try {
    Step '0. Pre-flight'
    if (-not (Test-Path -LiteralPath $vault)) {
        Stop-Setup "The vault ($vault) is not reachable. Connect to the internal network (office or VPN) and run this again."
    }
    Write-Host 'Vault reachable.'
    foreach ($f in 'Test-PackSanitisation.ps1', 'Compare-PackVersions.ps1', 'Invoke-PackRefresh.ps1', 'New-PackRefreshReminder.ps1',
                   'Register-BriefBuilderHeartbeat.sql', 'Sanitisation-Exclusions.txt', 'Sanitisation-Allow.txt') {
        if (-not (Test-Path -LiteralPath (Join-Path $b $f))) { Stop-Setup "Missing $b\$f (has OneDrive finished syncing?)" }
    }
    if (-not (Test-Path -LiteralPath $runner)) { Stop-Setup "SQL admin runner not found: $runner" }
    Write-Host 'All files present.'

    Step '1. Dry check of the regenerated pack (no email, no heartbeat)'
    & (Join-Path $b 'Invoke-PackRefresh.ps1') -NoSend -NoHeartbeat @refreshArgs
    if ($LASTEXITCODE -ne 0) { Stop-Setup 'The pack did not pass its checks (see the log lines above). Nothing was registered. Ask AVA to fix the pack, then re-run.' }

    Step '2. Register the heartbeat row (SQL admin runner on dmg-web)'
    $pre = & $runner -SqlFile $sql -Preview 2>&1 | Out-String
    Write-Host $pre
    if ($pre -notmatch 'sysadmin=1' -or $pre -notmatch 'PREVIEW:') { Stop-Setup 'SQL preview did not confirm a sysadmin connection. Nothing was run.' }
    $run = & $runner -SqlFile $sql 2>&1 | Out-String
    Write-Host $run
    if ($run -notmatch 'DONE:' -or $run -notmatch 'BriefBuilder_PackRefresh') { Stop-Setup 'SQL registration did not complete (no DONE / no read-back row).' }
    Write-Host '(Ignore the runner''s "Now run Verify-Fmt..." line - that is for the FMT project.)'

    Step '3. Create the monthly To Do reminder'
    & (Join-Path $b 'New-PackRefreshReminder.ps1')
    if ($LASTEXITCODE -ne 0) { Stop-Setup 'The To Do reminder was not created.' }

    Step '4. Refresh: email Ant + write the heartbeat'
    & (Join-Path $b 'Invoke-PackRefresh.ps1') @refreshArgs
    if ($LASTEXITCODE -ne 0) { Stop-Setup 'The refresh run failed (see the log lines above).' }

    Write-Host ''
    Write-Host 'Setup complete. Check your inbox for the "Brief Builder pack refreshed" email and follow its upload checklist.' -ForegroundColor Green
    Write-Host 'True'
    exit 0
}
catch {
    Stop-Setup $_.Exception.Message
}
