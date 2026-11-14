USE DMG;
/*  Register-BriefBuilderHeartbeat.sql - heartbeat row + monitor rules for BriefBuilder_PackRefresh.
    Idempotent - safe to re-run (it also moves an existing monthly row to the weekly rules).
    Riley, 14/11/2026 (Ant: option A, then weekly + git the same day).

    Run with the SQL admin runner (it connects to master, hence USE DMG above):
      Projects\FMT Fuel Management Integration\Scripts\Invoke-GrantFmtIntegration.ps1 -SqlFile <this file> [-Preview]
    Run-BriefBuilderSetup-20261114.ps1 does that for you (Preview first).

    Cadence: Invoke-PackRefresh.ps1 runs WEEKLY (Monday) from Ant's laptop after the pack is regenerated,
    and writes SUCCESS even when nothing changed. Writer = UPDATE ... WHERE ProcessName (it never inserts).
    MaxAgeMinutes 21600 = 15 days = two missed weekly runs + 1 day (the monitor's two-missed-runs rule).
    That also covers the Monday 28/12/2026 public-holiday closure. ActiveDays 'All', whole day.
    Columns are the real ones (01-Add-MonitorColumns.sql / 02-Seed-MonitorConfig.sql). MonitorNote is varchar(300).
*/
IF NOT EXISTS (SELECT 1 FROM dbo.IntegrationControl WHERE ProcessName = 'BriefBuilder_PackRefresh')
    INSERT INTO dbo.IntegrationControl (ProcessName, LastStatus, LastMessage, ModifiedAt, MonitorEnabled)
    VALUES ('BriefBuilder_PackRefresh', 'SEEDED', 'Seeded 14/11/2026 - awaiting first run', GETDATE(), 1);

UPDATE dbo.IntegrationControl
SET    MonitorEnabled = 1, MaxAgeMinutes = 21600, ActiveDays = 'All',
       ActiveFrom = NULL, ActiveTo = NULL,
       MonitorNote = 'Invoke-PackRefresh.ps1, weekly (Mon) from Ant''s laptop: scan, git push of the shared pack, email. SUCCESS even with no changes. 15 days = 2 missed runs + 1 day (covers 28/12 closure). Proves push, not the claude.ai sync. Projects\Brief Builder\Scripts.'
WHERE  ProcessName = 'BriefBuilder_PackRefresh';

SELECT ProcessName, MonitorEnabled, MaxAgeMinutes, ActiveDays, ActiveFrom, ActiveTo,
       LastStatus, LastSuccessfulRun, LastAttemptedRun, MonitorNote
FROM   dbo.IntegrationControl
WHERE  ProcessName = 'BriefBuilder_PackRefresh';
