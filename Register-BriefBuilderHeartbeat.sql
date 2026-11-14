USE DMG;
/*  Register-BriefBuilderHeartbeat.sql - heartbeat row + monitor rules for BriefBuilder_PackRefresh.
    Idempotent - safe to re-run. Riley, 14/11/2026 (Ant approved option A).

    Run with the SQL admin runner (it connects to master, hence USE DMG above):
      Projects\FMT Fuel Management Integration\Scripts\Invoke-GrantFmtIntegration.ps1 -SqlFile <this file> [-Preview]
    Run-BriefBuilderSetup-20261114.ps1 does that for you (Preview first).

    Cadence: Invoke-PackRefresh.ps1 runs once a month, on the first Monday, from Ant's laptop after
    the pack is regenerated in a Claude session. Writer = UPDATE ... WHERE ProcessName (it never inserts).
    MaxAgeMinutes 50400 = 35 days, ActiveDays 'All', whole day: the longest gap between two first
    Mondays is 35 days, so this alerts once a refresh is missed.
    Columns are the real ones (01-Add-MonitorColumns.sql / 02-Seed-MonitorConfig.sql). MonitorNote is varchar(300).
*/
IF NOT EXISTS (SELECT 1 FROM dbo.IntegrationControl WHERE ProcessName = 'BriefBuilder_PackRefresh')
    INSERT INTO dbo.IntegrationControl (ProcessName, LastStatus, LastMessage, ModifiedAt, MonitorEnabled)
    VALUES ('BriefBuilder_PackRefresh', 'SEEDED', 'Seeded 14/11/2026 - awaiting first run', GETDATE(), 1);

UPDATE dbo.IntegrationControl
SET    MonitorEnabled = 1, MaxAgeMinutes = 50400, ActiveDays = 'All',
       ActiveFrom = NULL, ActiveTo = NULL,
       MonitorNote = 'Invoke-PackRefresh.ps1, monthly (first Monday) from Ant''s laptop after the pack is regenerated. 35 days = longest gap between first Mondays. Proves generation, not the claude.ai upload. Projects\Brief Builder\Scripts.'
WHERE  ProcessName = 'BriefBuilder_PackRefresh';

SELECT ProcessName, MonitorEnabled, MaxAgeMinutes, ActiveDays, ActiveFrom, ActiveTo,
       LastStatus, LastSuccessfulRun, LastAttemptedRun, MonitorNote
FROM   dbo.IntegrationControl
WHERE  ProcessName = 'BriefBuilder_PackRefresh';
