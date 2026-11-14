# Knowledge Pack weekly refresh (Brief Builder + Ideas Capture)

Two claude.ai Projects share the pack: Brief Builder and Ideas Capture. The pack folder is `Projects\Brief Builder\Output\Knowledge Pack\`. It is a git repo pushed to the private `dwyers-automotive/ava-knowledge-pack` (branch `main`), and its allowlist `.gitignore` tracks only `0*.md`. Both Projects pick it up through GitHub "Sync now".

AVA/Riley regenerate the 8 files in a Claude session; that is the judgement step. These scripts then do everything deterministic. They are written for Windows PowerShell 5.1 and run on Ant's laptop, because the push, email and heartbeat steps need the internal network, the vault and GitHub. All paths are built from `$env:USERPROFILE`.

| File | What it does |
|---|---|
| `Invoke-PackRefresh.ps1` | The orchestrator. It stops at the first failure and prints True/False. Steps: (a) the pack is a git repo on `main` and all 8 files are present; (b) 00 and 04 carry today's date; (c) sanitisation passes, which is the **push gate**; (d) it diffs the working tree against `HEAD` (or against `origin/main` if an earlier push failed) and writes `Output\_Refresh Work\<date>\change_summary.md`; (e) it runs `git add 0*.md`, commits "Knowledge Pack refresh dd/mm/yyyy", pushes to `origin main` and verifies the push with `ls-remote`; (f) it emails Ant from ava@ (draft, then GET-verify, then send) with the change summary and the sync checklist; (g) it writes the `BriefBuilder_PackRefresh` heartbeat. If nothing changed, it skips the commit and push, sends a short "no changes" note and still writes the heartbeat. `-NoPush`, `-NoSend` and `-NoHeartbeat` are for testing. `-SnapshotDate dd/mm/yyyy` is for a run on a later day than the regeneration. |
| `Test-PackSanitisation.ps1` | The sanitisation gate. It scans every `.md` in a folder and reports `file:line:rule:snippet`, masking possible secrets. Exit 0 = clean, 1 = hits, 2 = could not run. |
| `Sanitisation-Exclusions.txt` | Names and phrases tied to HR, customer, personal and confidential matters. **Gitignored: it holds names.** It lives only in OneDrive. Add a line whenever a new matter of that kind is logged. PITWALL stays in it pending Ant's ruling. |
| `Sanitisation-Allow.txt` | Known-safe entries: the public website hosts (`host:`) and exact phrases. Keep it minimal. |
| `Compare-PackVersions.ps1` | The folder-to-folder diff. It reports added/removed/status-changed projects (04), added/removed roster names (03), added/removed lessons (06), and changed yes/no with line counts for each file. |
| `Register-BriefBuilderHeartbeat.sql` | Idempotent. It sets the IntegrationControl row to MonitorEnabled 1, MaxAge 21,600 min (15 days = two missed weekly runs + 1 day), All days, whole day. Run it through `Invoke-GrantFmtIntegration.ps1`. |
| `New-PackRefreshReminder.ps1` | Creates one recurring To Do task in Ant's default list for every Monday from 23/11/2026. It is idempotent. It reports the old monthly task if one exists but leaves it alone. |
| `Run-BriefBuilderSetup-20261114.ps1` | One-off setup: pre-flight checks (vault, git, repo), a dry check (`-NoSend -NoHeartbeat -NoPush`), the SQL registration, the reminder, then the first real refresh. |

## Weekly runbook
1. Every Monday the To Do reminder fires. Tell AVA: "AVA, refresh Brief Builder pack". AVA regenerates the 8 files in the pack folder, dated today, and does not commit them.
2. On the laptop, run: `powershell -NoProfile -ExecutionPolicy Bypass -File "$env:USERPROFILE\OneDrive\PKA_Claude\Projects\Brief Builder\Scripts\Invoke-PackRefresh.ps1"`
3. If it prints False, nothing unsafe was pushed. Hand `Output\_Refresh Work\<date>\refresh.log` (and `sanitisation_report.txt`) to AVA to fix, then re-run. A failed push is retried on the next run.
4. If it prints True and the email says "refreshed": in claude.ai, open Brief Builder and Ideas Capture in turn, go to Project knowledge, then GitHub, then Sync now. Then start a test brief or card in each.
5. If the email says "no changes", there is nothing to sync this week.
