<#
    Compare-PackVersions.ps1 - "what changed" between two Brief Builder Knowledge Pack versions.

      04_Project_Index : table rows keyed by project name -> added / removed / status changed (old -> new)
      03_Team_Roster   : roster names -> added / removed
      06_Lessons...    : bullet lines -> added / removed
      every other file : changed yes/no + old/new line counts (files missing on either side are listed)

    Writes a Markdown summary to -OutPath and returns the result object.
    ASCII only. Windows PowerShell 5.1. Riley, 14/11/2026.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $OldPath,
    [Parameter(Mandatory)][string] $NewPath,
    [Parameter(Mandatory)][string] $OutPath
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Utf8 = New-Object Text.UTF8Encoding($false)
$Star = [string][char]0x2B50   # the "reusable platform" marker in 04

function Read-Lines { param([string]$Path) if (Test-Path -LiteralPath $Path) { ,[IO.File]::ReadAllLines($Path, [Text.Encoding]::UTF8) } else { $null } }
function Find-File  { param([string]$Dir, [string]$Prefix) Get-ChildItem -LiteralPath $Dir -Filter "$Prefix*.md" -File | Select-Object -First 1 }

function Clean-Cell {
    param([string]$s)
    $s = $s -replace [regex]::Escape($Star), '' -replace '\*\*', '' -replace '\s+', ' '
    return $s.Trim()
}

function Get-TableRows {
    # Returns ordered dictionary: key -> @{ Name; Status; Section }. Header rows (followed by |---) are skipped.
    param([string[]]$Lines)
    $rows = [ordered]@{}
    if (-not $Lines) { return $rows }
    $section = ''
    for ($i = 0; $i -lt $Lines.Length; $i++) {
        $l = $Lines[$i].Trim()
        if ($l -match '^#{1,6}\s+(.+)$') { $section = Clean-Cell $Matches[1]; continue }
        if (-not $l.StartsWith('|')) { continue }
        if ($l -match '^\|\s*:?-{2,}') { continue }
        if ($i + 1 -lt $Lines.Length -and $Lines[$i + 1].Trim() -match '^\|\s*:?-{2,}') { continue }   # header row
        $cells = @($l.Trim('|').Split('|') | ForEach-Object { Clean-Cell $_ })
        if (-not $cells[0]) { continue }
        $key = $cells[0]
        if ($rows.Contains($key)) { $key = "$key [$section]" }
        $rows[$key] = @{ Name = $cells[0]; Status = $(if ($cells.Count -gt 1) { $cells[1] } else { '' }); Section = $section }
    }
    return $rows
}

function Get-Bullets {
    param([string[]]$Lines)
    $set = New-Object System.Collections.Generic.List[string]
    if (-not $Lines) { return ,$set }
    foreach ($l in $Lines) {
        if ($l -match '^\s*[-*]\s+(.+)$') { $t = ($Matches[1] -replace '\s+', ' ').Trim(); if (-not $set.Contains($t)) { $set.Add($t) } }
    }
    return ,$set
}

$oldFiles = @(Get-ChildItem -LiteralPath $OldPath -Filter '*.md' -File | ForEach-Object Name)
$newFiles = @(Get-ChildItem -LiteralPath $NewPath -Filter '*.md' -File | ForEach-Object Name)

$result = [ordered]@{
    OldPath = $OldPath; NewPath = $NewPath; Generated = (Get-Date).ToString('dd/MM/yyyy HH:mm')
    ProjectsAdded = @(); ProjectsRemoved = @(); StatusChanged = @(); ProjectRowsOld = 0; ProjectRowsNew = 0
    RosterAdded = @(); RosterRemoved = @()
    LessonsAdded = @(); LessonsRemoved = @()
    Files = @(); FilesAdded = @($newFiles | Where-Object { $_ -notin $oldFiles }); FilesRemoved = @($oldFiles | Where-Object { $_ -notin $newFiles })
    HasChanges = $false
}

# ---- 04 Project Index ----
$o04 = Find-File $OldPath '04_'; $n04 = Find-File $NewPath '04_'
$oRows = Get-TableRows $(if ($o04) { Read-Lines $o04.FullName } else { $null })
$nRows = Get-TableRows $(if ($n04) { Read-Lines $n04.FullName } else { $null })
$result.ProjectRowsOld = $oRows.Count; $result.ProjectRowsNew = $nRows.Count
$result.ProjectsAdded   = @(foreach ($k in $nRows.Keys) { if (-not $oRows.Contains($k)) { [pscustomobject]@{ Project = $nRows[$k].Name; Status = $nRows[$k].Status; Section = $nRows[$k].Section } } })
$result.ProjectsRemoved = @(foreach ($k in $oRows.Keys) { if (-not $nRows.Contains($k)) { [pscustomobject]@{ Project = $oRows[$k].Name; Status = $oRows[$k].Status; Section = $oRows[$k].Section } } })
$result.StatusChanged   = @(foreach ($k in $nRows.Keys) { if ($oRows.Contains($k) -and $oRows[$k].Status -ne $nRows[$k].Status) { [pscustomobject]@{ Project = $nRows[$k].Name; Old = $oRows[$k].Status; New = $nRows[$k].Status } } })

# ---- 03 Team Roster (first table column = name) ----
$o03 = Find-File $OldPath '03_'; $n03 = Find-File $NewPath '03_'
$oNames = @((Get-TableRows $(if ($o03) { Read-Lines $o03.FullName } else { $null })).Values | ForEach-Object { $_.Name })
$nNames = @((Get-TableRows $(if ($n03) { Read-Lines $n03.FullName } else { $null })).Values | ForEach-Object { $_.Name })
$result.RosterAdded   = @($nNames | Where-Object { $_ -notin $oNames })
$result.RosterRemoved = @($oNames | Where-Object { $_ -notin $nNames })

# ---- 06 Lessons (bullets) ----
$o06 = Find-File $OldPath '06_'; $n06 = Find-File $NewPath '06_'
$oB = Get-Bullets $(if ($o06) { Read-Lines $o06.FullName } else { $null })
$nB = Get-Bullets $(if ($n06) { Read-Lines $n06.FullName } else { $null })
$result.LessonsAdded   = @($nB | Where-Object { -not $oB.Contains($_) })
$result.LessonsRemoved = @($oB | Where-Object { -not $nB.Contains($_) })

# ---- every file: changed yes/no + line counts ----
$result.Files = @(foreach ($name in @($oldFiles + $newFiles | Sort-Object -Unique)) {
    $ol = Read-Lines (Join-Path $OldPath $name); $nl = Read-Lines (Join-Path $NewPath $name)
    $changed = if ($null -eq $ol -or $null -eq $nl) { $true } else { (($ol -join "`n") -ne ($nl -join "`n")) }
    [pscustomobject]@{ File = $name; Changed = $changed
        OldLines = $(if ($null -eq $ol) { 0 } else { $ol.Length }); NewLines = $(if ($null -eq $nl) { 0 } else { $nl.Length }) }
})

$result.HasChanges = [bool](@($result.Files | Where-Object { $_.Changed }).Count)

# ---- Markdown summary ----
$sb = New-Object System.Text.StringBuilder
function Out-Line { param([string]$s = '') [void]$sb.AppendLine($s) }
Out-Line '# Brief Builder Knowledge Pack - change summary'
Out-Line ''
Out-Line "Generated $($result.Generated). Previous: ``$(Split-Path -Leaf $OldPath)`` -> new: ``$(Split-Path -Leaf $NewPath)``."
Out-Line ''
if (-not $result.HasChanges) { Out-Line '**No changes** - the two versions are identical.'; Out-Line '' }
Out-Line "## 04 Project Index ($($result.ProjectRowsOld) -> $($result.ProjectRowsNew) rows)"
Out-Line ''
Out-Line "**Added ($($result.ProjectsAdded.Count))**"
if ($result.ProjectsAdded.Count) { foreach ($p in $result.ProjectsAdded) { Out-Line "- $($p.Project) - $($p.Status) ($($p.Section))" } } else { Out-Line '- none' }
Out-Line ''
Out-Line "**Removed ($($result.ProjectsRemoved.Count))**"
if ($result.ProjectsRemoved.Count) { foreach ($p in $result.ProjectsRemoved) { Out-Line "- $($p.Project) - was $($p.Status)" } } else { Out-Line '- none' }
Out-Line ''
Out-Line "**Status changed ($($result.StatusChanged.Count))**"
if ($result.StatusChanged.Count) { foreach ($p in $result.StatusChanged) { Out-Line "- $($p.Project): $($p.Old) -> $($p.New)" } } else { Out-Line '- none' }
Out-Line ''
Out-Line '## 03 Team Roster'
Out-Line ''
Out-Line "- Added: $(if ($result.RosterAdded.Count) { $result.RosterAdded -join ', ' } else { 'none' })"
Out-Line "- Removed: $(if ($result.RosterRemoved.Count) { $result.RosterRemoved -join ', ' } else { 'none' })"
Out-Line ''
Out-Line '## 06 Lessons & Pitfalls'
Out-Line ''
Out-Line "**Added ($($result.LessonsAdded.Count))**"
if ($result.LessonsAdded.Count) { foreach ($b in $result.LessonsAdded) { Out-Line "- $b" } } else { Out-Line '- none' }
Out-Line ''
Out-Line "**Removed ($($result.LessonsRemoved.Count))**"
if ($result.LessonsRemoved.Count) { foreach ($b in $result.LessonsRemoved) { Out-Line "- $b" } } else { Out-Line '- none' }
Out-Line ''
Out-Line '## All files'
Out-Line ''
Out-Line '| File | Changed | Lines (old -> new) |'
Out-Line '|---|---|---|'
foreach ($f in $result.Files) { Out-Line "| $($f.File) | $(if ($f.Changed) { 'yes' } else { 'no' }) | $($f.OldLines) -> $($f.NewLines) |" }
if ($result.FilesAdded.Count)   { Out-Line ''; Out-Line "New files: $($result.FilesAdded -join ', ')" }
if ($result.FilesRemoved.Count) { Out-Line ''; Out-Line "Files no longer present: $($result.FilesRemoved -join ', ')" }

$dir = Split-Path -Parent $OutPath
if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
[IO.File]::WriteAllText($OutPath, $sb.ToString(), $Utf8)
$result.SummaryPath = $OutPath
[pscustomobject]$result
