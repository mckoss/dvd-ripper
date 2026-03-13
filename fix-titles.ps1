<#
.SYNOPSIS
    Renames MKV and MP4 files to "Movie Title (Year).ext" using TMDb lookup.

.DESCRIPTION
    Scans MKV and MP4 archive/processing directories, extracts raw titles
    from MKV metadata or filenames, searches TMDb for the official title
    and release year, and renames after confirmation. Supports single and
    batch modes. Requires a TMDb API key and ffprobe (from FFmpeg).

.PARAMETER MoviesDir
    Root movies directory (default: G:\Movies).

.PARAMETER TmdbApiKey
    TMDb API key (v3). If not provided, reads from .tmdb-api-key file
    in the script directory, or prompts interactively.

.EXAMPLE
    .\fix-titles.ps1
    .\fix-titles.ps1 -MoviesDir "H:\Media"
#>
param(
    [string]$MoviesDir,
    [string]$TmdbApiKey
)

# ── Configuration ─────────────────────────────────────────────────────────────

. "$PSScriptRoot\tmdb-helpers.ps1"

if ([string]::IsNullOrWhiteSpace($MoviesDir)) {
    $DefaultMoviesDir = 'G:\Movies'
    $MoviesDir = Read-Host "Enter movies directory (default: $DefaultMoviesDir)"
    if ([string]::IsNullOrWhiteSpace($MoviesDir)) { $MoviesDir = $DefaultMoviesDir }
}

# Directories where MKVs can live
$MkvArchiveDir = Join-Path $MoviesDir 'MKVs'
$RippedDir     = Join-Path $MoviesDir 'processing\ripped-for-encoding'

# Directories where MP4s can live
$Mp4ArchiveDir = Join-Path $MoviesDir 'MP4s'
$EncodedDir    = Join-Path $MoviesDir 'processing\encoded-for-upload'

# All MKV search dirs
$mkvDirs = @($MkvArchiveDir, $RippedDir) | Where-Object { Test-Path $_ }
# All MP4 search dirs
$mp4Dirs = @($Mp4ArchiveDir, $EncodedDir) | Where-Object { Test-Path $_ }

if ($mkvDirs.Count -eq 0 -and $mp4Dirs.Count -eq 0) {
    Write-Host "No movie directories found under $MoviesDir" -ForegroundColor Red
    exit 1
}

# TMDb API key (v3)
$TmdbApiKey = Initialize-TmdbApiKey -TmdbApiKey $TmdbApiKey

# ── Phase 1: Scan MKVs and MP4s, group by base name ──────────────────────────

Write-Host ''
Write-Host 'Scanning for movie files...' -ForegroundColor Cyan
$dirList = $mkvDirs -join ', '
Write-Host "  MKV locations: $dirList" -ForegroundColor Gray
$dirList = $mp4Dirs -join ', '
Write-Host "  MP4 locations: $dirList" -ForegroundColor Gray

$allMkvs = @()
foreach ($dir in $mkvDirs) {
    $allMkvs += Get-ChildItem -Path $dir -Filter *.mkv -File
}

$allMp4s = @()
foreach ($dir in $mp4Dirs) {
    $allMp4s += Get-ChildItem -Path $dir -Filter *.mp4 -File
}

if ($allMkvs.Count -eq 0 -and $allMp4s.Count -eq 0) {
    Write-Host 'No MKV or MP4 files found.' -ForegroundColor Yellow
    exit 0
}

$fileGroups = @{}
$mp4Warnings = [System.Collections.Generic.List[string]]::new()

# Add MKVs first, then find matching MP4s
foreach ($mkv in $allMkvs) {
    $base = $mkv.BaseName
    if ($fileGroups.ContainsKey($base)) { continue }

    $mp4Match = $null
    foreach ($dir in $mp4Dirs) {
        $candidate = Join-Path $dir ($base + '.mp4')
        if (Test-Path $candidate) {
            $mp4Match = Get-Item $candidate
            break
        }
    }

    $fileGroups[$base] = @{
        mkv = $mkv
        mp4 = $mp4Match
    }

    if (-not $mp4Match) {
        $mp4Warnings.Add($base)
    }
}

# Add orphan MP4s (no matching MKV)
$orphanMp4Count = 0
foreach ($mp4 in $allMp4s) {
    $base = $mp4.BaseName
    if ($fileGroups.ContainsKey($base)) { continue }

    $fileGroups[$base] = @{
        mkv = $null
        mp4 = $mp4
    }
    $orphanMp4Count++
}

$totalCount = $fileGroups.Count
$mkvCount = ($fileGroups.Values | Where-Object { $_.mkv }) | Measure-Object | Select-Object -ExpandProperty Count
Write-Host ''
Write-Host "Found $totalCount title(s): $mkvCount with MKV, $orphanMp4Count MP4-only."
if ($mp4Warnings.Count -gt 0) {
    $warnCount = $mp4Warnings.Count
    Write-Host ''
    Write-Host "WARNING: No matching MP4 found for $warnCount MKV(s):" -ForegroundColor Yellow
    $checkedDirs = $mp4Dirs -join ', '
    foreach ($w in $mp4Warnings) {
        Write-Host "  ${w}.mkv (checked: $checkedDirs)" -ForegroundColor DarkYellow
    }
}
Write-Host ''

# ── Phase 2: For each file group, look up TMDb and rename ─────────────────────

function Rename-MovieFile {
    param([System.IO.FileInfo]$File, [string]$NewBase)
    $newName = $NewBase + $File.Extension
    $newPath = Join-Path $File.DirectoryName $newName
    if ($File.Name -eq $newName) {
        Write-Host "  Already named: $newName" -ForegroundColor DarkGreen
        return 'skip'
    }
    if (Test-Path $newPath) {
        Write-Host "  SKIPPED (target exists): $newName" -ForegroundColor Yellow
        return 'skip'
    }
    try {
        Rename-Item -Path $File.FullName -NewName $newName -ErrorAction Stop
        Write-Host "  Renamed: $($File.Name) -> $newName" -ForegroundColor Green
        return 'ok'
    } catch {
        Write-Host "  FAILED: $($File.Name) - $_" -ForegroundColor Red
        return 'fail'
    }
}

function Rename-MovieGroup {
    param($Group, [string]$NewBase)
    $anyFail = $false
    if ($Group.mkv) {
        $res = Rename-MovieFile -File $Group.mkv -NewBase $NewBase
        if ($res -eq 'fail') { $anyFail = $true }
    }
    if ($Group.mp4) {
        $res = Rename-MovieFile -File $Group.mp4 -NewBase $NewBase
        if ($res -eq 'fail') { $anyFail = $true }
    }
    return $anyFail
}

function Get-MovieSearchResults {
    param($Group, [string]$BaseName, [string]$ApiKey)
    $rawTitle = $null
    if ($Group.mkv) {
        $rawTitle = Get-MkvTitle -FilePath $Group.mkv.FullName
        if ($rawTitle) {
            Write-Host "  Metadata title: $rawTitle" -ForegroundColor Gray
        }
    }

    $primaryFile = if ($Group.mkv) { $Group.mkv.FullName } else { $Group.mp4.FullName }
    $fileDuration = Get-FileDurationMinutes -FilePath $primaryFile
    if ($fileDuration) {
        $durStr = Format-Duration $fileDuration
        Write-Host "  File duration: $durStr" -ForegroundColor Gray
    }

    $searchQuery = if ($rawTitle) { Format-RawTitle $rawTitle } else { Format-RawTitle $BaseName }
    $filenameQuery = Format-RawTitle $BaseName
    Write-Host "  Search query: $searchQuery" -ForegroundColor Gray

    $results = @(Search-TMDb -Query $searchQuery -ApiKey $ApiKey)

    if ($results.Count -eq 0 -and $rawTitle -and $filenameQuery -ne $searchQuery) {
        Write-Host '  No results from metadata. Trying filename...' -ForegroundColor Yellow
        Write-Host "  Search query: $filenameQuery" -ForegroundColor Gray
        $results = @(Search-TMDb -Query $filenameQuery -ApiKey $ApiKey)
    }

    if ($results.Count -gt 0) {
        Add-RuntimeToResults -Results $results -ApiKey $ApiKey
    }

    return @{
        results = $results
        fileDuration = $fileDuration
    }
}

function Show-MovieHeader {
    param($Group, [string]$BaseName)
    Write-Host '-------------------------------------------------------------' -ForegroundColor DarkGray
    if ($Group.mkv) {
        $mkvName = $Group.mkv.Name
        $mkvDir = $Group.mkv.DirectoryName
        Write-Host "  MKV: $mkvName" -ForegroundColor White
        Write-Host "       $mkvDir" -ForegroundColor DarkGray
    } else {
        Write-Host '  MKV: NOT FOUND (MP4-only)' -ForegroundColor Yellow
    }
    if ($Group.mp4) {
        $mp4Name = $Group.mp4.Name
        $mp4Dir = $Group.mp4.DirectoryName
        Write-Host "  MP4: $mp4Name" -ForegroundColor White
        Write-Host "       $mp4Dir" -ForegroundColor DarkGray
    } else {
        Write-Host '  MP4: NOT FOUND (only MKV will be renamed)' -ForegroundColor Yellow
    }
}

function Get-DefaultNewBase {
    param($Results)
    if ($Results.Count -eq 0) { return $null }
    $selected = $Results[0]
    $year = Get-YearFromResult -Result $selected
    return Format-MovieFilename -Title $selected.title -Year $year
}

$renamed = 0
$failed = 0
$alreadyOk = [System.Collections.Generic.List[string]]::new()
$skipped = [System.Collections.Generic.List[string]]::new()
$batchMode = $false

$sortedKeys = @($fileGroups.Keys | Sort-Object)
$keyIndex = 0

while ($keyIndex -lt $sortedKeys.Count) {
    $baseName = $sortedKeys[$keyIndex]
    $group = $fileGroups[$baseName]
    $keyIndex++

    # Check if already in "Title (Year)" format
    if ($baseName -match '^.+\(\d{4}\)$') {
        $alreadyOk.Add($baseName)
        continue
    }

    if ($batchMode) {
        # ── Batch mode: auto-select #1 for up to 10 movies ──────────────
        $batchItems = [System.Collections.Generic.List[hashtable]]::new()
        # Process current movie + up to 9 more
        $batchKeys = @($baseName)
        while ($batchKeys.Count -lt 10 -and $keyIndex -lt $sortedKeys.Count) {
            $nextKey = $sortedKeys[$keyIndex]
            if ($nextKey -match '^.+\(\d{4}\)$') {
                $alreadyOk.Add($nextKey)
                $keyIndex++
                continue
            }
            $batchKeys += $nextKey
            $keyIndex++
        }

        foreach ($bk in $batchKeys) {
            $bg = $fileGroups[$bk]
            Show-MovieHeader -Group $bg -BaseName $bk
            $searchData = Get-MovieSearchResults -Group $bg -BaseName $bk -ApiKey $TmdbApiKey

            if ($searchData.results.Count -gt 0) {
                Show-TmdbResults -Results $searchData.results -FileDurationMin $searchData.fileDuration
                $newBase = Get-DefaultNewBase -Results $searchData.results
                $anyFail = Rename-MovieGroup -Group $bg -NewBase $newBase
                if ($anyFail) { $failed++ } else { $renamed++ }
                $batchItems.Add(@{ Num = $batchItems.Count + 1; BaseName = $bk; NewBase = $newBase; Group = $bg; Results = $searchData.results; FileDuration = $searchData.fileDuration })
            } else {
                Write-Host '  No TMDb results. Skipped in batch.' -ForegroundColor Yellow
                $skipped.Add($bk)
                $batchItems.Add(@{ Num = $batchItems.Count + 1; BaseName = $bk; NewBase = $null; Group = $bg; Results = @(); FileDuration = $searchData.fileDuration })
            }
            Write-Host ''
        }

        # ── Batch review ─────────────────────────────────────────────────
        if ($batchItems.Count -gt 0) {
            Write-Host '=============================================================' -ForegroundColor Cyan
            Write-Host '  BATCH REVIEW' -ForegroundColor Cyan
            Write-Host '=============================================================' -ForegroundColor Cyan
            foreach ($bi in $batchItems) {
                $num = $bi.Num
                $oldName = $bi.BaseName
                if ($bi.NewBase) {
                    Write-Host "    [$num] $oldName -> $($bi.NewBase)" -ForegroundColor White
                } else {
                    Write-Host "    [$num] $oldName -> (skipped - no results)" -ForegroundColor Yellow
                }
            }
            Write-Host ''
            Write-Host '  Enter a number to correct, [B] next batch, [S] single mode, [Q] quit' -ForegroundColor DarkGray

            while ($true) {
                $reviewChoice = Read-Host '  Review'
                if ([string]::IsNullOrWhiteSpace($reviewChoice) -or $reviewChoice -match '^[Bb]$') {
                    break
                }
                if ($reviewChoice -match '^[Ss]$') {
                    $batchMode = $false
                    break
                }
                if ($reviewChoice -match '^[Qq]$') {
                    Write-Host '  Quitting...' -ForegroundColor Yellow
                    $keyIndex = $sortedKeys.Count
                    break
                }
                if ($reviewChoice -match '^\d+$') {
                    $corrIdx = [int]$reviewChoice - 1
                    if ($corrIdx -ge 0 -and $corrIdx -lt $batchItems.Count) {
                        $ci = $batchItems[$corrIdx]
                        Write-Host ''
                        Write-Host "  Correcting: $($ci.BaseName)" -ForegroundColor Cyan

                        # Undo previous rename if it happened
                        if ($ci.NewBase) {
                            if ($ci.Group.mkv) {
                                $oldMkvName = $ci.BaseName + '.mkv'
                                $newMkvName = $ci.NewBase + '.mkv'
                                $newMkvPath = Join-Path $ci.Group.mkv.DirectoryName $newMkvName
                                if (Test-Path $newMkvPath) {
                                    Rename-Item -Path $newMkvPath -NewName $oldMkvName -ErrorAction SilentlyContinue
                                    Write-Host "  Reverted: $newMkvName -> $oldMkvName" -ForegroundColor Yellow
                                }
                            }
                            if ($ci.Group.mp4) {
                                $oldMp4Name = $ci.BaseName + '.mp4'
                                $newMp4Name = $ci.NewBase + '.mp4'
                                $newMp4Path = Join-Path $ci.Group.mp4.DirectoryName $newMp4Name
                                if (Test-Path $newMp4Path) {
                                    Rename-Item -Path $newMp4Path -NewName $oldMp4Name -ErrorAction SilentlyContinue
                                    Write-Host "  Reverted: $newMp4Name -> $oldMp4Name" -ForegroundColor Yellow
                                }
                            }
                            $renamed--
                        }

                        # Show results again and let user pick
                        if ($ci.Results.Count -gt 0) {
                            Show-TmdbResults -Results $ci.Results -FileDurationMin $ci.FileDuration
                        }
                        Write-Host '    [S] Skip  [C] Custom search  [M] Manual entry' -ForegroundColor DarkGray
                        $corrChoice = Read-Host '  Choose'

                        if ($corrChoice -match '^[Ss]$') {
                            $skipped.Add($ci.BaseName)
                            $ci.NewBase = $null
                        }
                        elseif ($corrChoice -match '^[Cc]$') {
                            $customSearch = Read-Host '  Enter search query'
                            $newResults = @(Search-TMDb -Query $customSearch -ApiKey $TmdbApiKey)
                            if ($newResults.Count -gt 0) {
                                Add-RuntimeToResults -Results $newResults -ApiKey $TmdbApiKey
                                Show-TmdbResults -Results $newResults -FileDurationMin $ci.FileDuration
                                $pick = Read-Host ('  Choose (1-' + $newResults.Count + ')')
                                if ($pick -match '^\d+$') {
                                    $pIdx = [int]$pick - 1
                                    if ($pIdx -ge 0 -and $pIdx -lt $newResults.Count) {
                                        $sel = $newResults[$pIdx]
                                        $yr = Get-YearFromResult -Result $sel
                                        $nb = Format-MovieFilename -Title $sel.title -Year $yr
                                        $anyFail = Rename-MovieGroup -Group $ci.Group -NewBase $nb
                                        if ($anyFail) { $failed++ } else { $renamed++ }
                                        $ci.NewBase = $nb
                                    }
                                }
                            } else {
                                Write-Host '  No results.' -ForegroundColor Yellow
                                $skipped.Add($ci.BaseName)
                            }
                        }
                        elseif ($corrChoice -match '^[Mm]$') {
                            $manualTitle = Read-Host '  Enter title'
                            $manualYear = Read-Host '  Enter year'
                            $nb = Format-MovieFilename -Title $manualTitle -Year $manualYear
                            $anyFail = Rename-MovieGroup -Group $ci.Group -NewBase $nb
                            if ($anyFail) { $failed++ } else { $renamed++ }
                            $ci.NewBase = $nb
                        }
                        elseif ($corrChoice -match '^\d+$' -and $ci.Results.Count -gt 0) {
                            $pIdx = [int]$corrChoice - 1
                            if ($pIdx -ge 0 -and $pIdx -lt $ci.Results.Count) {
                                $sel = $ci.Results[$pIdx]
                                $yr = Get-YearFromResult -Result $sel
                                $nb = Format-MovieFilename -Title $sel.title -Year $yr
                                $anyFail = Rename-MovieGroup -Group $ci.Group -NewBase $nb
                                if ($anyFail) { $failed++ } else { $renamed++ }
                                $ci.NewBase = $nb
                            }
                        }

                        # Refresh review list
                        Write-Host ''
                        foreach ($bi2 in $batchItems) {
                            $n2 = $bi2.Num
                            $o2 = $bi2.BaseName
                            if ($bi2.NewBase) {
                                Write-Host "    [$n2] $o2 -> $($bi2.NewBase)" -ForegroundColor White
                            } else {
                                Write-Host "    [$n2] $o2 -> (skipped)" -ForegroundColor Yellow
                            }
                        }
                        Write-Host ''
                        Write-Host '  Enter a number to correct, [B] next batch, [S] single mode, [Q] quit' -ForegroundColor DarkGray
                    } else {
                        Write-Host '  Invalid number.' -ForegroundColor Yellow
                    }
                }
            }
        }
        continue
    }

    # ── Single mode ──────────────────────────────────────────────────────────
    Show-MovieHeader -Group $group -BaseName $baseName
    $primaryFile = if ($group.mkv) { $group.mkv.FullName } else { $group.mp4.FullName }
    $result = Invoke-TmdbRenamePrompt -FilePath $primaryFile -RawBaseName $baseName -ApiKey $TmdbApiKey `
        -ExtraOptionsDisplay '[B] Batch 10  [Q] Quit' -ExtraOptionsKeys 'BQ'

    if ($result.Choice -eq 'Q') {
        Write-Host '  Quitting early...' -ForegroundColor Yellow
        break
    }
    elseif ($result.Choice -eq 'B') {
        # Apply default for current movie, then enter batch mode
        $newBase = Get-DefaultNewBase -Results $result.Results
        if ($newBase) {
            $anyFail = Rename-MovieGroup -Group $group -NewBase $newBase
            if ($anyFail) { $failed++ } else { $renamed++ }
        }
        $batchMode = $true
        Write-Host '  Entering batch mode...' -ForegroundColor Cyan
        Write-Host ''
        continue
    }
    elseif ($result.NewBase) {
        $anyFail = Rename-MovieGroup -Group $group -NewBase $result.NewBase
        if ($anyFail) { $failed++ } else { $renamed++ }
    } else {
        $skipped.Add($baseName)
    }
    Write-Host ''
}

# ── Summary ───────────────────────────────────────────────────────────────────

Write-Host ''
Write-Host '-------------------------------------------------------------' -ForegroundColor DarkGray
$summary = '  Done: ' + $renamed + ' movie(s) renamed | ' + $alreadyOk.Count + ' already OK | ' + $skipped.Count + ' skipped | ' + $failed + ' failed'
$resultColor = if ($failed -gt 0) { 'Yellow' } else { 'Green' }
Write-Host $summary -ForegroundColor $resultColor
Write-Host '-------------------------------------------------------------' -ForegroundColor DarkGray
Write-Host ''
