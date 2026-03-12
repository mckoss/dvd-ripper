param (
    [string]$MoviesDir
)

. "$PSScriptRoot\tmdb-helpers.ps1"

if ([string]::IsNullOrWhiteSpace($MoviesDir)) {
    $DefaultMoviesDir = "G:\Movies"
    $MoviesDir = Read-Host "Enter movies directory (default: $DefaultMoviesDir)"
    if ([string]::IsNullOrWhiteSpace($MoviesDir)) {
        $MoviesDir = $DefaultMoviesDir
    }
}

$RippedDir = Join-Path $MoviesDir "processing\ripped-for-encoding"
$EncodedDir = Join-Path $MoviesDir "processing\encoded-for-upload"
$MkvArchiveDir = Join-Path $MoviesDir "MKVs"
$Mp4ArchiveDir = Join-Path $MoviesDir "MP4s"

if (-not (Test-Path $RippedDir)) {
    New-Item -ItemType Directory -Path $RippedDir -Force | Out-Null
}

# Ensure output directories exist
foreach ($dir in @($EncodedDir, $MkvArchiveDir, $Mp4ArchiveDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

# Helper: get relative path from a base directory (without extension)
function Get-RelativeBase {
    param([string]$FilePath, [string]$BaseDir)
    $rel = $FilePath.Substring($BaseDir.Length).TrimStart('\', '/')
    return [System.IO.Path]::ChangeExtension($rel, $null).TrimEnd('.')
}

# Helper: find files under a directory (always recurses into subdirectories)
function Find-Files {
    param([string]$Path, [string]$Filter)
    Get-ChildItem -Path $Path -Filter $Filter -File -Recurse
}

# Helper: remove empty subdirectories (deepest first)
function Remove-EmptyDirs {
    param([string]$Path)
    Get-ChildItem -Path $Path -Directory -Recurse | Sort-Object FullName -Descending | Where-Object {
        (Get-ChildItem -Path $_.FullName -Force).Count -eq 0
    } | ForEach-Object { Remove-Item $_.FullName -Force }
}

# Helper: prompt to rename MKVs needing title fixes in a given directory
function Invoke-TitleRenames {
    param([string]$Dir)

    # Find -check-title files
    $checkTitle = @(Get-ChildItem -Path $Dir -Filter '*-check-title.mkv' -File)
    # Find files missing a (Year)
    $missingYear = @(Get-ChildItem -Path $Dir -Filter '*.mkv' -File | Where-Object {
        $_.BaseName -notmatch '\(\d{4}\)' -and $_.BaseName -notmatch '-check-title$'
    })

    $toRename = @($checkTitle) + @($missingYear)
    if ($toRename.Count -eq 0) { return }

    if (-not $script:TmdbApiKey) { $script:TmdbApiKey = Initialize-TmdbApiKey }

    Write-Host "`n  $($toRename.Count) file(s) need title confirmation:" -ForegroundColor Cyan

    foreach ($mkv in $toRename) {
        $rawBase = if ($mkv.BaseName -match '-check-title$') {
            $mkv.BaseName -replace '-[A-Z]-check-title$', ''
        } else {
            $mkv.BaseName
        }

        Write-Host '  ---' -ForegroundColor DarkGray
        Write-Host "    File: $($mkv.Name)" -ForegroundColor White

        $result = Invoke-TmdbRenamePrompt -FilePath $mkv.FullName -RawBaseName $rawBase -ApiKey $script:TmdbApiKey

        if ($result.NewBase) {
            $newName = $result.NewBase + '.mkv'
            $newPath = Join-Path $mkv.DirectoryName $newName
            if (Test-Path $newPath) {
                Write-Host "    SKIPPED (target exists): $newName" -ForegroundColor Yellow
            } else {
                Rename-Item -Path $mkv.FullName -NewName $newName -ErrorAction Stop
                Write-Host "    Renamed: $($mkv.Name) -> $newName" -ForegroundColor Green
            }
        } else {
            Write-Host "    Kept as-is." -ForegroundColor Gray
        }
    }
}

# Build list of already-encoded MP4 relative paths (for archive and skip logic)
$existingMp4RelPaths = @()
if (Test-Path $EncodedDir) {
    $existingMp4RelPaths += Find-Files $EncodedDir '*.mp4' | ForEach-Object {
        Get-RelativeBase -FilePath $_.FullName -BaseDir $EncodedDir
    }
}
if (Test-Path $Mp4ArchiveDir) {
    $existingMp4RelPaths += Find-Files $Mp4ArchiveDir '*.mp4' | ForEach-Object {
        Get-RelativeBase -FilePath $_.FullName -BaseDir $Mp4ArchiveDir
    }
}

# Move already-encoded MKVs to the archive (preserving subfolder structure)
$alreadyEncoded = @(Find-Files $RippedDir '*.mkv' | Where-Object {
    $relBase = Get-RelativeBase -FilePath $_.FullName -BaseDir $RippedDir
    $relBase -in $existingMp4RelPaths
})

if ($alreadyEncoded.Count -gt 0) {
    Write-Host "Moving already-encoded MKVs to archive ($($alreadyEncoded.Count)):"
    $alreadyEncoded | ForEach-Object {
        $relPath = $_.FullName.Substring($RippedDir.Length).TrimStart('\', '/')
        $destPath = Join-Path $MkvArchiveDir $relPath
        $destDir = Split-Path $destPath -Parent
        if (-not (Test-Path $destDir)) {
            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        }
        try {
            Move-Item -Path $_.FullName -Destination $destPath -Force -ErrorAction Stop
            Write-Host "  Archived: $relPath"
        } catch {
            Write-Host "  Skipped (file in use): $relPath" -ForegroundColor Yellow
        }
    }
    Remove-EmptyDirs $RippedDir
    Write-Host ""
}

# Find all MKVs that still need encoding, grouped by directory
$needsEncoding = @(Find-Files $RippedDir '*.mkv' | Where-Object {
    $relBase = Get-RelativeBase -FilePath $_.FullName -BaseDir $RippedDir
    $relBase -notin $existingMp4RelPaths
})

if ($needsEncoding.Count -eq 0) {
    Write-Host "No MKV files need encoding."
} else {
    $dirGroups = @($needsEncoding | Group-Object DirectoryName)
    Write-Host "MKV files ready for encoding: $($needsEncoding.Count) file(s) in $($dirGroups.Count) folder(s)`n"

    for ($i = 0; $i -lt $dirGroups.Count; $i++) {
        $group = $dirGroups[$i]
        $sourceDir = $group.Name
        $relSubDir = $sourceDir.Substring($RippedDir.Length).TrimStart('\', '/')
        $outputDir = if ($relSubDir) { Join-Path $EncodedDir $relSubDir } else { $EncodedDir }

        Write-Host "-------------------------------------------------------------" -ForegroundColor DarkGray
        Write-Host "Folder $($i + 1) of $($dirGroups.Count): $($group.Count) file(s)" -ForegroundColor Cyan
        Write-Host "  Source:      $sourceDir" -ForegroundColor Gray
        Write-Host "  Destination: $outputDir" -ForegroundColor Gray
        $group.Group | ForEach-Object { Write-Host "    $($_.Name)" }

        $response = Read-Host "`nProcess this folder? (y/N)"
        if ($response -ne 'y') {
            Write-Host "  Skipped.`n"
            continue
        }

        # Rename titles needing confirmation in this directory
        Invoke-TitleRenames -Dir $sourceDir | Out-Null

        if (-not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        Write-Host ""
        Write-Host "  ** Before adding to the queue, set the default output folder **" -ForegroundColor Yellow
        Write-Host "  In HandBrake: Tools > Preferences > Output Files > Default Path" -ForegroundColor Yellow
        Write-Host "  Set to: $outputDir" -ForegroundColor Yellow
        Write-Host "  (Changing the preference updates the path for files not yet queued)" -ForegroundColor Yellow
        Start-Process "C:\Program Files\HandBrake\HandBrake.exe" -ArgumentList "`"$sourceDir`""

        if ($i -lt $dirGroups.Count - 1) {
            Read-Host "`nPress Enter when ready for the next folder"
        }
        Write-Host ""
    }
}
