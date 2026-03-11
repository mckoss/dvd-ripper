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

# Step 1: Rename -check-title MKVs using TMDb lookup
$needsRenaming = @(Find-Files $RippedDir '*-check-title.mkv')

if ($needsRenaming.Count -gt 0) {
    Write-Host "$($needsRenaming.Count) MKV(s) need title confirmation:" -ForegroundColor Cyan
    $needsRenaming | ForEach-Object { Write-Host "  $($_.Name)" -ForegroundColor Gray }
    Write-Host ''

    $TmdbApiKey = Initialize-TmdbApiKey
    $titleRenamed = 0
    $titleSkipped = 0

    foreach ($mkv in $needsRenaming) {
        # Strip drive letter + -check-title suffix to get raw disc title
        $rawBase = $mkv.BaseName -replace '-[A-Z]-check-title$', ''
        Write-Host '-------------------------------------------------------------' -ForegroundColor DarkGray
        Write-Host "  File: $($mkv.Name)" -ForegroundColor White
        Write-Host "       $($mkv.DirectoryName)" -ForegroundColor DarkGray

        $result = Invoke-TmdbRenamePrompt -FilePath $mkv.FullName -RawBaseName $rawBase -ApiKey $TmdbApiKey

        if ($result.NewBase) {
            $newName = $result.NewBase + '.mkv'
            $newPath = Join-Path $mkv.DirectoryName $newName
            if (Test-Path $newPath) {
                Write-Host "  SKIPPED (target exists): $newName" -ForegroundColor Yellow
                $titleSkipped++
            } else {
                Rename-Item -Path $mkv.FullName -NewName $newName -ErrorAction Stop
                Write-Host "  Renamed: $($mkv.Name) -> $newName" -ForegroundColor Green
                $titleRenamed++
            }
        } else {
            Write-Host '  Skipped. Rename this file manually and re-run.' -ForegroundColor Yellow
            $titleSkipped++
        }
        Write-Host ''
    }

    Write-Host ''
    Write-Host "Title confirmation: $titleRenamed renamed, $titleSkipped skipped." -ForegroundColor Cyan
    if ($titleSkipped -gt 0) {
        Write-Host "  $titleSkipped file(s) still need manual renaming before encoding." -ForegroundColor Yellow
    }
    Write-Host ''
}

# Re-scan after renaming (files may have new names now)
$needsRenaming = @(Find-Files $RippedDir '*-check-title.mkv')
if ($needsRenaming.Count -gt 0) {
    Write-Host "$($needsRenaming.Count) file(s) still have -check-title suffix:" -ForegroundColor Yellow
    $needsRenaming | ForEach-Object { Write-Host "  $($_.Name)" }
    Write-Host "Remove the suffix manually before encoding.`n" -ForegroundColor Yellow
}

# Step 1b: Catch MKVs missing a (Year) in the filename
$missingYear = @(Find-Files $RippedDir '*.mkv' | Where-Object {
    $_.BaseName -notmatch '\(\d{4}\)' -and $_.BaseName -notmatch '-check-title$'
})

if ($missingYear.Count -gt 0) {
    Write-Host "$($missingYear.Count) MKV(s) missing year - need title confirmation:" -ForegroundColor Cyan
    $missingYear | ForEach-Object { Write-Host "  $($_.Name)" -ForegroundColor Gray }
    Write-Host ''

    if (-not $TmdbApiKey) { $TmdbApiKey = Initialize-TmdbApiKey }
    if (-not $titleRenamed) { $titleRenamed = 0 }
    if (-not $titleSkipped) { $titleSkipped = 0 }

    foreach ($mkv in $missingYear) {
        Write-Host '-------------------------------------------------------------' -ForegroundColor DarkGray
        Write-Host "  File: $($mkv.Name)" -ForegroundColor White
        Write-Host "       $($mkv.DirectoryName)" -ForegroundColor DarkGray

        $result = Invoke-TmdbRenamePrompt -FilePath $mkv.FullName -RawBaseName $mkv.BaseName -ApiKey $TmdbApiKey

        if ($result.NewBase) {
            $newName = $result.NewBase + '.mkv'
            $newPath = Join-Path $mkv.DirectoryName $newName
            if (Test-Path $newPath) {
                Write-Host "  SKIPPED (target exists): $newName" -ForegroundColor Yellow
                $titleSkipped++
            } else {
                Rename-Item -Path $mkv.FullName -NewName $newName -ErrorAction Stop
                Write-Host "  Renamed: $($mkv.Name) -> $newName" -ForegroundColor Green
                $titleRenamed++
            }
        } else {
            Write-Host '  Skipped. Rename this file manually and re-run.' -ForegroundColor Yellow
            $titleSkipped++
        }
        Write-Host ''
    }

    Write-Host ''
    Write-Host "Title confirmation: $titleRenamed renamed, $titleSkipped skipped." -ForegroundColor Cyan
    if ($titleSkipped -gt 0) {
        Write-Host "  $titleSkipped file(s) still need manual renaming before encoding." -ForegroundColor Yellow
    }
    Write-Host ''
}

# Step 2: Check for MKVs that already have a corresponding MP4 in encoded-for-upload or MP4s archive
# Use relative paths so subfolder structure is respected (e.g. "Show Name\S01E01" matches across dirs)
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

# Step 3: Move already-encoded MKVs to the archive (preserving subfolder structure)
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
    # Clean up empty subdirectories left behind in ripped-for-encoding
    Remove-EmptyDirs $RippedDir
    Write-Host ""
}

# Step 4: Find MKVs that still need encoding
$needsEncoding = @(Find-Files $RippedDir '*.mkv' | Where-Object {
    $relBase = Get-RelativeBase -FilePath $_.FullName -BaseDir $RippedDir
    $relBase -notin $existingMp4RelPaths
})

if ($needsEncoding.Count -eq 0) {
    Write-Host "No MKV files need encoding."
} else {
    Write-Host "MKV files ready for encoding ($($needsEncoding.Count)):"
    $needsEncoding | ForEach-Object {
        $relPath = $_.FullName.Substring($RippedDir.Length).TrimStart('\', '/')
        Write-Host "  $relPath"
    }

    Write-Host "`nHandBrake output must mirror subfolder structure under:" -ForegroundColor Gray
    Write-Host "  $EncodedDir" -ForegroundColor Gray
    $response = Read-Host "`nLaunch HandBrake on the ripped-for-encoding folder? (y/N)"
    if ($response -eq 'y') {
        Write-Host "`nLaunching HandBrake..."
        Write-Host "Set HandBrake output to: $EncodedDir"
        Start-Process "C:\Program Files\HandBrake\HandBrake.exe" -ArgumentList "`"$RippedDir`""
    }
}
