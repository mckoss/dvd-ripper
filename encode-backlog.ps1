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

# Step 1: Rename -check-title MKVs using TMDb lookup
$needsRenaming = @(Get-ChildItem -Path $RippedDir -Filter *-check-title.mkv)

if ($needsRenaming.Count -gt 0) {
    Write-Host "$($needsRenaming.Count) MKV(s) need title confirmation:" -ForegroundColor Cyan
    $needsRenaming | ForEach-Object { Write-Host "  $($_.Name)" -ForegroundColor Gray }
    Write-Host ''

    $TmdbApiKey = Initialize-TmdbApiKey
    $titleRenamed = 0
    $titleSkipped = 0

    foreach ($mkv in $needsRenaming) {
        # Strip -check-title suffix to get raw disc title
        $rawBase = $mkv.BaseName -replace '-check-title$', ''
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
$needsRenaming = @(Get-ChildItem -Path $RippedDir -Filter *-check-title.mkv)
if ($needsRenaming.Count -gt 0) {
    Write-Host "$($needsRenaming.Count) file(s) still have -check-title suffix:" -ForegroundColor Yellow
    $needsRenaming | ForEach-Object { Write-Host "  $($_.Name)" }
    Write-Host "Remove the suffix manually before encoding.`n" -ForegroundColor Yellow
}

# Step 2: Check for MKVs that already have a corresponding MP4 in encoded-for-upload or MP4s archive
$existingMp4Titles = @()
if (Test-Path $EncodedDir) {
    $existingMp4Titles += Get-ChildItem -Path $EncodedDir -Filter *.mp4 | ForEach-Object { $_.BaseName }
}
if (Test-Path $Mp4ArchiveDir) {
    $existingMp4Titles += Get-ChildItem -Path $Mp4ArchiveDir -Filter *.mp4 | ForEach-Object { $_.BaseName }
}

# Step 3: Move already-encoded MKVs to the archive
$alreadyEncoded = Get-ChildItem -Path $RippedDir -Filter *.mkv | Where-Object {
    $_.BaseName -in $existingMp4Titles
}

if ($alreadyEncoded.Count -gt 0) {
    Write-Host "Moving already-encoded MKVs to archive ($($alreadyEncoded.Count)):"
    $alreadyEncoded | ForEach-Object {
        try {
            Move-Item -Path $_.FullName -Destination $MkvArchiveDir -Force -ErrorAction Stop
            Write-Host "  Archived: $($_.Name)"
        } catch {
            Write-Host "  Skipped (file in use): $($_.Name)" -ForegroundColor Yellow
        }
    }
    Write-Host ""
}

# Step 4: Find MKVs that still need encoding
$needsEncoding = Get-ChildItem -Path $RippedDir -Filter *.mkv | Where-Object {
    $_.BaseName -notin $existingMp4Titles
}

if ($needsEncoding.Count -eq 0) {
    Write-Host "No MKV files need encoding."
} else {
    Write-Host "MKV files ready for encoding ($($needsEncoding.Count)):"
    $needsEncoding | ForEach-Object { Write-Host "  $($_.Name)" }

    $response = Read-Host "`nLaunch HandBrake on the ripped-for-encoding folder? (y/N)"
    if ($response -eq 'y') {
        Write-Host "`nLaunching HandBrake..."
        Write-Host "Set HandBrake output to: $EncodedDir"
        Start-Process "C:\Program Files\HandBrake\HandBrake.exe" -ArgumentList "`"$RippedDir`""
    }
}
