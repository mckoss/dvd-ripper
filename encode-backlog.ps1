param (
    [string]$MoviesDir
)

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

# Step 1: Check for MKV files that still need title confirmation (ending in -check-title)
$needsRenaming = Get-ChildItem -Path $RippedDir -Filter *-check-title.mkv

if ($needsRenaming.Count -gt 0) {
    Write-Host "The following MKV files need their title confirmed."
    Write-Host "Remove the '-check-title' suffix (and correct the name if needed) before proceeding:`n"
    $needsRenaming | ForEach-Object { Write-Host "  $($_.Name)" }
    exit 1
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
        Move-Item -Path $_.FullName -Destination $MkvArchiveDir -Force
        Write-Host "  Archived: $($_.Name)"
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
