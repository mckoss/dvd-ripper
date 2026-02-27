param (
    [string]$MoviesDir,
    [switch]$Archive
)

if ([string]::IsNullOrWhiteSpace($MoviesDir)) {
    $DefaultMoviesDir = "G:\Movies"
    $MoviesDir = Read-Host "Enter movies directory (default: $DefaultMoviesDir)"
    if ([string]::IsNullOrWhiteSpace($MoviesDir)) {
        $MoviesDir = $DefaultMoviesDir
    }
}

$EncodedDir = Join-Path $MoviesDir "processing\encoded-for-upload"
$Mp4ArchiveDir = Join-Path $MoviesDir "MP4s"

# Ensure directories exist
foreach ($dir in @($EncodedDir, $Mp4ArchiveDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

function Invoke-RcloneUpload {
    param(
        [string]$SourceDir,
        [string[]]$ExcludeArgs = @()
    )
    $logFile = Join-Path $env:TEMP "rclone-upload.log"
    rclone copy $SourceDir gdrive:Movies --ignore-existing --progress -v --log-file $logFile @ExcludeArgs

    $script:lastUploaded = @()
    if (Test-Path $logFile) {
        $logContent = Get-Content $logFile
        $script:lastUploaded = @($logContent | Where-Object { $_ -match ': Copied \(new\)' } | ForEach-Object {
            if ($_ -match 'INFO\s+:\s+(.+?)\s*:\s+Copied') { $Matches[1] }
        })
        Remove-Item $logFile -Force
    }
}

if ($Archive) {
    # ── Archive mode: upload from MP4s archive folder, skip already-existing ──
    $files = @(Get-ChildItem -Path $Mp4ArchiveDir -Filter *.mp4 -File)
    if ($files.Count -eq 0) {
        Write-Host "No MP4 files found in archive: $Mp4ArchiveDir"
        exit 0
    }

    Write-Host "Archive mode: uploading from $Mp4ArchiveDir" -ForegroundColor Cyan
    Write-Host "Found $($files.Count) MP4(s) in archive." -ForegroundColor Gray

    # List remote files once to check for duplicates
    Write-Host "Checking remote for existing files..." -ForegroundColor Gray
    $remoteList = @(rclone lsf gdrive:Movies --files-only 2>$null)

    $toUpload = @()
    $alreadyRemote = @()
    foreach ($file in $files) {
        if ($file.Name -in $remoteList) {
            $alreadyRemote += $file.Name
        } else {
            $toUpload += $file
        }
    }

    if ($alreadyRemote.Count -gt 0) {
        Write-Host "`nAlready on remote ($($alreadyRemote.Count)):" -ForegroundColor DarkGreen
        $alreadyRemote | ForEach-Object { Write-Host "  $_" -ForegroundColor DarkGreen }
    }

    if ($toUpload.Count -eq 0) {
        Write-Host "`nAll files already exist on remote. Nothing to upload." -ForegroundColor Green
        exit 0
    }

    Write-Host "`nWill upload $($toUpload.Count) new file(s):" -ForegroundColor White
    $toUpload | ForEach-Object { Write-Host "  $($_.Name)" }

    # Build include filters so rclone only processes files we know are new
    $includeArgs = @()
    foreach ($f in $toUpload) {
        $includeArgs += "--include"
        $includeArgs += $f.Name
    }

    Write-Host "`nUploading to gdrive:Movies ..." -ForegroundColor Cyan
    Invoke-RcloneUpload -SourceDir $Mp4ArchiveDir -ExcludeArgs $includeArgs

    if ($script:lastUploaded.Count -gt 0) {
        Write-Host "`nUploaded ($($script:lastUploaded.Count)):" -ForegroundColor Green
        $script:lastUploaded | ForEach-Object { Write-Host "  $_" -ForegroundColor Green }
    } else {
        Write-Host "`nNo new files were uploaded." -ForegroundColor Yellow
    }
    exit 0
}

# ── Normal mode: upload from encoded-for-upload, then move to archive ─────────

# Build exclude list for files that are locked (actively being encoded)
$files = Get-ChildItem -Path $EncodedDir -Filter *.mp4
$excludeArgs = @()
$skipped = @()

foreach ($file in $files) {
    try {
        $stream = [System.IO.File]::Open($file.FullName, 'Open', 'Read', 'None')
        $stream.Close()
    } catch {
        $skipped += $file.Name
        $excludeArgs += "--exclude"
        $excludeArgs += $file.Name
    }
}

if ($files.Count -eq 0) {
    Write-Host "No MP4 files found in: $EncodedDir"
    exit 0
}

if ($skipped.Count -gt 0) {
    Write-Host "Skipping locked files (still being encoded):"
    $skipped | ForEach-Object { Write-Host "  $_" }
    Write-Host ""
}

$uploadCount = ($files | Where-Object { $_.Name -notin $skipped }).Count
Write-Host "Uploading $uploadCount file(s) to gdrive:Movies ..." -ForegroundColor Cyan
Invoke-RcloneUpload -SourceDir $EncodedDir -ExcludeArgs $excludeArgs

# Determine which unlocked files can be moved (uploaded or already on remote)
$unlocked = $files | Where-Object { $_.Name -notin $skipped }

if ($script:lastUploaded.Count -gt 0) {
    Write-Host "`nFiles uploaded ($($script:lastUploaded.Count)):"
    $script:lastUploaded | ForEach-Object { Write-Host "  $_" }
}

if ($unlocked.Count -gt 0) {
    $response = Read-Host "`nMove $($unlocked.Count) file(s) to $Mp4ArchiveDir? (y/N)"
    if ($response -eq 'y') {
        $unlocked | ForEach-Object {
            Move-Item -Path $_.FullName -Destination $Mp4ArchiveDir -Force
            Write-Host "  Moved: $($_.Name)"
        }
    }
} else {
    Write-Host "`nNo files ready to move (all locked or none found)."
}
