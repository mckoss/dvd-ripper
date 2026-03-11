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
    $files = @(Find-Files $Mp4ArchiveDir '*.mp4')
    if ($files.Count -eq 0) {
        Write-Host "No MP4 files found in archive: $Mp4ArchiveDir"
        exit 0
    }

    Write-Host "Archive mode: uploading from $Mp4ArchiveDir" -ForegroundColor Cyan
    Write-Host "Found $($files.Count) MP4(s) in archive." -ForegroundColor Gray

    # List remote files once to check for duplicates (recursive listing)
    Write-Host "Checking remote for existing files..." -ForegroundColor Gray
    $remoteList = @(rclone lsf gdrive:Movies --files-only --recursive 2>$null)

    $toUpload = @()
    $alreadyRemote = @()
    foreach ($file in $files) {
        $relPath = $file.FullName.Substring($Mp4ArchiveDir.Length).TrimStart('\', '/') -replace '\\', '/'
        if ($relPath -in $remoteList) {
            $alreadyRemote += $relPath
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
    $toUpload | ForEach-Object {
        $relPath = $_.FullName.Substring($Mp4ArchiveDir.Length).TrimStart('\', '/')
        Write-Host "  $relPath"
    }

    # Build include filters using relative paths so rclone only processes files we know are new
    $includeArgs = @()
    foreach ($f in $toUpload) {
        $relPath = $f.FullName.Substring($Mp4ArchiveDir.Length).TrimStart('\', '/') -replace '\\', '/'
        $includeArgs += "--include"
        $includeArgs += $relPath
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
$files = @(Find-Files $EncodedDir '*.mp4')
$excludeArgs = @()
$skipped = @()

foreach ($file in $files) {
    try {
        $stream = [System.IO.File]::Open($file.FullName, 'Open', 'Read', 'None')
        $stream.Close()
    } catch {
        $relPath = $file.FullName.Substring($EncodedDir.Length).TrimStart('\', '/') -replace '\\', '/'
        $skipped += $relPath
        $excludeArgs += "--exclude"
        $excludeArgs += $relPath
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
$unlocked = $files | Where-Object {
    $relPath = $_.FullName.Substring($EncodedDir.Length).TrimStart('\', '/') -replace '\\', '/'
    $relPath -notin $skipped
}

if ($script:lastUploaded.Count -gt 0) {
    Write-Host "`nFiles uploaded ($($script:lastUploaded.Count)):"
    $script:lastUploaded | ForEach-Object { Write-Host "  $_" }
}

if ($unlocked.Count -gt 0) {
    $response = Read-Host "`nMove $($unlocked.Count) file(s) to $Mp4ArchiveDir? (y/N)"
    if ($response -eq 'y') {
        $unlocked | ForEach-Object {
            $relPath = $_.FullName.Substring($EncodedDir.Length).TrimStart('\', '/')
            $destPath = Join-Path $Mp4ArchiveDir $relPath
            $destDir = Split-Path $destPath -Parent
            if (-not (Test-Path $destDir)) {
                New-Item -ItemType Directory -Path $destDir -Force | Out-Null
            }
            Move-Item -Path $_.FullName -Destination $destPath -Force
            Write-Host "  Moved: $relPath"
        }
        # Clean up empty subdirectories left behind
        Remove-EmptyDirs $EncodedDir
    }
} else {
    Write-Host "`nNo files ready to move (all locked or none found)."
}
