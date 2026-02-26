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

$EncodedDir = Join-Path $MoviesDir "processing\encoded-for-upload"
$Mp4ArchiveDir = Join-Path $MoviesDir "MP4s"

# Ensure directories exist
foreach ($dir in @($EncodedDir, $Mp4ArchiveDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

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

$logFile = Join-Path $env:TEMP "rclone-upload.log"
rclone copy $EncodedDir gdrive:Movies --ignore-existing --progress -v --log-file $logFile @excludeArgs

# Parse rclone log for actually transferred files
$uploaded = @()
if (Test-Path $logFile) {
    $logContent = Get-Content $logFile
    $uploaded = $logContent | Where-Object { $_ -match ': Copied \(new\)' } | ForEach-Object {
        if ($_ -match 'INFO\s+:\s+(.+?)\s*:\s+Copied') { $Matches[1] }
    }
    if ($uploaded.Count -eq 0 -and $files.Count -gt $skipped.Count) {
        # Check if all files already existed on remote
        $alreadyExisted = $logContent | Where-Object { $_ -match 'Skipped' -or $_ -match 'existing' }
        if ($alreadyExisted.Count -eq 0) {
            Write-Host "`nDEBUG: Log file contents:"
            $logContent | ForEach-Object { Write-Host "  $_" }
        }
    }
    Remove-Item $logFile -Force
}

# Determine which unlocked files can be moved (uploaded or already on remote)
$unlocked = $files | Where-Object { $_.Name -notin $skipped }

if ($uploaded.Count -gt 0) {
    Write-Host "`nFiles uploaded ($($uploaded.Count)):"
    $uploaded | ForEach-Object { Write-Host "  $_" }
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
