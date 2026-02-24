$SourceDir = "G:\Movies\MP4s\Upload"

# Build exclude list for files that are locked (actively being encoded)
$files = Get-ChildItem -Path $SourceDir -Filter *.mp4
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

if ($skipped.Count -gt 0) {
    Write-Host "Skipping locked files (still being encoded):"
    $skipped | ForEach-Object { Write-Host "  $_" }
    Write-Host ""
}

$logFile = Join-Path $env:TEMP "rclone-upload.log"
rclone copy $SourceDir gdrive:Movies --ignore-existing --progress --log-file $logFile --log-level INFO @excludeArgs

# Parse rclone log for actually transferred files
$uploaded = @()
if (Test-Path $logFile) {
    $logContent = Get-Content $logFile
    $uploaded = $logContent | Where-Object { $_ -match ': Copied \(new\)' } | ForEach-Object {
        if ($_ -match 'INFO\s+:\s+(.+?)\s*:\s+Copied') { $Matches[1] }
    }
    if ($uploaded.Count -eq 0) {
        Write-Host "`nDEBUG: Log file contents:"
        $logContent | ForEach-Object { Write-Host "  $_" }
    }
    Remove-Item $logFile -Force
}

if ($uploaded.Count -gt 0) {
    Write-Host "`nFiles uploaded ($($uploaded.Count)):"
    $uploaded | ForEach-Object { Write-Host "  $_" }

    $response = Read-Host "`nMove these file(s) to G:\Movies\MP4s? (y/N)"
    if ($response -eq 'y') {
        $DestDir = "G:\Movies\MP4s"
        $uploaded | ForEach-Object {
            $filePath = Join-Path $SourceDir $_
            if (Test-Path $filePath) {
                Move-Item -Path $filePath -Destination $DestDir -Force
                Write-Host "  Moved: $_"
            }
        }
    }
} else {
    Write-Host "`nNo new files were uploaded."
}
