$SourceDir = "G:\Movies\MP4s"

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

rclone copy $SourceDir gdrive:Movies --ignore-existing --progress @excludeArgs
