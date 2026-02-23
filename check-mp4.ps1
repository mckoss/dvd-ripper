param (
    [string]$OutputDir
)

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $DefaultOutputDir = "G:\Movies"
    $OutputDir = Read-Host "Enter movies directory (default: $DefaultOutputDir)"
    if ([string]::IsNullOrWhiteSpace($OutputDir)) {
        $OutputDir = $DefaultOutputDir
    }
}

$MkvDir = Join-Path $OutputDir "MKVs"
$Mp4Dir = Join-Path $OutputDir "MP4s"

if (-not (Test-Path $MkvDir)) {
    Write-Host "MKVs directory not found: $MkvDir"
    exit 1
}

# Get MP4 titles (basenames without extension)
$mp4Titles = @()
if (Test-Path $Mp4Dir) {
    $mp4Titles = Get-ChildItem -Path $Mp4Dir -Filter *.mp4 | ForEach-Object { $_.BaseName }
}

# Find MKVs without a corresponding MP4
$missing = Get-ChildItem -Path $MkvDir -Filter *.mkv | Where-Object {
    $_.BaseName -notin $mp4Titles
}

if ($missing.Count -eq 0) {
    Write-Host "All MKV files have corresponding MP4s."
} else {
    Write-Host "MKV files without a corresponding MP4 ($($missing.Count)):"
    $missing | ForEach-Object { Write-Host "  $($_.Name)" }
}
