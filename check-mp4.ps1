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

# Check for MKV files with timestamps in their names (e.g., TITLE-26-02-22-17-54.mkv)
$timestamped = Get-ChildItem -Path $MkvDir -Filter *.mkv | Where-Object {
    $_.BaseName -match '-\d{2}-\d{2}-\d{2}-\d{2}-\d{2}$'
}

if ($timestamped.Count -gt 0) {
    Write-Host "The following MKV files still have timestamps in their names."
    Write-Host "Please rename them with the correct movie title before proceeding:`n"
    $timestamped | ForEach-Object { Write-Host "  $($_.Name)" }
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

    $response = Read-Host "`nMove these files to a TBD folder and open HandBrake? (y/N)"
    if ($response -eq 'y') {
        $TbdDir = Join-Path $MkvDir "TBD"
        if (-not (Test-Path $TbdDir)) {
            New-Item -ItemType Directory -Path $TbdDir -Force | Out-Null
        }

        $missing | ForEach-Object {
            Move-Item -Path $_.FullName -Destination $TbdDir -Force
            Write-Host "  Moved: $($_.Name)"
        }

        Write-Host "`nLaunching HandBrake..."
        Start-Process "C:\Program Files\HandBrake\HandBrake.exe" -ArgumentList "`"$TbdDir`""
    }
}
