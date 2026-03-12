<#
.SYNOPSIS
    Launches parallel ripping sessions for all optical drives.

.DESCRIPTION
    Detects optical drives, lets you select which to use, then opens
    Windows Terminal with side-by-side panes running ripping-loop.ps1
    for each selected drive.

.PARAMETER MoviesDir
    Root movies directory (default: G:\Movies).

.PARAMETER MinLength
    Minimum title length in seconds to rip (default: 3600).

.EXAMPLE
    .\rip-all-drives.ps1
    .\rip-all-drives.ps1 -MinLength 1800
#>
param (
    [string]$MoviesDir,
    [int]$MinLength = 3600
)

if ([string]::IsNullOrWhiteSpace($MoviesDir)) {
    $DefaultMoviesDir = "G:\Movies"
    $MoviesDir = Read-Host "Enter movies directory (default: $DefaultMoviesDir)"
    if ([string]::IsNullOrWhiteSpace($MoviesDir)) {
        $MoviesDir = $DefaultMoviesDir
    }
}

# Enumerate optical drives via Windows API (works regardless of disc insertion)
Write-Host 'Scanning for optical drives...' -ForegroundColor Cyan
$cdroms = @(Get-CimInstance Win32_CDROMDrive | Sort-Object DeviceID)

if ($cdroms.Count -eq 0) {
    Write-Host 'No optical drives found.' -ForegroundColor Red
    exit 1
}

$drives = @()
for ($i = 0; $i -lt $cdroms.Count; $i++) {
    $drives += @{ Index = $i; Letter = $cdroms[$i].Drive; Name = $cdroms[$i].Name }
}

Write-Host ''
Write-Host "Found $($drives.Count) optical drive(s):" -ForegroundColor Green
foreach ($d in $drives) {
    Write-Host "  disc:$($d.Index) = $($d.Letter)  ($($d.Name))" -ForegroundColor White
}
Write-Host ''

# Ask which drives to use
$allLetters = ($drives | ForEach-Object { $_.Letter }) -join ', '
$response = Read-Host "Launch ripping on all drives ($allLetters)? [Y/n] or enter drive letters (e.g. D,F,H)"
if ([string]::IsNullOrWhiteSpace($response) -or $response -match '^[Yy]$') {
    $selectedDrives = $drives
} else {
    $requestedLetters = $response.ToUpper() -split '[,\s]+' | ForEach-Object { $_.TrimEnd(':') + ':' }
    $selectedDrives = $drives | Where-Object { $_.Letter -in $requestedLetters }
    if ($selectedDrives.Count -eq 0) {
        Write-Host 'No matching drives found.' -ForegroundColor Red
        exit 1
    }
}

Write-Host ''

# Layout: master pane (small strip at top), then drive panes split vertically (side by side) below.
# First split-pane -H creates a large bottom area (0.9 = 90% of height).
# Subsequent split-pane -V splits that bottom area into equal-width columns.
$scriptPath = Join-Path $PSScriptRoot 'ripping-loop.ps1'
$driveCount = $selectedDrives.Count
$wtParts = @()

for ($i = 0; $i -lt $driveCount; $i++) {
    $d = $selectedDrives[$i]
    $driveLetter = $d.Letter.TrimEnd(':')
    $discIdx = $d.Index
    $title = "Drive $($d.Letter)"
    $fileArgs = "powershell -NoExit -File `"$scriptPath`" -InputDrive `"$driveLetter`" -MoviesDir `"$MoviesDir`" -DiscIndex $discIdx -MinLength $MinLength"

    if ($i -eq 0) {
        # First drive: split horizontally from master, taking 90% of height
        $wtParts += "split-pane -H -s 0.9 --title `"$title`" $fileArgs"
    } else {
        # Subsequent drives: split vertically (side by side) within the bottom area
        $size = [math]::Round(1 - (1 / ($driveCount - $i + 1)), 4)
        $wtParts += "split-pane -V -s $size --title `"$title`" $fileArgs"
    }
    Write-Host "  $title (disc:$discIdx)" -ForegroundColor Green
}

$wtCmdLine = '-w 0 ' + ($wtParts -join ' ; ')
Start-Process wt -ArgumentList $wtCmdLine

Write-Host ''
Write-Host "Launched $driveCount ripping pane(s)." -ForegroundColor Cyan
Write-Host 'Resize panes: Alt+Shift+Arrow | Switch panes: Alt+Arrow' -ForegroundColor DarkGray
