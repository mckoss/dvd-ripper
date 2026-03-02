param (
    [string]$MoviesDir
)

$MakeMkvPath = "C:\Program Files (x86)\MakeMKV\makemkvcon64.exe"

if ([string]::IsNullOrWhiteSpace($MoviesDir)) {
    $DefaultMoviesDir = "G:\Movies"
    $MoviesDir = Read-Host "Enter movies directory (default: $DefaultMoviesDir)"
    if ([string]::IsNullOrWhiteSpace($MoviesDir)) {
        $MoviesDir = $DefaultMoviesDir
    }
}

# Enumerate disc indices via MakeMKV (one-time scan before any ripping starts)
Write-Host 'Scanning for optical drives...' -ForegroundColor Cyan
$infoOutput = & $MakeMkvPath -r info disc:9999 2>&1 | Out-String

$drives = @()
foreach ($line in $infoOutput -split "`n") {
    # DRV lines: DRV:index,visible,unknown,unknown,"Name","DevicePath","DiscLabel"
    # On Windows DevicePath is like \Device\CdRom0 — drive letter comes from the OS
    if ($line -match '^DRV:(\d+),(\d+),') {
        $idx = [int]$Matches[1]
        $visible = [int]$Matches[2]
        if ($visible -gt 0) {
            $drives += @{ Index = $idx; Line = $line.Trim() }
        }
    }
}

if ($drives.Count -eq 0) {
    Write-Host 'No optical drives found.' -ForegroundColor Red
    exit 1
}

# Map disc indices to drive letters using Win32_CDROMDrive (sorted by DeviceID to match MakeMKV order)
$cdroms = @(Get-CimInstance Win32_CDROMDrive | Sort-Object DeviceID)

Write-Host ''
Write-Host "Found $($drives.Count) optical drive(s):" -ForegroundColor Green
foreach ($d in $drives) {
    $letter = '??'
    if ($d.Index -lt $cdroms.Count) {
        $letter = $cdroms[$d.Index].Drive
    }
    $d.Letter = $letter
    Write-Host "  disc:$($d.Index) = $letter" -ForegroundColor White
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

# Launch a ripping-loop window for each selected drive
$scriptPath = Join-Path $PSScriptRoot 'ripping-loop.ps1'
Write-Host ''
foreach ($d in $selectedDrives) {
    $driveLetter = $d.Letter.TrimEnd(':')
    $discIdx = $d.Index
    $title = "Ripping - Drive $($d.Letter)"
    $cmd = "& '$scriptPath' -InputDrive '$driveLetter' -MoviesDir '$MoviesDir' -DiscIndex $discIdx"

    Write-Host "Launching: $title (disc:$discIdx)" -ForegroundColor Green
    Start-Process powershell -ArgumentList "-NoExit", "-Command", "(`$Host.UI.RawUI.WindowTitle = '$title'); $cmd"
}

Write-Host ''
Write-Host "Launched $($selectedDrives.Count) ripping window(s). You can close this window." -ForegroundColor Cyan
