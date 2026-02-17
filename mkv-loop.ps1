$MakeMkvPath = "C:\Program Files (x86)\MakeMKV\makemkvcon64.exe"
$DriveLetter = "D:"
$OutputDir = "C:\Users\Mike\Videos\Ripped Movies\MKVs"
$MinLength = 3600 # Minimum title length in seconds (1 hour) to isolate the main feature

# Ensure base output directory exists
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

Write-Host "Starting MakeMKV Automation Loop..."

while ($true) {
    Write-Host "Waiting for disc insertion in drive $DriveLetter..."

    # 1. Poll WMI until media is loaded
    while ($true) {
        $drive = Get-CimInstance Win32_CDROMDrive | Where-Object { $_.Drive -eq $DriveLetter }
        if ($drive.MediaLoaded) { break }
        Start-Sleep -Seconds 5
    }

    # 2. Extract and sanitize Volume Name
    $VolumeName = $drive.VolumeName
    if ([string]::IsNullOrWhiteSpace($VolumeName)) { $VolumeName = "UNKNOWN_DISC" }

    # Strip illegal characters for Windows file systems
    $SafeVolumeName = $VolumeName -replace '[\\/:*?"<>|]', '_'
    $TitleOutputDir = Join-Path $OutputDir $SafeVolumeName

    if (-not (Test-Path $TitleOutputDir)) {
        New-Item -ItemType Directory -Path $TitleOutputDir -Force | Out-Null
    }

    Write-Host "Disc detected: $VolumeName. Starting MakeMKV..."

    # 3. Execute MakeMKV
    # 'disc:0' targets the first optical drive. '--minlength' filters out extras.
    $ArgumentList = "mkv disc:0 all `"$TitleOutputDir`" --minlength=$MinLength"
    $process = Start-Process -FilePath $MakeMkvPath -ArgumentList $ArgumentList -Wait -NoNewWindow -PassThru

    if ($process.ExitCode -eq 0) {
        Write-Host "Extraction complete for $VolumeName. Please swap the disc."
    } else {
        Write-Host "MakeMKV exited with code $($process.ExitCode). Please check the disc or logs."
    }

    # 4. Loop sound until media is removed (tray opened)
    while ($true) {
        $checkDrive = Get-CimInstance Win32_CDROMDrive | Where-Object { $_.Drive -eq $DriveLetter }
        if (-not $checkDrive.MediaLoaded) { break }

        [System.Media.SystemSounds]::Beep.Play()
        Start-Sleep -Seconds 2
    }

    Write-Host "Disc removed. Resetting for next disc..."
    Start-Sleep -Seconds 5
}