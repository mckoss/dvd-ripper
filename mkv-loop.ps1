param (
    [string]$InputDrive,
    [string]$OutputDir
)

$MakeMkvPath = "C:\Program Files (x86)\MakeMKV\makemkvcon64.exe"

if ([string]::IsNullOrWhiteSpace($InputDrive)) {
    $DefaultDrive = "D"
    $InputDrive = Read-Host "Enter input drive letter (default: $DefaultDrive)"
    if ([string]::IsNullOrWhiteSpace($InputDrive)) {
        $InputDrive = $DefaultDrive
    }
}
$InputDriveLetter = "$($InputDrive.TrimEnd(':')):"

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $DefaultOutputDir = "G:\Movies\MKVs"
    $OutputDir = Read-Host "Enter output directory for ripped MKVs (default: $DefaultOutputDir)"
    if ([string]::IsNullOrWhiteSpace($OutputDir)) {
        $OutputDir = $DefaultOutputDir
    }
}
$MinLength = 3600
$AlertSoundPath = Join-Path $PSScriptRoot "alert.wav"

# Create a single sound player instance to reuse
$soundPlayer = $null
if (Test-Path $AlertSoundPath) {
    $soundPlayer = New-Object System.Media.SoundPlayer
    $soundPlayer.SoundLocation = $AlertSoundPath
}


if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

Write-Host "Starting MakeMKV Automation Loop..."

while ($true) {
    Write-Host "`nWaiting for disc insertion in drive $InputDriveLetter..."

    # Exponential backoff beep until media is inserted
    $delaySeconds = 2
    $maxDelaySeconds = 300
    $timeSinceLastBeep = 0

    while ($true) {
        if (Test-Path "$InputDriveLetter\") { break }

        if ($timeSinceLastBeep -ge $delaySeconds) {
            # Play WAV file if it exists, otherwise fall back to system beep
            if ($null -ne $soundPlayer) {
                $soundPlayer.PlaySync() # Play the sound synchronously (wait for completion)
            } else {
                [System.Media.SystemSounds]::Beep.Play()
            }
            $timeSinceLastBeep = 0

            $delaySeconds *= 2
            if ($delaySeconds -gt $maxDelaySeconds) {
                $delaySeconds = $maxDelaySeconds
            }
        }

        Start-Sleep -Seconds 2
        $timeSinceLastBeep += 2
    }

    # Extract and sanitize Volume Name (retry once if blank, to filter transient detections)
    $volume = Get-Volume -DriveLetter $InputDriveLetter.TrimEnd(':') -ErrorAction SilentlyContinue
    $VolumeName = $volume.FileSystemLabel
    if ([string]::IsNullOrWhiteSpace($VolumeName)) {
        Write-Host "No volume label detected. Waiting for disc to settle..."
        Start-Sleep -Seconds 5
        if (-not (Test-Path "$InputDriveLetter\")) {
            Write-Host "Drive no longer accessible. Retrying..."
            continue
        }
        $volume = Get-Volume -DriveLetter $InputDriveLetter.TrimEnd(':') -ErrorAction SilentlyContinue
        $VolumeName = $volume.FileSystemLabel
        if ([string]::IsNullOrWhiteSpace($VolumeName)) { $VolumeName = "UNKNOWN_DISC" }
    }

    # Add timestamp to ensure unique folder names
    $timestamp = Get-Date -Format "yy-MM-dd-HH-mm"
    $SafeVolumeName = $VolumeName -replace '[\\/:*?"<>|]', '_'
    $UniqueVolumeName = "$SafeVolumeName-$timestamp"
    $FinalOutputFile = Join-Path $OutputDir "$UniqueVolumeName.mkv"
    $TempOutputDir = Join-Path $OutputDir $UniqueVolumeName

    Write-Host "Disc detected: $VolumeName. Extracting to staging folder..."

    # Execute MakeMKV asynchronously using dev: to target specific drive
    if (-not (Test-Path $TempOutputDir)) {
        New-Item -ItemType Directory -Path $TempOutputDir -Force | Out-Null
    }
    $ArgumentList = "mkv dev:$InputDriveLetter all `"$TempOutputDir`" --minlength=$MinLength"
    $process = Start-Process -FilePath $MakeMkvPath -ArgumentList $ArgumentList -NoNewWindow -PassThru

    Write-Host "Extraction started. Monitoring file size..."

    # Monitor file size
    $progressCounter = 0
    $trackedMkvFile = $null
    while (-not $process.HasExited) {
        Start-Sleep -Seconds 2
        $progressCounter += 2

        # Write progress every 60 seconds (30 iterations of 2-second checks)
        if ($progressCounter -ge 60) {
            # Only search for MKV file if we haven't found one yet
            if ($null -eq $trackedMkvFile) {
                $mkvFile = Get-ChildItem -Path $TempOutputDir -Filter *.mkv | Sort-Object Length -Descending | Select-Object -First 1
                if ($mkvFile) {
                    $trackedMkvFile = $mkvFile
                    Write-Host "Detected MKV file: $($trackedMkvFile.Name)"
                }
            }

            # Show progress using the tracked file
            if ($null -ne $trackedMkvFile) {
                # Refresh the file info to get current size
                $trackedMkvFile.Refresh()
                $sizeMB = [math]::Round($trackedMkvFile.Length / 1MB, 2)
                Write-Host "Progress: $sizeMB MB written..."
            }
            $progressCounter = 0
        }
    }

    # Wait for process to fully exit and get final exit code
    $process.WaitForExit()
    $exitCode = $process.ExitCode

    # Move file and cleanup
    if ($null -eq $exitCode -or $exitCode -eq 0) {
        # If we didn't track an MKV during progress monitoring, scan the temp dir now
        if ($null -eq $trackedMkvFile -or -not (Test-Path $trackedMkvFile.FullName)) {
            $foundMkv = Get-ChildItem -Path $TempOutputDir -Filter *.mkv -ErrorAction SilentlyContinue |
                Sort-Object Length -Descending | Select-Object -First 1
            if ($foundMkv) {
                $trackedMkvFile = $foundMkv
                Write-Host "Found MKV file after completion: $($trackedMkvFile.Name)"
            }
        }

        if ($null -ne $trackedMkvFile -and (Test-Path $trackedMkvFile.FullName)) {
            Write-Host "MakeMKV completed successfully."
            # Use the MKV filename if it's more descriptive than the volume label
            $mkvBaseName = $trackedMkvFile.BaseName -replace '-[A-Z]\d+_t\d+$', ''  # Strip track suffix like -A5_t00
            $genericLabels = @('UNKNOWN_DISC', 'DVD_VIDEO', 'DVDVOLUME', 'DVD')
            if ($mkvBaseName.Length -gt 0 -and $genericLabels -contains $VolumeName) {
                $SafeMkvName = $mkvBaseName -replace '[\\/:*?"<>|]', '_'
                $FinalOutputFile = Join-Path $OutputDir "$SafeMkvName-$timestamp.mkv"
                Write-Host "Using MKV filename '$mkvBaseName' instead of volume label '$VolumeName'"
            }
            Move-Item -Path $trackedMkvFile.FullName -Destination $FinalOutputFile -Force
            Write-Host "Extraction complete. Saved as: $FinalOutputFile"

            # Only delete temp directory if it's empty
            if (Test-Path $TempOutputDir) {
                $remaining = Get-ChildItem -Path $TempOutputDir -ErrorAction SilentlyContinue
                if ($remaining.Count -eq 0) {
                    Remove-Item -Path $TempOutputDir -Force
                } else {
                    Write-Host "Temp folder not empty - additional files left for inspection: $TempOutputDir"
                    $remaining | ForEach-Object { Write-Host "  $($_.Name)" }
                }
            }
        } else {
            Write-Host "WARNING: MakeMKV produced no output files. Disc may be damaged or unreadable."
            # Clean up empty temp directory
            if ((Test-Path $TempOutputDir) -and (Get-ChildItem -Path $TempOutputDir -ErrorAction SilentlyContinue).Count -eq 0) {
                Remove-Item -Path $TempOutputDir -Force
            }
        }
    } else {
        Write-Host "MakeMKV exited with code $exitCode."
    }

    # Eject drive
    Write-Host "Ejecting drive $InputDriveLetter..."
    $shell = New-Object -ComObject Shell.Application
    $shellDrive = $shell.Namespace(17).ParseName($InputDriveLetter)
    if ($shellDrive) {
        $shellDrive.InvokeVerb("Eject")
    }
}