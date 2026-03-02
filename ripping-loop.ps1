param (
    [string]$InputDrive,
    [string]$MoviesDir
)

$MakeMkvPath = "C:\Program Files (x86)\MakeMKV\makemkvcon64.exe"

. "$PSScriptRoot\tmdb-helpers.ps1"
$TmdbApiKey = Initialize-TmdbApiKey

if ([string]::IsNullOrWhiteSpace($InputDrive)) {
    $DefaultDrive = "D"
    $InputDrive = Read-Host "Enter input drive letter (default: $DefaultDrive)"
    if ([string]::IsNullOrWhiteSpace($InputDrive)) {
        $InputDrive = $DefaultDrive
    }
}
$InputDriveLetter = "$($InputDrive.TrimEnd(':')):"

if ([string]::IsNullOrWhiteSpace($MoviesDir)) {
    $DefaultMoviesDir = "G:\Movies"
    $MoviesDir = Read-Host "Enter movies directory (default: $DefaultMoviesDir)"
    if ([string]::IsNullOrWhiteSpace($MoviesDir)) {
        $MoviesDir = $DefaultMoviesDir
    }
}
$OutputDir = Join-Path $MoviesDir "processing\ripped-for-encoding"
$MinLength = 3600
$AlertSoundPath = Join-Path $PSScriptRoot "alert.wav"

# Resolve drive letter to MakeMKV disc index to avoid scanning all drives
$DiscIndex = $null
Write-Host "Resolving disc index for drive $InputDriveLetter..."
$infoOutput = & $MakeMkvPath -r info disc:9999 2>&1 | Out-String
foreach ($line in $infoOutput -split "`n") {
    # Lines like: DRV:0,2,999,12,"BD-RE HL-DT-ST","/dev/sr0","DISC_LABEL"
    # On Windows the device field contains the drive letter like "D:"
    if ($line -match '^DRV:(\d+),\d+,\d+,\d+,' -and $line -match [regex]::Escape($InputDriveLetter.TrimEnd(':'))) {
        $DiscIndex = $Matches[1]
        break
    }
}
if ($null -ne $DiscIndex) {
    Write-Host "Mapped $InputDriveLetter to disc:$DiscIndex" -ForegroundColor Green
} else {
    Write-Host "Could not resolve disc index. Falling back to dev:$InputDriveLetter" -ForegroundColor Yellow
}

# Create a single sound player instance to reuse
$soundPlayer = $null
if (Test-Path $AlertSoundPath) {
    $soundPlayer = New-Object System.Media.SoundPlayer
    $soundPlayer.SoundLocation = $AlertSoundPath
}


if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

Write-Host "Starting Ripping Processing Loop..."
$lastRippedFile = $null

while ($true) {
    if ($lastRippedFile) {
        Write-Host "`nWaiting for disc in drive $InputDriveLetter... (press SPACE to confirm title of last rip)" -ForegroundColor Cyan
    } else {
        Write-Host "`nWaiting for disc insertion in drive $InputDriveLetter..."
    }

    # Exponential backoff beep until media is inserted
    $delaySeconds = 2
    $maxDelaySeconds = 900
    $timeSinceLastBeep = 0

    # Flush any buffered keystrokes
    while ([Console]::KeyAvailable) { [Console]::ReadKey($true) | Out-Null }

    while ($true) {
        if (Test-Path "$InputDriveLetter\") { break }

        # Check for SPACE to rename the last ripped file
        if ($lastRippedFile -and [Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            if ($key.Key -eq 'Spacebar') {
                Write-Host ''
                $rawBase = [System.IO.Path]::GetFileNameWithoutExtension($lastRippedFile) -replace '-check-title$', ''
                Write-Host '-------------------------------------------------------------' -ForegroundColor DarkGray
                Write-Host "  File: $([System.IO.Path]::GetFileName($lastRippedFile))" -ForegroundColor White
                Write-Host "       $([System.IO.Path]::GetDirectoryName($lastRippedFile))" -ForegroundColor DarkGray

                $tmdbResult = Invoke-TmdbRenamePrompt -FilePath $lastRippedFile -RawBaseName $rawBase -ApiKey $TmdbApiKey

                if ($tmdbResult.NewBase) {
                    $newName = $tmdbResult.NewBase + '.mkv'
                    $newPath = Join-Path $OutputDir $newName
                    if (Test-Path $newPath) {
                        Write-Host "  SKIPPED (target exists): $newName" -ForegroundColor Yellow
                    } else {
                        Rename-Item -Path $lastRippedFile -NewName $newName -ErrorAction Stop
                        Write-Host "  Renamed: $([System.IO.Path]::GetFileName($lastRippedFile)) -> $newName" -ForegroundColor Green
                    }
                } else {
                    Write-Host '  Skipped. Title can be confirmed later via encode-backlog.' -ForegroundColor Yellow
                }
                $lastRippedFile = $null
                Write-Host ''
                Write-Host "Waiting for disc insertion in drive $InputDriveLetter..."
            }
        }

        if ($timeSinceLastBeep -ge $delaySeconds) {
            # Skip alerts during quiet hours (11 PM - 7 AM)
            $hour = (Get-Date).Hour
            $isQuietHours = ($hour -ge 23 -or $hour -lt 7)

            if (-not $isQuietHours) {
                # Play WAV file if it exists, otherwise fall back to system beep
                if ($null -ne $soundPlayer) {
                    $soundPlayer.PlaySync() # Play the sound synchronously (wait for completion)
                } else {
                    [System.Media.SystemSounds]::Beep.Play()
                }
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

    # Use timestamp only for temp folder uniqueness during ripping
    $timestamp = Get-Date -Format "yy-MM-dd-HH-mm"
    $SafeVolumeName = $VolumeName -replace '[\\/:*?"<>|]', '_'
    $TempOutputDir = Join-Path $OutputDir "$SafeVolumeName-$($InputDrive.TrimEnd(':'))-$timestamp"

    Write-Host "Disc detected: $VolumeName. Extracting to staging folder..."

    # Execute MakeMKV targeting specific drive
    if (-not (Test-Path $TempOutputDir)) {
        New-Item -ItemType Directory -Path $TempOutputDir -Force | Out-Null
    }
    $driveArg = if ($null -ne $DiscIndex) { "disc:$DiscIndex" } else { "dev:$InputDriveLetter" }
    $ArgumentList = "mkv $driveArg all `"$TempOutputDir`" --minlength=$MinLength"
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
                    Write-Host "[$InputDriveLetter] Detected MKV file: $($trackedMkvFile.Name)"
                }
            }

            # Show progress using the tracked file
            if ($null -ne $trackedMkvFile) {
                # Refresh the file info to get current size
                $trackedMkvFile.Refresh()
                $sizeMB = [math]::Round($trackedMkvFile.Length / 1MB, 2)
                Write-Host "[$InputDriveLetter] Progress: $sizeMB MB written..."
            }
            $progressCounter = 0
        }
    }

    # Wait for process to fully exit and get final exit code
    $process.WaitForExit()
    $exitCode = $process.ExitCode

    # Move file and cleanup
    $shortTitleWarning = $false
    if ($null -eq $exitCode -or $exitCode -eq 0) {
        # Check if MakeMKV produced any output
        $mkvFiles = Get-ChildItem -Path $TempOutputDir -Filter *.mkv -ErrorAction SilentlyContinue
        if ($mkvFiles.Count -eq 0) {
            Write-Host "No titles found over $([math]::Round($MinLength / 60)) minutes. Retrying with no minimum length to get the longest title..."
            $shortTitleWarning = $true
            $ArgumentList = "mkv $driveArg all `"$TempOutputDir`" --minlength=0"
            $process = Start-Process -FilePath $MakeMkvPath -ArgumentList $ArgumentList -NoNewWindow -PassThru
            $process.WaitForExit()
        }

        # Find the largest MKV (the main feature) from the temp directory
        if ($shortTitleWarning -or $null -eq $trackedMkvFile -or -not (Test-Path $trackedMkvFile.FullName)) {
            $foundMkv = Get-ChildItem -Path $TempOutputDir -Filter *.mkv -ErrorAction SilentlyContinue |
                Sort-Object Length -Descending | Select-Object -First 1
            if ($foundMkv) {
                $trackedMkvFile = $foundMkv
                Write-Host "Found MKV file after completion: $($trackedMkvFile.Name)"
            }
        }

        if ($null -ne $trackedMkvFile -and (Test-Path $trackedMkvFile.FullName)) {
            if ($shortTitleWarning) {
                $sizeMB = [math]::Round($trackedMkvFile.Length / 1MB, 0)
                Write-Host "WARNING: No title over $([math]::Round($MinLength / 60)) min found. Kept longest title ($sizeMB MB): $($trackedMkvFile.Name)" -ForegroundColor Yellow
            }
            Write-Host "MakeMKV completed successfully."
            # Determine best name: prefer MKV filename over generic volume labels
            $mkvBaseName = $trackedMkvFile.BaseName -replace '-[A-Z]\d+_t\d+$', ''  # Strip track suffix like -A5_t00
            $genericLabels = @('UNKNOWN_DISC', 'DVD_VIDEO', 'DVDVOLUME', 'DVD')
            if ($mkvBaseName.Length -gt 0 -and $genericLabels -contains $SafeVolumeName) {
                $finalName = $mkvBaseName -replace '[\\/:*?"<>|]', '_'
                Write-Host "Using MKV filename '$mkvBaseName' instead of volume label '$VolumeName'"
            } else {
                $finalName = $SafeVolumeName
            }

            # Add -check-title suffix so encoding-loop knows the title hasn't been confirmed
            $FinalOutputFile = Join-Path $OutputDir "$finalName-check-title.mkv"
            $counter = 2
            while (Test-Path $FinalOutputFile) {
                $FinalOutputFile = Join-Path $OutputDir "$finalName ($counter)-check-title.mkv"
                $counter++
            }

            Move-Item -Path $trackedMkvFile.FullName -Destination $FinalOutputFile -Force
            Write-Host "Extraction complete. Saved as: $FinalOutputFile"
            $lastRippedFile = $FinalOutputFile

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
