param (
    [string]$InputDrive,
    [string]$MoviesDir,
    [int]$DiscIndex = -1,
    [int]$MinLength = 3600
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
$Host.UI.RawUI.WindowTitle = "Ripping - Drive $InputDriveLetter"

if ([string]::IsNullOrWhiteSpace($MoviesDir)) {
    $DefaultMoviesDir = "G:\Movies"
    $MoviesDir = Read-Host "Enter movies directory (default: $DefaultMoviesDir)"
    if ([string]::IsNullOrWhiteSpace($MoviesDir)) {
        $MoviesDir = $DefaultMoviesDir
    }
}
$OutputDir = Join-Path $MoviesDir "processing\ripped-for-encoding"
$AlertSoundPath = Join-Path $PSScriptRoot "alert.wav"

# Build MakeMKV drive argument: disc:N if provided, otherwise dev:X:
if ($DiscIndex -ge 0) {
    $driveArg = "disc:$DiscIndex"
    Write-Host "Using $driveArg for drive $InputDriveLetter" -ForegroundColor Green
} else {
    $driveArg = "dev:$InputDriveLetter"
    Write-Host "Using $driveArg (pass -DiscIndex N to use disc:N instead)" -ForegroundColor DarkGray
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
$lastMetaTitle = $null

while ($true) {
    if ($lastRippedFile) {
        $lastIsFolder = Test-Path $lastRippedFile -PathType Container
        $lastDisplayName = (Split-Path $lastRippedFile -Leaf) -replace '-[A-Z]-check-title$', ''
        if ($lastMetaTitle) {
            Write-Host "`n[$InputDriveLetter] Metadata title: $lastMetaTitle" -ForegroundColor Cyan
        } else {
            Write-Host "`n[$InputDriveLetter] Metadata title: (none)" -ForegroundColor DarkGray
        }
        $typeLabel = if ($lastIsFolder) { 'Collection folder' } else { 'Saved as' }
        Write-Host "[$InputDriveLetter] $($typeLabel): $lastDisplayName" -ForegroundColor White
        Write-Host "Waiting for disc in drive $InputDriveLetter... (press SPACE to rename: $lastDisplayName)" -ForegroundColor Cyan
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

        # Check for SPACE to rename the last ripped file or folder
        if ($lastRippedFile -and [Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            if ($key.Key -eq 'Spacebar') {
                Write-Host ''
                $isFolder = Test-Path $lastRippedFile -PathType Container
                $rawBase = (Split-Path $lastRippedFile -Leaf) -replace '-[A-Z]-check-title(\.mkv)?$', ''
                Write-Host '-------------------------------------------------------------' -ForegroundColor DarkGray
                if ($lastMetaTitle) {
                    Write-Host "  Metadata title: $lastMetaTitle" -ForegroundColor Cyan
                } else {
                    Write-Host "  Metadata title: (none)" -ForegroundColor DarkGray
                }

                if ($isFolder) {
                    $folderMkvs = @(Get-ChildItem -Path $lastRippedFile -Filter *.mkv)
                    Write-Host "  Folder: $(Split-Path $lastRippedFile -Leaf)  ($($folderMkvs.Count) titles)" -ForegroundColor Magenta
                    Write-Host "       $($lastRippedFile)" -ForegroundColor DarkGray

                    $tvResult = Invoke-TmdbTvRenamePrompt -FolderPath $lastRippedFile -RawBaseName $rawBase -ApiKey $TmdbApiKey

                    if ($tvResult.FolderName) {
                        $newFolderPath = Join-Path $OutputDir $tvResult.FolderName
                        if (Test-Path $newFolderPath) {
                            Write-Host "  SKIPPED (target exists): $($tvResult.FolderName)" -ForegroundColor Yellow
                        } else {
                            Rename-Item -Path $lastRippedFile -NewName $tvResult.FolderName -ErrorAction Stop
                            Write-Host "  Renamed folder: $(Split-Path $lastRippedFile -Leaf) -> $($tvResult.FolderName)" -ForegroundColor Green
                        }
                    } else {
                        Write-Host '  Skipped. Folder can be renamed later.' -ForegroundColor Yellow
                    }
                } else {
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
    if ($null -eq $exitCode -or $exitCode -eq 0) {
        # Check if MakeMKV produced any output
        $mkvFiles = @(Get-ChildItem -Path $TempOutputDir -Filter *.mkv -ErrorAction SilentlyContinue)
        if ($mkvFiles.Count -eq 0) {
            Write-Host "[$InputDriveLetter] No titles found over $([math]::Round($MinLength / 60)) minutes. Ejecting disc - reinsert to retry." -ForegroundColor Red
            # Clean up empty temp directory
            if ((Test-Path $TempOutputDir) -and (Get-ChildItem -Path $TempOutputDir -ErrorAction SilentlyContinue).Count -eq 0) {
                Remove-Item -Path $TempOutputDir -Force
            }
            # Eject and continue to wait loop
            Write-Host "Ejecting drive $InputDriveLetter..."
            $shell = New-Object -ComObject Shell.Application
            $shellDrive = $shell.Namespace(17).ParseName($InputDriveLetter)
            if ($shellDrive) { $shellDrive.InvokeVerb("Eject") }
            Start-Sleep -Seconds 8
            continue
        }

        Write-Host "MakeMKV completed successfully. $($mkvFiles.Count) title(s) ripped."

        if ($mkvFiles.Count -gt 1) {
            # --- COLLECTION MODE: multiple titles (TV episodes, cartoon collections, etc.) ---
            # Keep all MKVs in the folder, rename the folder instead
            $totalSizeMB = [math]::Round(($mkvFiles | Measure-Object Length -Sum).Sum / 1MB, 0)
            Write-Host "[$InputDriveLetter] Collection disc: $($mkvFiles.Count) titles ($totalSizeMB MB total)" -ForegroundColor Magenta
            $mkvFiles | Sort-Object Name | ForEach-Object {
                $sizeMB = [math]::Round($_.Length / 1MB, 0)
                Write-Host "  $($_.Name)  ($sizeMB MB)" -ForegroundColor Gray
            }

            # Determine best folder name: metadata title from largest file > volume label
            $largestMkv = $mkvFiles | Sort-Object Length -Descending | Select-Object -First 1
            $metaTitle = Get-MkvTitle -FilePath $largestMkv.FullName
            if ($metaTitle) {
                Write-Host "[$InputDriveLetter] Metadata title: $metaTitle" -ForegroundColor Cyan
            } else {
                Write-Host "[$InputDriveLetter] Metadata title: (none)" -ForegroundColor DarkGray
            }

            $genericLabels = @('UNKNOWN_DISC', 'DVD_VIDEO', 'DVDVOLUME', 'DVD')
            $isGenericVolume = $genericLabels -contains $SafeVolumeName
            if ($metaTitle) {
                $folderName = $metaTitle -replace '[\\/:*?"<>|]', '_'
            } elseif (-not $isGenericVolume) {
                $folderName = $SafeVolumeName
            } else {
                $folderName = "UNKNOWN_COLLECTION"
                Write-Host "WARNING: No metadata title and generic volume label." -ForegroundColor Yellow
            }

            # Rename folder with drive letter + check-title suffix
            $driveSuffix = $InputDrive.TrimEnd(':')
            $newFolderName = "$folderName-$driveSuffix-check-title"
            $newFolderPath = Join-Path $OutputDir $newFolderName
            $counter = 2
            while (Test-Path $newFolderPath) {
                $newFolderPath = Join-Path $OutputDir "$folderName ($counter)-$driveSuffix-check-title"
                $counter++
            }
            Rename-Item -Path $TempOutputDir -NewName (Split-Path $newFolderPath -Leaf)
            Write-Host "[$InputDriveLetter] Collection saved as folder: $(Split-Path $newFolderPath -Leaf)" -ForegroundColor Green

            $lastRippedFile = $newFolderPath
            $lastMetaTitle = $metaTitle
        } else {
            # --- SINGLE MOVIE MODE: one title ---
            # Find the MKV (may have been tracked during ripping, or find it now)
            if ($null -eq $trackedMkvFile -or -not (Test-Path $trackedMkvFile.FullName)) {
                $trackedMkvFile = $mkvFiles[0]
                Write-Host "Found MKV file after completion: $($trackedMkvFile.Name)"
            }

            # Determine best name: metadata title > MKV filename > volume label
            $metaTitle = Get-MkvTitle -FilePath $trackedMkvFile.FullName
            $mkvBaseName = $trackedMkvFile.BaseName -replace '-[A-Z]\d+_t\d+$', ''  # Strip track suffix like -A5_t00
            $genericLabels = @('UNKNOWN_DISC', 'DVD_VIDEO', 'DVDVOLUME', 'DVD')
            $isGenericFilename = $mkvBaseName -match '^[A-Z]\d+_t\d+$' -or $mkvBaseName -match '^title_t\d+$'
            $isGenericVolume = $genericLabels -contains $SafeVolumeName

            # Always show metadata title separately
            if ($metaTitle) {
                Write-Host "[$InputDriveLetter] Metadata title: $metaTitle" -ForegroundColor Cyan
            } else {
                Write-Host "[$InputDriveLetter] Metadata title: (none)" -ForegroundColor DarkGray
            }

            if ($metaTitle) {
                $finalName = $metaTitle -replace '[\\/:*?"<>|]', '_'
            } elseif (-not $isGenericFilename -and $isGenericVolume) {
                $finalName = $mkvBaseName -replace '[\\/:*?"<>|]', '_'
                Write-Host "Using MKV filename '$mkvBaseName' instead of volume label '$VolumeName'"
            } else {
                $finalName = $SafeVolumeName
                if ($isGenericVolume) {
                    Write-Host "WARNING: No metadata title, generic volume label and filename." -ForegroundColor Yellow
                }
            }

            # Add drive letter + -check-title suffix so encoding-loop knows the title hasn't been confirmed
            # Drive letter makes filenames unique per drive, preventing collisions
            $driveSuffix = $InputDrive.TrimEnd(':')
            $FinalOutputFile = Join-Path $OutputDir "$finalName-$driveSuffix-check-title.mkv"
            $counter = 2
            $moved = $false
            while (-not $moved) {
                while (Test-Path $FinalOutputFile) {
                    $FinalOutputFile = Join-Path $OutputDir "$finalName ($counter)-$driveSuffix-check-title.mkv"
                    $counter++
                }
                try {
                    Move-Item -Path $trackedMkvFile.FullName -Destination $FinalOutputFile -ErrorAction Stop
                    $moved = $true
                } catch {
                    # Another drive claimed this name between Test-Path and Move-Item - try next
                    $FinalOutputFile = Join-Path $OutputDir "$finalName ($counter)-$driveSuffix-check-title.mkv"
                    $counter++
                }
            }
            Write-Host "Extraction complete. Saved as: $FinalOutputFile"
            $lastRippedFile = $FinalOutputFile
            $lastMetaTitle = $metaTitle

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

    # Wait for drive to settle after eject to prevent false re-detection
    Start-Sleep -Seconds 8
}
