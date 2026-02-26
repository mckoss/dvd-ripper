# fix-faststart.ps1
# Checks MP4 files for moov atom placement and applies faststart if needed.

param(
    [string]$MoviesDir
)

if (-not $MoviesDir) {
    $defaultDir = "G:\Movies"
    $MoviesDir = Read-Host "Enter movies directory (default: $defaultDir)"
    if ([string]::IsNullOrWhiteSpace($MoviesDir)) { $MoviesDir = $defaultDir }
}

if (-not (Test-Path $MoviesDir)) {
    Write-Host "Directory not found: $MoviesDir" -ForegroundColor Red
    exit 1
}

$Mp4ArchiveDir = Join-Path $MoviesDir "MP4s"
$EncodedDir    = Join-Path $MoviesDir "processing\encoded-for-upload"
$TempDir       = Join-Path $MoviesDir "processing\temp"

# Verify ffmpeg is available
if (-not (Get-Command "ffmpeg" -ErrorAction SilentlyContinue)) {
    Write-Host "ffmpeg not found in PATH. Please install FFmpeg." -ForegroundColor Red
    exit 1
}

# Returns 'moov' or 'mdat' depending on which top-level atom appears first, or $null
function Get-FirstMediaAtom {
    param([string]$FilePath)
    $stream = [System.IO.File]::OpenRead($FilePath)
    try {
        $header = New-Object byte[] 8
        while ($stream.Position -lt $stream.Length) {
            $bytesRead = $stream.Read($header, 0, 8)
            if ($bytesRead -lt 8) { break }

            # Box size: big-endian uint32
            $size = ([uint64]$header[0] -shl 24) -bor ([uint64]$header[1] -shl 16) -bor ([uint64]$header[2] -shl 8) -bor [uint64]$header[3]
            $type = [System.Text.Encoding]::ASCII.GetString($header, 4, 4)

            if ($type -eq 'moov' -or $type -eq 'mdat') {
                return $type
            }

            # Handle extended size (size == 1 means 64-bit size follows)
            if ($size -eq 1) {
                $ext = New-Object byte[] 8
                if ($stream.Read($ext, 0, 8) -lt 8) { break }
                $size = [uint64]0
                for ($i = 0; $i -lt 8; $i++) { $size = ($size -shl 8) -bor [uint64]$ext[$i] }
                $stream.Position += ($size - 16)
            } elseif ($size -eq 0) {
                break  # box extends to end of file
            } else {
                $stream.Position += ($size - 8)
            }
        }
    } finally {
        $stream.Close()
    }
    return $null
}

$mp4Files = @()
$scanDirs = @($Mp4ArchiveDir, $EncodedDir) | Where-Object { Test-Path $_ }
foreach ($dir in $scanDirs) {
    $mp4Files += Get-ChildItem -Path $dir -Filter *.mp4 -File
}
if ($mp4Files.Count -eq 0) {
    Write-Host "No MP4 files found in scanned directories."
    exit 0
}

Write-Host "Scanning $($mp4Files.Count) MP4 file(s) across $($scanDirs.Count) folder(s)..."
foreach ($dir in $scanDirs) { Write-Host "  $dir" -ForegroundColor Gray }
Write-Host ""

$alreadyOk = @()
$needsFix = @()
$failed = @()
$fixed = @()

# Phase 1: Scan all files to determine moov atom placement
foreach ($file in $mp4Files) {
    $relDir = $file.DirectoryName.Replace($MoviesDir, '').TrimStart('\')
    Write-Host "Checking: [$relDir] $($file.Name)" -NoNewline

    $firstAtom = Get-FirstMediaAtom -FilePath $file.FullName

    if ($null -eq $firstAtom) {
        Write-Host " - Could not determine atom order, skipping" -ForegroundColor Yellow
        $failed += [PSCustomObject]@{ Name = "[$relDir] $($file.Name)"; Reason = "Could not find moov/mdat atoms" }
        continue
    }

    if ($firstAtom -eq 'moov') {
        Write-Host " - OK" -ForegroundColor Green
        $alreadyOk += [PSCustomObject]@{ Name = $file.Name; RelDir = $relDir }
        continue
    }

    Write-Host " - moov after mdat, needs fixing" -ForegroundColor Cyan
    $needsFix += $file
}

# Phase 2: Show summary and prompt before processing
Write-Host ""
if ($needsFix.Count -eq 0) {
    Write-Host "All files already have moov at head. Nothing to do." -ForegroundColor Green
} else {
    Write-Host "Files needing faststart fix ($($needsFix.Count)):" -ForegroundColor Cyan
    $needsFix | ForEach-Object {
        $relDir = $_.DirectoryName.Replace($MoviesDir, '').TrimStart('\')
        Write-Host "  [$relDir] $($_.Name) ($([math]::Round($_.Length / 1MB, 1)) MB)"
    }
    Write-Host ""
    $confirm = Read-Host "Proceed with fixing $($needsFix.Count) file(s)? (y/n)"
    if ($confirm -ne 'y') {
        Write-Host "Aborted." -ForegroundColor Yellow
        exit 0
    }

    Write-Host ""
    foreach ($file in $needsFix) {
        $relDir = $file.DirectoryName.Replace($MoviesDir, '').TrimStart('\')
        Write-Host "Processing: [$relDir] $($file.Name)..." -ForegroundColor Cyan
        if (-not (Test-Path $TempDir)) { New-Item -Path $TempDir -ItemType Directory -Force | Out-Null }
        $tempFile = Join-Path $TempDir $file.Name

        try {
            $ffmpegOutput = & ffmpeg -i $file.FullName -c copy -movflags +faststart $tempFile -y 2>&1 | Out-String
            if ($LASTEXITCODE -ne 0) {
                throw "ffmpeg exited with code $LASTEXITCODE`n$ffmpegOutput"
            }

            # Verify the temp file exists and is reasonable size (at least 50% of original)
            if (-not (Test-Path $tempFile)) {
                throw "Temp file was not created"
            }
            $origSize = $file.Length
            $tempSize = (Get-Item $tempFile).Length
            if ($tempSize -lt ($origSize * 0.5)) {
                throw "Output file is suspiciously small ($tempSize bytes vs $origSize bytes original)"
            }

            # Replace original with fixed file
            Remove-Item -Path $file.FullName -Force
            Move-Item -Path $tempFile -Destination $file.FullName -Force
            Write-Host "  Fixed: [$relDir] $($file.Name) ($([math]::Round($origSize / 1MB, 1)) MB)" -ForegroundColor Green
            $fixed += [PSCustomObject]@{ Name = $file.Name; RelDir = $relDir }
        }
        catch {
            Write-Host "  FAILED: $($file.Name) - $_" -ForegroundColor Red
            $failed += [PSCustomObject]@{ Name = $file.Name; RelDir = $relDir; Reason = $_.ToString() }
            # Leave temp file in processing\temp for inspection
        }
    }
}

# Summary
Write-Host "`n===== Summary ====="
Write-Host "Total scanned:   $($mp4Files.Count)"
Write-Host "Already OK:      $($alreadyOk.Count)" -ForegroundColor Green
Write-Host "Fixed:           $($fixed.Count)" -ForegroundColor Cyan
if ($failed.Count -gt 0) {
    Write-Host "Failed:          $($failed.Count)" -ForegroundColor Red
    $failed | ForEach-Object { Write-Host "  [$($_.RelDir)] $($_.Name): $($_.Reason)" -ForegroundColor Red }
}
Write-Host "==================`n"
