# Rename-Movies.ps1
# Renames local MP4 files to match Google Drive filenames (with release year).
# For each MP4 rename recognized, also renames the matching MKV if it exists.
# Fetches the authoritative filename list live from Drive using rclone.

param(
    [string]$MP4Folder    = "G:\Movies\MP4s",
    [string]$MKVFolder    = "G:\Movies\MKVs",
    [string]$RcloneRemote = "gdrive:Movies/MP4s"
)

# ── Helpers ───────────────────────────────────────────────────────────────────

function Strip-Year($name) {
    $name = [System.IO.Path]::GetFileNameWithoutExtension($name)
    $name = $name -replace '\s*\(\d{4}\)$', ''
    return $name.Trim()
}

function Normalize($name) {
    $n = $name.ToLower()
    $n = $n -replace "^(the |a |an )", ""
    $n = $n -replace "[^a-z0-9 ]", ""
    $n = $n -replace "\s+", " "
    return $n.Trim()
}

# ── Step 1: Fetch Drive file list via rclone ──────────────────────────────────

Write-Host "`n📡 Fetching file list from Drive ($RcloneRemote)..." -ForegroundColor Cyan

$rcloneOutput = rclone lsjson $RcloneRemote --files-only 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ rclone failed. Make sure rclone is configured and '$RcloneRemote' is correct." -ForegroundColor Red
    Write-Host $rcloneOutput
    exit 1
}

$driveFiles = $rcloneOutput | ConvertFrom-Json | Where-Object { $_.Name -match '\.mp4$' } | Select-Object -ExpandProperty Name
Write-Host "   Found $($driveFiles.Count) MP4 files on Drive." -ForegroundColor Green

# Build lookup: normalized-base-name → full Drive filename
$driveLookup = @{}
foreach ($f in $driveFiles) {
    $normKey = Normalize (Strip-Year $f)
    $driveLookup[$normKey] = $f
}

# ── Step 2: Scan local MP4 folder, check for companion MKVs ──────────────────

Write-Host "`n🗂  Scanning $MP4Folder..." -ForegroundColor Cyan

if (-not (Test-Path $MP4Folder)) {
    Write-Host "❌ MP4 folder not found: $MP4Folder" -ForegroundColor Red
    exit 1
}

$mp4Renames   = [System.Collections.Generic.List[hashtable]]::new()
$mkvRenames   = [System.Collections.Generic.List[hashtable]]::new()
$alreadyOk    = [System.Collections.Generic.List[string]]::new()
$unmatched    = [System.Collections.Generic.List[string]]::new()

$mp4Files = Get-ChildItem -Path $MP4Folder -Filter "*.mp4"
Write-Host "   Found $($mp4Files.Count) MP4 files locally." -ForegroundColor Green

foreach ($file in $mp4Files) {
    $normKey = Normalize (Strip-Year $file.Name)

    if ($driveLookup.ContainsKey($normKey)) {
        $driveFilename = $driveLookup[$normKey]
        $targetMp4     = $driveFilename  # e.g. "The Big Lebowski (1998).mp4"
        $targetMkv     = [System.IO.Path]::ChangeExtension($driveFilename, "mkv")

        if ($file.Name -eq $targetMp4) {
            $alreadyOk.Add($file.Name)
        } else {
            $mp4Renames.Add(@{ File = $file; Target = $targetMp4 })
        }

        # Check for a companion MKV (using the current MP4 base name)
        $currentMkvName = [System.IO.Path]::ChangeExtension($file.Name, "mkv")
        $currentMkvPath = Join-Path $MKVFolder $currentMkvName
        if ((Test-Path $currentMkvPath) -and ($currentMkvName -ne $targetMkv)) {
            $mkvRenames.Add(@{ File = Get-Item $currentMkvPath; Target = $targetMkv })
        }
    } else {
        $unmatched.Add($file.Name)
    }
}

# ── Step 3: Dry run output ────────────────────────────────────────────────────

Write-Host "`n─────────────────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "  DRY RUN — no files changed yet" -ForegroundColor Yellow
Write-Host "─────────────────────────────────────────────────────────" -ForegroundColor DarkGray

if ($alreadyOk.Count -gt 0) {
    Write-Host "`n✅ Already correct ($($alreadyOk.Count)):" -ForegroundColor Green
    foreach ($n in $alreadyOk) { Write-Host "   $n" -ForegroundColor DarkGreen }
}

if ($mp4Renames.Count -gt 0) {
    Write-Host "`n✏️  MP4s to rename ($($mp4Renames.Count)):" -ForegroundColor White
    foreach ($r in $mp4Renames) {
        Write-Host "   $($r.File.Name)" -ForegroundColor Gray
        Write-Host "   → $($r.Target)" -ForegroundColor White
    }
}

if ($mkvRenames.Count -gt 0) {
    Write-Host "`n✏️  Companion MKVs to rename ($($mkvRenames.Count)):" -ForegroundColor White
    foreach ($r in $mkvRenames) {
        Write-Host "   $($r.File.Name)" -ForegroundColor Gray
        Write-Host "   → $($r.Target)" -ForegroundColor White
    }
}

if ($unmatched.Count -gt 0) {
    Write-Host "`n⚠️  No Drive match ($($unmatched.Count)):" -ForegroundColor Yellow
    foreach ($n in $unmatched) { Write-Host "   $n" -ForegroundColor DarkYellow }
}

$totalRenames = $mp4Renames.Count + $mkvRenames.Count
Write-Host "`n─────────────────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "  Summary: $totalRenames to rename ($($mp4Renames.Count) MP4, $($mkvRenames.Count) MKV) | $($alreadyOk.Count) already correct | $($unmatched.Count) unmatched" -ForegroundColor White
Write-Host "─────────────────────────────────────────────────────────`n" -ForegroundColor DarkGray

# ── Step 4: Prompt ────────────────────────────────────────────────────────────

if ($totalRenames -eq 0) {
    Write-Host "Nothing to rename. You're all set! 🎬" -ForegroundColor Green
    exit 0
}

$confirm = Read-Host "Apply $totalRenames rename(s)? (y/N)"
if ($confirm -notmatch '^[Yy]') {
    Write-Host "Aborted. No files were changed." -ForegroundColor Yellow
    exit 0
}

# ── Step 5: Apply renames ─────────────────────────────────────────────────────

Write-Host "`n🚀 Renaming files...`n" -ForegroundColor Cyan
$success = 0
$failed  = 0

foreach ($r in ($mp4Renames + $mkvRenames)) {
    Write-Host "   [$([datetime]::Now.ToString('HH:mm:ss'))] Renaming:" -ForegroundColor DarkGray
    Write-Host "     FROM: $($r.File.FullName)" -ForegroundColor Gray
    Write-Host "     TO:   $($r.Target)" -ForegroundColor White
    try {
        Rename-Item -Path $r.File.FullName -NewName $r.Target -ErrorAction Stop
        Write-Host "     ✅ Done" -ForegroundColor Green
        $success++
    } catch {
        Write-Host "     ❌ Failed: $_" -ForegroundColor Red
        $failed++
    }
    Write-Host ""
}

Write-Host "─────────────────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "  Complete: $success renamed, $failed failed" -ForegroundColor $(if ($failed -gt 0) { "Yellow" } else { "Green" })
Write-Host "─────────────────────────────────────────────────────────`n" -ForegroundColor DarkGray
