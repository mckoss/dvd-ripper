# tmdb-helpers.ps1
# Shared TMDb lookup and movie file utility functions.
# Dot-source this file from other scripts: . "$PSScriptRoot\tmdb-helpers.ps1"
#
# Requires: ffprobe (from FFmpeg) on PATH for metadata extraction

function Get-MkvTitle {
    param([string]$FilePath)

    if (Get-Command 'ffprobe' -ErrorAction SilentlyContinue) {
        try {
            $output = & ffprobe -v quiet -show_format -of json $FilePath 2>$null | ConvertFrom-Json
            $title = $output.format.tags.title
            if (-not [string]::IsNullOrWhiteSpace($title)) { return $title }
        } catch {}
    }

    return $null
}

function Get-FileDurationMinutes {
    param([string]$FilePath)
    if (Get-Command 'ffprobe' -ErrorAction SilentlyContinue) {
        try {
            $output = & ffprobe -v quiet -show_format -of json $FilePath 2>$null | ConvertFrom-Json
            $seconds = [double]$output.format.duration
            return [int][Math]::Round($seconds / 60)
        } catch {}
    }
    return $null
}

function Format-Duration {
    param([int]$Minutes)
    $h = [Math]::Floor($Minutes / 60)
    $m = $Minutes % 60
    return ('{0}h {1:D2}m' -f $h, $m)
}

function Get-TmdbRuntime {
    param([int]$MovieId, [string]$ApiKey)
    $url = 'https://api.themoviedb.org/3/movie/' + $MovieId + '?api_key=' + $ApiKey
    try {
        $response = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop
        if ($response.runtime) { return [int]$response.runtime }
    } catch {}
    return $null
}

function Add-RuntimeToResults {
    param($Results, [string]$ApiKey)
    foreach ($r in $Results) {
        $rt = Get-TmdbRuntime -MovieId $r.id -ApiKey $ApiKey
        $r | Add-Member -NotePropertyName '_runtime' -NotePropertyValue $rt -Force
    }
}

function Format-RawTitle {
    param([string]$RawTitle)
    $t = $RawTitle
    $t = $t -replace '_', ' '
    $t = $t -replace '\s*(D|DISC|DISK)\s*\d+', ''
    $t = $t -replace '\s*T\d{2,}', ''
    $t = $t -replace '\s*\(?\d{4}\)?$', ''
    $t = $t -replace '\s+', ' '
    $t = $t.Trim()
    $t = (Get-Culture).TextInfo.ToTitleCase($t.ToLower())
    return $t
}

function Search-TMDb {
    param([string]$Query, [string]$ApiKey)
    $encoded = [System.Uri]::EscapeDataString($Query)
    $queryLower = $Query.ToLower().Trim()
    $allResults = @()
    for ($page = 1; $page -le 2; $page++) {
        $url = 'https://api.themoviedb.org/3/search/movie?api_key=' + $ApiKey + '&query=' + $encoded + '&page=' + $page
        try {
            $response = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop
            if ($response.results) { $allResults += $response.results }
            if ($page -ge $response.total_pages) { break }
        } catch {
            Write-Host "  TMDb API error: $_" -ForegroundColor Red
            return ,@()
        }
    }
    # Sort: exact title match first, then vote count
    return @($allResults | Sort-Object -Property @{Expression={if ($_.title.ToLower().Trim() -eq $queryLower) { 0 } else { 1 }}; Ascending=$true}, @{Expression='vote_count'; Descending=$true} | Select-Object -First 5)
}

function Format-MovieFilename {
    param([string]$Title, [string]$Year)
    $safe = $Title -replace '[/\\]', '-'
    $safe = $safe -replace '[<>:"|?*]', ''
    $safe = $safe -replace '\s+', ' '
    return ($safe + ' (' + $Year + ')')
}

function Show-TmdbResults {
    param($Results, [int]$FileDurationMin = 0)
    Write-Host '  TMDb results:' -ForegroundColor Cyan
    # Find closest runtime match
    $script:closestIdx = -1
    if ($FileDurationMin -gt 0) {
        $bestDiff = [int]::MaxValue
        for ($i = 0; $i -lt $Results.Count; $i++) {
            if ($Results[$i]._runtime) {
                $diff = [Math]::Abs($Results[$i]._runtime - $FileDurationMin)
                if ($diff -lt $bestDiff) {
                    $bestDiff = $diff
                    $script:closestIdx = $i
                }
            }
        }
        # Only highlight if within 10 minutes
        if ($bestDiff -gt 10) { $script:closestIdx = -1 }
    }
    for ($i = 0; $i -lt $Results.Count; $i++) {
        $r = $Results[$i]
        $year = '????'
        if ($r.release_date -and $r.release_date.Length -ge 4) {
            $year = $r.release_date.Substring(0, 4)
        }
        $num = $i + 1
        $rtStr = ''
        if ($r._runtime) {
            $rtStr = '  ' + (Format-Duration $r._runtime)
        }
        $votes = ''
        if ($r.vote_count) {
            $votes = '  ' + $r.vote_count.ToString('N0') + ' votes'
        }
        $marker = ''
        if ($i -eq $script:closestIdx) { $marker = '  << closest running time' }
        $rest = ' ' + $r.title + ' (' + $year + ')' + $rtStr + $votes + $marker
        if ($i -eq 0) {
            $line = '    [' + $num + ']' + $rest
            Write-Host $line -ForegroundColor Green
        } else {
            $line = '    [' + $num + ']' + $rest
            Write-Host $line -ForegroundColor White
        }
    }
}

function Get-YearFromResult {
    param($Result)
    if ($Result.release_date -and $Result.release_date.Length -ge 4) {
        return $Result.release_date.Substring(0, 4)
    }
    return '????'
}

function Initialize-TmdbApiKey {
    param([string]$TmdbApiKey)
    $ApiKeyFile = Join-Path $PSScriptRoot '.tmdb-api-key'
    if ([string]::IsNullOrWhiteSpace($TmdbApiKey)) {
        if (Test-Path $ApiKeyFile) {
            $TmdbApiKey = (Get-Content $ApiKeyFile -Raw).Trim()
        } else {
            $TmdbApiKey = Read-Host 'Enter TMDb API Key v3 (from https://www.themoviedb.org/settings/api)'
            if ([string]::IsNullOrWhiteSpace($TmdbApiKey)) {
                Write-Host 'TMDb API key is required.' -ForegroundColor Red
                exit 1
            }
            $saveKey = Read-Host "Save API key to ${ApiKeyFile} for future use? (y/n)"
            if ($saveKey -match '^[Yy]') {
                $TmdbApiKey | Set-Content $ApiKeyFile
                Write-Host '  Saved.' -ForegroundColor Green
            }
        }
    }
    return $TmdbApiKey
}

# Searches TMDb for a movie file, shows interactive results, and returns the chosen
# "Title (Year)" base name or $null if skipped. Handles the full search → display →
# prompt → custom search → manual entry flow.
#
# Parameters:
#   -FilePath     Path to the movie file (for metadata + duration extraction)
#   -RawBaseName  Fallback title derived from filename (without extension)
#   -ApiKey       TMDb v3 API key
#   -ExtraOptions (optional) Additional prompt options string (e.g. '[B] Batch 10  [Q] Quit')
#                 and their keys (e.g. 'BQ'). If the user picks one, returned as-is in .Choice
#
# Returns a hashtable:
#   @{ NewBase = 'Title (Year)' or $null; Choice = '' or 'B'/'Q'/etc; Results = @(...) }
function Invoke-TmdbRenamePrompt {
    param(
        [string]$FilePath,
        [string]$RawBaseName,
        [string]$ApiKey,
        [string]$ExtraOptionsDisplay = '',
        [string]$ExtraOptionsKeys = ''
    )

    # Extract metadata title
    $metaTitle = Get-MkvTitle -FilePath $FilePath
    if ($metaTitle) {
        Write-Host "  Metadata title: $metaTitle" -ForegroundColor Gray
    }

    # Get file duration
    $fileDuration = Get-FileDurationMinutes -FilePath $FilePath
    if ($fileDuration) {
        Write-Host "  File duration: $(Format-Duration $fileDuration)" -ForegroundColor Gray
    }

    # Build search query: metadata title first, filename fallback
    $searchQuery = if ($metaTitle) { Format-RawTitle $metaTitle } else { Format-RawTitle $RawBaseName }
    $filenameQuery = Format-RawTitle $RawBaseName
    Write-Host "  Search query: $searchQuery" -ForegroundColor Gray

    $results = @(Search-TMDb -Query $searchQuery -ApiKey $ApiKey)

    if ($results.Count -eq 0 -and $metaTitle -and $filenameQuery -ne $searchQuery) {
        Write-Host '  No results from metadata. Trying filename...' -ForegroundColor Yellow
        Write-Host "  Search query: $filenameQuery" -ForegroundColor Gray
        $results = @(Search-TMDb -Query $filenameQuery -ApiKey $ApiKey)
    }

    if ($results.Count -eq 0) {
        Write-Host '  No TMDb results found. Check for typos in the filename.' -ForegroundColor Yellow
        $customSearch = Read-Host '  Enter corrected search (or press Enter to skip)'
        if ([string]::IsNullOrWhiteSpace($customSearch)) {
            return @{ NewBase = $null; Choice = 'S'; Results = @() }
        }
        $results = @(Search-TMDb -Query $customSearch -ApiKey $ApiKey)
        if ($results.Count -eq 0) {
            Write-Host '  Still no results. Skipping.' -ForegroundColor Yellow
            return @{ NewBase = $null; Choice = 'S'; Results = @() }
        }
    }

    Add-RuntimeToResults -Results $results -ApiKey $ApiKey
    Show-TmdbResults -Results $results -FileDurationMin $fileDuration

    # Build prompt line
    $extraDisp = if ($ExtraOptionsDisplay) { '  ' + $ExtraOptionsDisplay } else { '' }
    Write-Host ('    [S] Skip  [C] Custom search  [M] Manual entry' + $extraDisp) -ForegroundColor DarkGray
    $extraKeys = if ($ExtraOptionsKeys) { '/' + (($ExtraOptionsKeys.ToCharArray() | ForEach-Object { $_.ToString().ToUpper() }) -join '/') } else { '' }
    Write-Host -NoNewline ('  Choose (1-' + $results.Count + '/S/C/M' + $extraKeys + ') [')
    Write-Host -NoNewline '1' -ForegroundColor Green
    Write-Host -NoNewline ']'
    $choice = Read-Host ' '
    if ([string]::IsNullOrWhiteSpace($choice)) { $choice = '1' }

    # Check for extra option keys (B, Q, etc.)
    if ($ExtraOptionsKeys -and $choice -match ('^[' + $ExtraOptionsKeys + ']$')) {
        return @{ NewBase = $null; Choice = $choice.ToUpper(); Results = $results; FileDuration = $fileDuration }
    }

    if ($choice -match '^[Ss]$') {
        return @{ NewBase = $null; Choice = 'S'; Results = $results; FileDuration = $fileDuration }
    }
    elseif ($choice -match '^[Cc]$') {
        $customSearch = Read-Host '  Enter search query'
        $results = @(Search-TMDb -Query $customSearch -ApiKey $ApiKey)
        if ($results.Count -eq 0) {
            Write-Host '  No results. Skipping.' -ForegroundColor Yellow
            return @{ NewBase = $null; Choice = 'S'; Results = @(); FileDuration = $fileDuration }
        }
        Add-RuntimeToResults -Results $results -ApiKey $ApiKey
        Show-TmdbResults -Results $results -FileDurationMin $fileDuration
        Write-Host -NoNewline ('  Choose (1-' + $results.Count + '/S) [')
        Write-Host -NoNewline '1' -ForegroundColor Green
        Write-Host -NoNewline ']'
        $choice = Read-Host ' '
        if ([string]::IsNullOrWhiteSpace($choice)) { $choice = '1' }
        if ($choice -match '^[Ss]$' -or -not ($choice -match '^\d+$')) {
            return @{ NewBase = $null; Choice = 'S'; Results = $results; FileDuration = $fileDuration }
        }
    }
    elseif ($choice -match '^[Mm]$') {
        $manualTitle = Read-Host '  Enter title'
        $manualYear = Read-Host '  Enter year'
        $nb = Format-MovieFilename -Title $manualTitle -Year $manualYear
        return @{ NewBase = $nb; Choice = 'M'; Results = $results; FileDuration = $fileDuration }
    }

    if ($choice -match '^\d+$') {
        $idx = [int]$choice - 1
        if ($idx -ge 0 -and $idx -lt $results.Count) {
            $selected = $results[$idx]
            $year = Get-YearFromResult -Result $selected
            $nb = Format-MovieFilename -Title $selected.title -Year $year
            return @{ NewBase = $nb; Choice = $choice; Results = $results; FileDuration = $fileDuration }
        } else {
            Write-Host '  Invalid choice. Skipping.' -ForegroundColor Yellow
        }
    }

    return @{ NewBase = $null; Choice = 'S'; Results = $results; FileDuration = $fileDuration }
}
