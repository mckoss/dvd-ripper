<#
.SYNOPSIS
    Deletes source files that are confirmed backed up in another folder.

.DESCRIPTION
    Compares files in Source against Backup by relative path and file size.
    Lists matching files and prompts before deleting from Source.

.PARAMETER Source
    Folder containing files to check and potentially delete.

.PARAMETER Backup
    Folder to verify files exist in (by matching path and size).

.EXAMPLE
    .\delete-if-backedup.ps1 -Source "G:\Movies\MP4s" -Backup "H:\Backup\MP4s"
#>
param (
    [string]$Source,
    [string]$Backup
)

if ([string]::IsNullOrWhiteSpace($Source)) {
    $Source = Read-Host "Enter source folder (files to check/delete)"
    if ([string]::IsNullOrWhiteSpace($Source)) { Write-Host "Source folder is required."; exit 1 }
}
if ([string]::IsNullOrWhiteSpace($Backup)) {
    $Backup = Read-Host "Enter backup folder (to verify against)"
    if ([string]::IsNullOrWhiteSpace($Backup)) { Write-Host "Backup folder is required."; exit 1 }
}

$safeToDelete = @()

Get-ChildItem -Path $Source -File -Recurse | ForEach-Object {
    $relativePath = $_.FullName.Substring($Source.Length + 1)
    $backupPath = Join-Path $Backup $relativePath

    if ((Test-Path -LiteralPath $backupPath) -and ($_.Length -eq (Get-Item -LiteralPath $backupPath).Length)) {
        Write-Host "Backed up: $relativePath"
        $safeToDelete += $_
    }
}

if ($safeToDelete.Count -gt 0) {
    $response = Read-Host "`nDelete these $($safeToDelete.Count) file(s) from $Source? (y/N)"
    if ($response -eq 'y') {
        $safeToDelete | ForEach-Object {
            $fileName = $_.Name
            try {
                Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
                Write-Host "  Deleted: $fileName"
            } catch {
                Write-Host "  Skipped (locked): $fileName"
            }
        }
    }
} else {
    Write-Host "No files in $Source are backed up to $Backup."
}