# DVD Rip-to-Cloud Pipeline

A set of PowerShell scripts that automate the full DVD archival workflow:
ripping discs to MKV, encoding to MP4 with HandBrake, and uploading to Google Drive.

## Workflow

### 1. Rip DVDs to MKV (`mkv-loop.ps1`)

Continuously monitors an optical drive, rips inserted discs using
[MakeMKV](https://www.makemkv.com/), ejects, and waits for the next disc.
Run multiple instances in separate terminals for parallel ripping from
different drives.

```powershell
.\mkv-loop.ps1                                        # prompted for drive and output dir
.\mkv-loop.ps1 -InputDrive D -OutputDir "G:\Movies\MKVs"  # non-interactive
```

Ripped files land in `MKVs/` with a timestamp suffix (e.g., `MOVIE-26-02-22-17-54.mkv`).
Rename them to the correct movie title before proceeding.

### 2. Prepare for HandBrake encoding (`check-mp4.ps1`)

Compares `MKVs/` against `MP4s/` to find MKVs that haven't been encoded yet.
Flags any files still carrying timestamp names that need renaming first.
Optionally moves the un-encoded files into `MKVs/TBD/` and launches HandBrake
pointed at that folder so you can queue them all at once.

```powershell
.\check-mp4.ps1                              # prompted for movies directory
.\check-mp4.ps1 -OutputDir "G:\Movies"       # non-interactive
```

In HandBrake, use the **DVD Ripped Archive** preset (included in
`Handbrake DVD Preset.json`), set the output to `MP4s/`, and start the queue.

### 3. Upload MP4s to Google Drive (`update-drive.ps1`)

Copies finalized MP4 files to Google Drive via rclone. Automatically skips
files that are still locked by HandBrake (actively encoding).

```powershell
.\update-drive.ps1
```

## Directory Structure

```
G:\Movies\                        # Root movies directory (configurable)
  MKVs\                           # Raw rips from MakeMKV
    Movie Title.mkv               # Renamed, ready for encoding
    UNKNOWN_DISC-26-02-22-17-54.mkv  # Needs renaming
    TBD\                          # Staged for HandBrake (created by check-mp4.ps1)
  MP4s\                           # Encoded output from HandBrake
    Movie Title.mp4               # Ready for upload
```

Files flow through the pipeline: **DVD --> MKVs/ --> (rename) --> TBD/ --> HandBrake --> MP4s/ --> Google Drive**

## Features

- **Continuous loop** -- Insert a disc, walk away, and come back to a folder full of MKV files.
- **Multi-drive support** -- Run multiple instances of the script on different drives to rip
  in parallel.
- **Exponential backoff alerts** -- Plays an audio alert (WAV file) when waiting for a disc,
  with increasing intervals (2s -> 4s -> 8s -> ... up to 5 minutes).
- **Smart file naming** -- Uses the disc volume label with a timestamp for unique filenames.
  Falls back to the MKV filename from MakeMKV if the volume label is generic
  (e.g., `DVD_VIDEO`).
- **Progress monitoring** -- Reports file size every 60 seconds while ripping.
- **Transient disc filtering** -- Detects and ignores brief drive accessibility during
  disc ejection/insertion to avoid false starts.
- **Safe cleanup** -- Temporary folders are only deleted if empty after the main file is moved.
- **Lock-aware uploads** -- Skips files still being encoded by HandBrake.

## Prerequisites

- [MakeMKV](https://www.makemkv.com/) installed
  (default path: `C:\Program Files (x86)\MakeMKV\makemkvcon64.exe`)
- [HandBrake](https://handbrake.fr/) installed
  (default path: `C:\Program Files\HandBrake\HandBrake.exe`)
- [rclone](https://rclone.org/) installed and configured with a `gdrive` remote
- Windows PowerShell 5.1 or later
- One or more optical disc drives

## mkv-loop.ps1 Parameters

| Parameter     | Type   | Default                        | Description                              |
|---------------|--------|--------------------------------|------------------------------------------|
| `-InputDrive` | String | *(prompted interactively)*     | The drive letter to monitor. Accepts with or without a colon (e.g., `F` or `F:`). If omitted, the script prompts for it with a default of `D`. |
| `-OutputDir`  | String | *(prompted interactively)*     | Directory where ripped MKV files are saved. If omitted, the script prompts for it with a default of `G:\Movies\MKVs`. |

## check-mp4.ps1 Parameters

| Parameter    | Type   | Default                        | Description                              |
|--------------|--------|--------------------------------|------------------------------------------|
| `-OutputDir` | String | *(prompted interactively)*     | Root movies directory containing `MKVs/` and `MP4s/` subfolders. Defaults to `G:\Movies`. |

## Configuration

The following variables can be modified at the top of `mkv-loop.ps1`:

| Variable          | Default                                           | Description                                      |
|-------------------|---------------------------------------------------|--------------------------------------------------|
| `$MakeMkvPath`    | `C:\Program Files (x86)\MakeMKV\makemkvcon64.exe` | Path to the MakeMKV command-line tool.           |
| `$MinLength`      | `3600` (seconds)                                  | Minimum title length to extract (skips menus, extras, etc.). |
