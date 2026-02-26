# DVD Rip-to-Cloud Pipeline

A set of PowerShell scripts that automate the full DVD archival workflow:
ripping discs to MKV, encoding to MP4 with HandBrake, and uploading to Google Drive.

Files flow through a folder-based pipeline where each file's location represents its
processing state. Each script can be interrupted and restarted independently — it
simply rescans its input folder to pick up where it left off.

## Quick Start

1. **Start ripping** — Open a PowerShell terminal for each optical drive and run
   `.\ripping-loop.ps1`. Each instance monitors one drive, rips the disc, ejects it,
   and plays an alert sound so you know to insert the next one. Just keep feeding discs.

2. **Confirm titles** — DVD discs don't carry reliable movie titles, so ripped files
   land with a `-check-title` suffix (e.g., `GRAVITY-check-title.mkv`). Browse to
   `processing\ripped-for-encoding\` and rename each file — correct the title if needed
   and remove the `-check-title` suffix (e.g., `Gravity.mkv`). This is the only manual
   step in the pipeline.

3. **Encode** — Run `.\encode-backlog.ps1` whenever you're ready. It shows which files
   need encoding, archives already-encoded MKVs, and launches HandBrake pointed at
   the ripped folder. In HandBrake, set the output to `processing\encoded-for-upload\`
   and start the queue.

4. **Upload** — Run `.\upload-mp4s.ps1` to push encoded MP4s to Google Drive via rclone.
   After uploading, it offers to move the files to the `MP4s\` archive folder.

Steps 1–2 can happen continuously while steps 3–4 are run on demand at any time.
Encoding and uploading are safe to run even while ripping is still in progress.

## Workflow Details

### 1. Rip DVDs to MKV (`ripping-loop.ps1`)

Continuously monitors an optical drive, rips inserted discs using
[MakeMKV](https://www.makemkv.com/), ejects, and waits for the next disc.
Run multiple instances in separate terminals for parallel ripping from
different drives.

```powershell
.\ripping-loop.ps1                                            # prompted for drive and movies dir
.\ripping-loop.ps1 -InputDrive D -MoviesDir "G:\Movies"        # non-interactive
```

Ripped files land in `processing/ripped-for-encoding/` with a `-check-title` suffix
(e.g., `MOVIE_TITLE-check-title.mkv`). Confirm or correct the title by removing the
`-check-title` suffix before proceeding to encoding.

### 2. Encode with HandBrake (`encode-backlog.ps1`)

Scans `processing/ripped-for-encoding/` for MKVs that need encoding.
Flags any files still carrying a `-check-title` suffix that need confirmation.
Automatically archives MKVs to `MKVs/` once a corresponding MP4 is found.
Launches HandBrake pointed at the ripped folder for remaining files.

```powershell
.\encode-backlog.ps1                              # prompted for movies directory
.\encode-backlog.ps1 -MoviesDir "G:\Movies"       # non-interactive
```

In HandBrake, use the **DVD Ripped Archive** preset (included in
`Handbrake DVD Preset.json`), set the output to `processing/encoded-for-upload/`,
and start the queue.

### 3. Upload MP4s to Google Drive (`upload-mp4s.ps1`)

Uploads encoded MP4 files from `processing/encoded-for-upload/` to Google Drive
via rclone. Automatically skips files still locked by HandBrake. After upload,
moves MP4s to the `MP4s/` archive.

```powershell
.\upload-mp4s.ps1                              # prompted for movies directory
.\upload-mp4s.ps1 -MoviesDir "G:\Movies"       # non-interactive
```

## Directory Structure

```
G:\Movies\                                    # Root movies directory (configurable)
  processing\                                 # Intermediate pipeline stages
    ripped-for-encoding\                      # MKVs from ripping, awaiting encoding
      Movie Title.mkv                         # Renamed, ready for HandBrake
      MOVIE_TITLE-check-title.mkv              # Needs title confirmation
    encoded-for-upload\                       # MP4s from HandBrake, awaiting upload
      Movie Title.mp4                         # Ready for upload to Google Drive
  MKVs\                                       # Archived source MKVs (post-encoding)
    Movie Title.mkv
  MP4s\                                       # Archived encoded files (post-upload)
    Movie Title.mp4
```

Files flow through the pipeline:

**DVD --> `ripped-for-encoding/` --> (rename) --> HandBrake --> `encoded-for-upload/` --> Google Drive + `MP4s/`**

MKVs are archived to `MKVs/` once encoding is confirmed.

## Features

- **Folder-based queue** -- A file's location represents its processing state.
  Each loop rescans its input folder, so they can be interrupted and restarted
  independently without losing track of progress.
- **Continuous loop** -- Insert a disc, walk away, and come back to a folder full of MKV files.
- **Multi-drive support** -- Run multiple instances of the script on different drives to rip
  in parallel.
- **Exponential backoff alerts** -- Plays an audio alert (WAV file) when waiting for a disc,
  with increasing intervals (2s -> 4s -> 8s -> ... up to 15 minutes). Alerts are
  silenced during quiet hours (11 PM – 7 AM).
- **Smart file naming** -- Uses the disc volume label for filenames, falling back to the
  MKV filename from MakeMKV if the volume label is generic (e.g., `DVD_VIDEO`).
  All ripped files get a `-check-title` suffix until confirmed.
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
  (see [rclone Setup](#rclone-setup) below)
- Windows PowerShell 5.1 or later
- One or more optical disc drives

## rclone Setup

The `upload-mp4s.ps1` script uses rclone to copy MP4 files to a folder called
`Movies` in your Google Drive. The remote is hardcoded as `gdrive:Movies`.

1. Install rclone: https://rclone.org/downloads/
2. Run `rclone config` and create a new remote named **`gdrive`** of type
   **Google Drive**. Follow the prompts to authorize access to your account.
3. Create a `Movies` folder in the root of your Google Drive (if it doesn't
   already exist). Uploaded MP4 files will appear here as flat files
   (e.g., `gdrive:Movies/Movie Title.mp4`).

To verify the setup, run:

```powershell
rclone lsl gdrive:Movies --human-readable
```

If you want to use a different remote name or destination folder, edit the
`rclone copy` line near the top of `upload-mp4s.ps1`.

## ripping-loop.ps1 Parameters

| Parameter     | Type   | Default                        | Description                              |
|---------------|--------|--------------------------------|------------------------------------------|
| `-InputDrive` | String | *(prompted interactively)*     | The drive letter to monitor. Accepts with or without a colon (e.g., `F` or `F:`). If omitted, the script prompts for it with a default of `D`. |
| `-MoviesDir`  | String | *(prompted interactively)*     | Root movies directory. Ripped files are saved to `processing\ripped-for-encoding\` under this path. Defaults to `G:\Movies`. |

## encode-backlog.ps1 Parameters

| Parameter    | Type   | Default                        | Description                              |
|--------------|--------|--------------------------------|------------------------------------------|
| `-MoviesDir` | String | *(prompted interactively)*     | Root movies directory containing `processing/`, `MKVs/`, and `MP4s/` subfolders. Defaults to `G:\Movies`. |

## upload-mp4s.ps1 Parameters

| Parameter    | Type   | Default                        | Description                              |
|--------------|--------|--------------------------------|------------------------------------------|
| `-MoviesDir` | String | *(prompted interactively)*     | Root movies directory containing `processing/` and `MP4s/` subfolders. Defaults to `G:\Movies`. |

## Configuration

The following variables can be modified at the top of `ripping-loop.ps1`:

| Variable          | Default                                           | Description                                      |
|-------------------|---------------------------------------------------|--------------------------------------------------|
| `$MakeMkvPath`    | `C:\Program Files (x86)\MakeMKV\makemkvcon64.exe` | Path to the MakeMKV command-line tool.           |
| `$MinLength`      | `3600` (seconds)                                  | Minimum title length to extract (skips menus, extras, etc.). |
