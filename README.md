# DVD Rip-to-Cloud Pipeline

A set of PowerShell scripts that automate the full DVD archival workflow:
ripping discs to MKV, encoding to MP4 with HandBrake, and uploading to Google Drive.

Files flow through a folder-based pipeline where each file's location represents its
processing state. Each script can be interrupted and restarted independently — it
simply rescans its input folder to pick up where it left off.

## Quick Start

1. **Start ripping** — Run `.\rip-all-drives.ps1` to launch parallel ripping
   sessions in Windows Terminal for all detected optical drives. Each pane monitors
   one drive, rips the disc, ejects it, and plays an alert sound so you know to
   insert the next one. Just keep feeding discs.

2. **Fix titles** — Run `.\fix-titles.ps1` to rename MKV and MP4 files to
   proper `Title (Year)` format using TMDb lookups. The script searches TMDb,
   shows the top matches with runtimes and vote counts, and lets you confirm
   or correct each one. Use batch mode (`B`) to process 10 at a time.

3. **Encode** — Run `.\encode-backlog.ps1` whenever you're ready. It shows which files
   need encoding, archives already-encoded MKVs, and for each folder prompts:
   **[Y]** rename titles then encode, **[E]** encode as-is (skip renaming), or
   **[S]** skip. Launches HandBrake pointed at the ripped folder. In HandBrake,
   set the output to `processing\encoded-for-upload\` and start the queue.

4. **Upload** — Run `.\upload-mp4s.ps1` to push encoded MP4s to Google Drive via rclone.
   After uploading, it offers to move the files to the `MP4s\` archive folder.

Steps 1–2 can happen continuously while steps 3–4 are run on demand at any time.
Encoding and uploading are safe to run even while ripping is still in progress.

## Workflow Details

### 1. Rip DVDs to MKV

#### Multi-drive (`rip-all-drives.ps1`)

Detects all optical drives, lets you select which to use, then launches
Windows Terminal with side-by-side panes running `ripping-loop.ps1` for each.

```powershell
.\rip-all-drives.ps1                              # prompted for movies dir
.\rip-all-drives.ps1 -MinLength 1800              # shorter minimum title length
```

#### Single drive (`ripping-loop.ps1`)

Continuously monitors an optical drive, rips inserted discs using
[MakeMKV](https://www.makemkv.com/) via `dev:X:` (direct device access — no
cross-drive scanning), ejects, and waits for the next disc.

```powershell
.\ripping-loop.ps1                                            # prompted for drive and movies dir
.\ripping-loop.ps1 -InputDrive D -MoviesDir "G:\Movies"        # non-interactive
```

Ripped files land in `processing/ripped-for-encoding/` with a `-check-title` suffix
(e.g., `MOVIE_TITLE-check-title.mkv`). Confirm or correct the title by removing the
`-check-title` suffix before proceeding to encoding.

### 2. Encode with HandBrake (`encode-backlog.ps1`)

Scans `processing/ripped-for-encoding/` for MKVs that need encoding.
Automatically archives MKVs to `MKVs/` once a corresponding MP4 is found.
For each folder, prompts with three options:
- **Y** (default) — Rename titles via TMDb lookup, then launch HandBrake
- **E** — Encode as-is, skipping all rename prompts
- **S** — Skip the folder entirely

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

Use `-Archive` to upload directly from the `MP4s/` archive folder. This is useful
after renaming files with `fix-titles.ps1` — it checks the remote for exact filename
matches and only uploads files that don't already exist.

```powershell
.\upload-mp4s.ps1 -Archive                     # upload renamed files from MP4s/
```

### Utility Scripts

#### Fix Titles (`fix-titles.ps1`)

Renames MKV and MP4 files to `Title (Year)` format using TMDb lookups. Scans
`MKVs/` and `MP4s/` folders, extracts metadata titles via ffprobe, searches TMDb,
and displays the top 5 matches sorted by exact title match and vote count. Each
result shows runtime and vote count for easy comparison.

Files already in `Title (Year)` format are skipped automatically on re-runs.
MP4-only files (no matching MKV) are also supported.

```powershell
.\fix-titles.ps1                               # prompted for movies dir and API key
.\fix-titles.ps1 -MoviesDir "G:\Movies"        # non-interactive (key from .tmdb-api-key)
```

Prompt options per movie:
- **1–5** — Select a TMDb result (default is `1`, press Enter to accept)
- **S** — Skip this movie
- **C** — Custom search (re-query TMDb with different terms)
- **M** — Manual entry (type title and year directly)
- **B** — Batch mode: auto-select #1 for the next 10 movies, then review
- **Q** — Quit and show summary

In batch mode, after processing 10 movies you get a numbered review list. Type a
number to correct any movie (reverts the rename and lets you re-pick), `B` for the
next batch, `S` to return to single mode, or `Q` to quit.

#### Fix Faststart (`fix-faststart.ps1`)

Checks MP4 files in `MP4s/` and `encoded-for-upload/` for moov atom placement.
Files with the moov atom after mdat (non-faststart) are re-muxed in place using
`ffmpeg -movflags +faststart`. Faststart placement is required for smooth streaming.

```powershell
.\fix-faststart.ps1                            # prompted for movies directory
.\fix-faststart.ps1 -MoviesDir "G:\Movies"     # non-interactive
```

#### Delete If Backed Up (`delete-if-backedup.ps1`)

Compares files in a source folder against a backup folder. Lists files that exist
in both locations with matching sizes, then offers to delete them from the source.
Useful for cleaning up local copies after confirming they've been backed up.

```powershell
.\delete-if-backedup.ps1 -Source "G:\Movies\MP4s" -Backup "D:\Backup\MP4s"
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
- **Multi-drive support** -- `rip-all-drives.ps1` launches side-by-side Windows Terminal
  panes for each optical drive. Uses `dev:X:` direct device access to avoid MakeMKV's
  cross-drive scanning.
- **Exponential backoff alerts** -- Plays an audio alert (WAV file) when waiting for a disc,
  with increasing intervals (2s -> 4s -> 8s -> ... up to 15 minutes). Alerts are
  silenced during quiet hours (11 PM – 7 AM).
- **Smart file naming** -- Uses the disc volume label for filenames, falling back to the
  MKV filename from MakeMKV if the volume label is generic (e.g., `DVD_VIDEO`).
  All ripped files get a `-check-title` suffix until confirmed.
- **Progress monitoring** -- Reports file size and transfer speed (MB/s) every 60 seconds
  while ripping. A speed of 0.0 MB/s indicates a stalled or retrying read.
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
- [FFmpeg](https://ffmpeg.org/) installed and on PATH — `fix-faststart.ps1` uses
  `ffmpeg` for moov atom fixing; `fix-titles.ps1` uses `ffprobe` for metadata
  extraction and runtime detection
- [TMDb API key](#tmdb-api-key-setup) (free — used by `fix-titles.ps1` for movie lookups)
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

## TMDb API Key Setup

The `fix-titles.ps1` script uses [The Movie Database (TMDb)](https://www.themoviedb.org/)
API to look up official movie titles and release years. A free API key is required.

1. Create a free account at https://www.themoviedb.org/signup
2. Go to **Settings > API** (https://www.themoviedb.org/settings/api)
3. Click **Create** under "Request an API Key" and choose the **Developer** plan (free)
4. Fill out the form:
   - **Application Name**: anything (e.g., "DVD Ripper")
   - **Application URL**: any URL (e.g., `https://github.com`)
   - **Type of Use**: Desktop Application
   - **Application Summary**: "Personal script to look up movie titles for renaming local media files"
5. Submit — your **API Key (v3 auth)** will appear on the API page

When you first run `fix-titles.ps1`, it will prompt for the key and offer to save it
to `.tmdb-api-key` in the script directory for future runs. You can also pass it
directly:

```powershell
.\fix-titles.ps1 -TmdbApiKey "your-key-here"
```

> **Note:** The `.tmdb-api-key` file is not checked into source control. Keep your
> key private.

## rip-all-drives.ps1 Parameters

| Parameter     | Type   | Default                        | Description                              |
|---------------|--------|--------------------------------|------------------------------------------|
| `-MoviesDir`  | String | *(prompted interactively)*     | Root movies directory. Defaults to `G:\Movies`. |
| `-MinLength`  | Int    | `3600`                         | Minimum title length in seconds to rip. Passed to each `ripping-loop.ps1` instance. |

## ripping-loop.ps1 Parameters

| Parameter     | Type   | Default                        | Description                              |
|---------------|--------|--------------------------------|------------------------------------------|
| `-InputDrive` | String | *(prompted interactively)*     | The drive letter to monitor. Accepts with or without a colon (e.g., `F` or `F:`). If omitted, the script prompts for it with a default of `D`. |
| `-MoviesDir`  | String | *(prompted interactively)*     | Root movies directory. Ripped files are saved to `processing\ripped-for-encoding\` under this path. Defaults to `G:\Movies`. |
| `-MinLength`  | Int    | `3600`                         | Minimum title length in seconds to rip (skips menus, extras, etc.). |

## encode-backlog.ps1 Parameters

| Parameter    | Type   | Default                        | Description                              |
|--------------|--------|--------------------------------|------------------------------------------|
| `-MoviesDir` | String | *(prompted interactively)*     | Root movies directory containing `processing/`, `MKVs/`, and `MP4s/` subfolders. Defaults to `G:\Movies`. |

## upload-mp4s.ps1 Parameters

| Parameter    | Type   | Default                        | Description                              |
|--------------|--------|--------------------------------|------------------------------------------|
| `-MoviesDir` | String | *(prompted interactively)*     | Root movies directory containing `processing/` and `MP4s/` subfolders. Defaults to `G:\Movies`. |
| `-Archive`   | Switch | Off                            | Upload from the `MP4s/` archive folder instead of `encoded-for-upload/`. Checks the remote for exact filename matches and skips files that already exist. Does not move any local files after upload. |

## fix-titles.ps1 Parameters

| Parameter      | Type   | Default                        | Description                              |
|----------------|--------|--------------------------------|------------------------------------------|
| `-MoviesDir`   | String | *(prompted interactively)*     | Root movies directory containing `MKVs/` and `MP4s/` subfolders. Defaults to `G:\Movies`. |
| `-TmdbApiKey`  | String | *(from `.tmdb-api-key` file)*  | TMDb v3 API key. If omitted, reads from `.tmdb-api-key` in the script directory, or prompts to enter and save one. |

## fix-faststart.ps1 Parameters

| Parameter    | Type   | Default                        | Description                              |
|--------------|--------|--------------------------------|------------------------------------------|
| `-MoviesDir` | String | *(prompted interactively)*     | Root movies directory containing `MP4s/` and `processing/encoded-for-upload/` subfolders. Defaults to `G:\Movies`. |

## delete-if-backedup.ps1 Parameters

| Parameter  | Type   | Default                        | Description                              |
|------------|--------|--------------------------------|------------------------------------------|
| `-Source`  | String | *(prompted interactively)*     | Folder containing files to check and potentially delete. |
| `-Backup`  | String | *(prompted interactively)*     | Folder to verify files exist in (with matching sizes) before allowing deletion. |

## Configuration

The following variables can be modified at the top of `ripping-loop.ps1`:

| Variable          | Default                                           | Description                                      |
|-------------------|---------------------------------------------------|--------------------------------------------------|
| `$MakeMkvPath`    | `C:\Program Files (x86)\MakeMKV\makemkvcon64.exe` | Path to the MakeMKV command-line tool.           |
| `$MinLength`      | `3600` (seconds)                                  | Minimum title length to extract (skips menus, extras, etc.). |
