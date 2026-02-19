# DVD Ripper Automation Script

A PowerShell script that automates DVD ripping using [MakeMKV](https://www.makemkv.com/).
It continuously monitors an optical drive, rips inserted discs to MKV files,
ejects the disc when done, and waits for the next one.

## Features

- **Continuous loop** — Insert a disc, walk away, and come back to a folder full of MKV files.
- **Multi-drive support** — Run multiple instances of the script on different drives to rip
  in parallel.
- **Exponential backoff alerts** — Plays an audio alert (WAV file) when waiting for a disc,
  with increasing intervals (2s → 4s → 8s → ... up to 5 minutes).
- **Smart file naming** — Uses the disc volume label with a timestamp for unique filenames.
  Falls back to the MKV filename from MakeMKV if the volume label is generic
  (e.g., `DVD_VIDEO`).
- **Progress monitoring** — Reports file size every 60 seconds while ripping, with a
  responsive 2-second polling loop for quick completion detection.
- **Transient disc filtering** — Detects and ignores brief drive accessibility during
  disc ejection/insertion to avoid false starts.
- **Safe cleanup** — Temporary folders are only deleted after a successful file move.

## Prerequisites

- [MakeMKV](https://www.makemkv.com/) installed
  (default path: `C:\Program Files (x86)\MakeMKV\makemkvcon64.exe`)
- Windows PowerShell 5.1 or later
- One or more optical disc drives

## Usage

The script expects to be run from a `Videos\Ripped Movies` folder in the current
user's home directory (e.g., `C:\Users\<username>\Videos\Ripped Movies`). Output
MKV files are saved to a `MKVs` subfolder within that directory.

```powershell
# Use default drive (D:)
.\mkv-loop.ps1

# Specify a drive letter
.\mkv-loop.ps1 -Drive F

# Run multiple drives in parallel (in separate terminals)
.\mkv-loop.ps1 -Drive D
.\mkv-loop.ps1 -Drive F
```

## Parameters

| Parameter | Type   | Default | Description                              |
|-----------|--------|---------|------------------------------------------|
| `-Drive`  | String | `D`     | The drive letter to monitor. Accepts with or without a colon (e.g., `F` or `F:`). |

## Configuration

The following variables can be modified at the top of the script:

| Variable          | Default                                           | Description                                      |
|-------------------|---------------------------------------------------|--------------------------------------------------|
| `$MakeMkvPath`    | `C:\Program Files (x86)\MakeMKV\makemkvcon64.exe` | Path to the MakeMKV command-line tool.           |
| `$MinLength`      | `3600` (seconds)                                  | Minimum title length to extract (skips menus, extras, etc.). |

## Directory Structure

```
~/Videos/Ripped Movies/
├── mkv-loop.ps1              # This script
├── alert.wav                 # Optional alert sound (played when waiting for disc)
├── MKVs/                     # Output directory for ripped MKV files
│   ├── MOVIE_NAME-26-02-18-15-45.mkv
│   ├── ANOTHER_MOVIE-26-02-18-16-30.mkv
│   └── MOVIE_NAME-26-02-18-17-00/   # Temporary folder (deleted after successful rip)
└── ...
```

## License

MIT
