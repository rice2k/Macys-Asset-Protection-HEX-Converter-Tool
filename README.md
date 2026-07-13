# Macy's Asset Protection China Grove Hex Converter Utility

![Version](https://img.shields.io/badge/version-1.1.1-e51b2d)
![Platform](https://img.shields.io/badge/platform-Windows-0b66c3)
![App Type](https://img.shields.io/badge/app-desktop%20utility-166534)
![Build](https://img.shields.io/badge/test%20EXE-available-7a4b00)
![Source History](https://img.shields.io/badge/source%20history-archived-667085)

Windows desktop utility for Macy's Asset Protection access-control conversion work at China Grove, North Carolina.

The utility converts access-control HEX values into Facility Code and Card Number, reverses FC/CN pairs back into HEX, cleans imported data, and exports professional reports for review.

## Quick Links

| Need | Link |
| --- | --- |
| Download latest test EXE | [Macys_AP_China_Grove_Hex_Utility.exe](https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/raw/main/dist/Macys_AP_China_Grove_Hex_Utility.exe) |
| Verify download checksum | [SHA-256 checksum](https://raw.githubusercontent.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/main/dist/Macys_AP_China_Grove_Hex_Utility.exe.sha256.txt) |
| Full documentation hub | [docs/README.md](docs/README.md) |
| User guide | [docs/notes/user-guide.md](docs/notes/user-guide.md) |
| Feature reference | [docs/notes/feature-reference.md](docs/notes/feature-reference.md) |
| Keyboard shortcuts | [docs/notes/keyboard-shortcuts.md](docs/notes/keyboard-shortcuts.md) |
| Version history | [docs/notes/version-history.md](docs/notes/version-history.md) |
| Original HTML/source archive | [docs/source-history/README.md](docs/source-history/README.md) |

## Current Test Build

Current version: `1.1.1`

Current SHA-256:

`15F7DA4D211292EC21002E760655EFEFBB3628F86A6EA9E9A10E8C04B46F3815`

Windows may show a SmartScreen warning because the EXE is not code-signed yet. For testing, choose **More info**, then **Run anyway**.

## Project Tags

`access-control` `asset-protection` `hex-converter` `facility-code` `card-number` `windows-desktop` `python` `tkinter` `excel-import` `report-export` `macy-style-ui` `china-grove`

## Conversion Rule

Facility Code (FC) is taken from the high 16 bits. Card Number (CN) is taken from the low 16 bits of the 32-bit HEX value.

Example:

| HEX | Facility Code | Card Number |
| --- | ---: | ---: |
| `88984717` | `34968` | `18199` |

## Main Features

- Batch converts many 8-character HEX IDs into Facility Code and Card Number.
- Converts one FC/CN pair or a batch of FC/CN pairs back into HEX.
- Imports TXT, CSV, TSV, XLS, XLSX, XLSM, XML Spreadsheet, HTML, and copied table data.
- Cleans Excel-style numeric IDs such as `88984765.0` and split IDs such as `8898-4765`.
- Previews cleaned numeric IDs before adding messy clipboard data to the queue.
- Highlights valid, warning, and invalid input rows.
- Shows Notes / Details only when a row was cleaned, duplicated, unusual, or invalid.
- Removes duplicates and keeps only valid rows during queue cleanup.
- Supports full-row copying, FC/CN copying, HEX copying, and right-click result actions.
- Exports Excel, CSV, TXT, and PDF reports.
- Includes Help, About, Settings, History, Recent Exports, desktop shortcut support, and copyable error reports.

## Screenshots

![Main screen](docs/screenshots/main.png)

![Settings](docs/screenshots/settings.png)

![Help](docs/screenshots/help.png)

![About](docs/screenshots/about.png)

![Export complete](docs/screenshots/export-complete.png)

## Source History

The original Access Control script document and earlier browser-based HTML versions are preserved so the project history is visible on GitHub:

[View original source history and HTML archive](docs/source-history/README.md)

The archived HTML versions are reference material only. The maintained app is the current Windows desktop EXE and `desktop_app.py`.

## Version Tags And Restore Points

| Version | Tag | Notes |
| --- | --- | --- |
| `1.1.1` | [`v1.1.1`](https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/tree/v1.1.1) | Current test build; improved Results alignment and Notes / Details wording. |
| `1.1.0` | [`v1.1.0`](https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/tree/v1.1.0) | Workflow polish, paste cleanup preview, right-click copy menus, recent export cleanup. |
| `1.0.9` | [`v1.0.9`](https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/tree/v1.0.9) | Almost-done baseline before v1.1.x workflow/documentation polish. |

More detail: [docs/notes/version-history.md](docs/notes/version-history.md)

## Run From Source

```powershell
python desktop_app.py
```

## Test

```powershell
python -m py_compile desktop_app.py tests\desktop_app_smoke.py
python tests\desktop_app_smoke.py
python desktop_app.py --self-test
```

## Build Windows EXE

```powershell
python -m PyInstaller --noconfirm --clean Macys_AP_China_Grove_Hex_Utility.spec
```

Expected output:

`dist/Macys_AP_China_Grove_Hex_Utility.exe`

## Credit

Made by Christopher Schumacher, Asset Protection FLO.

GitHub profile: [rice2k](https://github.com/rice2k)

## Release Notes

Releases use the built Windows EXE and SHA-256 checksum file from `dist`.

Before tagging a future version, add a matching `RELEASE_NOTES_vX.Y.Z.md` file so the GitHub Release has clean notes and restore details.

Automated GitHub release builds are not currently running, so the test EXE is linked directly from the repository for easy download and testing.

See [CHANGELOG.md](CHANGELOG.md) for version history.
