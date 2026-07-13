# Macy's Asset Protection China Grove Hex Converter Utility

Current version: `1.1.1`

Windows desktop utility for Macy's Asset Protection access-control conversion work at China Grove, North Carolina.

## Download Test Build

Latest test EXE:

[Download Macys_AP_China_Grove_Hex_Utility.exe](https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/raw/main/dist/Macys_AP_China_Grove_Hex_Utility.exe)

Checksum file:

[Download SHA-256 checksum](https://raw.githubusercontent.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/main/dist/Macys_AP_China_Grove_Hex_Utility.exe.sha256.txt)

Current SHA-256:

`15F7DA4D211292EC21002E760655EFEFBB3628F86A6EA9E9A10E8C04B46F3815`

Windows may show a SmartScreen warning because the EXE is not code-signed yet. For testing, users can choose More info, then Run anyway.

## What It Does

- Converts one or many 8-character HEX IDs into Facility Code and Card Number.
- Converts FC/CN pairs back into HEX IDs.
- Imports TXT, CSV, TSV, XLS, XLSX, XLSM, XML Spreadsheet, HTML, and copied table data.
- Cleans Excel-style numeric IDs such as `88984765.0` and split IDs such as `8898-4765`.
- Highlights valid, warning, and invalid input rows.
- Shows a live row count while editing the input queue.
- Previews cleaned numeric IDs before pasting messy Excel-style clipboard data.
- Shows Notes / Details only when a row was cleaned, duplicated, unusual, or invalid.
- Removes duplicates and keeps only valid rows when cleaning a queue.
- Supports full-row copying and right-click copy actions in result tables.
- Exports professional Excel, CSV, TXT, and PDF reports.
- Includes Help, About, Settings, History, recent exports, clearable recent export history, default export settings, copyable error reports, and desktop shortcut support.

## Screenshots

![Main screen](docs/screenshots/main.png)

![Settings](docs/screenshots/settings.png)

![Help](docs/screenshots/help.png)

![About](docs/screenshots/about.png)

![Export complete](docs/screenshots/export-complete.png)

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

The current built executable is:

`dist/Macys_AP_China_Grove_Hex_Utility.exe`

## Credit

Made by Christopher Schumacher, Asset Protection FLO.

GitHub project: https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool

## Release Notes

Releases use the built Windows EXE and SHA-256 checksum file from `dist`.
Before tagging a future version, add a matching `RELEASE_NOTES_vX.Y.Z.md` file so the GitHub Release has clean notes and restore details.

See `CHANGELOG.md` for version history.
