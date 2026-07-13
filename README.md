# Macy's Asset Protection China Grove Hex Converter Utility

Current version: `1.0.5`

Windows desktop utility for Macy's Asset Protection access-control conversion work at China Grove, North Carolina.

## What It Does

- Converts one or many 8-character HEX IDs into Facility Code and Card Number.
- Converts FC/CN pairs back into HEX IDs.
- Imports TXT, CSV, TSV, XLS, XLSX, XLSM, XML Spreadsheet, HTML, and copied table data.
- Highlights valid, warning, and invalid input rows.
- Removes duplicates and keeps only valid rows when cleaning a queue.
- Exports professional Excel, CSV, TXT, and PDF reports.
- Includes Help, About, Settings, History, default export settings, and desktop shortcut support.

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

GitHub: https://github.com/rice2k
