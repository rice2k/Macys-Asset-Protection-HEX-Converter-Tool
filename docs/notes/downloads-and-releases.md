# Downloads And Releases

## Current Test Download

Latest test EXE:

[Download Macys_AP_China_Grove_Hex_Utility.exe](https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/raw/main/dist/Macys_AP_China_Grove_Hex_Utility.exe)

Checksum file:

[Download SHA-256 checksum](https://raw.githubusercontent.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/main/dist/Macys_AP_China_Grove_Hex_Utility.exe.sha256.txt)

Current SHA-256:

`15F7DA4D211292EC21002E760655EFEFBB3628F86A6EA9E9A10E8C04B46F3815`

## Windows SmartScreen

Windows may show a SmartScreen warning because the test EXE is not code-signed yet.

For testing:

1. Choose More info.
2. Choose Run anyway.

Future improvement: code signing would make Windows trust the EXE more.

## Release Notes

Release-note files are kept in the repository root:

- `RELEASE_NOTES_v1.1.1.md`
- `RELEASE_NOTES_v1.1.0.md`
- `RELEASE_NOTES_v1.0.9.md`
- `RELEASE_NOTES_v1.0.8.md`
- `RELEASE_NOTES_v1.0.7.md`
- `RELEASE_NOTES_v1.0.6.md`
- `RELEASE_NOTES_v1.0.5.md`

Use `RELEASE_NOTES_TEMPLATE.md` before creating a new tag.

## Automation Status

Automated GitHub release builds are not currently running. This does not affect the app, the repository, or the direct EXE download.

What still works:

- Repository browsing.
- Direct EXE download.
- Direct checksum download.
- Git tags.
- Archive branches.
- Local build and local testing.

What is not automated right now:

- Automatic Windows build workflow.
- Automatic GitHub Release publishing workflow.

## Local Build

```powershell
python -m PyInstaller --noconfirm --clean Macys_AP_China_Grove_Hex_Utility.spec
```

Expected output:

`dist/Macys_AP_China_Grove_Hex_Utility.exe`

## Local Tests

```powershell
python -m py_compile desktop_app.py tests\desktop_app_smoke.py
python tests\desktop_app_smoke.py
python desktop_app.py --self-test
```
