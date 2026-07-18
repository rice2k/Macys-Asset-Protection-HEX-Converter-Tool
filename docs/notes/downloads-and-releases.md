# Downloads And Releases

## Current Test Download

Latest test EXE:

[Download Macys_AP_China_Grove_Hex_Utility.exe](https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/raw/main/dist/Macys_AP_China_Grove_Hex_Utility.exe)

Checksum file:

[Download SHA-256 checksum](https://raw.githubusercontent.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/main/dist/Macys_AP_China_Grove_Hex_Utility.exe.sha256.txt)

Current SHA-256:

`09402ce256e3e5a20e9526d8bd8e04ed3e8d081f3cfcbdc62976eaa77a3fa0f1`

## Important URLs

| Resource | URL |
| --- | --- |
| Repository | <https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool> |
| Direct EXE | <https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/raw/main/dist/Macys_AP_China_Grove_Hex_Utility.exe> |
| Direct checksum | <https://raw.githubusercontent.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/main/dist/Macys_AP_China_Grove_Hex_Utility.exe.sha256.txt> |
| Tags | <https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/tags> |
| Releases area | <https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/releases> |
| Release checklist | [release-checklist.md](release-checklist.md) |

## Verify The Download

Run this from the folder where the EXE was downloaded:

```powershell
Get-FileHash .\Macys_AP_China_Grove_Hex_Utility.exe -Algorithm SHA256
```

The hash should match:

```text
09402ce256e3e5a20e9526d8bd8e04ed3e8d081f3cfcbdc62976eaa77a3fa0f1
```

If the hash does not match, delete the downloaded EXE and download it again from the repository link above.

## Windows SmartScreen

Windows may show a SmartScreen warning because the test EXE is not code-signed yet.

For testing:

1. Choose More info.
2. Choose Run anyway.

Future improvement: code signing would make Windows trust the EXE more.

## Release Notes

Release-note files are kept in the repository root:

- `RELEASE_NOTES_v1.1.4.md`
- `RELEASE_NOTES_v1.1.3.md`
- `RELEASE_NOTES_v1.1.2.md`
- `RELEASE_NOTES_v1.1.1.md`
- `RELEASE_NOTES_v1.1.0.md`
- `RELEASE_NOTES_v1.0.9.md`
- `RELEASE_NOTES_v1.0.8.md`
- `RELEASE_NOTES_v1.0.7.md`
- `RELEASE_NOTES_v1.0.6.md`
- `RELEASE_NOTES_v1.0.5.md`

Use `RELEASE_NOTES_TEMPLATE.md` before creating a new tag.

Manual GitHub Releases should be created from:

<https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/releases>

Recommended attached files:

- `Macys_AP_China_Grove_Hex_Utility.exe`
- `Macys_AP_China_Grove_Hex_Utility.exe.sha256.txt`
- Matching `RELEASE_NOTES_vX.Y.Z.md`

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
