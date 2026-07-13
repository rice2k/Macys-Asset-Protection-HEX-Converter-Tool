# Release Checklist

Use this checklist before publishing a future release.

## 1. Update Version Files

- Update `VERSION`.
- Update `APP_VERSION` in `desktop_app.py`.
- Update README badges and current version text.
- Add `RELEASE_NOTES_vX.Y.Z.md`.
- Update [CHANGELOG.md](../../CHANGELOG.md).

## 2. Run Tests

```powershell
python -m py_compile desktop_app.py tests\desktop_app_smoke.py
python tests\desktop_app_smoke.py
python desktop_app.py --self-test
```

## 3. Build The EXE

```powershell
python -m PyInstaller --noconfirm --clean Macys_AP_China_Grove_Hex_Utility.spec
```

Expected output:

```text
dist/Macys_AP_China_Grove_Hex_Utility.exe
```

## 4. Create Checksum

```powershell
Get-FileHash dist\Macys_AP_China_Grove_Hex_Utility.exe -Algorithm SHA256
```

Update:

```text
dist/Macys_AP_China_Grove_Hex_Utility.exe.sha256.txt
```

## 5. Refresh Screenshots

Update:

- [../screenshots/main.png](../screenshots/main.png)
- [../screenshots/menu-options.png](../screenshots/menu-options.png)
- [../screenshots/right-click-results.png](../screenshots/right-click-results.png)
- [../screenshots/settings.png](../screenshots/settings.png)
- [../screenshots/help.png](../screenshots/help.png)
- [../screenshots/about.png](../screenshots/about.png)
- [../screenshots/export-complete.png](../screenshots/export-complete.png)

## 6. Commit And Tag

```powershell
git status --short
git add .
git commit -m "Release vX.Y.Z"
git tag vX.Y.Z
git push origin main
git push origin vX.Y.Z
```

## 7. Create Archive Branch

```powershell
git branch archive/vX.Y.Z vX.Y.Z
git push origin archive/vX.Y.Z
```

## 8. Create GitHub Release

GitHub Releases page:

<https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/releases>

Recommended release assets:

- `Macys_AP_China_Grove_Hex_Utility.exe`
- `Macys_AP_China_Grove_Hex_Utility.exe.sha256.txt`
- Matching `RELEASE_NOTES_vX.Y.Z.md`

Recommended release title:

```text
Macy's AP China Grove Hex Converter Utility vX.Y.Z
```

## 9. Update GitHub Page Settings

Suggested social preview image:

[../images/social-preview.png](../images/social-preview.png)

Suggested project description:

```text
Windows desktop utility for Macy's Asset Protection HEX, Facility Code, and Card Number conversion.
```

## 10. Recommended GitHub Topics

Add these in GitHub repository settings when topic editing is available:

```text
access-control
asset-protection
hex-converter
facility-code
card-number
windows-app
python
tkinter
excel-import
report-export
china-grove
```
