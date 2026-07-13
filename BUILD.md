# Build Macy's AP China Grove Hex Utility

The GitHub Actions workflow installs the Python dependencies, runs syntax and converter checks, builds the Windows executable with PyInstaller, and publishes the EXE plus a SHA-256 checksum as workflow artifacts.

Before creating a new tag, add `RELEASE_NOTES_vX.Y.Z.md` for that exact tag. The release workflow requires matching notes so every GitHub Release has a clean download page, checksum context, and restore notes.

Expected output:

`Macys_AP_China_Grove_Hex_Utility.exe`
