# Macy's AP China Grove Hex Utility v1.1.5

## Highlights

- Fixes the Windows download package so the EXE runs by itself.
- Removes the missing `_internal\\python313.dll` failure caused by uploading a folder-style EXE without its support folder.
- Restores the PyInstaller spec to true one-file mode with bundled Python DLLs, dependencies, and app assets.
- Keeps the v1.1.4 visual polish for dropdowns and sidebar strips.

## Verification

- Syntax check
- Desktop smoke checks
- Built-in self-test
- Windows EXE self-test
- Clean-folder standalone EXE self-test
- Windows EXE build with checksum

SHA-256:

`2c540f1af57b646adf83e45fa1d4f450900055a0e011c25289f045cb5fa57c5f`

## Restore Notes

- Previous restore tag: `v1.1.4`
- Previous archive branch: `archive/v1.1.4`
- This release should use the matching standalone EXE and SHA-256 checksum generated from `dist`.
