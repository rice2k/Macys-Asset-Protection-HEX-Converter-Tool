# Macy's AP China Grove Hex Utility v1.1.6

## Summary

Polish release focused on a cleaner application header, clearer sidebar guidance, and safer scanner/import cleanup for numeric ID workflows.

## Changes

- Replaced the long top banner image with a compact China Grove identity badge.
- Updated the custom dropdown buttons to use a plain readable caret.
- Changed the sidebar Quick Tip area into a Field Guide card with a title and short explanation.
- Improved import, scanner, and cleaned clipboard behavior so those paths pull numeric 8-character IDs only.
- Kept manual Batch Converter conversion flexible for typed 8-character HEX IDs.
- Added tests for numeric-only cleanup behavior.

## Verification

- `python -m py_compile desktop_app.py tests\desktop_app_smoke.py`
- `python tests\desktop_app_smoke.py`
- `python desktop_app.py --self-test`
- Standalone EXE self-test after rebuild

## Download

[Macys_AP_China_Grove_Hex_Utility.exe](https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/raw/main/dist/Macys_AP_China_Grove_Hex_Utility.exe)

## SHA-256

`6fe892d69486ac7adacfe5d3496ac451fafb5bcd2cfe3f83a5f9eed1f6947383`
