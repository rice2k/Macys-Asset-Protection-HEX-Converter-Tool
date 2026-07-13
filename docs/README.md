# Documentation Hub

This folder contains the detailed notes for the Macy's Asset Protection China Grove Hex Converter Utility.

## Start Here

| Page | Purpose |
| --- | --- |
| [User Guide](notes/user-guide.md) | How to use the current Windows desktop app. |
| [Feature Reference](notes/feature-reference.md) | Detailed list of app features and what each area does. |
| [Keyboard Shortcuts](notes/keyboard-shortcuts.md) | Shortcut keys, mouse actions, and copy actions. |
| [Version History](notes/version-history.md) | Current version, previous versions, restore tags, and archive branches. |
| [Downloads And Releases](notes/downloads-and-releases.md) | Test EXE download links, checksum, release notes, and automation status. |
| [Troubleshooting](notes/troubleshooting.md) | Common import, export, SmartScreen, and validation issues. |
| [Source History](source-history/README.md) | Original source document and earlier HTML versions. |

## Current Version

Current app version: `1.1.1`

Current EXE:

[Download Macys_AP_China_Grove_Hex_Utility.exe](https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/raw/main/dist/Macys_AP_China_Grove_Hex_Utility.exe)

Current SHA-256:

`15F7DA4D211292EC21002E760655EFEFBB3628F86A6EA9E9A10E8C04B46F3815`

## Project Structure

| Path | What It Contains |
| --- | --- |
| `desktop_app.py` | Current maintained Python/Tkinter desktop application. |
| `dist/` | Current built Windows EXE and checksum file. |
| `src/assets/` | App icon and custom visual assets. |
| `tests/` | Smoke checks for conversion, import cleanup, and app behavior. |
| `docs/screenshots/` | GitHub README screenshots. |
| `docs/source-history/` | Original script document and earlier HTML versions. |
| `docs/notes/` | Detailed GitHub documentation pages. |
| `RELEASE_NOTES_v*.md` | Release notes for version tags. |

## Conversion Rule

Facility Code (FC) is taken from the high 16 bits. Card Number (CN) is taken from the low 16 bits of the 32-bit HEX value.

## Documentation Notes

- The current maintained program is the Windows desktop app.
- The HTML files in `source-history` are historical reference versions.
- The direct EXE link is available even while automated release builds are not running.
- GitHub version tags and archive branches are kept so older project states can be recovered.
