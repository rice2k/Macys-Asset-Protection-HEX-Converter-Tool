# Original Source History

This folder preserves the original Access Control materials and earlier HTML versions that led to the current Windows EXE utility.

These files are archive/reference material only. The current maintained desktop application remains the Python/Tkinter utility in the repository root, and the current test EXE remains in `dist`.

## Preserved Conversion Rule

Facility Code (FC) is taken from the high 16 bits. Card Number (CN) is taken from the low 16 bits of the 32-bit HEX value.

## Original Planning / Script Document

| File | Purpose |
| --- | --- |
| [Access Control Script.docx](<Access Control Script.docx>) | Original source/planning document that started the project history. |

## HTML Version Archive

| Version | File | Feature Notes |
| --- | --- | --- |
| Current v29 | [AP Access Control - Current - v29 - 2026-03-29.html](<html-versions/AP Access Control - Current - v29 - 2026-03-29.html>) | Later browser-based version with Macy's Asset Protection naming, single and batch conversion, FC/CN to HEX, supported text and Excel-style imports, drag/drop import, formatted Excel export, TXT download, and built-in usage help. |
| v28 | [AP Access Control - v28 - 2026-03-29.html](<html-versions/AP Access Control - v28 - 2026-03-29.html>) | Late HTML version with Macy's Asset Protection naming, batch conversion, single/reverse conversion sections, import support, export options, and date-based export naming. |
| v27 | [AP Access Control - v27 - 2026-03-29.html](<html-versions/AP Access Control - v27 - 2026-03-29.html>) | Late AP Access Control browser version with text/Excel-style import, drag/drop import, formatted Excel export, TXT download, and single/reverse conversion. |
| v21 | [AP Access Control - v21 - 2026-03-29.html](<html-versions/AP Access Control - v21 - 2026-03-29.html>) | Browser version with HEX to FC/CN, FC/CN to HEX, import TXT/Excel, formatted Excel export, and conversion warning tracking. |
| v20 | [AP Access Control - v20 - 2026-03-29.html](<html-versions/AP Access Control - v20 - 2026-03-29.html>) | Browser version focused on TXT/Excel import, batch results, reverse conversion, and formatted Excel export. |
| v15 | [AP Access Control - v15 - 2026-03-29.html](<html-versions/AP Access Control - v15 - 2026-03-29.html>) | Added broader import wording for TXT/CSV/Excel, warning tracking, batch results, and formatted Excel export. |
| v14 | [AP Access Control - v14 - 2026-03-05.html](<html-versions/AP Access Control - v14 - 2026-03-05.html>) | Early browser-based converter with HEX to FC/CN, FC/CN to HEX, multi-result table, and formatted Excel export. |

## Current App Compared To HTML Versions

| Area | HTML Versions | Current Windows App |
| --- | --- | --- |
| App type | Browser-based single HTML files. | Windows desktop EXE. |
| Input | Text box and later file import/drop area. | Input Queue with live count, color validation, paste cleanup preview, file import, and optional drag/drop. |
| Conversion | HEX to FC/CN and FC/CN to HEX. | Batch HEX to FC/CN, Single Lookup, FC/CN to Hex, and Unconvert Batch. |
| Cleanup | Earlier extraction and warning behavior. | Excel numeric cleanup, split ID cleanup, full-text extraction, duplicate cleanup, Keep Valid, and detailed Notes / Details. |
| Export | Formatted Excel and TXT in later versions. | Excel, CSV, TXT, and PDF reports with app version and professional formatting. |
| Interface | HTML controls. | Macy's-style desktop GUI with sidebar, dropdown menus, Help, About, Settings, History, Recent Exports, and status strip. |
| Recoverability | Saved as archive files. | Git tags, archive branches, changelog, release notes, and source-history docs. |

## Notes

- The HTML files show the earlier browser-based versions made before the desktop EXE redesign.
- The current EXE is the intended test build for users.
- The historical files are kept so the project history is visible on GitHub without changing the current application.
- If a user wants to inspect an older browser version, open the matching HTML file from this folder.
