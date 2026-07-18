# Feature Reference

## Current App

Current maintained app: Windows desktop EXE built from `desktop_app.py`

Current version: `1.1.2`

Full option map: [Every Option Reference](every-option-reference.md)

Screenshots: [Screenshot Guide](../screenshots/README.md)

## Workspaces

| Workspace | Purpose |
| --- | --- |
| Batch Converter | Main workflow for converting many HEX IDs to FC/CN. |
| Single Hex Lookup | Fast one-ID conversion. |
| FC/CN to Hex | Reverse one Facility Code/Card Number pair back to HEX. |
| Unconvert Batch | Reverse multiple FC/CN pairs back to HEX. |
| History | Review recent local conversion runs. |

## Batch Converter Features

- Multi-line Input Queue.
- Scanner Input field for keyboard-style handheld scanners.
- File import.
- Clipboard paste cleanup preview.
- Drag/drop file import when optional drag/drop support is available.
- Live row count.
- Input validation color highlights.
- Remove Duplicates.
- Keep Valid.
- Numeric-only sample data.
- Conversion summary cards.
- Warning and invalid counts.
- Results search.
- Status filter.
- Sortable result columns.
- Row hover highlighting.
- Header tooltips.
- Right-click result menu.
- Copy All.
- Copy FC.
- Copy CN.
- Copy Pair.
- Copy Row.
- Clear Invalid.

## Scanner Features

- Batch Scanner Input accepts handheld scanner input that behaves like keyboard typing.
- Enter and Tab scanner suffixes submit the scanned value.
- Each new batch scan is cleaned and placed at the top of the Input Queue.
- The scanner strip tracks the current session scan count.
- F9 focuses Batch Scanner Input.
- F10 focuses Single Hex Lookup.
- Single Hex Lookup can auto-convert a scanned HEX ID and copy the FC,CN pair.

## Import Features

Supported file types:

| Type | Notes |
| --- | --- |
| TXT | Raw lines or delimited rows. |
| CSV | Delimited table import and raw fallback. |
| TSV | Tab-delimited table import. |
| XLS | Legacy Excel import when dependency support is available. |
| XLSX | Spreadsheet import. |
| XLSM | Macro-enabled workbook import, read as data. |
| XML Spreadsheet | Spreadsheet-style XML rows. |
| HTML / HTM | HTML table-style Excel exports. |

Cleanup behavior:

- Cleans `88984765.0` into `88984765`.
- Joins `8898-4765` into `88984765`.
- Joins `8898 4765` into `88984765`.
- Extracts an ID from full employee-style text.
- Detects likely Candidate Name and Colleague # columns.
- Deduplicates imported employee-ID rows.

## Conversion Features

HEX to FC/CN:

- Validates exactly 8 HEX characters.
- Converts locally on the workstation.
- Calculates FC from the high 16 bits.
- Calculates CN from the low 16 bits.
- Flags duplicates.
- Flags unusual values.

FC/CN to HEX:

- Accepts whole numbers.
- Requires 0 to 65535.
- Builds 8-character uppercase HEX.
- Warns on unusual values.

## Status And Notes

| Status | Meaning |
| --- | --- |
| Valid | Converted without warning. |
| Warning | Converted, but app noticed cleanup, duplicate, or unusual value. |
| Invalid | Could not convert. |

Notes / Details can come from:

- Excel numeric cleanup.
- Full-text extraction.
- Split ID cleanup.
- Duplicate detection.
- All zeros.
- All Fs.
- Numeric ID not starting with common `88` prefix.
- Facility Code 0.
- Card Number 0.
- Very high FC or CN value.
- Invalid characters.
- Too short or too long input.

## Export Features

| Export | Features |
| --- | --- |
| Excel Workbook | Summary sheet, Results sheet, frozen headers, filters, status colors, wrapped Notes / Details. |
| CSV Report | Spreadsheet-friendly rows with app version and converted timestamp. |
| TXT Report | Plain text report with aligned columns. |
| PDF Report | Landscape report with summary and status-highlighted rows. |

Export download and release notes: [Downloads And Releases](downloads-and-releases.md)

## Settings Features

- Default export type.
- Default export folder.
- Open default export folder.
- Create desktop shortcut.
- Clear Recent Exports.

## Reliability Features

- Copy Last Error Report for import/export/open failures.
- Recent Exports list skips missing files.
- Local History keeps recent conversion summaries.
- App version appears in About and exports.
- SHA-256 checksum is published for the EXE.

## Visual/UI Features

- Macy's-inspired red accent color.
- Custom app icon.
- Custom window icon for popups.
- Light corporate desktop layout.
- Sidebar workspace navigation.
- Dropdown menu buttons.
- Styled dropdown and right-click menus.
- Scrollbars in text areas and result tables.
- Full-view startup.
- Mouse-wheel scrolling.
- Table hover highlighting.
- Status strip for Ready / Needs Review / Exported states.
- Clickable sidebar and bottom status areas that jump to the related workspace.

## GitHub Visual Assets

- [README banner](../images/github-banner.png)
- [App icon](../images/app-icon.png)
- [Screenshot collage](../images/screenshot-collage.png)
- [Demo workflow GIF](../images/demo-workflow.gif)
- [Right-click menu reference](../images/right-click-menu-reference.png)
