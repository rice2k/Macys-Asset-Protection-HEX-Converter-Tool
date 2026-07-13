# Excel And Clipboard Tips

## Best Excel Workflow

1. Open the spreadsheet.
2. Select the column or rows that contain access-control IDs.
3. Copy the selection.
4. In the app, use Import > Paste Clipboard To Queue.
5. Review the Paste Cleanup Preview when it appears.
6. Choose Add Clean IDs to add only detected IDs.

## Supported Spreadsheet Files

| Format | Notes |
| --- | --- |
| XLS | Legacy Excel workbook import. |
| XLSX | Modern Excel workbook import. |
| XLSM | Macro workbook import for data reading. |
| CSV | Comma-separated table import. |
| TSV | Tab-separated table import. |
| XML Spreadsheet | Spreadsheet-style XML import. |
| HTML / HTM | Table exports from browser or system reports. |

## What The App Cleans

| Messy Value | Cleaned Value |
| --- | --- |
| `88984765.0` | `88984765` |
| `8898-4130` | `88984130` |
| `8898 4130` | `88984130` |
| `Active User, 88984717` | `88984717` |

## Import Preview

For structured files, the app may show Import Preview before adding rows.

Use this preview to confirm that the app found the right column or the right name/ID pairs.

## When Import Does Not Find IDs

Try these fallback steps:

- Copy the exact ID column from Excel and use Paste Clipboard To Queue.
- Save the sheet as CSV and import the CSV.
- Confirm IDs are visible as numbers or text.
- Remove hidden filters in Excel before copying.
- Use Keep Valid after importing to remove bad rows.

## Drag And Drop

Drag supported files onto the Input Queue box when drag/drop is available.

If drag/drop does not work on a workstation, use Import > Browse Files. The file parsing is the same after the file is selected.
