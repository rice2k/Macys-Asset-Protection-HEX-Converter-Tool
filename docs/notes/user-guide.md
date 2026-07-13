# User Guide

## Purpose

The Macy's Asset Protection China Grove Hex Converter Utility helps users convert access-control identifiers between:

- 8-character HEX ID
- Facility Code (FC)
- Card Number (CN)

The app is intended for quick operational review, cleanup, copying, and reporting.

## Conversion Rule

Facility Code (FC) is taken from the high 16 bits. Card Number (CN) is taken from the low 16 bits of the 32-bit HEX value.

Example:

| HEX | Facility Code | Card Number |
| --- | ---: | ---: |
| `88984717` | `34968` | `18199` |

## Basic Workflow

1. Open the EXE.
2. Use Batch Converter for most work.
3. Paste IDs, import files, or drag supported files into the Input Queue.
4. Review row colors before converting.
5. Use Convert.
6. Review Results and Notes / Details.
7. Copy values or export a report.

## Input Queue

The Input Queue accepts one value per line. It can also pull an 8-character ID from many full text lines.

Accepted examples:

```text
88984717
88984765.0
8898-4130
Active Test User, 88984717
```

Row colors:

| Color | Meaning |
| --- | --- |
| Green | Clean valid HEX ID. |
| Yellow | Valid, but cleaned, duplicate, or unusual. |
| Red | Invalid and needs review. |

## Importing Data

Use Import to add supported files to the queue.

Supported file types:

- TXT
- CSV
- TSV
- XLS
- XLSX
- XLSM
- XML Spreadsheet
- HTML / HTM table files

The app can detect employee-style columns such as Candidate Name and Colleague #, then place clean numeric IDs into the queue.

## Clipboard Cleanup

Use Paste Clipboard To Queue when copying data from Excel or another table.

If the copied content looks messy, the app shows a Paste Cleanup Preview. Choose Add Clean IDs to add only the extracted numeric ID values.

## Batch Converter

Batch Converter converts queued HEX IDs into Facility Code and Card Number.

Controls:

| Control | What It Does |
| --- | --- |
| Import | Browse for supported files. |
| Sample | Loads numeric sample IDs. |
| Convert | Converts all queued rows. |
| Remove Duplicates | Keeps the first valid matching ID and removes repeats. |
| Keep Valid | Removes rows that cannot be read as valid HEX IDs. |
| Clear | Clears the workspace. |

## Results

The Results table shows:

- Line
- Hex ID
- FC
- CN
- Status
- Notes / Details

Status meanings:

| Status | Meaning |
| --- | --- |
| Valid | Row converted normally. |
| Warning | Row converted, but the app noticed cleanup, duplicate, or unusual data. |
| Invalid | Row could not be converted. |

Notes / Details stays blank for clean rows. It appears when the app cleaned a row, found a duplicate, flagged an unusual value, or explains why input is invalid.

## Single Hex Lookup

Use Single Hex Lookup when checking one HEX ID.

The app converts it and copies the FC,CN pair to the clipboard.

## FC/CN To Hex

Use FC/CN to Hex to build one 8-character HEX value from a Facility Code and Card Number.

Both values must be whole numbers from 0 to 65535.

## Unconvert Batch

Use Unconvert Batch when you have multiple FC/CN pairs.

Accepted examples:

```text
34968,18199
34968 18277
FC 34968 CN 18199
```

## Exporting Reports

Export options:

- Excel Workbook
- CSV Report
- TXT Report
- PDF Report

After export, the Export Complete window lets you open the file or open the folder.

## Settings

Settings controls:

- Default export type
- Default export folder
- Desktop shortcut creation
- Clear Recent Exports list

## BlueWave

The BlueWave button opens the access-control site in the default browser.

## Help And About

Help shows a visual guide inside the app.

About shows app purpose, version, GitHub project, and contact/credit information.
