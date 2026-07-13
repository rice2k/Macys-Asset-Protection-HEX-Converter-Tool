Macy's Asset Protection - China Grove Hex Converter Utility
README

App Version: 1.1.1

Purpose
-------
The Macy's Asset Protection - China Grove Hex Converter Utility is a desktop
application built to help Asset Protection users convert access-control badge
or card identifiers between HEX, Facility Code, and Card Number formats.

The app is intended for quick operational review, cleanup, copying, and report
exporting. It is especially useful when working with lists of access-control
IDs from text files, spreadsheets, copied employee lines, or exported table
data.

Built For
---------
Macy's Asset Protection operations at the China Grove, North Carolina facility.

Built By
--------
Christopher Schumacher, Asset Protection FLO
Email: christopher.schumacher@macys.com
GitHub project:
https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool

Download Test Build
-------------------
Latest test EXE:
https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/raw/main/dist/Macys_AP_China_Grove_Hex_Utility.exe

Checksum file:
https://raw.githubusercontent.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/main/dist/Macys_AP_China_Grove_Hex_Utility.exe.sha256.txt

Current SHA-256:
15F7DA4D211292EC21002E760655EFEFBB3628F86A6EA9E9A10E8C04B46F3815

Windows may show a SmartScreen warning because the EXE is not code-signed yet.
For testing, users can choose More info, then Run anyway.

Source History
--------------
The original Access Control script document and earlier HTML versions are
preserved in:

docs\source-history

Conversion rule:
Facility Code (FC) is taken from the high 16 bits. Card Number (CN) is taken
from the low 16 bits of the 32-bit HEX value.

Documentation Hub
-----------------
Detailed GitHub documentation is organized in:

docs\README.md

Detailed notes are organized in:

docs\notes

Important documentation pages:
- docs\notes\user-guide.md
- docs\notes\feature-reference.md
- docs\notes\every-option-reference.md
- docs\notes\input-examples.md
- docs\notes\excel-clipboard-tips.md
- docs\notes\keyboard-shortcuts.md
- docs\notes\version-history.md
- docs\notes\restore-older-versions.md
- docs\notes\downloads-and-releases.md
- docs\notes\release-checklist.md
- docs\notes\roadmap-and-known-issues.md
- docs\notes\troubleshooting.md

Screenshot guide:
docs\screenshots\README.md

GitHub visual assets:
docs\images\README.md

Project tags:
access-control, asset-protection, hex-converter, facility-code, card-number,
windows-desktop, python, tkinter, excel-import, report-export, macy-style-ui,
china-grove

Main Functions
--------------
1. Batch Converter
   Converts one or many 8-character HEX IDs into:
   - Facility Code
   - Card Number
   - Status
   - Notes or warnings

2. Single Lookup
   Converts one HEX ID quickly and copies the FC,CN pair to the clipboard.

3. FC/CN to Hex
   Converts one Facility Code and Card Number pair back into an 8-character
   HEX ID.

4. Unconvert Batch
   Converts many FC/CN pairs back into HEX IDs.

5. History
   Shows recent conversion activity saved by the utility.

How The Conversion Works
------------------------
For an 8-character HEX value, the app reads the value as a 32-bit number:

- Facility Code = high 16 bits
- Card Number = low 16 bits

Example:
HEX: 88984717
Facility Code: 34968
Card Number: 18199

Supported Input
---------------
You can add data to the Input Queue several ways:

1. Paste IDs directly into the Input Queue.
2. Use Import > Browse Files.
3. Use Import > Paste Clipboard To Queue.
4. Drag supported files directly onto the Input Queue box.
5. Use Sample to load example data.

Excel and copied-table cleanup:
- Numeric spreadsheet cells such as 88984765.0 are cleaned to 88984765.
- Split IDs such as 8898-4765 or 8898 4765 are joined before conversion.
- Full employee lines are scanned for the first clean 8-digit ID.
- Messy Excel-style clipboard data shows a cleanup preview before it is added.
- Structured spreadsheet imports place clean numeric IDs into the queue.

Supported file types:
- TXT
- CSV
- TSV
- XLS
- XLSX
- XLSM
- XML Spreadsheet
- HTML / HTM table files

The app can also pull an 8-character ID out of a full employee-style line when
possible. For example:

Active Test User, 88984717

Input Queue Highlights
----------------------
The Input Queue uses row colors to make review easier before conversion:

- Green-tinted rows are clean valid HEX IDs.
- Yellow-tinted rows are valid, but include a warning such as extracted text,
  a duplicate ID, or an unusual value.
- Red-tinted rows are invalid and need review or cleanup.

The Input Queue also shows a live row count while typing, pasting, importing,
or clearing data.

Queue Cleanup Tools
-------------------
The Batch Converter includes cleanup buttons next to Convert:

- Remove Duplicates keeps the first matching valid HEX ID and removes later
  repeated valid IDs.
- Keep Valid removes rows that cannot be read as valid 8-character HEX IDs.

These tools work on the Input Queue before conversion.

Results And Warnings
--------------------
After conversion, the Results table shows:

- Line number
- HEX ID or raw input
- Facility Code
- Card Number
- Status
- Notes / Details

Statuses include:
- Valid: The row converted normally.
- Warning: The row converted, but the app noticed something unusual.
- Invalid: The row could not be converted.

Notes / Details stays blank for clean rows. It fills in when the app has a
reason to explain something about that row, such as cleanup, duplicates,
unusual values, or invalid input.

Examples of possible warnings or notes:
- Extracted an ID from a full text line
- Duplicate value found
- All zeros
- High or unusual value
- Too short or too long
- Invalid HEX character

Copy Tools
----------
The Results area includes tools for copying:

- All valid FC/CN pairs
- Selected Facility Code
- Selected Card Number
- Selected FC,CN pair
- Selected full result row
- Right-click copy actions for full row, HEX/raw, notes, FC, CN, and pair

The full right-click and menu option list is documented in:
docs\notes\every-option-reference.md

The Unconvert Batch area includes tools for copying HEX results.

Export Options
--------------
The Export menu creates professional reports in several formats:

1. Excel Workbook (.xlsx)
   Includes:
   - Summary sheet
   - Results sheet
   - Frozen header row
   - Filters
   - Status colors
   - Wrapped notes

2. CSV Report (.csv)
   Good for importing into other tools or spreadsheet programs.

3. TXT Report (.txt)
   Plain text report that opens easily in Notepad.

4. PDF Report (.pdf)
   Formatted report with:
   - Title
   - Summary section
   - Results table
   - Row striping
   - Warning and invalid highlighting

After an export finishes, the app shows an Export Complete window with:

- Open File
- Open Folder
- Saved file path
- App version and run summary

Recent Exports
--------------
Use Export > Recent Exports to reopen recently saved reports from this
workstation. The list keeps the newest reports first and skips files that no
longer exist.

Use File > Settings > Clear Recent Exports to clear the saved shortcuts without
deleting any exported files.

Default Export Folder
---------------------
Use File > Default Export Folder to choose where reports should start saving.
Use File > Open Export Folder to open that location.

Settings
--------
Use File > Settings to control:

- Default export type
- Default export folder
- Desktop shortcut creation
- Recent export list cleanup

Use Export > Export Default to export using the saved default report type.

Desktop Shortcut
----------------
Use File > Create Desktop Shortcut or File > Settings > Create Shortcut to add
a launcher to the Windows desktop.

BlueWave Link
-------------
Use the BlueWave button in the top bar to open the access-control site in your
browser.

Help And About
--------------
Use Help > How To Use for a visual guide inside the app.
Use Help > About for app purpose, contact information, and links.

Mouse And Keyboard
------------------
The app supports mouse-wheel scrolling in:

- Input Queue
- Unconvert input
- Results tables
- History
- Help window
- Import preview

Useful shortcuts:
- Ctrl+I: Import
- Ctrl+R: Convert batch
- Ctrl+E: Export Excel
- Ctrl+Shift+E: Export CSV
- Ctrl+P: Export PDF
- Ctrl+F: Jump to search
- Ctrl+L: Clear workspace

Important Notes
---------------
- The conversion is performed locally on the computer.
- The app does not replace official access-control systems or records.
- Always verify results against the proper system of record when needed.
- Reports should be handled according to company policy.
- The BlueWave and GitHub links open in your default web browser.

Troubleshooting
---------------
If import does not find IDs:
- Confirm the file contains 8-character HEX values.
- Try opening the file in Excel or Notepad to confirm the data is visible.
- Copy the ID column manually and paste it into the Input Queue.

If export says there is no report:
- Run a conversion first.
- Make sure there are rows in the Results table.

If drag and drop does not work:
- Use Import > Browse Files instead.
- Confirm the file type is supported.

If a row is invalid:
- Check that the HEX ID is exactly 8 characters.
- Check that it only contains 0-9 and A-F.

If import, export, or opening a saved report fails:
- Use Help > Copy Last Error Report.
- Paste the copied report into an email or support note so the file path,
  app version, time, and error details are available.

Current Executable
------------------
The built Windows executable is located at:

D:\Macys-Asset-Protection-HEX-Converter-Tool\dist\Macys_AP_China_Grove_Hex_Utility.exe

Release Notes
-------------
Releases use the built Windows EXE and SHA-256 checksum file from the dist
folder. Before creating a future tag, add a matching RELEASE_NOTES_vX.Y.Z.md
file so the GitHub Release has clean notes and restore details.

See CHANGELOG.md for version history.
