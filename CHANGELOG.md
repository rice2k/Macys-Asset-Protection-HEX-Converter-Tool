# Changelog

## 1.1.1

- Added Batch Scanner Input so keyboard-style handheld scanners can place each new scan at the top of the Input Queue.
- Added scanner-ready Single Lookup behavior so a scanned HEX ID can auto-convert and copy the FC,CN pair.
- Redesigned Single Lookup, FC/CN to Hex, Help, and About for cleaner spacing, clearer result cards, and more professional formatting.
- Made the sidebar status card and bottom status strip clickable so Needs Review messages can jump back to the related workspace.
- Polished dropdown and right-click menu styling with larger text and clearer hover states.
- Added right-click edit menus for Batch Converter, Unconvert Batch, search, Single Lookup, FC, CN, and Settings input fields.
- Added drag-to-resize handles for the Batch and Unconvert input boxes so long pasted data is easier to review.
- Tightened import cleanup so Batch imports only detected 8-character numeric IDs and Unconvert imports only detected FC/CN pairs.
- Scoped exports to the two batch workspaces and added Unconvert Batch report export support.
- Kept GitHub and Christopher Schumacher contact links on the About page only, with cleaner Help/footer wording.
- Rebuilt the Windows EXE and checksum after the input, import, export, and layout polish.
- Improved Results and Unconvert Results column alignment so headers and row values line up more clearly.
- Renamed the table display header to Notes / Details and left-aligned it with the note text.
- Added in-app Help and README wording explaining where Notes / Details messages come from.
- Updated CSV, Excel, PDF, and TXT report headers to use Notes / Details consistently.

## 1.1.0

- Added cleaned clipboard preview for messy Excel-style pasted data so only numeric ID rows can be added to the queue.
- Added live Input Queue row count while typing, pasting, importing, or clearing data.
- Added Copy Row plus right-click copy menus for Results and Unconvert Results tables.
- Added a small import working indicator for file reads and larger imports.
- Added Settings support for clearing saved recent export shortcuts without deleting exported files.
- Added release-note template guidance and a required matching release-notes file for future GitHub tags.
- Kept v1.0.9 as the almost-done baseline restore point; this version builds workflow polish on top of it.

## 1.0.9

- Removed the Salesforce toolbar shortcut and unused Salesforce icon asset.
- Removed unused legacy visual assets from earlier layouts.
- Cleaned the main toolbar so BlueWave and Help align cleanly on the right.
- Improved compact-window alignment so summary panels do not get clipped by action buttons.
- Performed a final visual polish pass on the main workspace, dialogs, status strip, and screenshots.
- Rebuilt the Windows EXE and checksum from the cleaned project state.

## 1.0.8

- Changed built-in samples to numeric-only data so Sample buttons do not insert employee text or bad-line examples.
- Reduced Input Queue and Unconvert Queue text size for easier scanning.
- Shortened Single Lookup and FC/CN input fields so they no longer stretch across the workspace.
- Stabilized the sidebar and footer status areas so button clicks do not shift the workspace/status layout.
- Removed a duplicate search-entry layout call that could make the Results filter row behave inconsistently.

## 1.0.7

- Updated the app and EXE icon to a cleaner Macy's AP China Grove badge style.
- Kept BlueWave as a main-app shortcut only.
- Reworked first-open status text so the bottom strip no longer reads like duplicate Ready states.
- Improved visibility for the search box, input queues, single lookup fields, and FC/CN fields.
- Made Results and Unconvert Results easier to read with larger rows and valid/warning/invalid row colors.
- Improved Single Lookup and FC/CN to Hex result text so each value is clearly separated.
- Applied the app icon to custom popup windows.
- Updated Help/About links so GitHub opens the project repository and BlueWave is not repeated in popups.
- Improved examples and Excel-style import cleanup for numeric IDs, split IDs, and full employee lines.

## 1.0.6

- Added a clearer bottom status strip with Ready, Needs Review, and Exported states.
- Added recent exports so saved reports can be reopened from the Export menu.
- Added copyable error reports for import, export, and open-file failures.
- Added table header tooltips explaining FC, CN, status, and related columns.
- Added subtle row hover highlighting in Results, Unconvert Results, and History tables.
- Added Macy's-red footer accent lines to app popups.
- Refreshed the application icon so it reads cleaner at small taskbar sizes.
- Refreshed the top banner artwork to better match the light corporate utility layout.
- Added README screenshot placeholders and release documentation for the next GitHub update.

## 1.0.5

- Polished Macy's-style application UI.
- Added the current light application layout, custom icons, improved Help/About windows, and Export Complete dialog.
- Added input validation colors, Remove Duplicates, Keep Valid, default export settings, and desktop shortcut support.
- Synced the polished desktop app source, EXE, checksum, and GitHub project files.

## 1.0.4

- Earlier GitHub build of the HEX conversion utility.
- Included Windows build output and the initial project structure used before the polished v1.0.5 app sync.

## Future

- Publish GitHub Releases with EXE downloads, checksums, notes, and restore points.
- Continue improving screenshots, help content, and visual consistency as the app evolves.
