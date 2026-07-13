# Macy's AP China Grove Hex Utility v1.0.7

This release focuses on making the desktop app easier to see, easier to use, and cleaner when launched as a Windows EXE.

## Highlights

- Cleaner Macy's AP China Grove app icon for the EXE, taskbar, and custom popup windows.
- Main toolbar now includes Salesforce and BlueWave shortcuts.
- BlueWave links were removed from Help/About so shortcuts are not repeated in multiple places.
- GitHub links now open the project repository.
- First-open status text no longer looks like duplicate Ready messages.
- Search, input queues, Single Lookup, and FC/CN fields have stronger borders and clearer text.
- Results and Unconvert Results use larger rows and clearer valid/warning/invalid highlights.
- Single Lookup and FC/CN to Hex output now displays as easy-to-read value lines.
- Import cleanup better handles Excel-style numeric IDs such as `88984765.0`, split IDs like `8898-4765`, and full employee lines.

## Validation

- Source compile check.
- Desktop smoke test.
- Built-in conversion/import self-test.
- Windows EXE rebuild with SHA256 checksum.
