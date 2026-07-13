# Troubleshooting

## The EXE Download Shows A Windows Warning

Reason: the EXE is not code-signed yet.

For testing:

1. Choose More info.
2. Choose Run anyway.

This is a Windows trust warning. It does not stop testing when Run anyway is available.

## GitHub Actions Are Not Running

This means the hosted build/release workflow is not currently producing a new automatic release.

It does not stop users from downloading the EXE directly from the repository.

Direct download:

[Macys_AP_China_Grove_Hex_Utility.exe](https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/raw/main/dist/Macys_AP_China_Grove_Hex_Utility.exe)

## Import Does Not Find IDs

Try these checks:

- Confirm the file contains 8-character HEX values or 8-digit IDs.
- Confirm the file is one of the supported types.
- Open the file in Excel or Notepad to confirm the data is visible.
- Copy the ID column manually and use Paste Clipboard To Queue.
- Use Keep Valid after import to remove bad rows.

## Pasted Excel Data Looks Messy

Use Import > Paste Clipboard To Queue.

If the copied data looks like a table, the app should show Paste Cleanup Preview. Choose Add Clean IDs to add only the detected ID values.

## A Row Is Invalid

Common reasons:

- Fewer than 8 characters.
- More than 8 characters without a clean 8-character token.
- Invalid HEX characters.
- Empty line.
- Mixed text with no usable ID.

## Notes / Details Is Blank

Blank is normal for clean rows.

Notes / Details only appears when the app has something to explain:

- Cleaned Excel numeric ID.
- Extracted ID from full text.
- Joined split ID.
- Duplicate.
- Unusual value.
- Invalid row reason.

## Export Says There Is No Report

Run a conversion first. Exports require rows in the Results table.

## Export Fails

Use Help > Copy Last Error Report.

The copied report includes:

- App version.
- Time.
- Error type.
- File path.
- Error details.

## Recent Export Does Not Open

The file may have been moved or deleted.

Use Settings > Clear Recent Exports to clear old shortcuts without deleting any exported files.

## Drag And Drop Does Not Work

Use Import > Browse Files instead.

Drag/drop depends on optional Windows/Tkinter drag/drop support.

## BlueWave Does Not Open

The BlueWave button opens the configured internal URL in the default browser. If it does not load, confirm the workstation has access to that internal site.
