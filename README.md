# AP Access Control Converter

Desktop version of the AP access control HEX to FC/CN converter.

## Commands

- `npm start` launches the app.
- `npm test` runs converter smoke checks.
- `npm run build` creates the portable Windows `.exe` in `dist`.

## Notes

- Import supports TXT, CSV, TSV, XLS, XLSX, XLSM, XML Spreadsheet, and HTML table files.
- Export creates formatted Excel XML (`.xls`) and plain text reports.
- Conversion uses the high 16 bits for Facility Code and low 16 bits for Card Number.
