# Sharp Titan TV. AQ. XML-XLS converter

## Version 2.x

Version 2.2 is published as a separate web application in `v2/`.

- Uses the internal `MAP2.0.xlsx` project data embedded in the app.
- Requires only the Dolby Tuning Tool XML input file.
- Reads only the `internal_speaker` endpoint and supported profiles: Movie, Music, Voice, and User Selectable.
- Creates a new XLSX workbook for manual Copy/Paste and a downloadable LOG workbook.
- Keeps the Version 1.x application available at the repository root.

## Version 1.x

Version 1.1e web application for updating TitanOS XLS audio settings with values extracted from a Dolby Tuning Tool XML file.

## How to run

Open `index.html` in a browser. The app is client-side only and keeps files on the local machine.

The page loads SheetJS from the official CDN to read and write `.xls`, `.xlsx`, `.xlsm`, and `.csv` workbooks. An internet connection is required unless the library is vendored locally later.

## Conversion rules

- Required inputs: Dolby XML file, input XLS workbook, MAP workbook, and a non-empty version number.
- XLS data is read and updated in worksheet `AQ_Tbl`.
- XLS column headers are read from Excel line 25.
- XLS data rows are read from Excel row 26 through row 4257.
- MAP data is read from worksheet `MAP`; its header row is detected automatically in the first 100 rows and must contain `DSP`, `Description`, `Parameters`, and `Mapped to DAP XML`.
- The target XLS workbook line 25 must contain `DSP`, `Description`, and `Parameters`.
- XML profile names are cross-linked to MAP DSP mode names: Movie to Entertainment Custom mode, Music to Music Custom mode, Voice to Dialog Custom mode, and User Selectable to Personal Custom mode. Game and Night are not used. Duplicate XML items are matched using the mapped XML item name plus DSP mode context.
- Rows with empty `Mapped to DAP XML` or `Not relevant` are skipped.
- The converter updates only the selected TV model GAIN columns: F, I, L, and/or O.
- The converter applies the AQ settings version to cell `B2` in the output workbook.
- The converted workbook preserves formatting for `.xlsx/.xlsm` files when supported; legacy `.xls` files fall back to value conversion and may lose formatting.
- The log can be downloaded in XLS format and contains DSP, Description, Parameters, TV model, GAIN column, previous value, new value, and XML mapping for each modified cell.
- The log includes a skipped rows section with reason details.
- The UI includes TV model selection, an optional shared 32/40/43 setting, and manual copy/paste output.
- The UI includes a bottom conversion progress indicator and a bottom application error, status, and debug message window.

If the XML contains repeated item names, the converter requires all relevant names to match: the mapped XML item name, `DSP`, and `Description`. Empty cells and cells marked `Not relevant` are excluded from XML matching.
