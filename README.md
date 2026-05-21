# XML-XLS converter

Version 1.0 web application for updating TitanOS XLS audio settings with values extracted from a Dolby Tuning Tool XML file.

## How to run

Open `index.html` in a browser. The app is client-side only and keeps files on the local machine.

The page loads SheetJS from the official CDN to read and write `.xls`, `.xlsx`, `.xlsm`, and `.csv` workbooks. An internet connection is required unless the library is vendored locally later.

## Conversion rules

- Required inputs: Dolby XML file, input XLS workbook, MAP workbook, and a non-empty version number.
- XLS data is read and updated in worksheet `AQ_Tbl`.
- XLS column headers are read from Excel line 24 or line 25.
- XLS data rows are read from the line immediately after the detected header row.
- MAP data is read from worksheet `MAP`; its header row is detected automatically in the first 100 rows and must contain `DSP`, `Description`, `Parameters`, and `Mapped to DAP XML`.
- The target XLS workbook line 24 or line 25 must contain `DSP`, `Description`, and `Parameters`.
- Rows with empty `Mapped to DAP XML` or `Not relevant` are skipped.
- The converter updates GAIN columns 5, 8, 11, and 14 with the same XML-derived value.
- The log contains DSP, Description, Parameters, target column number, previous value, new value, and XML mapping for each modified cell.
- The UI includes a bottom conversion progress indicator and a bottom application error, status, and debug message window.

If the XML contains repeated item names, the converter requires all relevant names to match: the mapped XML item name, `DSP`, and `Description`. Empty cells and cells marked `Not relevant` are excluded from XML matching.
