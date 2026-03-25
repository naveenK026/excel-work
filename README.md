# 1.orbit

Small webapp with two spreadsheet tools:

- Campaign P360 Generator
- AppsFlyer Duplicate Highlighter

## Supports

- `.xlsx` input
- `.csv` input
- Large-file friendly processing for the Campaign P360 tool
- Per-campaign output files like `EasyPaisa - 1210 - p360.xlsx`
- `PA - Install` and `PA - InApp` sheets
- Install-only output when a campaign is missing from the InApp file
- Sorted and highlighted AppsFlyer duplicate rows in a new `.xlsx` file

## Run

```bash
npm install
npm start
```

Open `http://localhost:3000`.

## Tools

### Campaign P360 Generator

- Upload one Install file and one InApp file
- Both files must contain a `Campaign` column
- Campaign numbers are extracted from the last numeric part of the campaign value
  - Example: `mobisaturn_1216` -> `1216`
- Downloaded ZIP: `Campaign - p360.zip`
- Inside the ZIP: one `.xlsx` file per campaign ID

### AppsFlyer Duplicate Highlighter

- Upload one source file
- The file must contain an `AppsFlyer ID` column
- The tool sorts rows by `AppsFlyer ID`
- Rows with duplicate `AppsFlyer ID` values are highlighted in the output workbook
