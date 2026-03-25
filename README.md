# Campaign P360 Generator

Small webapp to split large Install and InApp reports into per-campaign Excel files.

## Supports

- `.xlsx` input
- `.csv` input
- Large files by processing uploads from disk instead of holding everything in memory
- Per-campaign output files like `EasyPaisa - 1210 - p360.xlsx`
- `PA - Install` and `PA - InApp` sheets
- Install-only output when a campaign is missing from the InApp file

## Run

```bash
npm install
npm start
```

Open `http://localhost:3000`.

## Input rules

- Upload one Install file and one InApp file
- Both files must contain a `Campaign` column
- Campaign numbers are extracted from the last numeric part of the campaign value
  - Example: `mobisaturn_1216` -> `1216`

## Output

- Downloaded ZIP: `Campaign - p360.zip`
- Inside the ZIP: one `.xlsx` file per campaign ID
