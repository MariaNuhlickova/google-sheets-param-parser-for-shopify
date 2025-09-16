# Google Sheets Param Parser (Apps Script)

A small sidebar app for Google Sheets that parses URL query parameters from any column and builds a clean table on a new sheet. Works great with GA4 Exploration exports (e.g., header row at 7, “Landing page + query string”).

## Features
- Choose source sheet and header row
- Select the URL column header
- Keep any “carry” columns (e.g., `Event name`, `Event count`)
- Extract arbitrary parameters (`locale`, `surface_type`, `utm_*`, etc.)
- Clean output with filter, frozen header, banding, and auto-resized columns

## Install
1. In your Google Sheet: **Extensions → Apps Script**.
2. Create the project and add:
   - `src/Code.gs` (copy contents from this repo)
   - `src/Sidebar.html`
   - `appsscript.json` (overwrite manifest)
3. Save. Reload the sheet.

You’ll see a new menu: **Param Parser → Open parser**.

## Usage
1. Open the sidebar.
2. Pick the source sheet, set **Header row** (e.g., `7` for GA4 exports).
3. Set **URL column header** (default: `Landing page + query string`).
4. Set **Carry columns** (comma-separated, default: `Event name, Event count`).
5. Set **Params to extract**.
6. Click **Build table**. A new sheet `<source> - parsed` will be created.

## Notes
- If a parameter appears multiple times, the first occurrence wins.
- Numeric-looking values (e.g., `2`, `18`) are written as numbers for easier filtering/summing.
- You can safely re-run; the output sheet is overwritten each run.

## License
MIT
