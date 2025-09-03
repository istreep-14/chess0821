### Google Sheets daily roll and Game Data helpers

This Apps Script automates your daily sheet roll at local midnight, adds computed columns to `Game Data`, and provides a `RATING_CHANGE` function.

### What it does

- Inserts a new daily row at the top of `Daily` at 12:00 am local time with the date in column `A`
- Copies the prior top row's live formulas into the new row
- Flushes calculations and converts all rows below the new date into static values
- Adds/maintains in `Game Data` the following columns:
  - **Date**: local date floor of `End Time` (no time)
  - **ResultBinary**: 1 for win, 0.5 for draw, 0 for loss (robust text match)
  - **RatingChange**: your rating for that game minus the most recent prior rating within the same `Format`

### Files

- `google-apps-script/Code.gs`: Main script
- `google-apps-script/appsscript.json`: Manifest

### Setup

1. Open your Google Sheet.
2. Extensions → Apps Script.
3. In the script editor, create the same files and paste contents from this repo:
   - `Code.gs`
   - `appsscript.json` (replace existing manifest)
4. Adjust constants at the top of `Code.gs` if needed:
   - `DAILY_SHEET_NAME` (default `Daily`)
   - `GAME_DATA_SHEET_NAME` (default `Game Data`)
   - `DAILY_HEADER_ROW` (default `1`)
   - `DAILY_DATE_COLUMN_INDEX` (default `1` for column A)
5. Make sure `Game Data` has headers named exactly: `End Time`, `Result`, `Format`, and `Rating`.
6. In Apps Script, run `ensureGameDataComputedColumns` once (Authorize when prompted).
7. Run `installDailyTrigger` to schedule the midnight automation (uses spreadsheet time zone).

### Notes

- If today's row already exists at the top of `Daily`, the script is idempotent and will not duplicate it.
- `ResultBinary` recognizes common result text. If your sheet uses different labels, adjust the regex in `ensureGameDataComputedColumns`.
- `RATING_CHANGE` should be placed in row 2 of the `RatingChange` column by the setup function. It returns a column of deltas aligned with your data rows.
- The Date and ResultBinary columns use a single `ARRAYFORMULA` in row 1 to populate the entire column (including the header cell).

### Manual use

- Use the custom menu “Daily Automation” to:
  - Run the daily roll immediately
  - Install or remove the midnight trigger
  - Re-setup the `Game Data` helper columns

