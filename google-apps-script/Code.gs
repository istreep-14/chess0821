/**
 * Spreadsheet automation for daily rollup and Game Data computed columns.
 * - Inserts a new daily row at midnight local time with today's date in column A
 * - Copies live formulas from the previous top row into the new row
 * - Freezes all rows below the new date by converting formulas to values
 * - Ensures Game Data tab has Date, ResultBinary, and RatingChange columns
 * - Provides custom function RATING_CHANGE(formatCol, ratingCol, endTimeCol)
 */

const DAILY_SHEET_NAME = 'Daily';
const GAME_DATA_SHEET_NAME = 'Game Data';
const DAILY_HEADER_ROW = 1; // Row number of headers in Daily sheet
const DAILY_DATE_COLUMN_INDEX = 1; // Column A contains the Date in Daily sheet

/**
 * Add custom menu for manual actions.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Daily Automation')
    .addItem('Run Daily Roll (now)', 'dailyRoll')
    .addSeparator()
    .addItem('Install Midnight Trigger', 'installDailyTrigger')
    .addItem('Remove Midnight Triggers', 'removeDailyTriggers')
    .addSeparator()
    .addItem('Setup Game Data Columns', 'ensureGameDataComputedColumns')
    .addToUi();
}

/**
 * Create a daily time-based trigger to run at 00:00 in the spreadsheet time zone.
 */
function installDailyTrigger() {
  removeDailyTriggers();
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  ScriptApp.newTrigger('dailyRoll')
    .timeBased()
    .inTimezone(tz)
    .everyDays(1)
    .atHour(0)
    .create();
}

/**
 * Remove any existing triggers that call dailyRoll, to avoid duplicates.
 */
function removeDailyTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((t) => {
    if (t.getHandlerFunction && t.getHandlerFunction() === 'dailyRoll') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

/**
 * Main job: insert today's daily row at the top, copy formulas, update values.
 */
function dailyRoll() {
  ensureGameDataComputedColumns();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const daily = ss.getSheetByName(DAILY_SHEET_NAME);
  if (!daily) {
    throw new Error(`Sheet "${DAILY_SHEET_NAME}" not found. Create it or adjust DAILY_SHEET_NAME.`);
  }

  const tz = ss.getSpreadsheetTimeZone();
  const todayString = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const headerRow = DAILY_HEADER_ROW;
  const dateCol = DAILY_DATE_COLUMN_INDEX;

  // Ensure there is at least a header and one data row to copy formulas from.
  const lastRow = daily.getLastRow();
  const lastCol = daily.getLastColumn();
  if (lastRow < headerRow + 1) {
    // Not enough rows to copy formulas; just insert below header.
    daily.insertRowAfter(headerRow);
    const dateCell = daily.getRange(headerRow + 1, dateCol);
    dateCell.setValue(todayString);
    dateCell.setNumberFormat('yyyy-mm-dd');
    SpreadsheetApp.flush();
    return;
  }

  // If the top data row already has today's date, exit idempotently.
  const topDate = daily.getRange(headerRow + 1, dateCol).getDisplayValue();
  if (topDate === todayString) {
    return;
  }

  // Insert new row for today just below header
  daily.insertRowAfter(headerRow);

  // Set today's date (no time) in the date column
  const dateCell = daily.getRange(headerRow + 1, dateCol);
  dateCell.setValue(todayString);
  dateCell.setNumberFormat('yyyy-mm-dd');

  // Copy formulas from what is now row (headerRow + 2) into the new row, starting from column 2 to lastCol
  const sourceRowIndex = headerRow + 2; // Previously top row before insertion
  if (lastCol > 1) {
    const sourceFormulaRange = daily.getRange(sourceRowIndex, 2, 1, lastCol - 1);
    const formulasR1C1 = sourceFormulaRange.getFormulasR1C1();
    if (formulasR1C1 && formulasR1C1.length > 0) {
      daily.getRange(headerRow + 1, 2, 1, lastCol - 1).setFormulasR1C1(formulasR1C1);
    } else {
      // If no formulas present, copy values as a fallback
      const values = daily.getRange(sourceRowIndex, 2, 1, lastCol - 1).getValues();
      daily.getRange(headerRow + 1, 2, 1, lastCol - 1).setValues(values);
    }
  }

  // Force calculation and run quick update hook if available
  SpreadsheetApp.flush();
  tryQuickUpdate();

  const newLastRow = daily.getLastRow();
  if (newLastRow >= headerRow + 2) {
    const freezeStartRow = headerRow + 2; // Rows below the newly inserted row
    const numRowsToFreeze = newLastRow - freezeStartRow + 1;
    if (numRowsToFreeze > 0) {
      const range = daily.getRange(freezeStartRow, 1, numRowsToFreeze, lastCol);
      const values = range.getValues();
      range.setValues(values);
    }
  }
}

/**
 * Ensure Game Data tab has helper columns and formulas:
 * - Date: INT(End Time) for local date (rounded down)
 * - ResultBinary: 1 win, 0.5 draw, 0 loss
 * - RatingChange: current rating minus most recent prior rating in same Format
 */
function ensureGameDataComputedColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(GAME_DATA_SHEET_NAME);
  if (!sheet) {
    throw new Error(`Sheet "${GAME_DATA_SHEET_NAME}" not found. Create it or adjust GAME_DATA_SHEET_NAME.`);
  }

  const headerRow = 1;
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(headerRow, 1, 1, Math.max(1, lastCol)).getValues()[0];

  const getColIndexByHeader = (name) => headers.findIndex(h => String(h).trim().toLowerCase() === name.toLowerCase()) + 1;
  const endTimeCol = getColIndexByHeader('End Time');
  const resultCol = getColIndexByHeader('Result');
  const formatCol = getColIndexByHeader('Format');
  const ratingCol = getColIndexByHeader('Rating');

  // Create or find computed columns
  const dateColIdx = findOrCreateColumn(sheet, headers, 'Date');
  const resultBinColIdx = findOrCreateColumn(sheet, headers, 'ResultBinary');
  const ratingChangeColIdx = findOrCreateColumn(sheet, headers, 'RatingChange');

  // Re-read headers after potential insertions
  const newHeaders = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const currentHeaderText = (idx) => String(newHeaders[idx - 1] || '').trim();

  // Date formula (requires End Time)
  if (endTimeCol) {
    const endColLetter = columnToLetter(endTimeCol);
    const dateFormula = `=ARRAYFORMULA({"Date"; IF(LEN(${endColLetter}2:${endColLetter})=0, , INT(${endColLetter}2:${endColLetter}))})`;
    sheet.getRange(1, dateColIdx).setFormula(dateFormula);
  }

  // ResultBinary formula (requires Result)
  if (resultCol) {
    const resLetter = columnToLetter(resultCol);
    const resultBinFormula = `=ARRAYFORMULA({"ResultBinary"; IF(LEN(${resLetter}2:${resLetter})=0, , IF(REGEXMATCH(LOWER(${resLetter}2:${resLetter}), "win"), 1, IF(REGEXMATCH(LOWER(${resLetter}2:${resLetter}), "draw|1/2"), 0.5, IF(REGEXMATCH(LOWER(${resLetter}2:${resLetter}), "loss|lose"), 0, ""))))})`;
    sheet.getRange(1, resultBinColIdx).setFormula(resultBinFormula);
  }

  // RatingChange header and formula (requires Format, Rating, End Time)
  if (formatCol && ratingCol && endTimeCol) {
    if (currentHeaderText(ratingChangeColIdx) !== 'RatingChange') {
      sheet.getRange(1, ratingChangeColIdx).setValue('RatingChange');
    }
    const fmtLetter = columnToLetter(formatCol);
    const ratLetter = columnToLetter(ratingCol);
    const endLetter = columnToLetter(endTimeCol);
    const ratingChangeFormula = `=RATING_CHANGE(${fmtLetter}2:${fmtLetter}, ${ratLetter}2:${ratLetter}, ${endLetter}2:${endLetter})`;
    sheet.getRange(2, ratingChangeColIdx).setFormula(ratingChangeFormula);
  }
}

/**
 * Custom function: rating change per game grouped by Format, ordered by End Time.
 * Usage (place in row 2 of the RatingChange column):
 *   =RATING_CHANGE(Format2:Format, Rating2:Rating, EndTime2:EndTime)
 * Returns a single column array of deltas, same height as input ranges.
 * First row in the provided ranges is treated as data (no header expected).
 * If a game has no prior game in the same Format, its delta is 0.
 *
 * @param {any[][]} formatCol
 * @param {any[][]} ratingCol
 * @param {any[][]} endTimeCol
 * @return {any[][]}
 */
function RATING_CHANGE(formatCol, ratingCol, endTimeCol) {
  const toFlat = (arr) => {
    if (!Array.isArray(arr)) return [];
    if (Array.isArray(arr[0])) return arr.map(r => r[0]);
    return arr;
  };
  const fmt = toFlat(formatCol);
  const rat = toFlat(ratingCol).map(v => (v === '' || v == null ? null : Number(v)));
  const end = toFlat(endTimeCol).map(v => v instanceof Date ? v : (v ? new Date(v) : null));

  const n = Math.max(fmt.length, rat.length, end.length);
  const rows = [];
  for (let i = 0; i < n; i++) {
    rows.push({
      idx: i,
      format: fmt[i] == null ? '' : String(fmt[i]),
      rating: rat[i] == null || isNaN(rat[i]) ? null : rat[i],
      endTime: end[i] instanceof Date && !isNaN(end[i].getTime()) ? end[i] : null,
    });
  }

  // Filter rows that have sufficient data
  const valid = rows.map((r, i) => ({ ...r, origIndex: i }))
    .filter(r => r.format !== '' && r.rating != null && r.endTime != null);

  // Sort by endTime ascending, then by original index to stabilize
  valid.sort((a, b) => a.endTime - b.endTime || a.origIndex - b.origIndex);

  // Compute deltas grouped by format
  const lastByFormat = new Map();
  const deltaByOrigIndex = new Map();
  for (const r of valid) {
    const key = r.format;
    const prev = lastByFormat.get(key);
    const delta = prev == null ? 0 : r.rating - prev;
    deltaByOrigIndex.set(r.origIndex, delta);
    lastByFormat.set(key, r.rating);
  }

  // Build output aligned to original length; blank for rows lacking data
  const out = new Array(n).fill('');
  for (let i = 0; i < n; i++) {
    if (deltaByOrigIndex.has(i)) {
      out[i] = deltaByOrigIndex.get(i);
    } else {
      out[i] = '';
    }
  }

  // Return as 2D column array
  return out.map(v => [v]);
}

/**
 * Helpers
 */
function findOrCreateColumn(sheet, headersRowValues, headerName) {
  const idx = headersRowValues.findIndex(h => String(h).trim().toLowerCase() === headerName.toLowerCase());
  if (idx >= 0) {
    return idx + 1;
  }
  const lastCol = sheet.getLastColumn();
  const insertAt = lastCol + 1;
  sheet.getRange(1, insertAt).setValue(headerName);
  return insertAt;
}

function columnToLetter(column) {
  let temp = '';
  let col = column;
  while (col > 0) {
    let rem = (col - 1) % 26;
    temp = String.fromCharCode(rem + 65) + temp;
    col = Math.floor((col - rem) / 26);
  }
  return temp;
}

/**
 * If a global function named quickUpdate exists, call it to refresh external data.
 * This allows the daily roll to also trigger any user-defined fetch/update logic.
 */
function tryQuickUpdate() {
  try {
    // Using typeof avoids ReferenceError if quickUpdate is not declared
    if (typeof quickUpdate === 'function') {
      quickUpdate();
      SpreadsheetApp.flush();
    }
  } catch (err) {
    // Swallow errors to keep daily roll resilient
    try {
      console && console.warn && console.warn('quickUpdate failed:', err);
    } catch (e) {}
  }
}

