// Enhanced Chess.com Game Data Manager
// This script manages Chess.com game data across multiple sheets

// =================================================================
// SCRIPT CONFIGURATION
// =================================================================

const SHEETS = {
  CONFIG: 'Config',
  ARCHIVES: 'Archives',
  GAMES: 'Game Data',
  DAILY: 'Daily Data',
  STATS: 'Player Stats',
  PROFILE: 'Player Profile',
  LOGS: 'Execution Logs',
  OPENINGS_URLS: 'Opening URLS',
  ADD_OPENINGS: 'Add Openings',
  CHESS_ECO: 'Chess ECO',
  TRIGGERS: 'Triggers'
};

const HEADERS = {
  CONFIG: ['Setting', 'Value', 'Description'],
  ARCHIVES: ['Archive URL', 'Year-Month', 'Last Updated'],
  GAMES: [
    'Game URL', 'Time Control', 'Base Time (min)', 'Increment (sec)', 'Rated',
    'Time Class', 'Rules', 'Format', 'End Time', 'Game Duration (sec)',
    'My Rating', 'My Color', 'Opponent', 'Opponent Rating', 'Result',
    'Termination', 'Event', 'Site', 'Date', 'Round', 'Opening', 'ECO',
    'ECO URL', 'UTC Date', 'UTC Time', 'Start Time', 'End Date', 'End Time',
    'Current Position', 'Full PGN', 'Moves', 'Times', 'Moves Per Side',
    'Opening from URL', 'Opening from ECO'
  ],
  DAILY: [
    'Date',
    'Bullet Win', 'Bullet Loss', 'Bullet Draw', 'Bullet Rating', 'Bullet Change', 'Bullet Time',
    'Blitz Win', 'Blitz Loss', 'Blitz Draw', 'Blitz Rating', 'Blitz Change', 'Blitz Time',
    'Rapid Win', 'Rapid Loss', 'Rapid Draw', 'Rapid Rating', 'Rapid Change', 'Rapid Time',
    'Total Games', 'Total Wins', 'Total Losses', 'Total Draws', 'Rating Sum', 'Total Rating Change', 'Total Time (sec)', 'Avg Game Duration (sec)'
  ],
  LOGS: ['Timestamp', 'Function', 'Username', 'Status', 'Execution Time (ms)', 'Notes'],
  STATS: ['Pulled At'],
  PROFILE: ['Field', 'Value']
};

// Functions allowed to be scheduled via Triggers sheet
const TRIGGERABLE_FUNCTIONS = [
  'quickUpdate',
  'completeUpdate',
  'refreshRecentData',
  'updateArchives',
  'updateDailyData'
];

const CONFIG_KEYS = {
  USERNAME: 'username',
  BASE_API: 'base_api_url'
};

const DEFAULT_CONFIG = [
  [CONFIG_KEYS.USERNAME, '', 'Your Chess.com username (lowercase)'],
  [CONFIG_KEYS.BASE_API, 'https://api.chess.com/pub', 'Base API URL']
];

// =================================================================
// MENU & UI
// =================================================================

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Chess.com Manager')
      .addItem('▶️ Quick Update (Recommended)', 'quickUpdate')
      .addItem('🔄 Complete Update (All Games)', 'completeUpdate')
      .addSeparator()
      .addItem('📝 Categorize Openings from URL', 'categorizeOpeningsFromUrl')
      .addItem('📊 Categorize Openings from ECO', 'categorizeOpeningsFromEco')
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Trigger Management')
          .addItem('Set up Triggers Sheet', 'setupTriggersSheet')
          .addItem('✅ Apply Trigger Settings', 'applyTriggers')
          .addItem('❌ Delete All Triggers', 'deleteAllTriggers'))
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Individual Actions')
          .addItem('Fetch & Process Current Month', 'refreshRecentData')
          .addItem('Update Archives List', 'updateArchives')
          .addItem('Fetch Current Month Games', 'fetchCurrentMonthGames')
          .addItem('Process Daily Data', 'updateDailyData')
          .addItem('Update Player Stats', 'updateStats')
          .addItem('Update Player Profile', 'updateProfile'))
      .addSeparator()
      .addItem('Run Initial Setup (Once)', 'setupSheets')
      .addToUi();
}

function onInstall() {
  onOpen();
}

// =================================================================
// SETUP & CONFIG
// =================================================================

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheetWithHeaders(ss, SHEETS.CONFIG, HEADERS.CONFIG);
  ensureSheetWithHeaders(ss, SHEETS.ARCHIVES, HEADERS.ARCHIVES);
  ensureSheetWithHeaders(ss, SHEETS.GAMES, HEADERS.GAMES);
  ensureSheetWithHeaders(ss, SHEETS.DAILY, HEADERS.DAILY);
  ensureSheetWithHeaders(ss, SHEETS.STATS, HEADERS.STATS);
  ensureSheetWithHeaders(ss, SHEETS.PROFILE, HEADERS.PROFILE);
  ensureSheetWithHeaders(ss, SHEETS.LOGS, HEADERS.LOGS);
  ensureSheetWithHeaders(ss, SHEETS.OPENINGS_URLS, [['Family', 'Base URL']]);
  ensureSheetWithHeaders(ss, SHEETS.ADD_OPENINGS, [['URLs to Categorize']]);
  ensureSheetWithHeaders(ss, SHEETS.CHESS_ECO, [['ECO', 'Opening Name']]);

  const configSheet = ss.getSheetByName(SHEETS.CONFIG);
  if (configSheet.getLastRow() < 2) {
    configSheet.getRange(2, 1, DEFAULT_CONFIG.length, 3).setValues(
      DEFAULT_CONFIG.map(([k, v, d]) => [k, v, d])
    );
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Sheets initialized. Fill in Config → username.', 'Setup Complete', 8);
}

function ensureSheetWithHeaders(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.clear();
  const headerRow = Array.isArray(headers[0]) ? headers[0] : headers;
  sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function getUsername() {
  const username = getConfig(CONFIG_KEYS.USERNAME);
  if (!username) throw new Error('Config username is missing. Open Config sheet and set "username".');
  return username.toString().trim().toLowerCase();
}

function updateConfig(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.CONFIG);
  const values = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 2).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === key) {
      sheet.getRange(i + 2, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value, '']);
}

function getConfig(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.CONFIG);
  const values = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 2).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === key) return values[i][1];
  }
  return '';
}

function logExecution(fnName, username, status, execMs, notes) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.LOGS);
    sheet.appendRow([
      new Date(),
      fnName,
      username || '',
      status,
      execMs,
      notes || ''
    ]);
  } catch (e) {
    // Swallow log errors to avoid masking root cause
    console.log('Log error: ' + e.message);
  }
}

// =================================================================
// CHESS.COM API HELPERS
// =================================================================

function buildApiUrl(path) {
  const base = getConfig(CONFIG_KEYS.BASE_API) || 'https://api.chess.com/pub';
  return base.replace(/\/$/, '') + '/' + path.replace(/^\//, '');
}

function fetchJson(url) {
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, validateHttpsCertificates: true });
  const code = resp.getResponseCode();
  if (code >= 200 && code < 300) return JSON.parse(resp.getContentText());
  throw new Error('HTTP ' + code + ' for ' + url);
}

// =================================================================
// PARSERS & MAPPERS
// =================================================================

function parseTimeControl(tc) {
  if (!tc) return { baseMinutes: '', incrementSeconds: '' };
  // Formats observed: "600+5", "300", "1/259200" (daily)
  if (tc.indexOf('/') >= 0) {
    // daily: X moves per Y seconds (e.g., 1 move per 259200 seconds)
    const parts = tc.split('/');
    const perSeconds = parseInt(parts[1], 10);
    return { baseMinutes: Math.round(perSeconds / 60), incrementSeconds: 0 };
  }
  const [base, inc] = tc.split('+');
  const baseMinutes = base ? Math.round(parseInt(base, 10) / 60) : '';
  const incrementSeconds = inc ? parseInt(inc, 10) : '';
  return { baseMinutes, incrementSeconds };
}

function formatDateTime(epochSeconds) {
  if (!epochSeconds && epochSeconds !== 0) return '';
  const d = new Date(epochSeconds * 1000);
  return d;
}

function formatDate(epochSeconds) {
  if (!epochSeconds && epochSeconds !== 0) return '';
  return new Date(epochSeconds * 1000);
}

function computeGameDuration(startEpoch, endEpoch) {
  if (!startEpoch || !endEpoch) return '';
  const sec = Math.max(0, endEpoch - startEpoch);
  return sec;
}

function computeFormat(rules) {
  if (!rules) return '';
  return rules.toLowerCase() === 'chess960' ? 'Chess960' : 'Standard';
}

function getPlayerPerspective(game, username) {
  const whiteUsername = (game.white && game.white.username) ? game.white.username.toLowerCase() : '';
  const blackUsername = (game.black && game.black.username) ? game.black.username.toLowerCase() : '';
  if (whiteUsername === username) return 'white';
  if (blackUsername === username) return 'black';
  return '';
}

function simplifyResult(result) {
  if (!result) return '';
  const r = result.toLowerCase();
  if (r === 'win') return 'Win';
  if (r === 'agreed' || r === 'repetition' || r === 'stalemate' || r === 'insufficient' || r === '50move' || r === 'timevsinsufficient') return 'Draw';
  return 'Loss';
}

function extractPgnTags(pgn) {
  if (!pgn) return {};
  const tags = {};
  const tagPattern = /\[(\w+)\s+"([\s\S]*?)"\]/g;
  let m;
  while ((m = tagPattern.exec(pgn)) !== null) {
    tags[m[1]] = m[2];
  }
  return tags;
}

function extractMovesFromPgn(pgn) {
  if (!pgn) return { moves: '', times: '', movesPerSide: '' };
  const parts = pgn.split(/\n\n/);
  const movesText = parts.length > 1 ? parts[parts.length - 1] : '';
  // Very light parsing; leave times blank by default
  const moves = movesText.replace(/\{[^}]*\}/g, '').trim();
  return { moves, times: '', movesPerSide: '' };
}

function gameToRow(game, username) {
  const headers = HEADERS.GAMES;
  const tc = parseTimeControl(game.time_control);
  const myColor = getPlayerPerspective(game, username);
  const my = myColor === 'white' ? game.white : game.black;
  const opp = myColor === 'white' ? game.black : game.white;
  const myRating = my && my.rating ? my.rating : '';
  const oppRating = opp && opp.rating ? opp.rating : '';
  const resultRaw = myColor === 'white' ? game.white_result : game.black_result;
  const result = simplifyResult(resultRaw);
  const pgn = game.pgn || '';
  const tags = extractPgnTags(pgn);
  const mv = extractMovesFromPgn(pgn);

  const endEpoch = game.end_time || (tags.EndTime ? Date.parse(tags.EndDate + ' ' + tags.EndTime) / 1000 : '');
  const startEpoch = game.start_time || '';

  const row = [];
  const push = (v) => row.push(v === undefined ? '' : v);

  push(game.url || ''); // Game URL
  push(game.time_control || ''); // Time Control
  push(tc.baseMinutes);
  push(tc.incrementSeconds);
  push(game.rated === true ? 'TRUE' : 'FALSE'); // Rated
  push(game.time_class || '');
  push(game.rules || '');
  push(computeFormat(game.rules));
  push(formatDateTime(endEpoch));
  push(computeGameDuration(startEpoch, endEpoch));
  push(myRating);
  push(myColor);
  push(opp && opp.username ? opp.username : '');
  push(oppRating);
  push(result);
  push(tags.Termination || '');
  push(tags.Event || '');
  push(tags.Site || '');
  push(tags.Date || '');
  push(tags.Round || '');
  push(tags.Opening || '');
  push(tags.ECO || '');
  push(tags.ECOUrl || '');
  push(tags.UTCDate || '');
  push(tags.UTCTime || '');
  push(tags.StartTime || '');
  push(tags.EndDate || '');
  push(tags.EndTime || '');
  push(game.fen || '');
  push(pgn);
  push(mv.moves);
  push(mv.times);
  push(mv.movesPerSide);
  push(''); // Opening from URL (to be filled by categorize)
  push(''); // Opening from ECO (to be filled by categorize)

  // Ensure row matches headers length
  while (row.length < headers.length) row.push('');
  if (row.length > headers.length) row.length = headers.length;
  return row;
}

// =================================================================
// ARCHIVES & GAMES
// =================================================================

function fetchArchives(username) {
  const url = buildApiUrl('player/' + encodeURIComponent(username) + '/games/archives');
  const data = fetchJson(url);
  return (data && data.archives) ? data.archives : [];
}

function updateArchivesSheet(username, archives) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ARCHIVES);
  if (archives.length === 0) return;
  const now = new Date();
  const values = archives.map(u => {
    const ym = u.split('/').slice(-2).join('-'); // YYYY-MM
    return [u, ym, now];
  });
  // Overwrite below header
  sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
  if (sheet.getLastRow() > values.length + 1) {
    sheet.getRange(values.length + 2, 1, sheet.getLastRow() - values.length - 1, sheet.getLastColumn()).clearContent();
  }
}

function getAllArchivesFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ARCHIVES);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(r => r[0]).filter(Boolean);
}

function getCurrentArchive(username) {
  const d = new Date();
  const y = d.getUTCFullYear();
  const m = (d.getUTCMonth() + 1).toString().padStart(2, '0');
  return buildApiUrl('player/' + encodeURIComponent(username) + '/games/' + y + '/' + m);
}

function fetchGamesFromArchive(archiveUrl) {
  const data = fetchJson(archiveUrl);
  return (data && data.games) ? data.games : [];
}

function getExistingGameUrls() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GAMES);
  if (!sheet || sheet.getLastRow() < 2) return new Set();
  const urls = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(r => r[0]).filter(Boolean);
  return new Set(urls);
}

function addNewGames(rows) {
  if (!rows || rows.length === 0) return 0;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GAMES);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  return rows.length;
}

function fetchCurrentMonthGames() {
  const startMs = new Date().getTime();
  const username = getUsername();
  const archiveUrl = getCurrentArchive(username);
  const games = fetchGamesFromArchive(archiveUrl);
  const existing = getExistingGameUrls();
  const rows = [];
  for (let i = 0; i < games.length; i++) {
    const g = games[i];
    if (!g.url || existing.has(g.url)) continue;
    rows.push(gameToRow(g, username));
  }
  const added = addNewGames(rows);
  logExecution('fetchCurrentMonthGames', username, 'SUCCESS', new Date().getTime() - startMs, 'Added ' + added + ' new games');
}

function fetchAllGames() {
  const startMs = new Date().getTime();
  const username = getUsername();
  const archives = getAllArchivesFromSheet();
  const existing = getExistingGameUrls();
  let added = 0;
  for (let i = 0; i < archives.length; i++) {
    const archiveUrl = archives[i];
    try {
      const games = fetchGamesFromArchive(archiveUrl);
      const rows = [];
      for (let j = 0; j < games.length; j++) {
        const g = games[j];
        if (!g.url || existing.has(g.url)) continue;
        rows.push(gameToRow(g, username));
        existing.add(g.url);
      }
      added += addNewGames(rows);
      Utilities.sleep(200); // be nice to API
    } catch (e) {
      console.log('Archive fetch error: ' + archiveUrl + ' => ' + e.message);
    }
  }
  logExecution('fetchAllGames', username, 'SUCCESS', new Date().getTime() - startMs, 'Added ' + added + ' new games');
}

function updateArchives() {
  const startMs = new Date().getTime();
  const username = getUsername();
  const archives = fetchArchives(username);
  updateArchivesSheet(username, archives);
  logExecution('updateArchives', username, 'SUCCESS', new Date().getTime() - startMs, 'Archives: ' + archives.length);
}

// =================================================================
// DAILY AGGREGATION
// =================================================================

function updateDailyData() {
  const startMs = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gameSheet = ss.getSheetByName(SHEETS.GAMES);
  const dailySheet = ss.getSheetByName(SHEETS.DAILY);
  if (!gameSheet || gameSheet.getLastRow() < 2) return;

  const data = gameSheet.getRange(1, 1, gameSheet.getLastRow(), gameSheet.getLastColumn()).getValues();
  const header = data[0];
  const idx = {
    endTime: header.indexOf('End Time'),
    timeClass: header.indexOf('Time Class'),
    result: header.indexOf('Result'),
    myRating: header.indexOf('My Rating'),
    duration: header.indexOf('Game Duration (sec)')
  };

  const dailyMap = new Map();

  function keyFor(date) { return date.toISOString().slice(0, 10); }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dt = row[idx.endTime];
    const timeClass = (row[idx.timeClass] || '').toString().toLowerCase();
    const result = (row[idx.result] || '').toString();
    const rating = Number(row[idx.myRating] || 0);
    const dur = Number(row[idx.duration] || 0);
    if (!dt || !(dt instanceof Date)) continue;
    const dKey = keyFor(dt);
    if (!dailyMap.has(dKey)) {
      dailyMap.set(dKey, {
        bullet: { w: 0, l: 0, d: 0, rating: 0, change: 0, time: 0 },
        blitz:  { w: 0, l: 0, d: 0, rating: 0, change: 0, time: 0 },
        rapid:  { w: 0, l: 0, d: 0, rating: 0, change: 0, time: 0 },
        totals: { g: 0, w: 0, l: 0, d: 0, ratingSum: 0, change: 0, time: 0 }
      });
    }
    const agg = dailyMap.get(dKey);
    const bucket = (timeClass === 'bullet' || timeClass === 'blitz' || timeClass === 'rapid') ? timeClass : null;
    if (bucket) {
      if (result === 'Win') agg[bucket].w += 1;
      else if (result === 'Loss') agg[bucket].l += 1;
      else if (result === 'Draw') agg[bucket].d += 1;
      // naive: set rating as last seen of the day per class
      agg[bucket].rating = rating || agg[bucket].rating;
      agg[bucket].time += dur;
    }
    agg.totals.g += 1;
    if (result === 'Win') agg.totals.w += 1;
    else if (result === 'Loss') agg.totals.l += 1;
    else if (result === 'Draw') agg.totals.d += 1;
    agg.totals.ratingSum += rating;
    agg.totals.time += dur;
  }

  const rows = [];
  const keys = Array.from(dailyMap.keys()).sort();
  for (let i = 0; i < keys.length; i++) {
    const k = keys[i];
    const a = dailyMap.get(k);
    const totalChange = a.totals.change; // keep 0 for now
    const avgDur = a.totals.g > 0 ? a.totals.time / a.totals.g : 0;
    rows.push([
      new Date(k),
      a.bullet.w, a.bullet.l, a.bullet.d, a.bullet.rating, a.bullet.change, a.bullet.time,
      a.blitz.w,  a.blitz.l,  a.blitz.d,  a.blitz.rating,  a.blitz.change,  a.blitz.time,
      a.rapid.w,  a.rapid.l,  a.rapid.d,  a.rapid.rating,  a.rapid.change,  a.rapid.time,
      a.totals.g, a.totals.w, a.totals.l, a.totals.d, a.totals.ratingSum, totalChange, a.totals.time, avgDur
    ]);
  }

  // Overwrite below header
  if (rows.length > 0) {
    dailySheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    if (dailySheet.getLastRow() > rows.length + 1) {
      dailySheet.getRange(rows.length + 2, 1, dailySheet.getLastRow() - rows.length - 1, dailySheet.getLastColumn()).clearContent();
    }
  }

  logExecution('updateDailyData', getUsername(), 'SUCCESS', new Date().getTime() - startMs, 'Days: ' + rows.length);
}

// =================================================================
// PROFILE & STATS
// =================================================================

function fetchPlayerStats(username) {
  const url = buildApiUrl('player/' + encodeURIComponent(username) + '/stats');
  return fetchJson(url);
}

function updateStats() {
  const startMs = new Date().getTime();
  const username = getUsername();
  const stats = fetchPlayerStats(username);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.STATS);
  const flat = flattenObject(stats);
  const values = [['Pulled At', new Date()]];
  Object.keys(flat).sort().forEach(k => values.push([k, flat[k]]));
  sheet.clear();
  sheet.getRange(1, 1, 1, 1).setValue('Pulled At').setFontWeight('bold');
  sheet.getRange(2, 1, values.length - 1, 2).setValues(values.slice(1));
  logExecution('updateStats', username, 'SUCCESS', new Date().getTime() - startMs, 'Stats fields: ' + (values.length - 1));
}

function fetchPlayerProfile(username) {
  const url = buildApiUrl('player/' + encodeURIComponent(username));
  return fetchJson(url);
}

function updateProfile() {
  const startMs = new Date().getTime();
  const username = getUsername();
  const profile = fetchPlayerProfile(username);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PROFILE);
  const entries = Object.entries(profile);
  sheet.clear();
  sheet.getRange(1, 1, 1, 2).setValues([[HEADERS.PROFILE[0], HEADERS.PROFILE[1]]]).setFontWeight('bold');
  if (entries.length > 0) sheet.getRange(2, 1, entries.length, 2).setValues(entries);
  logExecution('updateProfile', username, 'SUCCESS', new Date().getTime() - startMs, 'Profile fields: ' + entries.length);
}

function flattenObject(obj, prefix = '', res = {}) {
  if (obj === null || obj === undefined) return res;
  Object.keys(obj).forEach(k => {
    const v = obj[k];
    const p = prefix ? prefix + '.' + k : k;
    if (typeof v === 'object' && v !== null && !Array.isArray(v)) flattenObject(v, p, res);
    else res[p] = Array.isArray(v) ? JSON.stringify(v) : v;
  });
  return res;
}

// =================================================================
// CATEGORIZATION FUNCTIONS (WITH CACHING)
// =================================================================

function categorizeOpeningsFromUrl() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gameSheet = ss.getSheetByName(SHEETS.GAMES);
  const lookupSheet = ss.getSheetByName(SHEETS.OPENINGS_URLS);
  let addOpeningsSheet = ss.getSheetByName(SHEETS.ADD_OPENINGS);

  if (!lookupSheet || lookupSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('The "Opening URLS" sheet is missing or empty.');
    return;
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = 'openingUrlMap';
  let openingMap;

  const cachedMap = cache.get(cacheKey);
  if (cachedMap) {
    openingMap = new Map(JSON.parse(cachedMap));
    console.log('Loaded Opening URL map from cache.');
  } else {
    console.log('Building Opening URL map from sheet...');
    const lookupData = lookupSheet.getRange(2, 1, lookupSheet.getLastRow() - 1, 2).getValues();
    openingMap = new Map(lookupData.map(row => [row[1], row[0]]));
    cache.put(cacheKey, JSON.stringify(Array.from(openingMap.entries())), 21600);
  }

  if (!addOpeningsSheet) {
    addOpeningsSheet = ss.insertSheet(SHEETS.ADD_OPENINGS);
    addOpeningsSheet.getRange(1, 1).setValue('URLs to Categorize').setFontWeight('bold');
  }

  const existingAddUrls = new Set(addOpeningsSheet.getLastRow() > 1 ? addOpeningsSheet.getRange(2, 1, addOpeningsSheet.getLastRow() - 1, 1).getValues().flat() : []);

  const gameData = gameSheet.getRange(1, 1, gameSheet.getLastRow(), gameSheet.getLastColumn()).getValues();
  const headers = gameData[0];
  const ecoUrlIndex = headers.indexOf('ECO URL');
  const targetColIndex = headers.indexOf('Opening from URL');

  if (ecoUrlIndex === -1 || targetColIndex === -1) { return; }

  let categorizedCount = 0;
  const newUrlsToAdd = [];

  for (let i = 1; i < gameData.length; i++) {
    const gameEcoUrl = gameData[i][ecoUrlIndex];
    if (!gameEcoUrl) continue;
    let foundFamily = '';
    for (const [baseUrl, familyName] of openingMap.entries()) {
      if (gameEcoUrl.toString().startsWith(baseUrl)) {
        foundFamily = familyName;
        categorizedCount++;
        break;
      }
    }
    gameData[i][targetColIndex] = foundFamily;
    if (!foundFamily && !existingAddUrls.has(gameEcoUrl)) {
      newUrlsToAdd.push([gameEcoUrl]);
      existingAddUrls.add(gameEcoUrl);
    }
  }

  gameSheet.getRange(1, 1, gameData.length, gameData[0].length).setValues(gameData);
  if (newUrlsToAdd.length > 0) {
    addOpeningsSheet.getRange(addOpeningsSheet.getLastRow() + 1, 1, newUrlsToAdd.length, 1).setValues(newUrlsToAdd);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('URL Categorization complete! Updated ' + categorizedCount + ' games.', 'Success', 8);
}

function categorizeOpeningsFromEco() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gameSheet = ss.getSheetByName(SHEETS.GAMES);
  const lookupSheet = ss.getSheetByName(SHEETS.CHESS_ECO);

  if (!lookupSheet || lookupSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('The "Chess ECO" sheet is missing or empty.');
    return;
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = 'ecoMap';
  let ecoMap;

  const cachedMap = cache.get(cacheKey);
  if (cachedMap) {
    ecoMap = new Map(JSON.parse(cachedMap));
    console.log('Loaded ECO map from cache.');
  } else {
    console.log('Building ECO map from sheet...');
    const lookupData = lookupSheet.getRange(2, 1, lookupSheet.getLastRow() - 1, 2).getValues();
    ecoMap = new Map(lookupData.map(row => [row[0].toString().trim(), row[1]]));
    cache.put(cacheKey, JSON.stringify(Array.from(ecoMap.entries())), 21600);
  }

  const gameData = gameSheet.getRange(1, 1, gameSheet.getLastRow(), gameSheet.getLastColumn()).getValues();
  const headers = gameData[0];
  const ecoIndex = headers.indexOf('ECO');
  const targetColIndex = headers.indexOf('Opening from ECO');

  if (ecoIndex === -1 || targetColIndex === -1) { return; }

  let categorizedCount = 0;
  for (let i = 1; i < gameData.length; i++) {
    const gameEcoCode = gameData[i][ecoIndex];
    const key = gameEcoCode ? gameEcoCode.toString().trim() : '';
    if (key && ecoMap.has(key)) {
      gameData[i][targetColIndex] = ecoMap.get(key);
      categorizedCount++;
    }
  }

  gameSheet.getRange(1, 1, gameData.length, gameData[0].length).setValues(gameData);
  SpreadsheetApp.getActiveSpreadsheet().toast('ECO categorization complete! Updated ' + categorizedCount + ' games.', 'Success', 8);
}

// =================================================================
// TRIGGER MANAGEMENT
// =================================================================

function setupTriggersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let triggerSheet = ss.getSheetByName(SHEETS.TRIGGERS);
  if (!triggerSheet) {
    triggerSheet = ss.insertSheet(SHEETS.TRIGGERS);
  }
  triggerSheet.clear().setFrozenRows(1);

  const headers = ['Function to Run', 'Frequency', 'Day of the Week (for Weekly)', 'Time of Day (for Daily/Weekly)', 'Status', 'Trigger ID'];
  triggerSheet.getRange('A1:F1').setValues([headers]).setFontWeight('bold');

  const functionRule = SpreadsheetApp.newDataValidation().requireValueInList(TRIGGERABLE_FUNCTIONS, true).build();

  const frequencyRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([
        'Every hour', 'Every 2 hours', 'Every 4 hours', 'Every 6 hours', 'Every 12 hours',
        'Daily', 'Weekly'
    ], true).build();

  const dayOfWeekRule = SpreadsheetApp.newDataValidation().requireValueInList(['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'], true).build();

  triggerSheet.getRange('A2:A').setDataValidation(functionRule);
  triggerSheet.getRange('B2:B').setDataValidation(frequencyRule);
  triggerSheet.getRange('C2:C').setDataValidation(dayOfWeekRule);

  triggerSheet.autoResizeColumns(1, 4);
  SpreadsheetApp.getActiveSpreadsheet().toast('Triggers sheet has been set up!', 'Success', 5);
}

function applyTriggers() {
  const allTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of allTriggers) {
    if (TRIGGERABLE_FUNCTIONS.includes(trigger.getHandlerFunction())) {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triggerSheet = ss.getSheetByName(SHEETS.TRIGGERS);
  if (!triggerSheet) { throw new Error('Triggers sheet not found.'); }

  const settings = triggerSheet.getRange(2, 1, Math.max(triggerSheet.getLastRow() - 1, 0), 6).getValues();

  settings.forEach((row, index) => {
    const [functionName, frequency, dayOfWeek, timeOfDay] = row;
    if (!functionName || !frequency) {
      triggerSheet.getRange(index + 2, 5, 1, 2).setValues([['Inactive', '']]);
      return;
    }
    try {
      let newTriggerBuilder;

      if (frequency.startsWith('Every')) {
        const hoursMatch = frequency.match(/\d+/);
        const hours = hoursMatch ? parseInt(hoursMatch[0]) : 1;
        newTriggerBuilder = ScriptApp.newTrigger(functionName).timeBased().everyHours(hours);
      }
      else if (frequency === 'Daily' && timeOfDay) {
        const [hour, minute] = timeOfDay.toString().split(':');
        newTriggerBuilder = ScriptApp.newTrigger(functionName).timeBased().everyDays(1).atHour(parseInt(hour)).nearMinute(parseInt(minute));
      }
      else if (frequency === 'Weekly' && dayOfWeek && timeOfDay) {
        const [hour, minute] = timeOfDay.toString().split(':');
        const weekDay = ScriptApp.WeekDay[dayOfWeek.toUpperCase()];
        newTriggerBuilder = ScriptApp.newTrigger(functionName).timeBased().onWeekDay(weekDay).atHour(parseInt(hour)).nearMinute(parseInt(minute));
      } else {
        throw new Error('Invalid settings for frequency: ' + frequency);
      }

      const newTrigger = newTriggerBuilder.create();
      triggerSheet.getRange(index + 2, 5, 1, 2).setValues([['Active', newTrigger.getUniqueId()]]);

    } catch (e) {
      console.error('Error on row ' + (index + 2) + ': ' + e.message);
      triggerSheet.getRange(index + 2, 5).setValue('Error: ' + e.message);
    }
  });
  SpreadsheetApp.getActiveSpreadsheet().toast('Triggers have been updated successfully!', 'Triggers Synced', 8);
}

function deleteAllTriggers() {
  const allTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of allTriggers) {
    ScriptApp.deleteTrigger(trigger);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triggerSheet = ss.getSheetByName(SHEETS.TRIGGERS);
  if (triggerSheet && triggerSheet.getLastRow() > 1) {
    triggerSheet.getRange(2, 5, triggerSheet.getLastRow() - 1, 2).clearContent();
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('Deleted ' + allTriggers.length + ' triggers.', 'All Triggers Removed', 5);
}

// =================================================================
// FOCUSED RECENT REFRESH
// =================================================================

function refreshRecentData() {
  const startTime = new Date().getTime();
  let username = '';
  const results = [];
  try {
    username = getUsername();
    console.log('Starting recent data refresh for user: ' + username);

    try {
      fetchCurrentMonthGames();
      results.push('✓ Current month games fetched');
    } catch (error) {
      results.push('✗ Current month games failed: ' + error.message);
    }

    try {
      updateDailyData();
      results.push('✓ Daily data processed');
    } catch (error) {
      results.push('✗ Daily data failed: ' + error.message);
    }

    const executionTime = new Date().getTime() - startTime;
    logExecution('refreshRecentData', username, 'SUCCESS', executionTime, results.join(', '));
    SpreadsheetApp.getActiveSpreadsheet().toast('Recent data refresh complete!\n\n' + results.join('\n'), 'Refresh Complete', 8);

  } catch (error) {
    const executionTime = new Date().getTime() - startTime;
    logExecution('refreshRecentData', username, 'ERROR', executionTime, error.message);
    throw error;
  }
}

// =================================================================
// HIGH-LEVEL WORKFLOWS
// =================================================================

function quickUpdate() {
  const startMs = new Date().getTime();
  const username = getUsername();
  const results = [];
  try {
    updateArchives();
    results.push('Archives updated');
  } catch (e) {
    results.push('Archives failed: ' + e.message);
  }
  try {
    fetchCurrentMonthGames();
    results.push('Current month games fetched');
  } catch (e) {
    results.push('Fetch current month failed: ' + e.message);
  }
  try {
    updateDailyData();
    results.push('Daily data updated');
  } catch (e) {
    results.push('Daily update failed: ' + e.message);
  }
  try {
    updateStats();
    results.push('Stats updated');
  } catch (e) {
    results.push('Stats failed: ' + e.message);
  }
  try {
    updateProfile();
    results.push('Profile updated');
  } catch (e) {
    results.push('Profile failed: ' + e.message);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(results.join('\n'), 'Quick Update', 8);
  logExecution('quickUpdate', username, 'SUCCESS', new Date().getTime() - startMs, results.join(', '));
}

function completeUpdate() {
  const startMs = new Date().getTime();
  const username = getUsername();
  const results = [];
  try {
    updateArchives();
    results.push('Archives updated');
  } catch (e) {
    results.push('Archives failed: ' + e.message);
  }
  try {
    fetchAllGames();
    results.push('All games fetched');
  } catch (e) {
    results.push('Fetch all games failed: ' + e.message);
  }
  try {
    updateDailyData();
    results.push('Daily data updated');
  } catch (e) {
    results.push('Daily update failed: ' + e.message);
  }
  try {
    updateStats();
    results.push('Stats updated');
  } catch (e) {
    results.push('Stats failed: ' + e.message);
  }
  try {
    updateProfile();
    results.push('Profile updated');
  } catch (e) {
    results.push('Profile failed: ' + e.message);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(results.join('\n'), 'Complete Update', 8);
  logExecution('completeUpdate', username, 'SUCCESS', new Date().getTime() - startMs, results.join(', '));
}

