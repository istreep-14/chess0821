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
    'Termination', 'Winner', 'Event', 'Site', 'Date', 'Round', 'Opening', 'ECO',
    'ECO URL', 'UTC Date', 'UTC Time', 'PGN Start Time', 'PGN End Date', 'PGN End Time',
    'Current Position', 'Full PGN', 'Moves', 'Times', 'Moves Per Side',
    'PGN Timezone', 'My Timezone', 'Local Start Time (Live)', 'Hour Differential (hrs)',
    'My Username', 'Rating Change', 'My Player ID', 'My UUID',
    'Opponent Color', 'Opponent Username', 'Opponent Player ID', 'Opponent UUID',
    'My Accuracy', 'Opponent Accuracy',
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
  STATS: ['Field', 'Value'],
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
  const headerRow = Array.isArray(headers[0]) ? headers[0] : headers;
  // Only set header row; do not clear existing data
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
  const maxAttempts = 5;
  const baseSleepMs = 500;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, validateHttpsCertificates: true });
      const code = resp.getResponseCode();
      if (code >= 200 && code < 300) return JSON.parse(resp.getContentText());
      if (code === 429 || (code >= 500 && code < 600)) {
        const sleepMs = baseSleepMs * Math.pow(2, attempt - 1) + Math.floor(Math.random() * 250);
        Utilities.sleep(sleepMs);
        continue;
      }
      throw new Error('HTTP ' + code + ' for ' + url);
    } catch (e) {
      if (attempt === maxAttempts) throw e;
      Utilities.sleep(baseSleepMs * Math.pow(2, attempt - 1));
    }
  }
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
  if (r === 'draw') return 'Draw';
  return 'Loss';
}

function normalizeMethod(causeCandidate, terminationTag) {
  const val = (causeCandidate || '').toString().toLowerCase();
  const term = (terminationTag || '').toString().toLowerCase();
  const map = {
    checkmated: 'Checkmated',
    checkmate: 'Checkmated',
    resigned: 'Resigned',
    timeout: 'Timeout',
    abandoned: 'Abandoned',
    stalemate: 'Stalemate',
    agreed: 'Draw agreed',
    repetition: 'Draw by repetition',
    insufficient: 'Insufficient material',
    '50move': 'Draw by 50-move rule',
    timevsinsufficient: 'Draw by timeout vs insufficient material',
    kingofthehill: 'Opponent king reached the hill',
    threecheck: 'Checked for the 3rd time',
    bughousepartnerlose: 'Bughouse partner lost',
    '': ''
  };
  if (map[val]) return map[val];
  // Try to infer from termination phrase
  const keys = Object.keys(map);
  for (let i = 0; i < keys.length; i++) {
    if (keys[i] && term.indexOf(keys[i]) !== -1) return map[keys[i]];
  }
  return term ? term.charAt(0).toUpperCase() + term.slice(1) : '';
}

function parseTerminationTag(terminationTag) {
  const raw = (terminationTag || '').toString();
  const t = raw.trim();
  if (!t) return { winner: '', cause: '' };
  const wonMatch = t.match(/^\s*([^\s].*?)\s+won\s+by\s+(.+?)\s*$/i);
  if (wonMatch) {
    return { winner: wonMatch[1], cause: wonMatch[2] };
  }
  const drawMatch = t.match(/\bdrawn\s+by\s+(.+?)\s*$/i);
  if (drawMatch) {
    return { winner: '', cause: drawMatch[1] };
  }
  // Fallback: try to use the whole phrase as cause
  return { winner: '', cause: t };
}

function normalizePgnDateToDate(pgnDate) {
  if (!pgnDate) return '';
  if (pgnDate instanceof Date) return pgnDate;
  const s = pgnDate.toString();
  // Expect formats like YYYY.MM.DD or sometimes 023.08.29 (3-digit year)
  const parts = s.split('.');
  if (parts.length !== 3) return '';
  let [y, m, d] = parts;
  if (!/^[0-9?]+$/.test(y) || !/^[0-9?]+$/.test(m) || !/^[0-9?]+$/.test(d)) return '';
  if (y.includes('?') || m.includes('?') || d.includes('?')) return '';
  if (y.length === 3) y = '2' + y; // e.g., 023 -> 2023
  if (y.length === 2) y = '20' + y; // fallback for 2-digit year
  if (y.length !== 4) return '';
  const iso = y + '-' + m.padStart(2, '0') + '-' + d.padStart(2, '0');
  const dt = new Date(iso);
  return isNaN(dt.getTime()) ? '' : dt;
}

function normalizePgnTimeToHms(pgnTime) {
  if (!pgnTime) return '';
  const s = pgnTime.toString().trim();
  if (!s || s.indexOf('?') !== -1) return '';
  const match = s.match(/^(\d{1,2})(?::(\d{1,2}))?(?::(\d{1,2})(?:\.\d+)?)?$/);
  if (!match) return '';
  let hh = parseInt(match[1], 10);
  let mm = match[2] !== undefined ? parseInt(match[2], 10) : 0;
  let ss = match[3] !== undefined ? parseInt(match[3], 10) : 0;
  if (isNaN(hh) || isNaN(mm) || isNaN(ss)) return '';
  if (hh < 0 || hh > 23 || mm < 0 || mm > 59 || ss < 0 || ss > 59) return '';
  const hStr = String(hh).padStart(2, '0');
  const mStr = String(mm).padStart(2, '0');
  const sStr = String(ss).padStart(2, '0');
  return hStr + ':' + mStr + ':' + sStr;
}

function combinePgnDateAndTimeToDate(pgnDate, pgnTime, assumeUtc) {
  try {
    if (!pgnDate || !pgnTime) return '';
    const ds = pgnDate.toString();
    const parts = ds.split('.');
    if (parts.length !== 3) return '';
    let [y, m, d] = parts;
    if (!/^[0-9]+$/.test(y) || !/^[0-9]+$/.test(m) || !/^[0-9]+$/.test(d)) return '';
    if (y.length === 3) y = '2' + y; // 3-digit year -> 20xx
    if (y.length === 2) y = '20' + y; // 2-digit year -> 20xx
    if (y.length !== 4) return '';
    const mm = m.padStart(2, '0');
    const dd = d.padStart(2, '0');
    const hms = normalizePgnTimeToHms(pgnTime);
    if (!hms) return '';
    const iso = y + '-' + mm + '-' + dd + 'T' + hms + (assumeUtc ? 'Z' : '');
    const dt = new Date(iso);
    return isNaN(dt.getTime()) ? '' : dt;
  } catch (e) {
    return '';
  }
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

function getLocalTimezoneName() {
  const date = new Date();
  const match = date.toString().match(/\(([^)]+)\)$/);
  return match && match[1] ? match[1] : '';
}

function getLocalTimezoneOffsetHours(date) {
  const offsetMinutes = date.getTimezoneOffset();
  return -offsetMinutes / 60; // JavaScript offset is minutes behind UTC; invert to get UTC+/- hours
}

function parsePgnTimezone(tags) {
  // PGN often provides UTCDate/UTCTime; assume 'UTC' when present
  if (tags && (tags.UTCDate || tags.UTCTime)) return 'UTC';
  // Some PGNs may include TimeZone or Zone tags
  if (tags && (tags.TimeZone || tags.Zone)) return tags.TimeZone || tags.Zone;
  return '';
}

function computeLocalStartTimeForLive(game, tags) {
  // Live games: prefer game.start_time (epoch seconds). If missing, derive from PGN UTCDate/UTCTime
  try {
    if (game && game.time_class && game.time_class !== 'daily') {
      if (game.start_time) return new Date(game.start_time * 1000);
      if (tags && tags.UTCDate && tags.UTCTime) {
        const iso = tags.UTCDate + 'T' + tags.UTCTime + 'Z';
        const d = new Date(iso);
        if (!isNaN(d.getTime())) return d; // local Date object will reflect local tz automatically when displayed in Sheets
      }
      if (tags && tags.Date && tags.StartTime) {
        // Fallback if only Date and StartTime without zone; treat as UTC to be safe
        const d = new Date(tags.Date + 'T' + tags.StartTime + 'Z');
        if (!isNaN(d.getTime())) return d;
      }
    }
  } catch (e) {}
  return '';
}

function computeHourDifferential(pgnTz, localDate) {
  try {
    if (!localDate || !(localDate instanceof Date)) return '';
    const localOffset = getLocalTimezoneOffsetHours(localDate);
    const pgnOffset = (pgnTz && pgnTz.toUpperCase() === 'UTC') ? 0 : null;
    if (pgnOffset === null) return '';
    const diff = localOffset - pgnOffset;
    return Math.round(diff * 100) / 100;
  } catch (e) { return ''; }
}

function extractMovesFromPgn(pgn) {
  if (!pgn) return { moves: '', times: '', movesPerSide: '' };
  const blankLineIndex = pgn.indexOf('\n\n');
  if (blankLineIndex === -1) return { moves: '', times: '', movesPerSide: '' };
  let movesText = pgn.slice(blankLineIndex + 2).trim();
  movesText = movesText.replace(/\s+(1-0|0-1|1\/2-1\/2|\*)\s*$/, '');
  movesText = movesText.replace(/\([^)]*\)/g, '');
  movesText = movesText.replace(/\$\d+/g, '');

  const segments = [];
  const pairRegex = /(\d+)\.\s*(?!\.\.\.)([^\s{}]+)(?:\s*\{([^}]*)\})?(?:\s+(?:(?:\d+)?\.{3}\s*)?([^\s{}]+)(?:\s*\{([^}]*)\})?)?/g;
  let m;
  while ((m = pairRegex.exec(movesText)) !== null) {
    const moveNo = m[1];
    const whiteSan = m[2];
    const whiteComment = m[3] || '';
    const blackSan = m[4];
    const blackComment = m[5] || '';
    const wClkMatch = whiteComment.match(/\[%clk\s+([0-9:\.]+)\]/);
    const bClkMatch = (blackComment || '').match(/\[%clk\s+([0-9:\.]+)\]/);
    const wStr = whiteSan ? (whiteSan + (wClkMatch ? ' [' + wClkMatch[1] + ']' : '')) : '';
    const bStr = blackSan ? (blackSan + (bClkMatch ? ' [' + bClkMatch[1] + ']' : '')) : '';
    if (wStr || bStr) {
      segments.push(moveNo + '. ' + [wStr, bStr].filter(Boolean).join(' '));
    }
  }
  // Handle black-only notation like: 23... c5 {[%clk 0:00:42]}
  const blackOnlyRegex = /(\d+)\.\s*\.\.\.\s*([^\s{}]+)(?:\s*\{([^}]*)\})?/g;
  while ((m = blackOnlyRegex.exec(movesText)) !== null) {
    const moveNo = m[1];
    const blackSan = m[2];
    const blackComment = m[3] || '';
    const bClkMatch = blackComment.match(/\[%clk\s+([0-9:\.]+)\]/);
    const bStr = blackSan + (bClkMatch ? ' [' + bClkMatch[1] + ']' : '');
    segments.push(moveNo + '... ' + bStr);
  }

  const moves = segments.join(' ');
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
  const parsedTerm = parseTerminationTag(tags.Termination || '');
  const termination = normalizeMethod(parsedTerm.cause || resultRaw, tags.Termination || '');
  // Winner determination
  let winner = '';
  const myUsername = my && my.username ? my.username : '';
  if (parsedTerm.winner) {
    winner = parsedTerm.winner;
  } else if (result === 'Win') {
    winner = myUsername || (myColor || '').charAt(0).toUpperCase() + (myColor || '').slice(1);
  } else if (result === 'Loss') {
    winner = opp && opp.username ? opp.username : (myColor === 'white' ? 'Black' : (myColor === 'black' ? 'White' : ''));
  } else {
    winner = 'Draw';
  }

  const startUtc = (tags && tags.UTCDate && tags.UTCTime)
    ? combinePgnDateAndTimeToDate(tags.UTCDate, tags.UTCTime, true)
    : (game.start_time ? new Date(game.start_time * 1000) : '');
  const endCombined = (tags && tags.EndDate && tags.EndTime)
    ? combinePgnDateAndTimeToDate(tags.EndDate, tags.EndTime, true)
    : (game.end_time ? new Date(game.end_time * 1000) : '');
  const durationSec = (startUtc && endCombined)
    ? Math.max(0, Math.round((endCombined.getTime() - startUtc.getTime()) / 1000))
    : '';
  // Timezone/context values
  const pgnTimezone = parsePgnTimezone(tags);
  const localStartTime = computeLocalStartTimeForLive(game, tags);
  const myTimezone = getLocalTimezoneName();
  const hourDifferential = computeHourDifferential(pgnTimezone, localStartTime || (endCombined ? endCombined : null));

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
  push(endCombined || '');
  push(durationSec);
  push(myRating);
  push(myColor);
  push(opp && opp.username ? opp.username : '');
  push(oppRating);
  // Build normalized Result phrase using usernames
  let resultPhrase = '';
  if (winner && winner.toLowerCase() !== 'draw') {
    resultPhrase = winner + ' won by ' + (parsedTerm.cause || termination || '').toString().toLowerCase();
  } else {
    const drawCause = (parsedTerm.cause || termination || '').toString().toLowerCase();
    resultPhrase = drawCause ? ('drawn by ' + drawCause) : 'drawn';
  }
  push(resultPhrase);
  push(termination);
  push(winner);
  push(tags.Event || '');
  push(tags.Site || '');
  push(normalizePgnDateToDate(tags.Date || ''));
  push(tags.Round || '');
  push(tags.Opening || '');
  push(tags.ECO || '');
  push(tags.ECOUrl || '');
  push(normalizePgnDateToDate(tags.UTCDate || ''));
  push(tags.UTCTime || '');
  push(startUtc || '');
  push(normalizePgnDateToDate(tags.EndDate || ''));
  push(tags.EndTime || '');
  push(game.fen || '');
  push(pgn);
  push(mv.moves);
  push(mv.times);
  push(mv.movesPerSide);
  // New timezone/identity columns
  push(pgnTimezone);
  push(myTimezone);
  push(localStartTime || '');
  push(hourDifferential);

  // Identity fields and stats
  const myId = my && (my.player_id || my.playerId || my.id) ? (my.player_id || my.playerId || my.id) : '';
  const myUuid = my && (my.uuid || my.uuid4 || my.guid) ? (my.uuid || my.uuid4 || my.guid) : '';
  const oppColor = myColor === 'white' ? 'black' : (myColor === 'black' ? 'white' : '');
  const oppUsername = opp && opp.username ? opp.username : '';
  const oppId = opp && (opp.player_id || opp.playerId || opp.id) ? (opp.player_id || opp.playerId || opp.id) : '';
  const oppUuid = opp && (opp.uuid || opp.uuid4 || opp.guid) ? (opp.uuid || opp.uuid4 || opp.guid) : '';

  // Rating change (not always available in archives) – leave blank if unknown
  let ratingChange = '';
  // Accuracies from PGN tags if present
  const myAccuracy = (myColor === 'white') ? (tags.WhiteAccuracy || '') : (tags.BlackAccuracy || '');
  const oppAccuracy = (myColor === 'white') ? (tags.BlackAccuracy || '') : (tags.WhiteAccuracy || '');

  push(myUsername);
  push(ratingChange);
  push(myId);
  push(myUuid);
  push(oppColor);
  push(oppUsername);
  push(oppId);
  push(oppUuid);
  push(myAccuracy);
  push(oppAccuracy);
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
    winner: header.indexOf('Winner'),
    myUsername: header.indexOf('My Username'),
    myRating: header.indexOf('My Rating'),
    duration: header.indexOf('Game Duration (sec)')
  };

  const dailyMap = new Map();

  function keyFor(date) { return date.toISOString().slice(0, 10); }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dt = row[idx.endTime];
    const timeClass = (row[idx.timeClass] || '').toString().toLowerCase();
    const winnerVal = (row[idx.winner] || '').toString();
    const myName = (row[idx.myUsername] || '').toString();
    let result = 'Draw';
    if (winnerVal && winnerVal.toLowerCase() !== 'draw') {
      result = (myName && winnerVal.toLowerCase() === myName.toLowerCase()) ? 'Win' : 'Loss';
    }
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
  const rows = [['Field', 'Value']];
  rows.push(['Pulled At', new Date()]);
  Object.keys(flat).sort().forEach(k => rows.push([k, flat[k]]));
  sheet.clear();
  sheet.getRange(1, 1, rows.length, 2).setValues(rows).setFontWeight('bold');
  sheet.getRange(2, 1, rows.length - 1, 2).setFontWeight('normal');
  logExecution('updateStats', username, 'SUCCESS', new Date().getTime() - startMs, 'Stats fields: ' + (rows.length - 2));
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

  const targetValues = [];
  for (let i = 1; i < gameData.length; i++) {
    const gameEcoUrl = gameData[i][ecoUrlIndex];
    let foundFamily = '';
    if (gameEcoUrl) {
      for (const [baseUrl, familyName] of openingMap.entries()) {
        if (gameEcoUrl.toString().startsWith(baseUrl)) {
          foundFamily = familyName;
          categorizedCount++;
          break;
        }
      }
      if (!foundFamily && !existingAddUrls.has(gameEcoUrl)) {
        newUrlsToAdd.push([gameEcoUrl]);
        existingAddUrls.add(gameEcoUrl);
      }
    }
    targetValues.push([foundFamily]);
  }

  if (targetValues.length > 0) {
    gameSheet.getRange(2, targetColIndex + 1, targetValues.length, 1).setValues(targetValues);
  }
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
  const targetValues = [];
  for (let i = 1; i < gameData.length; i++) {
    const gameEcoCode = gameData[i][ecoIndex];
    const key = gameEcoCode ? gameEcoCode.toString().trim() : '';
    let val = '';
    if (key && ecoMap.has(key)) {
      val = ecoMap.get(key);
      categorizedCount++;
    }
    targetValues.push([val]);
  }

  if (targetValues.length > 0) {
    gameSheet.getRange(2, targetColIndex + 1, targetValues.length, 1).setValues(targetValues);
  }
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
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { throw new Error('Another run is in progress'); }
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
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// =================================================================
// HIGH-LEVEL WORKFLOWS
// =================================================================

function quickUpdate() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { throw new Error('Another run is in progress'); }
  const startMs = new Date().getTime();
  const username = getUsername();
  const results = [];
  try {
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
  } catch (error) {
    logExecution('quickUpdate', username, 'ERROR', new Date().getTime() - startMs, error.message);
    throw error;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function completeUpdate() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { throw new Error('Another run is in progress'); }
  const startMs = new Date().getTime();
  const username = getUsername();
  const results = [];
  try {
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
  } catch (error) {
    logExecution('completeUpdate', username, 'ERROR', new Date().getTime() - startMs, error.message);
    throw error;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

