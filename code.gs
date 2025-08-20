// Enhanced Chess.com Game Data Manager
// This script manages Chess.com game data across multiple sheets

// Configuration
const SHEETS = {
  CONFIG: 'Config',
  ARCHIVES: 'Archives',
  GAMES: 'Game Data',
  DAILY: 'Daily Data',
  STATS: 'Player Stats',
  PROFILE: 'Player Profile',
  LOGS: 'Execution Logs'
};

const HEADERS = {
  CONFIG: ['Setting', 'Value', 'Description'],
  ARCHIVES: ['Archive URL', 'Year-Month', 'Last Updated'],
  GAMES: [
    'Game URL',
    'Time Control',
    'Base Time (min)',
    'Increment (sec)',
    'Rated',
    'Time Class',
    'Rules',
    'Format',
    'End Time',
    'Game Duration (sec)',
    'My Rating',
    'My Color',
    'Opponent',
    'Opponent Rating',
    'Result',
    'Termination',
    'Event',
    'Site',
    'Date',
    'Round',
    'Opening',
    'ECO',
    'ECO URL',
    'UTC Date',
    'UTC Time',
    'Start Time',
    'End Date',
    'End Time',
    'Current Position',
    'Full PGN'
  ],
  DAILY: [
    'Date',
    'Bullet Win', 'Bullet Loss', 'Bullet Draw', 'Bullet Rating', 'Bullet Change', 'Bullet Time',
    'Blitz Win', 'Blitz Loss', 'Blitz Draw', 'Blitz Rating', 'Blitz Change', 'Blitz Time',
    'Rapid Win', 'Rapid Loss', 'Rapid Draw', 'Rapid Rating', 'Rapid Change', 'Rapid Time',
    'Total Games', 'Total Wins', 'Total Losses', 'Total Draws', 'Rating Sum', 'Total Rating Change', 'Total Time (sec)', 'Avg Game Duration (sec)'
  ],
  LOGS: ['Timestamp', 'Function', 'Username', 'Status', 'Execution Time (ms)', 'Notes'],
  STATS: ['Pulled At'], // Dynamic headers will be added based on API response
  PROFILE: ['Field', 'Value'] // Will be populated dynamically
};

/**
 * Initial setup - creates all necessary sheets and headers
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let headerChanges = [];
  
  // Create Config sheet
  let configSheet = ss.getSheetByName(SHEETS.CONFIG);
  let existingUsername = '';
  
  if (!configSheet) {
    configSheet = ss.insertSheet(SHEETS.CONFIG);
    headerChanges.push('Config sheet created');
  } else {
    // Preserve existing username if it exists
    const existingData = configSheet.getLastRow() > 0 ? configSheet.getRange(1, 1, configSheet.getLastRow(), 2).getValues() : [];
    const usernameRow = existingData.find(row => row[0] === 'Username');
    if (usernameRow && usernameRow[1] && usernameRow[1] !== 'Enter your username here') {
      existingUsername = usernameRow[1];
    }
  }
  
  // Setup config data
  configSheet.clear();
  configSheet.getRange(1, 1, 1, HEADERS.CONFIG.length).setValues([HEADERS.CONFIG]);
  configSheet.getRange(1, 1, 1, HEADERS.CONFIG.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  
  const configData = [
    ['Username', existingUsername || 'Enter your username here', 'Your Chess.com username'],
    ['Last Daily Fetch', '', 'Last date daily data was fetched']
  ];
  
  configSheet.getRange(2, 1, configData.length, 3).setValues(configData);
  
  // Create other sheets and check for header changes
  const sheetsToCheck = [
    { name: SHEETS.ARCHIVES, headers: HEADERS.ARCHIVES, color: '#34a853' },
    { name: SHEETS.GAMES, headers: HEADERS.GAMES, color: '#4285f4' },
    { name: SHEETS.DAILY, headers: HEADERS.DAILY, color: '#ff6d01' },
    { name: SHEETS.STATS, headers: HEADERS.STATS, color: '#fbbc04' },
    { name: SHEETS.PROFILE, headers: HEADERS.PROFILE, color: '#9c27b0' },
    { name: SHEETS.LOGS, headers: HEADERS.LOGS, color: '#34a853' }
  ];
  
  sheetsToCheck.forEach(sheetInfo => {
    let sheet = ss.getSheetByName(sheetInfo.name);
    let isNewSheet = false;
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetInfo.name);
      isNewSheet = true;
      headerChanges.push(`${sheetInfo.name} sheet created`);
    }
    
    // Check existing headers
    const existingHeaders = sheet.getLastRow() > 0 ? 
      sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0] : [];
    
    // Compare headers
    if (!isNewSheet && existingHeaders.length > 0) {
      const existingHeadersStr = existingHeaders.join('|');
      const newHeadersStr = sheetInfo.headers.join('|');
      
      if (existingHeadersStr !== newHeadersStr) {
        headerChanges.push(`${sheetInfo.name} headers updated`);
        
        // Find specific changes
        const newHeaders = sheetInfo.headers.filter(h => !existingHeaders.includes(h));
        const removedHeaders = existingHeaders.filter(h => !sheetInfo.headers.includes(h));
        
        if (newHeaders.length > 0) {
          headerChanges.push(`  - New headers: ${newHeaders.join(', ')}`);
        }
        if (removedHeaders.length > 0) {
          headerChanges.push(`  - Removed headers: ${removedHeaders.join(', ')}`);
        }
      }
    }
    
    // Set headers
    sheet.getRange(1, 1, 1, sheetInfo.headers.length).setValues([sheetInfo.headers]);
    sheet.getRange(1, 1, 1, sheetInfo.headers.length)
      .setFontWeight('bold')
      .setBackground(sheetInfo.color)
      .setFontColor(sheetInfo.name === SHEETS.STATS ? 'black' : 'white');
  });
  
  // Show header changes notification
  if (headerChanges.length > 0) {
    const message = 'Sheet Setup Complete!\n\nChanges made:\n' + headerChanges.join('\n');
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Setup Complete', 10);
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast('All sheets are up to date!', 'Setup Complete', 3);
  }
}

/**
 * Gets username from the config sheet
 */
function getUsername() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SHEETS.CONFIG);
  if (!configSheet) {
    throw new Error('Config sheet not found. Please run setupSheets() first.');
  }
  
  const configData = configSheet.getLastRow() > 1 ? configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 2).getValues() : [];
  const usernameRow = configData.find(row => row[0] === 'Username');
  
  if (!usernameRow || !usernameRow[1] || usernameRow[1] === 'Enter your username here') {
    throw new Error('Please enter your Chess.com username in the Config sheet.');
  }
  
  return usernameRow[1].toString().trim();
}

/**
 * Updates a config value
 */
function updateConfig(setting, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SHEETS.CONFIG);
  if (!configSheet) return;
  
  const configData = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 2).getValues();
  const rowIndex = configData.findIndex(row => row[0] === setting);
  
  if (rowIndex !== -1) {
    configSheet.getRange(rowIndex + 2, 2).setValue(value);
  }
}

/**
 * Gets a config value
 */
function getConfig(setting) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SHEETS.CONFIG);
  if (!configSheet) return null;
  
  const configData = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 2).getValues();
  const row = configData.find(row => row[0] === setting);
  return row ? row[1] : null;
}

/**
 * Fetches all archives for a user
 */
function fetchArchives(username) {
  const url = `https://api.chess.com/pub/player/${username}/games/archives`;
  
  try {
    const response = UrlFetchApp.fetch(url);
    if (response.getResponseCode() !== 200) {
      throw new Error(`Failed to fetch archives. Response code: ${response.getResponseCode()}`);
    }
    
    const data = JSON.parse(response.getContentText());
    return data.archives || [];
  } catch (error) {
    console.error('Error fetching archives:', error);
    throw new Error(`Failed to fetch archives: ${error.message}`);
  }
}

/**
 * Updates the archives sheet with archive URLs
 */
function updateArchivesSheet(archives) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archivesSheet = ss.getSheetByName(SHEETS.ARCHIVES);
  
  // Get existing archives
  const existingData = archivesSheet.getLastRow() > 1 ? 
    archivesSheet.getRange(2, 1, archivesSheet.getLastRow() - 1, 3).getValues() : [];
  
  const existingArchives = new Set(existingData.map(row => row[0]));
  const newArchives = [];
  const currentDate = new Date().toISOString().split('T')[0];
  
  // Process each archive
  archives.forEach((archiveUrl) => {
    const urlParts = archiveUrl.split('/');
    const yearMonth = urlParts[urlParts.length - 2] + '-' + urlParts[urlParts.length - 1];
    
    if (!existingArchives.has(archiveUrl)) {
      newArchives.push([archiveUrl, yearMonth, currentDate]);
    } else {
      // Update last updated date for existing archives
      const existingRowIndex = existingData.findIndex(row => row[0] === archiveUrl);
      if (existingRowIndex !== -1) {
        archivesSheet.getRange(existingRowIndex + 2, 3).setValue(currentDate);
      }
    }
  });
  
  // Add new archives
  if (newArchives.length > 0) {
    archivesSheet.getRange(archivesSheet.getLastRow() + 1, 1, newArchives.length, 3)
      .setValues(newArchives);
  }
  
  return archives;
}

/**
 * Gets all archives from the archives sheet
 */
function getAllArchives() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archivesSheet = ss.getSheetByName(SHEETS.ARCHIVES);
  
  if (archivesSheet.getLastRow() <= 1) return [];
  
  const archivesData = archivesSheet.getRange(2, 1, archivesSheet.getLastRow() - 1, 3).getValues();
  return archivesData.map(row => row[0]);
}

/**
 * Gets the most recent archive (current month)
 */
function getCurrentArchive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archivesSheet = ss.getSheetByName(SHEETS.ARCHIVES);
  
  if (archivesSheet.getLastRow() <= 1) return null;
  
  const archivesData = archivesSheet.getRange(2, 1, archivesSheet.getLastRow() - 1, 3).getValues();
  
  // Sort by year-month and get the most recent
  const sortedArchives = archivesData.sort((a, b) => {
    const [yearA, monthA] = a[1].split('-').map(Number);
    const [yearB, monthB] = b[1].split('-').map(Number);
    
    if (yearA !== yearB) return yearB - yearA;
    return monthB - monthA;
  });
  
  return sortedArchives.length > 0 ? sortedArchives[0][0] : null;
}

/**
 * Fetches games from a specific archive
 */
function fetchGamesFromArchive(archiveUrl) {
  try {
    const response = UrlFetchApp.fetch(archiveUrl);
    if (response.getResponseCode() !== 200) {
      throw new Error(`Failed to fetch games from ${archiveUrl}. Response code: ${response.getResponseCode()}`);
    }
    
    const data = JSON.parse(response.getContentText());
    return data.games || [];
  } catch (error) {
    console.error(`Error fetching games from ${archiveUrl}:`, error);
    return [];
  }
}

/**
 * Logs execution details to the logs sheet
 */
function logExecution(functionName, username, status, executionTime, notes = '') {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logsSheet = ss.getSheetByName(SHEETS.LOGS);
    
    if (!logsSheet) return;
    
    const logRow = [
      new Date(),
      functionName,
      username,
      status,
      executionTime,
      notes
    ];
    
    // Insert at row 2 to keep most recent logs at top
    if (logsSheet.getLastRow() > 1) {
      logsSheet.insertRows(2);
    }
    
    logsSheet.getRange(logsSheet.getLastRow() > 1 ? 2 : logsSheet.getLastRow() + 1, 1, 1, HEADERS.LOGS.length)
      .setValues([logRow]);
      
    // Color code status
    const statusCell = logsSheet.getRange(logsSheet.getLastRow() > 1 ? 2 : logsSheet.getLastRow(), 4);
    if (status === 'SUCCESS') {
      statusCell.setBackground('#d9ead3');
    } else if (status === 'ERROR') {
      statusCell.setBackground('#f4c7c3');
    } else if (status === 'WARNING') {
      statusCell.setBackground('#fce5cd');
    }
    
  } catch (error) {
    console.error('Error logging execution:', error);
  }
}

/**
 * Parses time control string to extract base time and increment
 */
function parseTimeControl(timeControl) {
  if (!timeControl) return { baseTime: '', increment: '' };
  
  try {
    const timeStr = String(timeControl);
    
    if (timeStr.includes('+')) {
      const parts = timeStr.split('+');
      const baseSeconds = parseInt(parts[0]) || 0;
      const incrementSeconds = parseInt(parts[1]) || 0;
      
      return {
        baseTime: baseSeconds / 60,
        increment: incrementSeconds
      };
    } else {
      const baseSeconds = parseInt(timeStr) || 0;
      return {
        baseTime: baseSeconds / 60,
        increment: 0
      };
    }
  } catch (error) {
    console.error('Error parsing time control:', error);
    return { baseTime: '', increment: '' };
  }
}

/**
 * Parses PGN string to extract metadata
 */
function parsePGN(pgnString) {
  if (!pgnString) return {};
  
  const result = {
    event: '',
    site: '',
    date: '',
    round: '',
    opening: '',
    eco: '',
    ecoUrl: '',
    termination: '',
    utcDate: '',
    utcTime: '',
    startTime: '',
    endDate: '',
    endTime: '',
    currentPosition: ''
  };
  
  try {
    const lines = pgnString.split('\n');
    
    for (const line of lines) {
      const trimmedLine = line.trim();
      
      if (trimmedLine.startsWith('[') && trimmedLine.endsWith(']')) {
        const match = trimmedLine.match(/\[(\w+)\s+"([^"]+)"\]/);
        if (match) {
          const key = match[1].toLowerCase();
          const value = match[2];
          
          if (key === 'event') result.event = value;
          else if (key === 'site') result.site = value;
          else if (key === 'date') result.date = value;
          else if (key === 'round') result.round = value;
          else if (key === 'opening') result.opening = value;
          else if (key === 'eco') result.eco = value;
          else if (key === 'ecourl' || key === 'openingurl' || key === 'link') result.ecoUrl = value;
          else if (key === 'termination') result.termination = value;
          else if (key === 'utcdate') result.utcDate = value;
          else if (key === 'utctime') result.utcTime = value;
          else if (key === 'starttime') result.startTime = value;
          else if (key === 'enddate') result.endDate = value;
          else if (key === 'endtime') result.endTime = value;
          else if (key === 'currentposition') result.currentPosition = value;
        }
      }
    }
    
  } catch (error) {
    console.error('Error parsing PGN:', error);
  }
  
  return result;
}

/**
 * Computes game duration in seconds from PGN start/end times
 */
function computeGameDuration(startTime, endTime, startDate, endDate) {
  if (!startTime || !endTime) return '';
  
  try {
    let startDateTime;
    if (startDate && startTime) {
      startDateTime = new Date(`${startDate} ${startTime}`);
    } else if (startTime) {
      startDateTime = new Date(startTime);
    } else {
      return '';
    }
    
    let endDateTime;
    if (endDate && endTime) {
      endDateTime = new Date(`${endDate} ${endTime}`);
    } else if (endTime) {
      endDateTime = new Date(endTime);
    } else {
      return '';
    }
    
    if (endDateTime < startDateTime) {
      endDateTime.setDate(endDateTime.getDate() + 1);
    }
    
    const durationMs = endDateTime.getTime() - startDateTime.getTime();
    return Math.round(durationMs / 1000);
    
  } catch (error) {
    console.error('Error computing game duration:', error);
    return '';
  }
}

/**
 * Computes a normalized game format based on rules and time class
 */
function computeFormat(rules, timeClass) {
  const normalizedRules = (rules || '').toLowerCase();
  const normalizedTimeClass = (timeClass || '').toLowerCase();
  const isChess960 = normalizedRules.includes('960');
  const isStandardChess = normalizedRules === 'chess' || normalizedRules === '';
  
  if (isChess960 && normalizedTimeClass === 'daily') {
    return 'daily960';
  }
  if (isStandardChess) {
    return normalizedTimeClass || 'unknown';
  }
  return normalizedRules;
}

/**
 * Formats a Date to M/D/YYYY HH:MM:SS (24h)
 */
function formatDateTime(dateObj) {
  if (!(dateObj instanceof Date) || isNaN(dateObj.getTime())) return '';
  const pad2 = (n) => String(n).padStart(2, '0');
  const m = dateObj.getMonth() + 1;
  const d = dateObj.getDate();
  const y = dateObj.getFullYear();
  const hh = pad2(dateObj.getHours());
  const mm = pad2(dateObj.getMinutes());
  const ss = pad2(dateObj.getSeconds());
  return `${m}/${d}/${y} ${hh}:${mm}:${ss}`;
}

/**
 * Formats date to M/D/YYYY (no time)
 */
function formatDate(dateObj) {
  if (!(dateObj instanceof Date) || isNaN(dateObj.getTime())) return '';
  const m = dateObj.getMonth() + 1;
  const d = dateObj.getDate();
  const y = dateObj.getFullYear();
  return `${m}/${d}/${y}`;
}

/**
 * Determines player perspective data from game
 */
function getPlayerPerspective(game, username) {
  const usernameLower = username.toLowerCase();
  const whiteUser = String(game.white?.username || '').toLowerCase();
  const blackUser = String(game.black?.username || '').toLowerCase();
  
  const isWhite = whiteUser === usernameLower;
  const isBlack = blackUser === usernameLower;
  
  if (!isWhite && !isBlack) {
    return {
      myRating: '',
      myColor: '',
      opponent: '',
      opponentRating: '',
      result: '',
      termination: ''
    };
  }
  
  const myColor = isWhite ? 'White' : 'Black';
  const myRating = isWhite ? (game.white?.rating || '') : (game.black?.rating || '');
  const opponent = isWhite ? (game.black?.username || '') : (game.white?.username || '');
  const opponentRating = isWhite ? (game.black?.rating || '') : (game.white?.rating || '');
  
  const myResult = isWhite ? (game.white?.result || '') : (game.black?.result || '');
  const opponentResult = isWhite ? (game.black?.result || '') : (game.white?.result || '');
  
  let result = '';
  let termination = '';
  
  const myResultLower = String(myResult).toLowerCase();
  const opponentResultLower = String(opponentResult).toLowerCase();
  
  if (myResultLower === 'win') {
    result = 'Win';
    termination = opponentResultLower;
  } else if (myResultLower === 'checkmated' || myResultLower === 'resigned' || myResultLower === 'timeout') {
    result = 'Loss';
    termination = myResultLower;
  } else {
    result = 'Draw';
    termination = myResultLower || opponentResultLower || 'draw';
  }
  
  return {
    myRating,
    myColor,
    opponent,
    opponentRating,
    result,
    termination
  };
}

/**
 * Converts a game object to a spreadsheet row
 */
function gameToRow(game, username) {
  const pgn = game.pgn || '';
  const metadata = parsePGN(pgn);
  const rules = (game.rules || '').toLowerCase();
  const timeClass = (game.time_class || '').toLowerCase();
  const format = computeFormat(rules, timeClass);
  
  const timeControlData = parseTimeControl(game.time_control);
  const playerData = getPlayerPerspective(game, username);
  const duration = computeGameDuration(metadata.startTime, metadata.endTime, metadata.utcDate, metadata.endDate);
  
  return [
    game.url || '',
    game.time_control || '',
    timeControlData.baseTime,
    timeControlData.increment,
    game.rated || false,
    game.time_class || '',
    game.rules || '',
    format,
    game.end_time ? new Date(game.end_time * 1000) : '',
    duration || '',
    playerData.myRating,
    playerData.myColor,
    playerData.opponent,
    playerData.opponentRating,
    playerData.result,
    playerData.termination,
    metadata.event || '',
    metadata.site || '',
    metadata.date || '',
    metadata.round || '',
    metadata.opening || '', // Just use the opening name from PGN
    metadata.eco || '',
    metadata.ecoUrl || '',
    metadata.utcDate || '',
    metadata.utcTime || '',
    metadata.startTime || '',
    metadata.endDate || '',
    metadata.endTime || '',
    metadata.currentPosition || '',
    pgn
  ];
}

/**
 * Gets existing game URLs to prevent duplicates
 */
function getExistingGameUrls() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  
  if (gamesSheet.getLastRow() <= 1) return new Set();
  
  const gameUrls = gamesSheet.getRange(2, 1, gamesSheet.getLastRow() - 1, 1).getValues();
  return new Set(gameUrls.flat().filter(url => url));
}

/**
 * Adds new games to the top of the games sheet
 */
function addNewGames(games, username) {
  if (!games || games.length === 0) return 0;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  const existingUrls = getExistingGameUrls();
  
  const newGameRows = games
    .filter(game => !existingUrls.has(game.url))
    .map(game => gameToRow(game, username))
    .reverse();
  
  if (newGameRows.length === 0) {
    console.log('No new games to add (all games already exist)');
    return 0;
  }
  
  if (gamesSheet.getLastRow() > 1) {
    gamesSheet.insertRows(2, newGameRows.length);
  }
  
  gamesSheet.getRange(2, 1, newGameRows.length, HEADERS.GAMES.length)
    .setValues(newGameRows);
  
  console.log(`Added ${newGameRows.length} new games`);
  return newGameRows.length;
}

/**
 * Processes daily data from game data with complete date sequence and forward-filled ratings
 */
function processDailyData(startDate = null, endDate = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  const dailySheet = ss.getSheetByName(SHEETS.DAILY);
  const profileSheet = ss.getSheetByName(SHEETS.PROFILE);
  
  if (gamesSheet.getLastRow() <= 1) {
    console.log('No games data to process');
    return 0;
  }
  
  // Get player join date from profile sheet
  let playerJoinDate = null;
  if (profileSheet && profileSheet.getLastRow() > 1) {
    const profileData = profileSheet.getRange(2, 1, profileSheet.getLastRow() - 1, 2).getValues();
    const joinedRow = profileData.find(row => row[0] === 'joined');
    if (joinedRow && joinedRow[1] instanceof Date) {
      playerJoinDate = new Date(joinedRow[1].getFullYear(), joinedRow[1].getMonth(), joinedRow[1].getDate());
    }
  }
  
  // Set date range
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const actualStartDate = startDate || playerJoinDate || new Date('2020-01-01'); // Fallback if no profile
  actualStartDate.setHours(0, 0, 0, 0);
  
  const actualEndDate = endDate || today;
  actualEndDate.setHours(0, 0, 0, 0);
  
  console.log(`Processing daily data from ${formatDate(actualStartDate)} to ${formatDate(actualEndDate)}`);
  
  // Get all games data
  const gamesData = gamesSheet.getRange(2, 1, gamesSheet.getLastRow() - 1, HEADERS.GAMES.length).getValues();
  const gamesByDate = {};
  
  // Group games by date
  gamesData.forEach(row => {
    const endTime = row[8]; // End Time column
    if (!endTime || !(endTime instanceof Date)) return;
    
    const dateStr = formatDate(endTime);
    const gameDate = new Date(endTime.getFullYear(), endTime.getMonth(), endTime.getDate());
    
    // Filter by date range
    if (gameDate < actualStartDate || gameDate > actualEndDate) return;
    
    if (!gamesByDate[dateStr]) {
      gamesByDate[dateStr] = [];
    }
    gamesByDate[dateStr].push(row);
  });
  
  // Clear the daily sheet and rebuild completely for proper ordering
  dailySheet.clear();
  dailySheet.getRange(1, 1, 1, HEADERS.DAILY.length).setValues([HEADERS.DAILY]);
  dailySheet.getRange(1, 1, 1, HEADERS.DAILY.length)
    .setFontWeight('bold')
    .setBackground('#ff6d01')
    .setFontColor('white');
  
  // Generate complete date sequence (reverse chronological)
  const allDates = [];
  const currentDate = new Date(actualEndDate);
  
  while (currentDate >= actualStartDate) {
    allDates.push(new Date(currentDate));
    currentDate.setDate(currentDate.getDate() - 1);
  }
  
  console.log(`Generated ${allDates.length} dates to process`);
  
  // First pass: Build ratings history going forward in time to establish last known ratings
  const ratingHistory = {};
  let globalLastKnownRatings = { bullet: '', blitz: '', rapid: '' };
  
  // Process games chronologically to build rating history
  const sortedDates = allDates.slice().reverse(); // Forward chronological order
  sortedDates.forEach(date => {
    const dateStr = formatDate(date);
    const dayGames = gamesByDate[dateStr] || [];
    
    const dayRatings = { bullet: '', blitz: '', rapid: '' };
    
    dayGames.forEach(game => {
      const timeClass = String(game[5] || '').toLowerCase();
      const myRating = parseFloat(game[10]) || 0;
      
      if (myRating > 0) {
        if (timeClass === 'bullet') dayRatings.bullet = myRating;
        else if (timeClass === 'blitz') dayRatings.blitz = myRating;
        else if (timeClass === 'rapid') dayRatings.rapid = myRating;
      }
    });
    
    // Update global last known ratings
    if (dayRatings.bullet !== '') globalLastKnownRatings.bullet = dayRatings.bullet;
    if (dayRatings.blitz !== '') globalLastKnownRatings.blitz = dayRatings.blitz;
    if (dayRatings.rapid !== '') globalLastKnownRatings.rapid = dayRatings.rapid;
    
    // Store the ending ratings for this date (either from games or carried forward)
    ratingHistory[dateStr] = {
      bullet: dayRatings.bullet !== '' ? dayRatings.bullet : globalLastKnownRatings.bullet,
      blitz: dayRatings.blitz !== '' ? dayRatings.blitz : globalLastKnownRatings.blitz,
      rapid: dayRatings.rapid !== '' ? dayRatings.rapid : globalLastKnownRatings.rapid
    };
  });
  
  const dailyRows = [];
  
  // Second pass: Process each date in reverse chronological order for display
  allDates.forEach((date, index) => {
    const dateStr = formatDate(date);
    const dayGames = gamesByDate[dateStr] || [];
    
    const stats = {
      date: new Date(date),
      bullet: { win: 0, loss: 0, draw: 0, ratings: [], time: 0 },
      blitz: { win: 0, loss: 0, draw: 0, ratings: [], time: 0 },
      rapid: { win: 0, loss: 0, draw: 0, ratings: [], time: 0 },
      total: { games: dayGames.length, wins: 0, losses: 0, draws: 0, ratings: [], time: 0 }
    };
    
    // Process games for this date
    dayGames.forEach(game => {
      const timeClass = String(game[5] || '').toLowerCase();
      const result = String(game[14] || '').toLowerCase();
      const myRating = parseFloat(game[10]) || 0;
      const duration = parseFloat(game[9]) || 0;
      
      let category = null;
      if (timeClass === 'bullet') category = 'bullet';
      else if (timeClass === 'blitz') category = 'blitz';
      else if (timeClass === 'rapid') category = 'rapid';
      
      if (category) {
        // Count results
        if (result === 'win') {
          stats[category].win++;
          stats.total.wins++;
        } else if (result === 'loss') {
          stats[category].loss++;
          stats.total.losses++;
        } else {
          stats[category].draw++;
          stats.total.draws++;
        }
        
        if (myRating > 0) {
          stats[category].ratings.push(myRating);
          stats.total.ratings.push(myRating);
        }
        
        if (duration > 0) {
          stats[category].time += duration;
          stats.total.time += duration;
        }
      }
    });
    
    // Get ratings from history (these are already forward-filled)
    const currentRatings = ratingHistory[dateStr] || { bullet: '', blitz: '', rapid: '' };
    const bulletRating = currentRatings.bullet;
    const blitzRating = currentRatings.blitz;
    const rapidRating = currentRatings.rapid;
    
    // Calculate rating changes: today's rating - yesterday's rating
    let bulletChange = '';
    let blitzChange = '';
    let rapidChange = '';
    let totalRatingChange = '';
    
    // Get previous day's ratings for comparison
    if (index < allDates.length - 1) {
      const previousDate = allDates[index + 1];
      const previousDateStr = formatDate(previousDate);
      const previousRatings = ratingHistory[previousDateStr] || { bullet: '', blitz: '', rapid: '' };
      
      if (bulletRating !== '' && previousRatings.bullet !== '') {
        bulletChange = bulletRating - previousRatings.bullet;
        // Only show change if it's not zero
        if (bulletChange === 0) bulletChange = '';
      }
      if (blitzRating !== '' && previousRatings.blitz !== '') {
        blitzChange = blitzRating - previousRatings.blitz;
        if (blitzChange === 0) blitzChange = '';
      }
      if (rapidRating !== '' && previousRatings.rapid !== '') {
        rapidChange = rapidRating - previousRatings.rapid;
        if (rapidChange === 0) rapidChange = '';
      }
      
      // Calculate total rating change: sum of all individual changes
      const bulletChangeVal = typeof bulletChange === 'number' ? bulletChange : 0;
      const blitzChangeVal = typeof blitzChange === 'number' ? blitzChange : 0;
      const rapidChangeVal = typeof rapidChange === 'number' ? rapidChange : 0;
      const totalChange = bulletChangeVal + blitzChangeVal + rapidChangeVal;
      
      if (totalChange !== 0) {
        totalRatingChange = totalChange;
      }
    }
    
    // Calculate rating sum: simple sum of current ratings
    const bulletRatingVal = typeof bulletRating === 'number' ? bulletRating : 0;
    const blitzRatingVal = typeof blitzRating === 'number' ? blitzRating : 0;
    const rapidRatingVal = typeof rapidRating === 'number' ? rapidRating : 0;
    const ratingSum = bulletRatingVal + blitzRatingVal + rapidRatingVal;
    
    const avgGameDuration = stats.total.time > 0 ? Math.round(stats.total.time / stats.total.games) : '';
    
    // Create the daily row
    const dailyRow = [
      stats.date,
      stats.bullet.win, stats.bullet.loss, stats.bullet.draw, bulletRating, bulletChange, stats.bullet.time,
      stats.blitz.win, stats.blitz.loss, stats.blitz.draw, blitzRating, blitzChange, stats.blitz.time,
      stats.rapid.win, stats.rapid.loss, stats.rapid.draw, rapidRating, rapidChange, stats.rapid.time,
      stats.total.games, stats.total.wins, stats.total.losses, stats.total.draws, 
      ratingSum, totalRatingChange, stats.total.time, avgGameDuration
    ];
    
    dailyRows.push(dailyRow);
  });
  
  // Add all daily rows to the sheet
  if (dailyRows.length > 0) {
    dailySheet.getRange(2, 1, dailyRows.length, HEADERS.DAILY.length)
      .setValues(dailyRows);
  }
  
  console.log(`Processed daily data: ${dailyRows.length} daily summaries (${Object.keys(gamesByDate).length} days with games)`);
  return dailyRows.length;
}

/**
 * Fetches player statistics from Chess.com API
 */
function fetchPlayerStats(username) {
  const url = `https://api.chess.com/pub/player/${username}/stats`;
  
  try {
    const response = UrlFetchApp.fetch(url);
    if (response.getResponseCode() !== 200) {
      throw new Error(`Failed to fetch player stats. Response code: ${response.getResponseCode()}`);
    }
    
    const data = JSON.parse(response.getContentText());
    return data;
  } catch (error) {
    console.error('Error fetching player stats:', error);
    throw new Error(`Failed to fetch player stats: ${error.message}`);
  }
}

/**
 * Updates player stats sheet with current stats
 */
function updatePlayerStats(stats) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const statsSheet = ss.getSheetByName(SHEETS.STATS);
  
  // Flatten the stats object for spreadsheet format
  const flattenedStats = {};
  const timestamp = new Date();
  
  function flattenObject(obj, prefix = '') {
    for (const key in obj) {
      if (obj.hasOwnProperty(key)) {
        const value = obj[key];
        const newKey = prefix ? `${prefix}.${key}` : key;
        
        if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
          flattenObject(value, newKey);
        } else {
          flattenedStats[newKey] = value;
        }
      }
    }
  }
  
  flattenObject(stats);
  
  // Add timestamp
  flattenedStats['pulled_at'] = timestamp;
  
  // Get existing headers or create new ones
  const existingHeaders = statsSheet.getLastRow() > 0 ? 
    statsSheet.getRange(1, 1, 1, statsSheet.getLastColumn() || 1).getValues()[0] : [];
  
  const allKeys = ['pulled_at', ...Object.keys(flattenedStats).filter(key => key !== 'pulled_at').sort()];
  
  // Update headers if needed
  if (existingHeaders.length === 0 || JSON.stringify(existingHeaders) !== JSON.stringify(allKeys)) {
    statsSheet.clear();
    statsSheet.getRange(1, 1, 1, allKeys.length).setValues([allKeys]);
    statsSheet.getRange(1, 1, 1, allKeys.length)
      .setFontWeight('bold')
      .setBackground('#fbbc04')
      .setFontColor('black');
  }
  
  // Create data row
  const dataRow = allKeys.map(key => flattenedStats[key] || '');
  
  // Add the new row
  statsSheet.getRange(statsSheet.getLastRow() + 1, 1, 1, allKeys.length)
    .setValues([dataRow]);
  
  console.log('Player stats updated successfully');
}

/**
 * Fetches player profile from Chess.com API
 */
function fetchPlayerProfile(username) {
  const url = `https://api.chess.com/pub/player/${username}`;
  
  try {
    const response = UrlFetchApp.fetch(url);
    if (response.getResponseCode() !== 200) {
      throw new Error(`Failed to fetch player profile. Response code: ${response.getResponseCode()}`);
    }
    
    const data = JSON.parse(response.getContentText());
    return data;
  } catch (error) {
    console.error('Error fetching player profile:', error);
    throw new Error(`Failed to fetch player profile: ${error.message}`);
  }
}

/**
 * Updates player profile sheet
 */
function updatePlayerProfile(profile) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const profileSheet = ss.getSheetByName(SHEETS.PROFILE);
  
  // Clear existing data
  profileSheet.clear();
  profileSheet.getRange(1, 1, 1, HEADERS.PROFILE.length).setValues([HEADERS.PROFILE]);
  profileSheet.getRange(1, 1, 1, HEADERS.PROFILE.length)
    .setFontWeight('bold')
    .setBackground('#9c27b0')
    .setFontColor('white');
  
  // Convert profile to field-value pairs
  const profileData = [];
  profileData.push(['Updated At', new Date()]);
  
  for (const key in profile) {
    if (profile.hasOwnProperty(key)) {
      const value = profile[key];
      let displayValue = value;
      
      // Format specific fields
      if (key === 'joined' && typeof value === 'number') {
        displayValue = new Date(value * 1000);
      } else if (key === 'last_login' && typeof value === 'number') {
        displayValue = new Date(value * 1000);
      } else if (typeof value === 'object' && value !== null) {
        displayValue = JSON.stringify(value);
      }
      
      profileData.push([key, displayValue]);
    }
  }
  
  if (profileData.length > 1) {
    profileSheet.getRange(2, 1, profileData.length, 2).setValues(profileData);
  }
  
  console.log('Player profile updated successfully');
}

/**
 * Fetches games from current month archive only
 */
function fetchCurrentMonthGames() {
  const startTime = new Date().getTime();
  let username = '';
  
  try {
    username = getUsername();
    console.log(`Starting current month game fetch for user: ${username}`);
    
    const currentArchive = getCurrentArchive();
    if (!currentArchive) {
      logExecution('fetchCurrentMonthGames', username, 'WARNING', new Date().getTime() - startTime, 'No current archive found');
      SpreadsheetApp.getActiveSpreadsheet().toast('No current archive found. Run updateArchives() first.', 'Warning', 5);
      return;
    }
    
    console.log(`Fetching games from current archive: ${currentArchive}`);
    const games = fetchGamesFromArchive(currentArchive);
    console.log(`Retrieved ${games.length} games from current archive`);
    
    // Add new games to the sheet
    const totalNewGames = addNewGames(games, username);
    
    const executionTime = new Date().getTime() - startTime;
    const notes = `Current archive: ${currentArchive}, Total games: ${games.length}, New games: ${totalNewGames}`;
    
    logExecution('fetchCurrentMonthGames', username, 'SUCCESS', executionTime, notes);
    
    const message = `Current month game fetch completed!\n\nArchive: ${currentArchive}\nTotal games retrieved: ${games.length}\nNew games added: ${totalNewGames}\nExecution time: ${Math.round(executionTime/1000)}s`;
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Fetch Complete', 8);
    
  } catch (error) {
    const executionTime = new Date().getTime() - startTime;
    console.error('Error in fetchCurrentMonthGames:', error);
    logExecution('fetchCurrentMonthGames', username, 'ERROR', executionTime, error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Error: ${error.message}`, 'Fetch Error', 10);
    throw error;
  }
}

/**
 * Main function to fetch all games from all archives
 */
function fetchAllGames() {
  const startTime = new Date().getTime();
  let username = '';
  
  try {
    username = getUsername();
    console.log(`Starting full game fetch for user: ${username}`);
    
    // Get all archives
    const allArchives = getAllArchives();
    console.log(`Found ${allArchives.length} total archives`);
    
    if (allArchives.length === 0) {
      logExecution('fetchAllGames', username, 'WARNING', new Date().getTime() - startTime, 'No archives found');
      SpreadsheetApp.getActiveSpreadsheet().toast('No archives found. Run updateArchives() first.', 'Warning', 5);
      return;
    }
    
    let totalNewGames = 0;
    const allGames = [];
    
    // Fetch games from each archive
    for (const archiveUrl of allArchives) {
      console.log(`Fetching games from: ${archiveUrl}`);
      const games = fetchGamesFromArchive(archiveUrl);
      allGames.push(...games);
      console.log(`Retrieved ${games.length} games from archive`);
      
      // Add small delay to be respectful to the API
      Utilities.sleep(200);
    }
    
    console.log(`Total games retrieved: ${allGames.length}`);
    
    // Add new games to the sheet
    totalNewGames = addNewGames(allGames, username);
    
    const executionTime = new Date().getTime() - startTime;
    const notes = `Archives checked: ${allArchives.length}, Total games: ${allGames.length}, New games: ${totalNewGames}`;
    
    logExecution('fetchAllGames', username, 'SUCCESS', executionTime, notes);
    
    const message = `Full game fetch completed!\n\nArchives checked: ${allArchives.length}\nTotal games retrieved: ${allGames.length}\nNew games added: ${totalNewGames}\nExecution time: ${Math.round(executionTime/1000)}s`;
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Fetch Complete', 8);
    
  } catch (error) {
    const executionTime = new Date().getTime() - startTime;
    console.error('Error in fetchAllGames:', error);
    logExecution('fetchAllGames', username, 'ERROR', executionTime, error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Error: ${error.message}`, 'Fetch Error', 10);
    throw error;
  }
}

/**
 * Updates archives list from Chess.com API
 */
function updateArchives() {
  const startTime = new Date().getTime();
  let username = '';
  
  try {
    username = getUsername();
    console.log(`Updating archives for user: ${username}`);
    
    const archives = fetchArchives(username);
    console.log(`Retrieved ${archives.length} archives`);
    
    updateArchivesSheet(archives);
    console.log(`Archives sheet updated with ${archives.length} archives`);
    
    const executionTime = new Date().getTime() - startTime;
    const notes = `Total archives: ${archives.length}`;
    
    logExecution('updateArchives', username, 'SUCCESS', executionTime, notes);
    
    const message = `Archives updated!\n\nTotal archives: ${archives.length}\nExecution time: ${Math.round(executionTime/1000)}s`;
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Archives Updated', 5);
    
  } catch (error) {
    const executionTime = new Date().getTime() - startTime;
    console.error('Error in updateArchives:', error);
    logExecution('updateArchives', username, 'ERROR', executionTime, error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Error: ${error.message}`, 'Archives Error', 10);
    throw error;
  }
}

/**
 * Updates player statistics
 */
function updateStats() {
  const startTime = new Date().getTime();
  let username = '';
  
  try {
    username = getUsername();
    console.log(`Updating stats for user: ${username}`);
    
    const stats = fetchPlayerStats(username);
    updatePlayerStats(stats);
    
    const executionTime = new Date().getTime() - startTime;
    logExecution('updateStats', username, 'SUCCESS', executionTime, 'Player stats updated');
    
    SpreadsheetApp.getActiveSpreadsheet().toast('Player statistics updated successfully!', 'Stats Updated', 3);
    
  } catch (error) {
    const executionTime = new Date().getTime() - startTime;
    console.error('Error in updateStats:', error);
    logExecution('updateStats', username, 'ERROR', executionTime, error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Error: ${error.message}`, 'Stats Error', 10);
    throw error;
  }
}

/**
 * Updates player profile
 */
function updateProfile() {
  const startTime = new Date().getTime();
  let username = '';
  
  try {
    username = getUsername();
    console.log(`Updating profile for user: ${username}`);
    
    const profile = fetchPlayerProfile(username);
    updatePlayerProfile(profile);
    
    const executionTime = new Date().getTime() - startTime;
    logExecution('updateProfile', username, 'SUCCESS', executionTime, 'Player profile updated');
    
    SpreadsheetApp.getActiveSpreadsheet().toast('Player profile updated successfully!', 'Profile Updated', 3);
    
  } catch (error) {
    const executionTime = new Date().getTime() - startTime;
    console.error('Error in updateProfile:', error);
    logExecution('updateProfile', username, 'ERROR', executionTime, error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Error: ${error.message}`, 'Profile Error', 10);
    throw error;
  }
}

/**
 * Processes daily data from existing game data
 */
function updateDailyData() {
  const startTime = new Date().getTime();
  let username = '';
  
  try {
    username = getUsername();
    console.log(`Processing daily data for user: ${username}`);
    
    const newDailyRows = processDailyData();
    
    const executionTime = new Date().getTime() - startTime;
    const notes = `New daily summaries: ${newDailyRows}`;
    
    logExecution('updateDailyData', username, 'SUCCESS', executionTime, notes);
    
    const message = `Daily data processing complete!\n\nNew daily summaries: ${newDailyRows}\nExecution time: ${Math.round(executionTime/1000)}s`;
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Daily Data Updated', 5);
    
  } catch (error) {
    const executionTime = new Date().getTime() - startTime;
    console.error('Error in updateDailyData:', error);
    logExecution('updateDailyData', username, 'ERROR', executionTime, error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Error: ${error.message}`, 'Daily Data Error', 10);
    throw error;
  }
}

/**
 * Quick update - runs current month games, daily data, stats, and profile
 */
function quickUpdate() {
  const startTime = new Date().getTime();
  let username = '';
  const results = [];
  
  try {
    username = getUsername();
    console.log(`Starting quick update for user: ${username}`);
    
    // Update archives first
    try {
      updateArchives();
      results.push('✓ Archives updated');
    } catch (error) {
      results.push('✗ Archives failed: ' + error.message);
    }
    
    // Fetch current month games
    try {
      fetchCurrentMonthGames();
      results.push('✓ Current month games fetched');
    } catch (error) {
      results.push('✗ Current month games failed: ' + error.message);
    }
    
    // Update daily data
    try {
      updateDailyData();
      results.push('✓ Daily data processed');
    } catch (error) {
      results.push('✗ Daily data failed: ' + error.message);
    }
    
    // Update stats
    try {
      updateStats();
      results.push('✓ Stats updated');
    } catch (error) {
      results.push('✗ Stats failed: ' + error.message);
    }
    
    // Update profile
    try {
      updateProfile();
      results.push('✓ Profile updated');
    } catch (error) {
      results.push('✗ Profile failed: ' + error.message);
    }
    
    const executionTime = new Date().getTime() - startTime;
    const notes = results.join(', ');
    
    logExecution('quickUpdate', username, 'SUCCESS', executionTime, notes);
    
    const message = `Quick update finished!\n\n${results.join('\n')}\n\nTotal execution time: ${Math.round(executionTime/1000)}s`;
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Quick Update Done', 10);
    
  } catch (error) {
    const executionTime = new Date().getTime() - startTime;
    console.error('Error in quickUpdate:', error);
    logExecution('quickUpdate', username, 'ERROR', executionTime, error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Error: ${error.message}`, 'Quick Update Error', 10);
    throw error;
  }
}

/**
 * Complete update - runs all update functions including full game fetch
 */
function completeUpdate() {
  const startTime = new Date().getTime();
  let username = '';
  const results = [];
  
  try {
    username = getUsername();
    console.log(`Starting complete update for user: ${username}`);
    
    // Update archives
    try {
      updateArchives();
      results.push('✓ Archives updated');
    } catch (error) {
      results.push('✗ Archives failed: ' + error.message);
    }
    
    // Fetch all games
    try {
      fetchAllGames();
      results.push('✓ All games fetched');
    } catch (error) {
      results.push('✗ All games failed: ' + error.message);
    }
    
    // Update daily data
    try {
      updateDailyData();
      results.push('✓ Daily data processed');
    } catch (error) {
      results.push('✗ Daily data failed: ' + error.message);
    }
    
    // Update stats
    try {
      updateStats();
      results.push('✓ Stats updated');
    } catch (error) {
      results.push('✗ Stats failed: ' + error.message);
    }
    
    // Update profile
    try {
      updateProfile();
      results.push('✓ Profile updated');
    } catch (error) {
      results.push('✗ Profile failed: ' + error.message);
    }
    
    const executionTime = new Date().getTime() - startTime;
    const notes = results.join(', ');
    
    logExecution('completeUpdate', username, 'SUCCESS', executionTime, notes);
    
    const message = `Complete update finished!\n\n${results.join('\n')}\n\nTotal execution time: ${Math.round(executionTime/1000)}s`;
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Complete Update Done', 10);
    
  } catch (error) {
    const executionTime = new Date().getTime() - startTime;
    console.error('Error in completeUpdate:', error);
    logExecution('completeUpdate', username, 'ERROR', executionTime, error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Error: ${error.message}`, 'Complete Update Error', 10);
    throw error;
  }
}
