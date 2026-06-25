var CONFIG = {
  WAIVER_SHEET_ID: '1_YHAQLY49p9jxwWPONZi2w6OpmlNiZb_p7K-4E-YlUQ',
  WAIVER_TAB: 'Sheet1'
};

var SOURCES = [
  {
    name: 'active spreadsheet',
    type: 'active',
    tab: '',
    columns: { team: 1, first: 13, last: 14 },
    headers: {
      team: ['Sponsoring Organization - Team Name'],
      first: ['Participant’s First Name'],
      last: ['Participants Last Name']
    }
  },
  {
    name: 'configured waiver spreadsheet',
    type: 'config',
    tab: CONFIG.WAIVER_TAB,
    columns: { team: 1, first: 3, last: 4 },
    headers: {
      team: ['Sponsoring Organization - Team Name'],
      first: ['Participant’s First Name'],
      last: ['Participants Last Name']
    }
  }
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Milwaukee Dragon Boat Festival 2026 Check-in')
    .setFaviconUrl('https://horizons-cdn.hostinger.com/e22d90ec-431c-427c-8f3e-ca594edfa29b/f27fbe95b079b69cee049756ab048b60.png')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSheetData() {
  return getWaiverData();
}

function getWaiverData() {
  var attempts = SOURCES.map(loadSource_);
  for (var i = 0; i < attempts.length; i++) {
    if (attempts[i].ok && Object.keys(attempts[i].teams).length) {
      return attempts[i].teams;
    }
  }
  throw new Error('Waiver list could not be loaded. Please ask check-in staff to verify the waiver spreadsheet connection.');
}

function loadSource_(source) {
  try {
    var sheet = getSheet_(source);
    var values = sheet.getDataRange().getDisplayValues();
    var teams = parseWaivers_(values, source);
    return { ok: true, source: source.name, teams: teams };
  } catch (e) {
    return { ok: false, source: source.name, error: e.message || String(e), teams: {} };
  }
}

function getSheet_(source) {
  var ss = source.type === 'active' ? SpreadsheetApp.getActiveSpreadsheet() : openConfiguredSpreadsheet_();
  if (source.tab) {
    var sheet = ss.getSheetByName(source.tab);
    if (!sheet) throw new Error('Tab "' + source.tab + '" was not found.');
    return sheet;
  }
  var sheets = ss.getSheets();
  if (!sheets.length) throw new Error('Spreadsheet has no tabs.');
  return sheets[0];
}

function openConfiguredSpreadsheet_() {
  if (!CONFIG.WAIVER_SHEET_ID || CONFIG.WAIVER_SHEET_ID === 'YOUR_SHEET_ID_HERE') {
    throw new Error('CONFIG.WAIVER_SHEET_ID is not set.');
  }
  return SpreadsheetApp.openById(CONFIG.WAIVER_SHEET_ID);
}

function parseWaivers_(values, source) {
  var teams = {};
  if (!values || values.length < 2) return teams;

  var header = [];
  for (var i = 0; i < values[0].length; i++) {
    header[i] = clean_(values[0][i]);
  }

  var teamCol = columnIndex_(header, source.headers.team, source.columns.team);
  var firstCol = columnIndex_(header, source.headers.first, source.columns.first);
  var lastCol = columnIndex_(header, source.headers.last, source.columns.last);

  for (var rowIndex = 1; rowIndex < values.length; rowIndex++) {
    var row = values[rowIndex];
    var team = clean_(row[teamCol]);
    var first = clean_(row[firstCol]);
    var last = clean_(row[lastCol]);
    if (!team || !first || !last) continue;
    teams[team] = teams[team] || {};
    var personKey = (first + '|' + last).toLowerCase();
    if (!teams[team][personKey]) {
      teams[team][personKey] = titleCase_(first) + ' ' + titleCase_(last);
    }
  }

  var sorted = {};
  var teamNames = [];
  for (var teamName in teams) {
    if (teams.hasOwnProperty(teamName)) {
      teamNames.push(teamName);
    }
  }
  teamNames.sort();

  for (var i = 0; i < teamNames.length; i++) {
    var nameList = [];
    var teamName2 = teamNames[i];
    for (var person in teams[teamName2]) {
      if (teams[teamName2].hasOwnProperty(person)) {
        nameList.push(teams[teamName2][person]);
      }
    }
    nameList.sort();
    sorted[teamName2] = nameList;
  }
  return sorted;
}

function columnIndex_(header, names, fallback) {
  if (!header || !header.length) return fallback;
  for (var i = 0; i < header.length; i++) {
    var cell = clean_(header[i]).toLowerCase();
    for (var j = 0; j < names.length; j++) {
      if (cell === clean_(names[j]).toLowerCase()) {
        return i;
      }
    }
  }
  return fallback;
}

function clean_(value) {
  return String(value || '').replace(/[’‘]/g, "'").replace(/\s+/g, ' ').trim();
}

function titleCase_(value) {
  var parts = clean_(value).split(' ');
  for (var i = 0; i < parts.length; i++) {
    var part = parts[i];
    if (!part) continue;
    parts[i] = part.charAt(0).toUpperCase() + part.slice(1).toLowerCase();
  }
  return parts.join(' ');
}
