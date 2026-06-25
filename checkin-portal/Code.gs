const CONFIG = {
  WAIVER_SHEET_ID: 'YOUR_SHEET_ID_HERE',
  WAIVER_TAB: 'Sheet1'
};

const SOURCES = [
  {
    name: 'active spreadsheet',
    type: 'active',
    tab: '',
    columns: { team: 1, first: 13, last: 14 }, // B, N, O
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
    columns: { team: 1, first: 3, last: 4 }, // B, D, E
    headers: {
      team: ['team name', 'team/company name', 'team/company', 'team', 'company/organization', 'organization'],
      first: ['first name', 'firstname', 'given name'],
      last: ['last name', 'lastname', 'surname', 'family name']
    }
  }
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('MDBF 2026 Check-in')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSheetData() {
  return getWaiverData();
}

function getWaiverData() {
  const attempts = SOURCES.map(loadSource_);
  const usable = attempts.find(attempt => attempt.ok && Object.keys(attempt.teams).length > 0);

  if (usable) {
    if (attempts[0] !== usable) console.warn('Using fallback waiver source:', JSON.stringify(summarizeAttempts_(attempts)));
    return usable.teams;
  }

  console.error('No usable waiver source:', JSON.stringify(summarizeAttempts_(attempts)));
  throw new Error('Waiver list could not be loaded. Please ask check-in staff to verify the waiver spreadsheet connection.');
}

function debugCheckinPortal() {
  const attempts = SOURCES.map(loadSource_);
  const selected = attempts.find(attempt => attempt.ok && Object.keys(attempt.teams).length > 0);
  const debug = {
    generatedAt: new Date().toISOString(),
    selectedSource: selected ? selected.source : null,
    teamCount: selected ? Object.keys(selected.teams).length : 0,
    teamsPreview: selected ? Object.keys(selected.teams).slice(0, 20) : [],
    sources: attempts.map(attempt => ({
      source: attempt.source,
      ok: attempt.ok,
      error: attempt.error || null,
      spreadsheetName: attempt.spreadsheetName || null,
      spreadsheetId: attempt.spreadsheetId || null,
      sheetName: attempt.sheetName || null,
      rows: attempt.rows || 0,
      columns: attempt.columns || 0,
      selectedColumns: attempt.selectedColumns || null,
      teamCount: Object.keys(attempt.teams || {}).length,
      teamsPreview: Object.keys(attempt.teams || {}).slice(0, 20),
      preview: attempt.preview || [],
      warnings: attempt.warnings || []
    }))
  };

  console.log(JSON.stringify(debug, null, 2));
  return debug;
}

function summarizeAttempts_(attempts) {
  return attempts.map(attempt => ({
    source: attempt.source,
    ok: attempt.ok,
    error: attempt.error || null,
    sheetName: attempt.sheetName || null,
    rowCount: attempt.rows || 0,
    teamCount: Object.keys(attempt.teams || {}).length,
    warnings: attempt.warnings || []
  }));
}

function loadSource_(source) {
  try {
    const sheet = getSheet_(source);
    const values = sheet.getDataRange().getDisplayValues();
    const parsed = parseWaivers_(values, source);

    return {
      ok: true,
      source: source.name,
      spreadsheetName: sheet.getParent().getName(),
      spreadsheetId: sheet.getParent().getId(),
      sheetName: sheet.getName(),
      rows: sheet.getLastRow(),
      columns: sheet.getLastColumn(),
      selectedColumns: parsed.columns,
      preview: preview_(sheet),
      teams: parsed.teams,
      warnings: parsed.warnings
    };
  } catch (err) {
    return {
      ok: false,
      source: source.name,
      error: message_(err),
      teams: {},
      warnings: []
    };
  }
}

function getSheet_(source) {
  const ss = source.type === 'active' ? activeSpreadsheet_() : configuredSpreadsheet_();

  if (source.tab) {
    const sheet = ss.getSheetByName(source.tab);
    if (!sheet) throw new Error(`Tab "${source.tab}" was not found.`);
    return sheet;
  }

  const sheets = ss.getSheets();
  if (!sheets.length) throw new Error('Spreadsheet has no tabs.');
  return sheets[0];
}

function activeSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet is linked to this Apps Script project.');
  return ss;
}

function configuredSpreadsheet_() {
  if (!CONFIG.WAIVER_SHEET_ID || CONFIG.WAIVER_SHEET_ID === 'YOUR_SHEET_ID_HERE') {
    throw new Error('CONFIG.WAIVER_SHEET_ID is not set.');
  }
  return SpreadsheetApp.openById(CONFIG.WAIVER_SHEET_ID);
}

function parseWaivers_(values, source) {
  const teams = {};
  const warnings = [];

  if (!values || values.length < 2) {
    return { teams, columns: source.columns, warnings: ['No response rows found.'] };
  }

  const header = values[0].map(value => clean_(value));
  const columns = {
    team: column_(header, source.headers.team, source.columns.team),
    first: column_(header, source.headers.first, source.columns.first),
    last: column_(header, source.headers.last, source.columns.last)
  };

  Object.keys(columns).forEach(key => {
    if (columns[key] < 0 || columns[key] >= header.length) {
      warnings.push(`${key} column is outside the sheet range.`);
    }
  });

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const team = clean_(row[columns.team]);
    const first = clean_(row[columns.first]);
    const last = clean_(row[columns.last]);

    if (!team || !first || !last) continue;
    if (!teams[team]) teams[team] = {};

    const key = normalizeKey_(`${first}|${last}`);
    if (!teams[team][key]) teams[team][key] = `${titleCase_(first)} ${titleCase_(last)}`;
  }

  const result = {};
  Object.keys(teams).sort((a, b) => a.localeCompare(b)).forEach(team => {
    result[team] = Object.values(teams[team]).sort((a, b) => a.localeCompare(b));
  });

  if (Object.keys(result).length === 0) warnings.push('No complete team/name rows found.');
  return { teams: result, columns, warnings };
}

function column_(header, names, fallback) {
  const normalizedHeader = header.map(normalizeHeader_);
  const normalizedNames = names.map(normalizeHeader_);

  for (let i = 0; i < normalizedHeader.length; i++) {
    if (normalizedNames.includes(normalizedHeader[i])) return i;
  }

  for (let i = 0; i < normalizedHeader.length; i++) {
    if (normalizedNames.some(name => name && normalizedHeader[i].includes(name))) return i;
  }

  return fallback;
}

function preview_(sheet) {
  const rows = Math.min(sheet.getLastRow(), 5);
  const columns = Math.min(sheet.getLastColumn(), 16);
  if (!rows || !columns) return [];
  return sheet.getRange(1, 1, rows, columns).getDisplayValues();
}

function clean_(value) {
  return String(value || '').replace(/\s+/g, ' ').trim();
}

function titleCase_(value) {
  return clean_(value).split(' ').map(part => {
    return part ? part.charAt(0).toUpperCase() + part.slice(1).toLowerCase() : '';
  }).join(' ');
}

function normalizeHeader_(value) {
  return clean_(value)
    .toLowerCase()
    .replace(/[’‘]/g, "'")
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function normalizeKey_(value) {
  return clean_(value).toLowerCase();
}

function message_(err) {
  return err && err.message ? err.message : String(err);
}
