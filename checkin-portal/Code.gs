const CONFIG = {
  // Optional: paste the waiver spreadsheet ID here if this Apps Script project is
  // not bound directly to the waiver spreadsheet. If left blank, the script will
  // use the active/bound spreadsheet.
  WAIVER_SHEET_ID: '',
  // Optional: set to the exact waiver response tab name. If left blank, the first
  // sheet in the waiver spreadsheet is used.
  WAIVER_TAB: ''
};

// serves webpage
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('MDBF 2026 Check-in')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1'); // mobile friendly
}

// fetches and organizes all the waiver data
function getSheetData() {
  return getWaiverData();
}

function getWaiverData() {
  const sheet = getWaiverSheet_();
  const data = sheet.getDataRange().getDisplayValues();
  const teams = {};
  if (data.length < 2) return teams;

  const header = data[0].map(value => String(value || '').trim().toLowerCase());
  const teamCol = findColumn_(header, ['team name', 'team/company name', 'team/company', 'team', 'company/organization', 'organization'], 1);
  const firstCol = findColumn_(header, ['first name', 'firstname', 'given name'], 3);
  const lastCol = findColumn_(header, ['last name', 'lastname', 'surname', 'family name'], 4);

  // loop through rows (skip header)
  for (let i = 1; i < data.length; i++) {
    const team = String(data[i][teamCol] || '').trim();
    const first = String(data[i][firstCol] || '').trim();
    const last = String(data[i][lastCol] || '').trim();
    // capitalize names properly
    const formatName = (str) => String(str || '').trim().split(/\s+/).map(part =>
      part.charAt(0).toUpperCase() + part.slice(1).toLowerCase()
    ).join(' ');
    const fullName = first && last ? `${formatName(first)} ${formatName(last)}` : '';
    if (team && fullName) {
      // Initialize the team in the object if it doesn't exist
      if (!teams[team]) {
        teams[team] = new Set(); // Use Set to avoid duplicates
      }
      teams[team].add(fullName); // add the full name to the team's Set
    }
  }
  // convert Sets to Arrays, sort them alphabetically, and remove duplicates
  const result = {};
  Object.keys(teams).sort().forEach(t => {
    result[t] = Array.from(teams[t]).sort();
  });
  return result;
}

function getWaiverSheet_() {
  const spreadsheet = CONFIG.WAIVER_SHEET_ID
    ? SpreadsheetApp.openById(CONFIG.WAIVER_SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error('No active waiver spreadsheet found. Either bind this Apps Script project to the waiver response spreadsheet, or set CONFIG.WAIVER_SHEET_ID in Code.gs.');
  }

  const sheet = CONFIG.WAIVER_TAB
    ? spreadsheet.getSheetByName(CONFIG.WAIVER_TAB)
    : spreadsheet.getSheets()[0];

  if (!sheet) {
    throw new Error(CONFIG.WAIVER_TAB
      ? `Waiver tab "${CONFIG.WAIVER_TAB}" was not found.`
      : 'No sheets were found in the waiver spreadsheet.');
  }

  return sheet;
}

function findColumn_(header, names, fallbackIndex) {
  const exact = names.map(name => name.toLowerCase());
  for (let i = 0; i < header.length; i++) {
    if (exact.includes(header[i])) return i;
  }
  for (let i = 0; i < header.length; i++) {
    if (names.some(name => header[i].includes(name.toLowerCase()))) return i;
  }
  return fallbackIndex;
}
