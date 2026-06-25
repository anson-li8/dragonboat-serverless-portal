// serves webpage
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('MDBF 2026 Check-in')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1'); // mobile friendly
}

// fetches and organizes all the waiver data
function getSheetData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; 
  const data = sheet.getDataRange().getValues();
  const teams = {};
  // loop through rows (skip header)
  for (let i = 1; i < data.length; i++) {
    // use columns from before: Team = Col B(1), First = Col D(3), Last = Col E(4)
    const team = String(data[i][1] || '').trim();
    const first = String(data[i][3] || '').trim();
    const last = String(data[i][4] || '').trim();
    // capitalize names properly
    const formatName = (str) => str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
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