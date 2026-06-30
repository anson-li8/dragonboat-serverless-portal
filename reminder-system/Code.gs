const CONFIG = {
  REGISTRATION_SHEET_ID : 'YOUR_SHEET_ID_HERE',
  WAIVER_SHEET_ID       : 'YOUR_SHEET_ID_HERE',
  REGISTRATION_TAB      : 'Form Responses 2',
  WAIVER_TAB            : 'Sheet1',
  STATUS_TAB            : 'Status Dashboard',   
  FESTIVAL_DATE         : 'July 11th, 2026',
  WAIVER_FORM_LINK      : 'YOUR_FORM_LINK_HERE',
  PRACTICE_SIGNUP_LINK  : 'YOUR_FORM_LINK_HERE',
  REPLY_TO_EMAIL        : 'info@milwaukeedragonboatfest.org',
  SENDER_NAME           : 'Milwaukee Dragon Boat Festival',
  SEND_HOUR_CENTRAL     : 8,     // Hour to trigger daily automation (24-hr format)
  WAIVER_THRESHOLD      : 21,    // Minimum waivers needed per team
  VALID_YEAR            : 2026,  // Only process data from this calendar year
  CREATE_DRAFTS_ONLY    : true,  // Keep as true to test in your Gmail Drafts
  WAIVER_SUBJECT        : '[Milwaukee Dragon Boat Festival] Waiver Reminder – {TEAM_NAME} ({COUNT}/{THRESHOLD} done)',
  PRACTICE_SUBJECT      : '[Milwaukee Dragon Boat Festival] Practice Registration Reminder – {TEAM_NAME}',
  // NEW CAMPSITE CONSTANTS
  CAMPSITE_SIGNUP_LINK  : 'WAITING_FOR_LINK', 
  CAMPSITE_SUBJECT      : '[Milwaukee Dragon Boat Festival] Team Campsite Signup Reminder – {TEAM_NAME}'
};

// depends on the structure of your registration and waiver sheets
const REG_COLS = {
  DATE: 1,         
  SPONSOR: 2,      
  TEAM: 3,         
  MANAGER: 4,      
  EMAIL: 5,        
  PRACTICE: 11,    
  WAIVERS_DONE: 12,
  CAMPSITE: 13       // NEW: Column N (0-indexed = 13)
};
const WAV_COLS = {
  TIMESTAMP: 0,    
  TEAM_FIELD: 1,   
  FIRST: 3,        
  LAST: 4          
};

// google sheets dropdown menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('DO NOT CLICK IF NOT ON INFO ACCOUNT')
    .addItem('1. Sync Dashboard & Preview Status', 'previewStatus')
    .addItem('2. Run Automation (Send or Draft Emails)', 'runDailyCheckFromMenu')
    .addSeparator()
    .addItem('Enable Daily Automated Trigger', 'setupDailyTrigger')
    .addItem('Disable Daily Automated Trigger', 'removeAllTriggers')
    .addToUi();
}

// safe to show UI pop-ups
function previewStatus() { runEngine(true, false); }
function runDailyCheckFromMenu() { 
  const actions = getDailyEmailActions();
  runEngine(false, false, actions);
}

// time-driven trigger execution
function runDailyCheckTrigger() { 
  const actions = getDailyEmailActions();
  // we run the engine every day to update dashboard, 
  // but pass the strict rules on what emails are allowed out today.
  runEngine(false, true, actions); 
}

function getDailyEmailActions() {
  const today = new Date();
  const dateString = Utilities.formatDate(today, 'America/Chicago', 'yyyy-MM-dd');
  const [year, month, day] = dateString.split('-').map(Number);
  const currentDate = new Date(year, month - 1, day);
  const startSending = new Date(CONFIG.VALID_YEAR, 5, 21);  // June 20 (Month 5 is June)
  const practiceStart = new Date(CONFIG.VALID_YEAR, 5, 30); // June 30 (Forcing practice start)
  const dailyStart = new Date(CONFIG.VALID_YEAR, 6, 5);     // July 5 (Day after July 4)
  let actions = { sendWaivers: false, sendPractice: false, sendCampsite: true };
  if (currentDate < startSending) {
    actions = { sendWaivers: false, sendPractice: false, sendCampsite: true };
  } else if (currentDate >= dailyStart) {
    actions = { sendWaivers: true, sendPractice: true, sendCampsite: true };
  } else {
    // btw June 20 and July 4: Staggered sending
    const diffTime = Math.abs(currentDate - startSending);
    const diffDays = Math.round(diffTime / (1000 * 60 * 60 * 24));
    if (diffDays % 7 === 0) {
      actions = { sendWaivers: true, sendPractice: false, sendCampsite: true };
    } else if (diffDays % 7 === 1) {
      actions = { sendWaivers: false, sendPractice: true, sendCampsite: true };
    }
  }
  // strict rule override: Practice reminders CANNOT go out before June 30th
  if (currentDate < practiceStart) {
    actions.sendPractice = false;
  }
  return actions;
}

function runEngine(isDryRun, isTrigger, actions = { sendWaivers: true, sendPractice: true, sendCampsite: true }) {  try {
    const ss = SpreadsheetApp.openById(CONFIG.REGISTRATION_SHEET_ID);
    const dashSheet = ss.getSheetByName(CONFIG.STATUS_TAB) || ss.insertSheet(CONFIG.STATUS_TAB);
    const overrides = loadDashboardOverrides(dashSheet);
    const teams = loadRegistrationTeams(ss);
    const waiverMap = loadWaiverData();
    matchTeams(teams, waiverMap, overrides);
    updateStatusDashboard(dashSheet, teams, overrides);
    updateRegistrationSheet(ss, teams);
    if (!isDryRun) {
      // pass actions into processEmails
      processEmails(teams, overrides, actions);
      if (!isTrigger) {
        const modeText = CONFIG.CREATE_DRAFTS_ONLY ? "Drafts created! Open your Gmail 'Drafts' folder to review them." : "Live emails successfully sent.";
        SpreadsheetApp.getUi().alert(`Execution Complete.\n\n${modeText}`);
      }
    } else {
      if (!isTrigger) {
        SpreadsheetApp.getUi().alert(`Preview complete!\n\nCheck the "${CONFIG.STATUS_TAB}" tab. Fill out any "Manual Match Link" fields if auto-matching failed.`);
      }
    }
  } catch (err) {
    if (!isTrigger) {
      SpreadsheetApp.getUi().alert("Execution Error: " + err.message);
    } else {
      console.error("Execution Error in background trigger: " + err.message + "\n" + err.stack);
    }
  }
}

function loadDashboardOverrides(dashSheet) {
  const data = dashSheet.getDataRange().getValues();
  const overrides = {};
  if (data.length <= 1) return overrides;
  for (let i = 1; i < data.length; i++) {
    const sponsor = String(data[i][0] || '').trim();
    const teamName = String(data[i][1] || '').trim();
    const key = norm(sponsor + '|' + teamName);
    overrides[key] = {
      manualMatch: String(data[i][9] || '').trim(),
      skipEmail: data[i][10] === true
    };
  }
  return overrides;
}

function loadRegistrationTeams(ss) {
  const sheet = ss.getSheetByName(CONFIG.REGISTRATION_TAB);
  if (!sheet) throw new Error(`Could not find registration tab: "${CONFIG.REGISTRATION_TAB}"`);
  const data = sheet.getDataRange().getValues();
  const teams = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!isValidYear(row[REG_COLS.DATE])) continue;
    const sponsor = String(row[REG_COLS.SPONSOR] || '').trim();
    const teamName = String(row[REG_COLS.TEAM] || '').trim();
    if (!sponsor && !teamName) continue; 
    teams.push({
      rowIndex: i + 1,
      sponsor: sponsor,
      teamName: teamName,
      key: norm(sponsor + '|' + teamName),
      manager: String(row[REG_COLS.MANAGER] || '').trim(),
      emails: parseEmails(row[REG_COLS.EMAIL]),
      practiceReg: (row[REG_COLS.PRACTICE] == 1 || row[REG_COLS.PRACTICE] === true),
      campsiteReg: (row[REG_COLS.CAMPSITE] == 1 || row[REG_COLS.CAMPSITE] === true), // NEW
      waiverCount: 0,
      matchConfidence: 'UNMATCHED',
      matchedWaiverKey: '',
      notes: []
    });
  }
  return teams;
}

function loadWaiverData() {
  const ss = SpreadsheetApp.openById(CONFIG.WAIVER_SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.WAIVER_TAB);
  if (!sheet) throw new Error("Waiver tab not found.");
  const data = sheet.getDataRange().getValues();
  const waiverMap = {}; 
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!isValidYear(row[WAV_COLS.TIMESTAMP])) continue;
    const teamField = String(row[WAV_COLS.TEAM_FIELD] || '').trim();
    const first = String(row[WAV_COLS.FIRST] || '').trim().toLowerCase();
    const last = String(row[WAV_COLS.LAST] || '').trim().toLowerCase();
    if (!teamField || (!first && !last)) continue;
    const normKey = norm(teamField);
    if (!waiverMap[normKey]) {
      waiverMap[normKey] = { rawName: teamField, people: new Set() };
    }
    waiverMap[normKey].people.add(first + '|' + last);
  }
  return waiverMap;
}

function matchTeams(teams, waiverMap, overrides) {
  const waiverKeys = Object.keys(waiverMap);
  for (const team of teams) {
    const teamOverride = overrides[team.key];
    if (teamOverride && teamOverride.manualMatch) {
      const normManual = norm(teamOverride.manualMatch);
      if (waiverMap[normManual]) {
        team.waiverCount = waiverMap[normManual].people.size;
        team.matchConfidence = 'MANUAL';
        team.matchedWaiverKey = waiverMap[normManual].rawName;
        continue;
      } else {
        team.notes.push(`Manual match name "${teamOverride.manualMatch}" wasn't found on Waiver Sheet.`);
      }
    }
    const s = norm(team.sponsor);
    const t = norm(team.teamName);
    const exactCombined = norm(s + ' - ' + t);
    if (waiverMap[exactCombined]) {
      team.waiverCount = waiverMap[exactCombined].people.size;
      team.matchConfidence = 'EXACT';
      team.matchedWaiverKey = waiverMap[exactCombined].rawName;
      continue;
    }
    if (t && waiverMap[t]) {
      team.waiverCount = waiverMap[t].people.size;
      team.matchConfidence = 'EXACT';
      team.matchedWaiverKey = waiverMap[t].rawName;
      continue;
    }
    if (t && t.length > 3) {
      const hits = waiverKeys.filter(wk => wk.includes(t) || t.includes(wk));
      if (hits.length === 1) {
        const hitKey = hits[0];
        team.waiverCount = waiverMap[hitKey].people.size;
        team.matchConfidence = 'FUZZY';
        team.matchedWaiverKey = waiverMap[hitKey].rawName;
        team.notes.push(`Fuzzy match used. Paste exact name to override.`);
        continue;
      }
    }
    team.matchConfidence = 'UNMATCHED';
    team.notes.push('Could not auto-match. Paste exact name from Waiver Sheet into Manual Match Link.');
  }
}

function updateStatusDashboard(dashSheet, teams, overrides) {
  dashSheet.clear();
  const DASH_COLS = [
    'Sponsor', 'Team Name', 'Manager', 'Email(s)',
    'Practice', 'Campsite', 'Waivers Signed', 'Complete', 'Status',
    'Match Confidence', 'Manual Match Link', 'Skip Email?', 'Notes', 'Last Updated'
  ];
  const headerRange = dashSheet.getRange(1, 1, 1, DASH_COLS.length);
  headerRange.setValues([DASH_COLS]).setBackground('#1C3557').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10).setFontFamily('Arial');
  dashSheet.setFrozenRows(1);
    const widths = [150, 180, 140, 220, 70, 70, 95, 70, 150, 120, 180, 85, 250, 130];
  widths.forEach((w, i) => dashSheet.setColumnWidth(i + 1, w));
  if (teams.length === 0) return;
  const now = new Date().toLocaleString('en-US', { timeZone: 'America/Chicago' });
  const rows = [];
  for (const t of teams) {
    const isComplete = t.waiverCount >= CONFIG.WAIVER_THRESHOLD;
    let statusText = 'NOT STARTED';
    if (t.waiverCount >= CONFIG.WAIVER_THRESHOLD) statusText = 'COMPLETE';
    else if (t.waiverCount > 0) statusText = `IN PROGRESS (${t.waiverCount}/${CONFIG.WAIVER_THRESHOLD})`;
    if (t.matchConfidence === 'UNMATCHED') statusText = '⚠️ UNMATCHED';
    const teamOverride = overrides[t.key] || { manualMatch: '', skipEmail: false };
    rows.push([
      t.sponsor, t.teamName, t.manager, t.emails.join(', '),
      t.practiceReg ? 'Yes' : 'No', 
      t.campsiteReg ? 'Yes' : 'No', // NEW CAMPSITE FIELD
      t.waiverCount, isComplete ? 'Yes' : 'No',
      statusText, t.matchConfidence, teamOverride.manualMatch, teamOverride.skipEmail,
      t.notes.join(' | '), now
    ]);
  }
  
  const dataRange = dashSheet.getRange(2, 1, rows.length, DASH_COLS.length);
  dataRange.setValues(rows).setFontFamily('Arial').setFontSize(10);
  
  // 4. Shift the Checkbox column from 11 to 12 because we added a column
  dashSheet.getRange(2, 12, rows.length, 1).insertCheckboxes();
  
  // 5. Shift the Status color-checker from index 7 to 8
  for (let i = 0; i < rows.length; i++) {
    const status = rows[i][8]; 
    let color = '#FFFFFF';
    if (status === 'COMPLETE') color = '#C6EFCE'; 
    else if (status.includes('IN PROGRESS')) color = '#BDD7EE'; 
    else if (status.includes('UNMATCHED')) color = '#FCE4D6'; 
    else if (status === 'NOT STARTED') color = '#FFC7CE'; 
    dashSheet.getRange(i + 2, 1, 1, DASH_COLS.length).setBackground(color);
  }
}

function updateRegistrationSheet(ss, teams) {
  const sheet = ss.getSheetByName(CONFIG.REGISTRATION_TAB);
  const colM = REG_COLS.WAIVERS_DONE + 1;
  for (const team of teams) {
    if (!team.rowIndex) continue;
    const valueToWrite = (team.waiverCount >= CONFIG.WAIVER_THRESHOLD) ? 1 : 0;
    sheet.getRange(team.rowIndex, colM).setValue(valueToWrite);
  }
}

function processEmails(teams, overrides, actions) {
  for (const team of teams) {
    const teamOverride = overrides[team.key] || { manualMatch: '', skipEmail: false };
    if (teamOverride.skipEmail || team.emails.length === 0) continue;
    // REMOVED the "team.matchConfidence !== 'UNMATCHED'" restriction
    // Now, if they are unmatched, their count is naturally 0, and they get the 0/21 email!
    if (actions.sendWaivers && team.waiverCount < CONFIG.WAIVER_THRESHOLD) {
      const subject = CONFIG.WAIVER_SUBJECT
        .replace('{TEAM_NAME}', team.teamName || team.sponsor)
        .replace('{COUNT}', team.waiverCount)
        .replace('{THRESHOLD}', CONFIG.WAIVER_THRESHOLD);
      sendOrDraft(team.emails, subject, buildWaiverEmailBody(team));
    }
    // Only send if schedule allows Practice
    if (actions.sendPractice && !team.practiceReg) {
      const subject = CONFIG.PRACTICE_SUBJECT.replace('{TEAM_NAME}', team.teamName || team.sponsor);
      sendOrDraft(team.emails, subject, buildPracticeEmailBody(team));
    }
    // Only send if schedule allows Campsite
    if (actions.sendCampsite && !team.campsiteReg) {
      const subject = CONFIG.CAMPSITE_SUBJECT.replace('{TEAM_NAME}', team.teamName || team.sponsor);
      sendOrDraft(team.emails, subject, buildCampsiteEmailBody(team));
    }
  }
}

function sendOrDraft(emails, subject, body) {
  const mailOptions = {
    replyTo: CONFIG.REPLY_TO_EMAIL,
    name: CONFIG.SENDER_NAME
  };
  if (emails.length > 1) {
    mailOptions.cc = emails.slice(1).join(',');
  }
  if (CONFIG.CREATE_DRAFTS_ONLY) {
    GmailApp.createDraft(emails[0], subject, body, mailOptions);
  } else {
    GmailApp.sendEmail(emails[0], subject, body, mailOptions);
  }
}

function buildWaiverEmailBody(team) {
  const firstName = team.manager.split(/[\s(]/)[0] || 'Team Captain';
  const teamLabel = team.teamName || team.sponsor;
  const needed = CONFIG.WAIVER_THRESHOLD - team.waiverCount;
  return `Hi ${firstName},\n\n` +
         `This is a reminder from the Milwaukee Dragon Boat Festival organizers.\n\n` +
         `Your team, ${teamLabel}, currently has ${team.waiverCount} of ${CONFIG.WAIVER_THRESHOLD} required waivers on file. You still need ${needed} more waiver${needed !== 1 ? 's' : ''} before your team can board. Waivers are to be completed prior to July 2nd, 2026.\n\n` +
         `Please share the waiver form link with any remaining team members:\n${CONFIG.WAIVER_FORM_LINK}\n\n` +
         `Festival Date: ${CONFIG.FESTIVAL_DATE}\n\n` +
         `Please reply to this email if you have any questions.\n\n` +
         `Thank you,\n${CONFIG.SENDER_NAME}\n${CONFIG.REPLY_TO_EMAIL}`;
}

function buildPracticeEmailBody(team) {
  const firstName = team.manager.split(/[\s(]/)[0] || 'Team Captain';
  const teamLabel = team.teamName || team.sponsor;
  return `Hi ${firstName},\n\n` +
         `This is a reminder from the Milwaukee Dragon Boat Festival organizers.\n\n` +
         `Our records show that ${teamLabel} has not yet registered for a practice session.\n\n` +
         `Please register for your practice track here:\n${CONFIG.PRACTICE_SIGNUP_LINK}\n\n` +
         `Festival Date: ${CONFIG.FESTIVAL_DATE}\n\n` +
         `Thank you,\n${CONFIG.SENDER_NAME}\n${CONFIG.REPLY_TO_EMAIL}`;
}

function buildCampsiteEmailBody(team) {
  const firstName = team.manager.split(/[\s(]/)[0] || 'Team Captain';
  const teamLabel = team.teamName || team.sponsor;
  return `Hi ${firstName},\n\n` +
         `This is a reminder from the Milwaukee Dragon Boat Festival organizers.\n\n` +
         `Our records show that ${teamLabel} has not yet signed up for a team campsite.\n\n` +
         `Please complete your campsite signup by the deadline on Tuesday, July 2nd, 2026, using the link below:\n${CONFIG.CAMPSITE_SIGNUP_LINK}\n\n` +
         `Festival Date: ${CONFIG.FESTIVAL_DATE}\n\n` +
         `Thank you,\n${CONFIG.SENDER_NAME}\n${CONFIG.REPLY_TO_EMAIL}`;
}

function isValidYear(val) {
  if (!val) return false;
  if (val instanceof Date) return val.getFullYear() === CONFIG.VALID_YEAR;
  const parsedDate = new Date(val);
  if (!isNaN(parsedDate.getTime())) return parsedDate.getFullYear() === CONFIG.VALID_YEAR;
  return String(val).includes(String(CONFIG.VALID_YEAR));
}

function norm(str) {
  return String(str || '').toLowerCase().replace(/\s+/g, ' ').trim();
}

function parseEmails(raw) {
  return String(raw || '').split(/[;,\s]+/).map(e => e.trim()).filter(e => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e));
}

function setupDailyTrigger() {
  removeAllTriggers();
  ScriptApp.newTrigger('runDailyCheckTrigger')
    .timeBased().everyDays(1).atHour(CONFIG.SEND_HOUR_CENTRAL).inTimezone('America/Chicago').create();
  SpreadsheetApp.getUi().alert(`Automation Trigger Enabled! The system will run every day around ${CONFIG.SEND_HOUR_CENTRAL}:00 AM Central Time.`);
}

function removeAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
}