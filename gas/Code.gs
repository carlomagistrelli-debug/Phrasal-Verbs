// Deploy: Extensions → Apps Script → Deploy → New deployment
//   Type:       Web app
//   Execute as: User accessing the web app   ← required for per-user progress
//   Access:     Anyone with a Google account

const SPREADSHEET_ID = '1v8h1SJ-n5OW0ntrwnmcLsx_7_OhGmsZDl19g0hSM1Q0';
const CARDS_SHEET    = '150 Most Common Phrasal Verbs';
const PROGRESS_SHEET = 'Progress';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Phrasal Verbs – Flashcards')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Called once on page load – returns everything in a single round-trip
function getInitialData() {
  const email = Session.getActiveUser().getEmail() || 'anonymous';
  const row   = findUserRow_(email);
  return {
    cards:    getCards_(),
    srs:      parseCell_(row, 1, {}),
    newToday: parseCell_(row, 2, {}),
    settings: parseCell_(row, 3, {}),
    user:     email,
  };
}

function saveSRS(json) {
  const sheet = getProgressSheet_();
  const row   = findOrCreateUserRow_(getEmail_(), sheet);
  sheet.getRange(row, 2).setValue(JSON.stringify(json));
}

function saveNewToday(json) {
  const sheet = getProgressSheet_();
  const row   = findOrCreateUserRow_(getEmail_(), sheet);
  sheet.getRange(row, 3).setValue(JSON.stringify(json));
}

function saveSettings(json) {
  const sheet = getProgressSheet_();
  const row   = findOrCreateUserRow_(getEmail_(), sheet);
  sheet.getRange(row, 4).setValue(JSON.stringify(json));
}

function resetProgress() {
  const sheet = getProgressSheet_();
  const row   = findOrCreateUserRow_(getEmail_(), sheet);
  sheet.getRange(row, 2, 1, 2).setValues([['', '']]);
}

// ── Private helpers ────────────────────────────────────────────────────────────

function getEmail_() {
  return Session.getActiveUser().getEmail() || 'anonymous';
}

function getCards_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  // The cards tab may be named anything (Google often leaves it "Untitled").
  // Prefer the configured name, else fall back to the first non-Progress tab.
  const sheet = ss.getSheetByName(CARDS_SHEET)
             || ss.getSheets().filter(s => s.getName() !== PROGRESS_SHEET)[0];
  if (!sheet) return [];
  const data   = sheet.getDataRange().getValues();
  const header = data[0].map(h => String(h).toLowerCase().trim());
  const iV = header.findIndex(h => h.includes('verb'));
  const iM = header.findIndex(h => h.includes('mean'));
  const iE = header.findIndex(h => h.includes('ex'));
  const iC = header.findIndex(h => h.includes('cat'));
  return data.slice(1)
    .map(row => ({
      verb:     String(row[iV] || '').trim(),
      meaning:  iM >= 0 ? String(row[iM] || '').trim() : '',
      example:  iE >= 0 ? String(row[iE] || '').trim() : '',
      category: iC >= 0 ? String(row[iC] || '').trim() : '',
    }))
    .filter(c => c.verb);
}

function getProgressSheet_() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let   sheet = ss.getSheetByName(PROGRESS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PROGRESS_SHEET);
    sheet.appendRow(['Email', 'SRS Data', 'New Today', 'Settings']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 400);
    sheet.setColumnWidth(3, 200);
    sheet.setColumnWidth(4, 200);
  }
  return sheet;
}

// Returns the row data array, or null if not found
function findUserRow_(email) {
  const sheet = getProgressSheet_();
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === email.toLowerCase()) return data[i];
  }
  return null;
}

// Returns the 1-indexed sheet row number, creating the row if needed
function findOrCreateUserRow_(email, sheet) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === email.toLowerCase()) return i + 1;
  }
  const newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1).setValue(email);
  return newRow;
}

// row is the raw values array: [email, srsJson, newTodayJson, settingsJson]
// colIdx: 1 = SRS, 2 = newToday, 3 = settings
function parseCell_(row, colIdx, fallback) {
  if (!row) return fallback;
  const val = row[colIdx];
  try { return val ? JSON.parse(String(val)) : fallback; } catch (e) { return fallback; }
}
