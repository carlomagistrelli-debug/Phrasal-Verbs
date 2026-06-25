// Deploy: Extensions → Apps Script → Deploy → New deployment
//   Type:       Web app
//   Execute as: Me (the deploying user)      ← so external users never hit the
//                                              authorization / unverified-app wall
//   Access:     Anyone, even anonymous       ← no Google sign-in required
//
// Identity is name-based: the client sends a userKey (a name the visitor picks,
// stored in their browser) with every call. The script runs as the owner, so it
// can always reach the spreadsheet regardless of who is visiting.

const SPREADSHEET_ID = '1v8h1SJ-n5OW0ntrwnmcLsx_7_OhGmsZDl19g0hSM1Q0';
const CARDS_SHEET      = 'Common Phrasal Verbs';
const PROGRESS_SHEET   = 'Progress';
const CATEGORIES_SHEET = 'Categories';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Phrasal Verbs – Flashcards')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Called once on page load – returns everything in a single round-trip
function getInitialData(userKey) {
  const key = normalizeKey_(userKey);
  const row = findUserRow_(key);
  return {
    cards:      getCards_(),
    categories: getCategories_(),
    srs:        parseCell_(row, 1, {}),
    newToday:   parseCell_(row, 2, {}),
    settings:   parseCell_(row, 3, {}),
    user:       key,
  };
}

function saveSRS(json, userKey) {
  const sheet = getProgressSheet_();
  const row   = findOrCreateUserRow_(normalizeKey_(userKey), sheet);
  sheet.getRange(row, 2).setValue(JSON.stringify(json));
}

function saveNewToday(json, userKey) {
  const sheet = getProgressSheet_();
  const row   = findOrCreateUserRow_(normalizeKey_(userKey), sheet);
  sheet.getRange(row, 3).setValue(JSON.stringify(json));
}

function saveSettings(json, userKey) {
  const sheet = getProgressSheet_();
  const row   = findOrCreateUserRow_(normalizeKey_(userKey), sheet);
  sheet.getRange(row, 4).setValue(JSON.stringify(json));
}

function resetProgress(userKey) {
  const sheet = getProgressSheet_();
  const row   = findOrCreateUserRow_(normalizeKey_(userKey), sheet);
  sheet.getRange(row, 2, 1, 2).setValues([['', '']]);
}

// Append a new phrasal verb to the shared cards sheet.
// Rejects duplicates (case-insensitive on the phrasal verb itself).
// Returns { ok: true, card } on success, or { ok: false, reason } otherwise.
function addCard(card, userKey) {
  const verb     = String(card && card.verb     || '').trim();
  const meaning  = String(card && card.meaning  || '').trim();
  const example  = String(card && card.example  || '').trim();
  const category = String(card && card.category || '').trim();
  if (!verb)    return { ok: false, reason: 'The phrasal verb is required.' };
  if (!meaning) return { ok: false, reason: 'The meaning is required.' };

  const sheet = getCardsSheet_();
  if (!sheet) return { ok: false, reason: 'Cards sheet not found.' };

  const data   = sheet.getDataRange().getValues();
  const header = data[0].map(h => String(h).toLowerCase().trim());
  const iV = header.findIndex(h => h.includes('verb'));
  const iM = header.findIndex(h => h.includes('mean'));
  const iE = header.findIndex(h => h.includes('ex'));
  const iC = header.findIndex(h => h.includes('cat'));

  // Duplicate check — authoritative, server-side, case-insensitive on the verb.
  const dup = data.slice(1).some(r =>
    String(r[iV] || '').trim().toLowerCase() === verb.toLowerCase());
  if (dup) return { ok: false, reason: 'duplicate' };

  const rowArr = new Array(header.length).fill('');
  if (iV >= 0) rowArr[iV] = verb;
  if (iM >= 0) rowArr[iM] = meaning;
  if (iE >= 0) rowArr[iE] = example;
  if (iC >= 0) rowArr[iC] = category;
  sheet.appendRow(rowArr);
  if (category) ensureCategory_(category);

  return {
    ok: true,
    card: { verb: verb, meaning: meaning, example: example, category: category },
    categories: getCategories_(),
  };
}

// Canonical category list — column A of the "Categories" tab.
// Drops a header cell if present; trims, dedupes case-insensitively, sorts.
function getCategories_() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CATEGORIES_SHEET);
  if (!sheet || sheet.getLastRow() === 0) return [];
  const col  = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  const out  = [];
  const seen = {};
  col.forEach((r, i) => {
    const v = String(r[0] || '').trim();
    if (!v) return;
    if (i === 0 && (v.toLowerCase() === 'category' || v.toLowerCase() === 'categories')) return; // header
    const k = v.toLowerCase();
    if (seen[k]) return;
    seen[k] = true;
    out.push(v);
  });
  return out.sort((a, b) => a.localeCompare(b));
}

// Append a category to the "Categories" tab unless it's already there
// (case-insensitive). Creates the tab with a header if it doesn't exist.
function ensureCategory_(category) {
  const cat = String(category || '').trim();
  if (!cat) return;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CATEGORIES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CATEGORIES_SHEET);
    sheet.appendRow(['Category']);
    sheet.getRange(1, 1).setFontWeight('bold');
  }
  const last = sheet.getLastRow();
  if (last > 0) {
    const col = sheet.getRange(1, 1, last, 1).getValues();
    if (col.some(r => String(r[0] || '').trim().toLowerCase() === cat.toLowerCase())) return;
  }
  sheet.appendRow([cat]);
}

// ── Private helpers ────────────────────────────────────────────────────────────

// The client-supplied name is the identity key. Trim it, cap its length, and
// fall back to 'anonymous' so a blank key can never address an empty-named row.
function normalizeKey_(userKey) {
  const k = String(userKey == null ? '' : userKey).trim().slice(0, 60);
  return k || 'anonymous';
}

// The cards tab may be named anything (Google often leaves it "Untitled").
// Prefer the configured name, else fall back to the first non-Progress tab.
function getCardsSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(CARDS_SHEET)
      || ss.getSheets().filter(s => s.getName() !== PROGRESS_SHEET)[0]
      || null;
}

function getCards_() {
  const sheet = getCardsSheet_();
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
    sheet.appendRow(['User', 'SRS Data', 'New Today', 'Settings']);
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
