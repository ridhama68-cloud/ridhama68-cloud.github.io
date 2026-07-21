/*************************************************************************
 * Reactivify CRM — Code_v2.gs
 * Backend: Web App API + Reminder/Automation engine
 * Runtime: Google Apps Script (V8). Bound to the "Reactivify CRM" Sheet.
 *
 * Deploy: Deploy > Manage deployments > pencil > Version: "New version"
 *         (keeps the same /exec URL). First deploy asks for consent —
 *         that click is human-only (Google security).
 *
 * One-time setup after first deploy: run setupTriggers() once from the
 * editor (Run menu) to install the time-based triggers.
 *************************************************************************/

/* ============================ CONFIG ============================ */
const CONFIG = {
  OWNER_EMAIL: 'ridhama68@gmail.com',           // you (the founder) — reminders/digests go here
  FROM_EMAIL: 'growth@reactivify.com',          // doctor/reminder emails are SENT FROM this (must be a verified Gmail "Send mail as" alias in the owner account)
  COMPANY_NAME: 'Reactivify',
  SHEET_LEADS: 'Leads',
  SHEET_INTERNS: 'Interns',
  SHEET_SESSIONS: 'Sessions',                   // intern login / idle / logout time tracking (auto-created)

  // Which logged call outcomes count as an "answered call" that MUST have a
  // recording — this is what the logout reconciliation check enforces.
  // No Show / Dumped / Not Answering are excluded (nobody was actually spoken to).
  // Want No Show or Dumped to require a recording too? Just add them to this list.
  RECORDING_REQUIRED_OUTCOMES: ['Interested', 'Not Interested', 'Meeting Booked', 'Schedule Follow-up', 'Landed'],
  REC_MISSING_COL: 'Rec Missing Reason',        // stores the intern's reason when a recording genuinely doesn't exist

  // ---- Reminder timing model --------------------------------------
  // System emails the DOCTOR directly at each offset below.
  // Intern gets a ready-to-send wa.me/mailto "message the doctor" link.
  // INTERN_NUDGE_MODE controls WHEN the intern link fires relative to
  // each client reminder:  'PRIOR_15' | 'SAME_TIME' | 'BOTH'
  INTERN_NUDGE_MODE: 'PRIOR_15',               // <-- your chosen default
  INTERN_NUDGE_PRIOR_MIN: 15,                  // minutes before the tick

  EMAIL_DOCTOR: true,                          // system emails doctor directly
  EMAIL_INTERN_LINK: true,                     // intern gets send-the-message link

  // ---- Google Calendar auto-booking ------------------------------
  // When a meeting is booked, the backend (running as YOU) creates the
  // event on your Google Calendar with the intern + doctor as guests.
  // Needs the Calendar permission — the next "New version" deploy will
  // ask you to re-approve consent (one click).
  CREATE_CALENDAR_EVENTS: true,
  CALENDAR_ID: 'primary',                      // 'primary' = your team@reactivify.com calendar
  MEETING_DURATION_MIN: 30,
  SEND_CALENDAR_INVITES: false,                // true = email invites to guests too

  // ---- Call recordings + AI coaching (Groq) ----------------------
  // PASTE YOUR GROQ KEY BELOW (console.groq.com -> API Keys). Backend only — never goes to the public site.
  GROQ_API_KEY: 'PASTE_YOUR_GROQ_KEY_HERE',
  RECORDINGS_FOLDER_ID: '12iPlYoZv5w2CxqMzsWoNF_DpG-q2KA4D',   // Google Drive folder with the recordings (interns upload into their own name/date subfolders inside it)
  GROQ_TRANSCRIBE_MODEL: 'whisper-large-v3-turbo',
  GROQ_COACH_MODEL: 'llama-3.3-70b-versatile',
  RECORD_COL: 'Recording', TRANSCRIPT_COL: 'Transcript', COACH_COL: 'Coaching',
  COACH_BATCH: 6,                              // max calls coached per button press (Apps Script 6-min limit)

  // Meeting reminder offsets (minutes before meeting start)
  // Doctor + team reminder offsets (minutes before the meeting).
  MEETING_OFFSETS: [
    { key: 'DayBefore Sent', minutes: 24 * 60, label: '24 hours'   },
    { key: '1Hr Sent',       minutes: 240,     label: '4 hours'    },
    { key: '30Min Sent',     minutes: 45,      label: '45 minutes' }
  ],

  // Follow-up (cols V/W) reminder: fire once when due, notify intern+owner
  FOLLOWUP_SENT_FLAG: 'Followup Sent',         // NEW column X (add to sheet)

  // How close (minutes) a trigger tick must be to an offset to count as "due"
  TICK_TOLERANCE_MIN: 10,                       // triggers run every 10 min

  // WhatsApp/message template sent to the intern's link
  TEMPLATE: function (doctorName, clinic, dateStr, timeStr) {
    const dr = doctorName ? ('Dr. ' + doctorName) : 'Doctor';
    return 'Hi ' + dr + ', this is a friendly reminder about our meeting for '
      + (clinic || 'your clinic') + ' on ' + dateStr + ' at ' + timeStr
      + '. Looking forward to speaking with you! — ' + CONFIG.COMPANY_NAME;
  }
};

/* ===================== SHEET / COLUMN HELPERS ==================== */
// Canonical Leads header order (row 1 of the Leads tab).
// Cols A..X. Followup Date=V, Followup Time=W, Followup Sent=X (new).
const LEAD_HEADERS = [
  'ID','Timestamp','Name','Clinic','Phone','Email','City','Rating','Reviews',
  'Internal Rating','Status','Assigned To','Notes','Meeting Date','Meeting Time',
  'Last Follow-up Date','Follow-up Count','Conflict Flag','DayBefore Sent',
  '1Hr Sent','30Min Sent','Followup Date','Followup Time','Followup Sent',
  'Recording','Coaching','Rec Missing Reason'
];

// Session log (intern time tracking). Auto-created — no manual sheet work.
const SESSION_HEADERS = [
  'Session ID','Date','Intern','Login','Last Active','Status',
  'Work Min','Idle Min','Logout','Convos','Recordings','Reconciled','Notes'
];

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function leadsSheet_() { return ss_().getSheetByName(CONFIG.SHEET_LEADS); }
function internsSheet_() { return ss_().getSheetByName(CONFIG.SHEET_INTERNS); }

function headerIndex_(sheet) {
  const hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  hdr.forEach(function (h, i) { map[String(h).trim()] = i; });
  return map;
}

// Make sure the given header names exist on a sheet; append any that are
// missing (expanding the sheet if it has run out of columns). This is what
// lets the app add "Recording" / "Rec Missing Reason" etc. WITHOUT you having
// to insert columns by hand. Returns a fresh header index.
function ensureColumns_(sheet, names) {
  let idx = headerIndex_(sheet);
  const missing = names.filter(function (n) { return idx[n] === undefined; });
  if (!missing.length) return idx;
  const lastCol = sheet.getLastColumn();
  const need = lastCol + missing.length;
  if (sheet.getMaxColumns() < need) sheet.insertColumnsAfter(sheet.getMaxColumns(), need - sheet.getMaxColumns());
  sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  return headerIndex_(sheet);
}

// Read all leads as array of objects keyed by header name.
function readLeads_() {
  const sh = leadsSheet_();
  const last = sh.getLastRow();
  if (last < 2) return { rows: [], idx: headerIndex_(sh), values: [] };
  const idx = headerIndex_(sh);
  const values = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  const rows = values.map(function (r, i) {
    const o = { _row: i + 2 };
    Object.keys(idx).forEach(function (h) { o[h] = r[idx[h]]; });
    return o;
  });
  return { rows: rows, idx: idx, values: values };
}

function readInterns_() {
  const sh = internsSheet_();
  if (!sh || sh.getLastRow() < 2) return [];
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  return vals.filter(function (r) { return r[0]; })
             .map(function (r) { return { name: String(r[0]).trim(), email: String(r[1]).trim() }; });
}

function internEmail_(name) {
  const m = readInterns_().filter(function (i) { return i.name.toLowerCase() === String(name).toLowerCase(); });
  return m.length ? m[0].email : '';
}

// ---- Team (intern) management — editable from the app, no manual sheet work
function internsSheetEnsure_() {
  let sh = ss_().getSheetByName(CONFIG.SHEET_INTERNS);
  if (!sh) { sh = ss_().insertSheet(CONFIG.SHEET_INTERNS); sh.getRange(1, 1, 1, 2).setValues([['Name', 'Email']]); }
  else if (sh.getLastRow() < 1) { sh.getRange(1, 1, 1, 2).setValues([['Name', 'Email']]); }
  return sh;
}
function addIntern_(name, email) {
  name = String(name || '').trim();
  if (!name) return readInterns_();
  const sh = internsSheetEnsure_();
  const last = sh.getLastRow();
  const names = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
  let found = false;
  for (let i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim().toLowerCase() === name.toLowerCase()) { sh.getRange(i + 2, 2).setValue(email || ''); found = true; }
  }
  if (!found) sh.appendRow([name, email || '']);
  return readInterns_();
}
function removeIntern_(name) {
  name = String(name || '').trim();
  const sh = internsSheetEnsure_();
  const last = sh.getLastRow();
  if (last < 2) return readInterns_();
  const names = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = names.length - 1; i >= 0; i--) {
    if (String(names[i][0]).trim().toLowerCase() === name.toLowerCase()) sh.deleteRow(i + 2);
  }
  return readInterns_();
}

// ---- One-time self-heal: force date/time columns to plain-text so Google
// Sheets stops mangling times into 1899 serial values.
function ensureTextFormats_() {
  const sh = leadsSheet_();
  const idx = headerIndex_(sh);
  // Phone/Email/Clinic/Name forced to text so a leading "+" or "=" is never
  // interpreted as a spreadsheet formula (the cause of #ERROR! phone cells).
  const cols = ['Meeting Date', 'Meeting Time', 'Followup Date', 'Followup Time', 'Last Follow-up Date', 'Phone', 'Email', 'Clinic', 'Name'];
  const rows = Math.max(sh.getMaxRows(), 2);
  cols.forEach(function (h) { if (idx[h] !== undefined) sh.getRange(1, idx[h] + 1, rows, 1).setNumberFormat('@'); });
}

// Recover Phone cells that Sheets turned into #ERROR! formulas (numbers that
// began with "+"). The original digits are still inside the formula text.
function repairPhones_() {
  const sh = leadsSheet_();
  const idx = headerIndex_(sh);
  if (idx['Phone'] === undefined) return;
  const last = sh.getLastRow();
  if (last < 2) return;
  const rng = sh.getRange(2, idx['Phone'] + 1, last - 1, 1);
  rng.setNumberFormat('@');
  const fmls = rng.getFormulas();
  const vals = rng.getValues();
  let changed = false;
  for (let i = 0; i < fmls.length; i++) {
    if (fmls[i][0]) { vals[i][0] = String(fmls[i][0]).replace(/\D/g, ''); changed = true; }
  }
  if (changed) rng.setValues(vals);
}

function ensureSetup_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('FMT_SET_V3') === '1') return;
  try { ensureTextFormats_(); repairPhones_(); props.setProperty('FMT_SET_V3', '1'); } catch (e) { Logger.log('fmt setup: ' + e); }
}

// ---- Accurate cumulative stats (per intern), incremented as calls are logged.
function bumpStat_(intern, outcome) {
  if (!intern || !outcome) return;
  const props = PropertiesService.getScriptProperties();
  const all = JSON.parse(props.getProperty('STATS') || '{}');
  const s = all[intern] || { dials: 0, conversations: 0, meetings: 0, followups: 0, landed: 0 };
  s.dials++;
  if (/interested|landed/i.test(outcome)) s.conversations++;   // Interested / Not Interested / Landed
  if (/meeting booked/i.test(outcome)) s.meetings++;
  if (/follow-up/i.test(outcome)) s.followups++;
  if (/landed/i.test(outcome)) s.landed++;
  all[intern] = s;
  props.setProperty('STATS', JSON.stringify(all));
}

/* ============================ WEB API =========================== */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'list';
  try {
    ensureSetup_();
    // One-shot load: everything the app needs in a SINGLE round-trip (much
    // faster than calling list + getStats + getSessions separately).
    if (action === 'bootstrap')   return json_({ ok: true, leads: apiListLeads_(), interns: readInterns_(), stats: apiGetStats_(), sessions: getSessions_('') });
    if (action === 'list')        return json_({ ok: true, leads: apiListLeads_(), interns: readInterns_() });
    if (action === 'getStats')    return json_({ ok: true, stats: apiGetStats_() });
    if (action === 'getSessions') return json_({ ok: true, sessions: getSessions_((e && e.parameter && e.parameter.date) || '') });
    if (action === 'getDailyCalls') return json_({ ok: true, daily: getDaily_() });
    return json_({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, error: 'Bad JSON' }); }
  const action = body.action;

  // Serialize writes so 4 people using the tool at once can't corrupt the sheet.
  const lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (e) { return json_({ ok: false, error: 'Server busy, please retry' }); }
  try {
    ensureSetup_();
    switch (action) {
      case 'upsert':       return json_({ ok: true, lead: upsertLead_(body.lead, body.logStat) });
      case 'bulkImport':   return json_(Object.assign({ ok: true }, bulkImport_(body.leads || [], body.updateExisting)));
      case 'bulkAssign':   return json_({ ok: true, assigned: bulkAssign_(body) });
      case 'bulkField':    return json_({ ok: true, updated: bulkField_(body) });
      case 'delete':       return json_({ ok: true, deleted: deleteLead_(body.id) });
      case 'clearAll':     return json_({ ok: true, cleared: clearAll_() });
      case 'saveStats':    return json_({ ok: true, saved: saveStats_(body.stats) });
      case 'addIntern':    return json_({ ok: true, interns: addIntern_(body.name, body.email) });
      case 'removeIntern': return json_({ ok: true, interns: removeIntern_(body.name) });
      case 'attachRecordings': return json_(Object.assign({ ok: true }, attachRecordings_()));
      case 'coachCalls':       return json_(Object.assign({ ok: true }, coachCalls_(body.onlyToday !== false)));
      // ---- Intern time tracking + logout reconciliation ----
      case 'sessionStart': return json_(Object.assign({ ok: true }, sessionStart_(body)));
      case 'sessionPing':  return json_(Object.assign({ ok: true }, sessionPing_(body)));
      case 'sessionEnd':   return json_(Object.assign({ ok: true }, sessionEnd_(body)));
      case 'reconcile':    return json_(Object.assign({ ok: true }, reconcile_(body)));
      case 'markNoRecording': return json_(Object.assign({ ok: true }, markNoRecording_(body)));
      default:             return json_({ ok: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ========================= API HANDLERS ======================== */
const DATE_COLS = { 'Meeting Date':1, 'Followup Date':1, 'Last Follow-up Date':1, 'Timestamp':1 };
const TIME_COLS = { 'Meeting Time':1, 'Followup Time':1 };

// Sheets returns date/time cells as Date objects, which JSON-serialize in a
// way the frontend can't reliably parse. Normalize them to clean strings.
function fmtCell_(h, v) {
  if (v instanceof Date) {
    const tz = Session.getScriptTimeZone();
    if (TIME_COLS[h]) return Utilities.formatDate(v, tz, 'HH:mm');
    if (DATE_COLS[h]) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  return v;
}

// Ensure every lead row has a stable ID (older/hand-entered rows may lack one).
// Without this, upserts can duplicate rows and assignment-by-id silently misses.
function ensureIds_() {
  const sh = leadsSheet_();
  const last = sh.getLastRow();
  if (last < 2) return;
  const idx = headerIndex_(sh);
  if (idx['ID'] === undefined) return;
  const col = idx['ID'] + 1;
  const ids = sh.getRange(2, col, last - 1, 1).getValues();
  let changed = false;
  for (let i = 0; i < ids.length; i++) {
    if (!ids[i][0]) { ids[i][0] = 'L' + Utilities.getUuid(); changed = true; }
  }
  if (changed) sh.getRange(2, col, ids.length, 1).setValues(ids);
}

function apiListLeads_() {
  ensureIds_();
  const { rows } = readLeads_();
  return rows.map(function (r) {
    const o = {};
    LEAD_HEADERS.forEach(function (h) { o[h] = fmtCell_(h, r[h] === undefined ? '' : r[h]); });
    o._row = r._row;
    return o;
  });
}

// Insert or update a single lead (by ID). Returns the stored lead.
// logStat (optional): { intern, outcome } — increments cumulative call stats.
function upsertLead_(lead, logStat) {
  const sh = leadsSheet_();
  const idx = headerIndex_(sh);
  ensureId_(lead);
  if (lead['Phone'] !== undefined) lead['Phone'] = sanitizePhone_(lead['Phone']);
  // Any logged call outcome marks the lead as "worked today" so the logout
  // reconciliation can find today's answered calls even if no note was typed.
  if (logStat && logStat.outcome) {
    lead['Last Follow-up Date'] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const found = findRowById_(sh, idx, lead.ID);
  const rowArr = buildRowArray_(lead, idx, found ? found.current : null);
  if (found) {
    sh.getRange(found.row, 1, 1, rowArr.length).setValues([rowArr]);
  } else {
    if (!lead.Timestamp) rowArr[idx['Timestamp']] = new Date();
    sh.appendRow(rowArr);
  }
  const stored = objFromRow_(rowArr, idx);
  try { syncCalendar_(stored); } catch (e) { Logger.log('cal sync: ' + e); }
  try { maybeSendDoctorBookingEmail_(stored); } catch (e) { Logger.log('doc booking email: ' + e); }
  if (logStat) {
    try { bumpStat_(logStat.intern, logStat.outcome); } catch (e) { Logger.log('stat: ' + e); }
    try { bumpDaily_(logStat.intern, logStat.outcome); } catch (e) { Logger.log('daily: ' + e); }
  }
  return stored;
}

// Create/update/delete a Google Calendar event to mirror the lead's meeting.
// Booked -> create/update. Not booked (cancelled/rescheduled away) -> delete.
function syncCalendar_(lead) {
  if (!CONFIG.CREATE_CALENDAR_EVENTS) return;
  const props = PropertiesService.getScriptProperties();
  const propKey = 'CALEVT_' + lead.ID;
  const booked = /booked/i.test(String(lead['Status'] || ''));
  const at = booked ? toDateTime_(lead['Meeting Date'], lead['Meeting Time']) : null;

  const cal = CONFIG.CALENDAR_ID === 'primary'
    ? CalendarApp.getDefaultCalendar()
    : CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if (!cal) return;

  const existingId0 = props.getProperty(propKey);
  if (!booked || !at) {
    // meeting cancelled or moved off "booked" — remove any existing event
    if (existingId0) { try { const ev = cal.getEventById(existingId0); if (ev) ev.deleteEvent(); } catch (e) {} props.deleteProperty(propKey); }
    return;
  }

  const end = new Date(at.getTime() + CONFIG.MEETING_DURATION_MIN * 60000);
  const title = 'Reactivify: ' + (lead['Clinic'] || lead['Name'] || 'Meeting');
  const guests = [];
  if (lead['Email']) guests.push(lead['Email']);
  const internEm = internEmail_(lead['Assigned To']);
  if (internEm) guests.push(internEm);

  const existingId = existingId0;
  const desc = 'Auto-created by Reactivify CRM.\nAssigned: ' + (lead['Assigned To'] || '')
    + '\nPhone: ' + (lead['Phone'] || '') + '\nClinic: ' + (lead['Clinic'] || '');
  try {
    if (existingId) {
      const ev = cal.getEventById(existingId);
      if (ev) { ev.setTime(at, end); ev.setTitle(title); ev.setDescription(desc); return; }
    }
    const ev = cal.createEvent(title, at, end, {
      description: desc,
      guests: guests.join(','),
      sendInvites: CONFIG.SEND_CALENDAR_INVITES
    });
    PropertiesService.getScriptProperties().setProperty(propKey, ev.getId());
  } catch (e) { Logger.log('createEvent fail: ' + e); }
}

function ensureId_(lead) {
  // Utilities.getUuid() is globally unique — prevents the ID collisions that
  // happened when a big import generated hundreds of IDs in the same millisecond.
  if (!lead.ID) lead.ID = 'L' + Utilities.getUuid();
  return lead.ID;
}

// Strip leading characters Google Sheets treats as a formula (+ = - @) so a
// hand-typed "+91..." phone number never turns into #ERROR!.
function sanitizePhone_(v) {
  if (v === undefined || v === null) return v;
  return String(v).trim().replace(/^[+=\-@]+/, '');
}

function findRowById_(sh, idx, id) {
  if (!id) return null;
  const last = sh.getLastRow();
  if (last < 2) return null;
  const col = idx['ID'] + 1;
  const ids = sh.getRange(2, col, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      const current = sh.getRange(i + 2, 1, 1, sh.getLastColumn()).getValues()[0];
      return { row: i + 2, current: current };
    }
  }
  return null;
}

function buildRowArray_(lead, idx, current) {
  const width = Object.keys(idx).length;
  const arr = current ? current.slice() : new Array(width).fill('');
  Object.keys(idx).forEach(function (h) {
    if (lead[h] !== undefined) arr[idx[h]] = lead[h];
  });
  return arr;
}

function objFromRow_(arr, idx) {
  const o = {};
  Object.keys(idx).forEach(function (h) { o[h] = arr[idx[h]]; });
  return o;
}

// CSV / scraper import. Accepts already-mapped lead objects from frontend.
// Returns { imported, skipped, overlaps:[names] }.
// Duplicates (same phone, else same name) are NOT re-added — the existing
// row is kept intact so Notes/Status/Assignment/meeting history survive.
// If updateExisting is true, refresh only the scraped fields on overlaps
// (rating/reviews/email/city/phone) while still preserving notes etc.
function bulkImport_(leads, updateExisting) {
  const sh = leadsSheet_();
  const idx = headerIndex_(sh);
  const byKey = {};
  readLeads_().rows.forEach(function (r) { byKey[dedupeKey_(r.Phone, r.Name)] = r; });

  const toAppend = [];
  let imported = 0, skipped = 0;
  const overlaps = [];

  leads.forEach(function (l) {
    const key = dedupeKey_(l.Phone || l.phone, l.Name || l.name);
    const hit = byKey[key];
    if (hit && hit.ID) {                       // overlap with an existing lead
      overlaps.push(l.Name || l.name || l.Clinic || l.clinic || '(unnamed)');
      if (updateExisting) {
        const patch = { ID: hit.ID };
        ['Rating', 'Reviews', 'Email', 'City', 'Phone'].forEach(function (f) {
          const v = l[f] !== undefined ? l[f] : l[f.toLowerCase()];
          if (v !== undefined && v !== '') patch[f] = v;
        });
        upsertLead_(patch);                    // notes/status/assignment untouched
      }
      skipped++;
      return;
    }
    byKey[key] = { ID: 'pending' };            // guard against dupes within this batch
    ensureId_(l);
    if (l['Phone'] !== undefined) l['Phone'] = sanitizePhone_(l['Phone']);
    if (!l.Timestamp) l.Timestamp = new Date();
    if (!l.Status) l.Status = 'New';
    if (!l['Assigned To']) l['Assigned To'] = 'Team';   // land in owner's account first
    toAppend.push(buildRowArray_(l, idx, null));
    imported++;
  });
  if (toAppend.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, toAppend[0].length).setValues(toAppend);
  }
  return { imported: imported, skipped: skipped, overlaps: overlaps.slice(0, 300) };
}

function dedupeKey_(phone, name) {
  const p = String(phone || '').replace(/\D/g, '');
  if (p) return 'p:' + p.slice(-10);
  return 'n:' + String(name || '').trim().toLowerCase();
}

// Bulk assign. Modes: {mode:'auto'} round-robin unassigned;
// {mode:'range', from, to, intern} assign serial rows from..to.
function bulkAssign_(body) {
  const sh = leadsSheet_();
  const { rows, idx } = readLeads_();
  const interns = readInterns_().map(function (i) { return i.name; });
  let n = 0;

  // "unassigned" = blank OR still sitting in the Team/owner account
  function isPool_(v){ return !v || String(v).trim().toLowerCase() === 'team'; }

  if (body.mode === 'auto') {
    if (!interns.length) return 0;
    let k = 0;
    rows.forEach(function (r) {
      if (isPool_(r['Assigned To'])) {
        sh.getRange(r._row, idx['Assigned To'] + 1).setValue(interns[k % interns.length]);
        k++; n++;
      }
    });
  } else if (body.mode === 'ids') {
    // Preferred path: frontend sends the exact lead IDs it wants assigned,
    // computed from the sorted list the admin is actually looking at.
    const who = body.intern;
    const idSet = {};
    (body.ids || []).forEach(function (x) { idSet[String(x)] = true; });
    rows.forEach(function (r) {
      if (idSet[String(r['ID'])]) {
        sh.getRange(r._row, idx['Assigned To'] + 1).setValue(who);
        n++;
      }
    });
  } else if (body.mode === 'range') {
    // Legacy serial-range fallback (sheet order).
    const from = Math.max(1, parseInt(body.from, 10));
    const to = parseInt(body.to, 10);
    const intern = body.intern;
    rows.forEach(function (r, i) {
      const serial = i + 1;
      if (serial >= from && serial <= to) {
        sh.getRange(r._row, idx['Assigned To'] + 1).setValue(intern);
        n++;
      }
    });
  }
  return n;
}

// Set any field on a set of sheet rows (by 1-based sheet row number).
// Used for assignment and city cleanup — no dependency on lead IDs.
function bulkField_(body) {
  const sh = leadsSheet_();
  const idx = headerIndex_(sh);
  const field = body.field;
  if (idx[field] === undefined) return 0;
  const value = body.value;
  const rows = body.rows || [];
  let n = 0;
  rows.forEach(function (rw) {
    const r = parseInt(rw, 10);
    if (r >= 2) { sh.getRange(r, idx[field] + 1).setValue(value); n++; }
  });
  return n;
}

function deleteLead_(id) {
  const sh = leadsSheet_();
  const idx = headerIndex_(sh);
  const found = findRowById_(sh, idx, id);
  if (found) { sh.deleteRow(found.row); return 1; }
  return 0;
}

function clearAll_() {
  const sh = leadsSheet_();
  const last = sh.getLastRow();
  if (last > 1) sh.deleteRows(2, last - 1);
  return 1;
}

/* ===================== STATS (PropertiesService) ================ */
function apiGetStats_() {
  const raw = PropertiesService.getScriptProperties().getProperty('STATS');
  return raw ? JSON.parse(raw) : {};
}
function saveStats_(stats) {
  PropertiesService.getScriptProperties().setProperty('STATS', JSON.stringify(stats || {}));
  return true;
}

// Per-day, per-intern call log — every logged dial (including redials to the
// same lead). Kept in a Script Property, pruned to the last 120 days.
function bumpDaily_(intern, outcome) {
  if (!intern) return;
  const day = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const props = PropertiesService.getScriptProperties();
  const all = JSON.parse(props.getProperty('DAILY') || '{}');
  const d = all[day] || {};
  const rec = d[intern] || { calls: 0 };
  rec.calls = (rec.calls || 0) + 1;
  if (outcome) rec[outcome] = (rec[outcome] || 0) + 1;
  d[intern] = rec; all[day] = d;
  const keys = Object.keys(all).sort();
  while (keys.length > 120) { delete all[keys.shift()]; }
  props.setProperty('DAILY', JSON.stringify(all));
}
function getDaily_() { return JSON.parse(PropertiesService.getScriptProperties().getProperty('DAILY') || '{}'); }

/* ================================================================
 *            INTERN TIME TRACKING (Sessions sheet)
 * ==============================================================
 * The frontend clocks interns in on login, watches for 2 min of
 * inactivity (idle), and clocks out on logout. The client is the
 * source of truth for the work/idle timers; the backend just records
 * them so the admin can see who logged in when, and for how long.
 * ============================================================== */
function sessionsSheetEnsure_() {
  let sh = ss_().getSheetByName(CONFIG.SHEET_SESSIONS);
  if (!sh) { sh = ss_().insertSheet(CONFIG.SHEET_SESSIONS); sh.getRange(1, 1, 1, SESSION_HEADERS.length).setValues([SESSION_HEADERS]); }
  else if (sh.getLastRow() < 1) { sh.getRange(1, 1, 1, SESSION_HEADERS.length).setValues([SESSION_HEADERS]); }
  // Keep everything as plain text so times/dates aren't mangled into serials.
  try { sh.getRange(1, 1, Math.max(sh.getMaxRows(), 2), SESSION_HEADERS.length).setNumberFormat('@'); } catch (e) {}
  return sh;
}

function findSessionRow_(sh, idx, sid) {
  const last = sh.getLastRow();
  if (last < 2 || !sid) return 0;
  const col = idx['Session ID'] + 1;
  const vals = sh.getRange(2, col, last - 1, 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) { if (String(vals[i][0]) === String(sid)) return i + 2; }
  return 0;
}

function sessionStart_(b) {
  const sh = sessionsSheetEnsure_();
  const idx = headerIndex_(sh);
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  // If this session id already has a row (re-issued start), just refresh it.
  const existing = findSessionRow_(sh, idx, b.sid);
  if (existing) {
    sh.getRange(existing, idx['Last Active'] + 1).setValue(Utilities.formatDate(now, tz, 'HH:mm:ss'));
    sh.getRange(existing, idx['Status'] + 1).setValue('Working');
    return { session: b.sid };
  }
  const row = new Array(SESSION_HEADERS.length).fill('');
  row[idx['Session ID']]  = b.sid || ('S' + now.getTime());
  row[idx['Date']]        = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  row[idx['Intern']]      = b.intern || '';
  row[idx['Login']]       = Utilities.formatDate(now, tz, 'HH:mm:ss');
  row[idx['Last Active']] = Utilities.formatDate(now, tz, 'HH:mm:ss');
  row[idx['Status']]      = 'Working';
  row[idx['Work Min']]    = 0;
  row[idx['Idle Min']]    = 0;
  sh.appendRow(row);
  return { session: b.sid };
}

function sessionPing_(b) {
  const sh = sessionsSheetEnsure_();
  const idx = headerIndex_(sh);
  const row = findSessionRow_(sh, idx, b.sid);
  if (!row) return { updated: false };
  const tz = Session.getScriptTimeZone();
  sh.getRange(row, idx['Last Active'] + 1).setValue(Utilities.formatDate(new Date(), tz, 'HH:mm:ss'));
  if (b.status) sh.getRange(row, idx['Status'] + 1).setValue(b.status);
  if (b.workSec !== undefined) sh.getRange(row, idx['Work Min'] + 1).setValue(Math.round(b.workSec / 60));
  if (b.idleSec !== undefined) sh.getRange(row, idx['Idle Min'] + 1).setValue(Math.round(b.idleSec / 60));
  return { updated: true };
}

function sessionEnd_(b) {
  const sh = sessionsSheetEnsure_();
  const idx = headerIndex_(sh);
  const row = findSessionRow_(sh, idx, b.sid);
  if (!row) return { updated: false };
  const tz = Session.getScriptTimeZone();
  sh.getRange(row, idx['Logout'] + 1).setValue(Utilities.formatDate(new Date(), tz, 'HH:mm:ss'));
  sh.getRange(row, idx['Status'] + 1).setValue('Logged out');
  if (b.workSec !== undefined) sh.getRange(row, idx['Work Min'] + 1).setValue(Math.round(b.workSec / 60));
  if (b.idleSec !== undefined) sh.getRange(row, idx['Idle Min'] + 1).setValue(Math.round(b.idleSec / 60));
  if (b.convos !== undefined) sh.getRange(row, idx['Convos'] + 1).setValue(b.convos);
  if (b.recordings !== undefined) sh.getRange(row, idx['Recordings'] + 1).setValue(b.recordings);
  sh.getRange(row, idx['Reconciled'] + 1).setValue(b.reconciled ? 'Yes' : 'No');
  if (b.notes) sh.getRange(row, idx['Notes'] + 1).setValue(String(b.notes).slice(0, 2000));
  return { updated: true };
}

// Return session rows for a given day (default: today), formatted as strings.
function getSessions_(dateStr) {
  const sh = sessionsSheetEnsure_();
  const idx = headerIndex_(sh);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const tz = Session.getScriptTimeZone();
  const today = dateStr || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const vals = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  const out = [];
  vals.forEach(function (r) {
    const o = {};
    Object.keys(idx).forEach(function (h) {
      let v = r[idx[h]];
      if (v instanceof Date) {
        v = (h === 'Date') ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : Utilities.formatDate(v, tz, 'HH:mm:ss');
      }
      o[h] = v;
    });
    if (String(o['Date']).indexOf(today) >= 0) out.push(o);
  });
  return out;
}

/* ================================================================
 *      LOGOUT RECONCILIATION (convos today vs recordings)
 * ==============================================================
 * On logout we: (1) rescan the Drive recordings folder RECURSIVELY
 * (interns upload into their own name/date subfolders) and attach any
 * new recordings by phone match, then (2) list this intern's answered
 * calls from today that still have no recording and no "doesn't exist"
 * reason. The frontend hard-blocks logout until each is resolved.
 * ============================================================== */
function reconcile_(b) {
  const intern = String(b.intern || '').trim();
  const sh = leadsSheet_();
  ensureColumns_(sh, [CONFIG.RECORD_COL, CONFIG.REC_MISSING_COL]);

  // Pull in anything freshly uploaded to Drive (recursive, matches by phone).
  let att = { attached: 0, scanned: 0, unmatched: 0 };
  try { att = attachRecordings_(); } catch (e) { Logger.log('reconcile attach: ' + e); }

  const { rows } = readLeads_();
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const req = {};
  CONFIG.RECORDING_REQUIRED_OUTCOMES.forEach(function (s) { req[String(s).toLowerCase()] = 1; });

  const mine = rows.filter(function (r) {
    return String(r['Assigned To'] || '').trim().toLowerCase() === intern.toLowerCase()
      && String(r['Last Follow-up Date'] || '').indexOf(today) >= 0
      && req[String(r['Status'] || '').trim().toLowerCase()];
  });

  let recordings = 0;
  const missing = [];
  mine.forEach(function (r) {
    const hasRec = String(r[CONFIG.RECORD_COL] || '').trim() !== '';
    const hasReason = String(r[CONFIG.REC_MISSING_COL] || '').trim() !== '';
    if (hasRec) { recordings++; return; }
    if (hasReason) { recordings++; return; }   // resolved as "no recording exists" — counts as accounted-for
    missing.push({ id: r['ID'], clinic: r['Clinic'] || r['Name'] || '(unnamed)', name: r['Name'] || '', phone: r['Phone'] || '' });
  });

  return {
    convos: mine.length,
    recordings: recordings,
    missing: missing,
    scanned: att.scanned,
    attached: att.attached,
    unmatched: att.unmatched      // Drive files that matched no lead ("misaligned")
  };
}

// Record the intern's reason that a recording genuinely doesn't exist, so the
// lead stops being flagged and you can see why in the sheet.
function markNoRecording_(b) {
  const sh = leadsSheet_();
  ensureColumns_(sh, [CONFIG.REC_MISSING_COL]);
  const idx = headerIndex_(sh);
  const found = findRowById_(sh, idx, b.id);
  if (!found) return { updated: false, error: 'Lead not found' };
  const tz = Session.getScriptTimeZone();
  const stamp = '[' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd') + ' · ' + (b.intern || '') + '] '
    + String(b.reason || '').slice(0, 500);
  sh.getRange(found.row, idx[CONFIG.REC_MISSING_COL] + 1).setValue(stamp);
  return { updated: true };
}

/* ========================================================================
 *                       AUTOMATION / REMINDER ENGINE
 * ===================================================================== */

// Install all time-based triggers. Run ONCE from the editor.
function setupTriggers() {
  ensureTextFormats_();                        // make date/time cols plain-text
  // Clear existing to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });

  ScriptApp.newTrigger('checkMeetingReminders').timeBased().everyMinutes(10).create();
  ScriptApp.newTrigger('checkConflicts').timeBased().everyMinutes(10).create();
  ScriptApp.newTrigger('checkFollowUpReminders').timeBased().everyMinutes(10).create();
  ScriptApp.newTrigger('sendFollowUpDigest').timeBased().atHour(9).everyDays(1).create();

  return 'Triggers installed.';
}

// Parse a Sheet date + time into a JS Date in the script timezone.
function toDateTime_(dateVal, timeVal) {
  if (!dateVal) return null;
  const tz = Session.getScriptTimeZone();
  let d;
  if (dateVal instanceof Date) d = new Date(dateVal.getTime());
  else { d = new Date(dateVal); if (isNaN(d)) return null; }

  let hh = 0, mm = 0;
  if (timeVal instanceof Date) { hh = timeVal.getHours(); mm = timeVal.getMinutes(); }
  else if (timeVal) {
    const m = String(timeVal).match(/(\d{1,2}):(\d{2})\s*(AM|PM)?/i);
    if (m) {
      hh = parseInt(m[1], 10); mm = parseInt(m[2], 10);
      const ap = (m[3] || '').toUpperCase();
      if (ap === 'PM' && hh < 12) hh += 12;
      if (ap === 'AM' && hh === 12) hh = 0;
    }
  }
  d.setHours(hh, mm, 0, 0);
  return d;
}

// Core "is this offset due right now?" test, shared + unit-tested.
// Returns true if `now` is within tolerance of (eventTime - offsetMin).
function isDue_(now, eventTime, offsetMin, toleranceMin) {
  if (!eventTime) return false;
  const target = eventTime.getTime() - offsetMin * 60000;
  const diff = now.getTime() - target;              // >=0 means we've reached it
  return diff >= 0 && diff < toleranceMin * 60000;
}

// Build the intern "message the doctor" links (WhatsApp + email).
function internNudgeLinks_(lead) {
  const phone = String(lead['Phone'] || '').replace(/\D/g, '');
  const dateStr = fmtDate_(lead['Meeting Date']);
  const timeStr = fmtTime_(lead['Meeting Time']);
  const msg = CONFIG.TEMPLATE(lead['Name'], lead['Clinic'], dateStr, timeStr);
  const wa = phone ? ('https://wa.me/' + normalizePhone_(phone) + '?text=' + encodeURIComponent(msg)) : '';
  const mail = lead['Email'] ? ('mailto:' + lead['Email']
      + '?subject=' + encodeURIComponent(CONFIG.COMPANY_NAME + ' — meeting reminder')
      + '&body=' + encodeURIComponent(msg)) : '';
  return { wa: wa, mail: mail, msg: msg };
}

function normalizePhone_(digits) {
  // Default to India country code if a bare 10-digit number.
  if (digits.length === 10) return '91' + digits;
  return digits;
}

function fmtDate_(v) {
  if (!v) return '';
  const tz = Session.getScriptTimeZone();
  const d = (v instanceof Date) ? v : new Date(v);
  return isNaN(d) ? String(v) : Utilities.formatDate(d, tz, 'EEE, d MMM yyyy');
}
function fmtTime_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'h:mm a');
  return String(v);
}

/* ---------------------- Meeting reminders ---------------------- */
function checkMeetingReminders() {
  const now = new Date();
  const sh = leadsSheet_();
  const { rows, idx } = readLeads_();

  rows.forEach(function (lead) {
    const meetingAt = toDateTime_(lead['Meeting Date'], lead['Meeting Time']);
    if (!meetingAt) return;
    if (meetingAt.getTime() < now.getTime() - 60 * 60000) return;   // already past
    const status = String(lead['Status'] || '');
    if (/dumped|not interested/i.test(status)) return;

    CONFIG.MEETING_OFFSETS.forEach(function (off) {
      const alreadySent = String(lead[off.key]) === 'true' || lead[off.key] === true;
      if (alreadySent) return;

      // (a) client + owner + intern at the offset
      if (isDue_(now, meetingAt, off.minutes, CONFIG.TICK_TOLERANCE_MIN)) {
        fireMeetingReminder_(lead, off, 'SAME_TIME');
        sh.getRange(lead._row, idx[off.key] + 1).setValue('true');   // mark sent
      }
    });

    // Intern 15-min-prior nudge: fires before each offset if mode calls for it
    if (CONFIG.EMAIL_INTERN_LINK &&
        (CONFIG.INTERN_NUDGE_MODE === 'PRIOR_15' || CONFIG.INTERN_NUDGE_MODE === 'BOTH')) {
      CONFIG.MEETING_OFFSETS.forEach(function (off) {
        const priorFlag = off.key + ' InternPrior';
        // stored in a Script Property keyed per lead+offset (no extra column)
        if (getSentFlag_(lead.ID, priorFlag)) return;
        const priorOffset = off.minutes + CONFIG.INTERN_NUDGE_PRIOR_MIN;
        if (isDue_(now, meetingAt, priorOffset, CONFIG.TICK_TOLERANCE_MIN)) {
          emailInternNudge_(lead, off, 'PRIOR_15');
          setSentFlag_(lead.ID, priorFlag);
        }
      });
    }
  });
}

function fireMeetingReminder_(lead, off, when) {
  const assigned = lead['Assigned To'];
  const dateStr = fmtDate_(lead['Meeting Date']);
  const timeStr = fmtTime_(lead['Meeting Time']);

  // (a) client / doctor — reminder with the calendar link (24h / 4h / 45min)
  if (CONFIG.EMAIL_DOCTOR && lead['Email']) {
    sendDoctorEmail_(lead, off.label, toDateTime_(lead['Meeting Date'], lead['Meeting Time']));
  }

  // (b) owner — always notified at 1hr/30min
  if (off.minutes <= 60) {
    safeEmail_(CONFIG.OWNER_EMAIL,
      '[Reactivify] Meeting in ' + off.label + ': ' + (lead['Clinic'] || lead['Name']),
      'Meeting with Dr. ' + (lead['Name'] || '') + ' (' + (lead['Clinic'] || '') + ') at '
      + timeStr + ' on ' + dateStr + '. Assigned: ' + (assigned || 'unassigned') + '.');
  }

  // (c) intern — same-time link (if BOTH or SAME_TIME mode)
  if (CONFIG.EMAIL_INTERN_LINK &&
      (CONFIG.INTERN_NUDGE_MODE === 'SAME_TIME' || CONFIG.INTERN_NUDGE_MODE === 'BOTH')) {
    emailInternNudge_(lead, off, 'SAME_TIME');
  }
}

function emailInternNudge_(lead, off, when) {
  const to = internEmail_(lead['Assigned To']);
  if (!to) return;
  const links = internNudgeLinks_(lead);
  const dateStr = fmtDate_(lead['Meeting Date']);
  const timeStr = fmtTime_(lead['Meeting Time']);
  const whenLabel = when === 'PRIOR_15'
    ? ('in ~' + (off.minutes + CONFIG.INTERN_NUDGE_PRIOR_MIN) + ' min the doctor gets their ' + off.label + ' reminder')
    : ('the doctor is being reminded now (' + off.label + ')');

  const plain =
    'Hi ' + lead['Assigned To'] + ', time to message Dr. ' + (lead['Name'] || '') + ' (' + (lead['Clinic'] || '') + ').\n' +
    'Meeting: ' + dateStr + ' at ' + timeStr + ' (' + whenLabel + ')\n' +
    'WhatsApp: ' + (links.wa || 'n/a') + '\nEmail: ' + (links.mail || 'n/a') + '\nMessage: "' + links.msg + '"';

  const html = emailShell_(
    '📞 Message Dr. ' + (lead['Name'] || ''),
    (lead['Clinic'] || '') + ' &middot; Meeting <b>' + dateStr + '</b> at <b>' + timeStr + '</b><br>' + whenLabel,
    '<div style="background:#f1f5f9;padding:12px;border-radius:10px;margin-bottom:6px;font-size:14px">'
      + 'Suggested message:<br><i>&ldquo;' + links.msg + '&rdquo;</i></div>'
      + leadButtonsHtml_(lead));

  htmlEmail_(to, '⏰ Message Dr. ' + (lead['Name'] || '') + ' now — ' + (lead['Clinic'] || 'meeting'), plain, html);
}

/* ---------------------- Conflict detection --------------------- */
function checkConflicts() {
  const sh = leadsSheet_();
  const { rows, idx } = readLeads_();
  const slots = {};
  rows.forEach(function (lead) {
    const at = toDateTime_(lead['Meeting Date'], lead['Meeting Time']);
    if (!at) return;
    const key = Utilities.formatDate(at, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    (slots[key] = slots[key] || []).push(lead);
  });
  Object.keys(slots).forEach(function (key) {
    const group = slots[key];
    const conflict = group.length > 1;
    group.forEach(function (lead) {
      const cur = String(lead['Conflict Flag']) === 'true' || lead['Conflict Flag'] === true;
      if (conflict !== cur) sh.getRange(lead._row, idx['Conflict Flag'] + 1).setValue(conflict ? 'true' : '');
    });
    if (conflict) {
      safeEmail_(CONFIG.OWNER_EMAIL,
        '[Reactivify] ⚠ Double-booking at ' + key,
        'These leads share the slot ' + key + ':\n' +
        group.map(function (l) { return '• ' + (l['Clinic'] || l['Name']) + ' (' + (l['Assigned To'] || '') + ')'; }).join('\n'));
    }
  });
}

/* -------- Follow-up (cols V/W) reminders + daily digest -------- */
// One-shot reminder when a scheduled follow-up (Followup Date/Time) comes due.
function checkFollowUpReminders() {
  const now = new Date();
  const sh = leadsSheet_();
  const { rows, idx } = readLeads_();
  const hasSentCol = idx[CONFIG.FOLLOWUP_SENT_FLAG] !== undefined;

  rows.forEach(function (lead) {
    const at = toDateTime_(lead['Followup Date'], lead['Followup Time']);
    if (!at) return;
    const sent = hasSentCol ? (String(lead[CONFIG.FOLLOWUP_SENT_FLAG]) === 'true')
                            : getSentFlag_(lead.ID, 'FollowupDue');
    if (sent) return;
    const status = String(lead['Status'] || '');
    if (/dumped|landed|meeting booked/i.test(status)) return;

    // Due when now has reached the follow-up time (small look-back window)
    if (isDue_(now, at, 0, CONFIG.TICK_TOLERANCE_MIN)) {
      const to = internEmail_(lead['Assigned To']);
      const dateStr = fmtDate_(lead['Followup Date']);
      const timeStr = fmtTime_(lead['Followup Time']);
      const links = internNudgeLinks_(lead);
      if (to) {
        const plain = 'Follow-up due with Dr. ' + (lead['Name'] || '') + ' (' + (lead['Clinic'] || '') + '), '
          + dateStr + ' ' + timeStr + '. WhatsApp: ' + (links.wa || 'n/a') + ' Email: ' + (links.mail || 'n/a');
        const html = emailShell_(
          '🔔 Follow-up due: ' + (lead['Clinic'] || lead['Name']),
          'Dr. ' + (lead['Name'] || '') + ' &middot; scheduled ' + dateStr + ' ' + timeStr,
          '<div style="margin-bottom:6px;font-size:14px">Reach out now — one tap:</div>' + leadButtonsHtml_(lead));
        htmlEmail_(to, '⏰ Follow up now: ' + (lead['Clinic'] || lead['Name']), plain, html);
      }
      safeEmail_(CONFIG.OWNER_EMAIL, '[Reactivify] Follow-up due: ' + (lead['Clinic'] || lead['Name']),
        'Follow-up due for ' + (lead['Clinic'] || lead['Name']) + ' — assigned ' + (lead['Assigned To'] || 'nobody') + '.');

      if (hasSentCol) sh.getRange(lead._row, idx[CONFIG.FOLLOWUP_SENT_FLAG] + 1).setValue('true');
      else setSentFlag_(lead.ID, 'FollowupDue');
    }
  });
}

// Daily 9am: email each intern their overdue/cold leads.
function sendFollowUpDigest() {
  const { rows } = readLeads_();
  const now = new Date();
  const byIntern = {};
  rows.forEach(function (lead) {
    const status = String(lead['Status'] || '');
    if (/meeting booked|landed|dumped/i.test(status)) return;
    const assigned = lead['Assigned To'];
    if (!assigned) return;
    // "cold" = has a follow-up date in the past, OR no activity + not new-today
    const fu = toDateTime_(lead['Followup Date'], lead['Followup Time']);
    const lastFu = lead['Last Follow-up Date'] ? new Date(lead['Last Follow-up Date']) : null;
    const overdue = (fu && fu.getTime() < now.getTime()) ||
                    (lastFu && (now.getTime() - lastFu.getTime()) > 3 * 24 * 3600 * 1000);
    if (overdue) (byIntern[assigned] = byIntern[assigned] || []).push(lead);
  });

  Object.keys(byIntern).forEach(function (name) {
    const to = internEmail_(name);
    if (!to) return;
    const leads = byIntern[name];
    const plain = 'Good morning ' + name + ', ' + leads.length + ' leads need a follow-up:\n' +
      leads.map(function (l) { return '• ' + (l['Clinic'] || l['Name']) + ' — ' + (l['Phone'] || ''); }).join('\n');
    const cards = leads.map(function (l) {
      return '<div style="border:1px solid #e2e8f0;border-radius:12px;padding:12px;margin-bottom:10px">'
        + '<div style="font-weight:bold;font-size:15px">' + (l['Clinic'] || l['Name'] || 'Lead') + '</div>'
        + '<div style="color:#64748b;font-size:13px;margin-bottom:4px">' + (l['Phone'] || '') + ' &middot; ' + (l['City'] || '') + '</div>'
        + leadButtonsHtml_(l) + '</div>';
    }).join('');
    const html = emailShell_('☀️ Good morning, ' + name,
      '<b>' + leads.length + '</b> lead' + (leads.length === 1 ? '' : 's') + ' going cold — knock them out today.',
      cards);
    htmlEmail_(to, '📋 ' + leads.length + ' follow-ups waiting — ' + name, plain, html);
  });
}

/* ----------------------- shared utilities ---------------------- */
// Per-lead "sent" flags kept in Script Properties (avoids extra columns).
function getSentFlag_(id, key) {
  return PropertiesService.getScriptProperties().getProperty('SENT_' + id + '_' + key) === '1';
}
function setSentFlag_(id, key) {
  PropertiesService.getScriptProperties().setProperty('SENT_' + id + '_' + key, '1');
}

// Returns FROM_EMAIL only if it's actually a verified alias in this account,
// so a mis-typed/unverified alias never blocks sending.
function fromAlias_() {
  try {
    const aliases = GmailApp.getAliases();
    if (CONFIG.FROM_EMAIL && aliases.indexOf(CONFIG.FROM_EMAIL) >= 0) return CONFIG.FROM_EMAIL;
  } catch (e) {}
  return '';
}

function safeEmail_(to, subject, body) {
  if (!to) return;
  const from = fromAlias_();
  try {
    if (from) GmailApp.sendEmail(to, subject, body, { from: from, name: CONFIG.COMPANY_NAME });
    else MailApp.sendEmail(to, subject, body, { name: CONFIG.COMPANY_NAME });
  } catch (err) { Logger.log('email fail ' + to + ': ' + err); try { MailApp.sendEmail(to, subject, body); } catch (e) {} }
}

/* ------------------- Rich HTML reminder emails ------------------- */
function htmlEmail_(to, subject, plain, html) {
  if (!to) return;
  const from = fromAlias_();
  try {
    if (from) GmailApp.sendEmail(to, subject, plain || '', { from: from, htmlBody: html, name: CONFIG.COMPANY_NAME });
    else MailApp.sendEmail(to, subject, plain || '', { htmlBody: html, name: CONFIG.COMPANY_NAME });
  } catch (err) { Logger.log('html email fail ' + to + ': ' + err); try { MailApp.sendEmail(to, subject, plain || '', { htmlBody: html, name: CONFIG.COMPANY_NAME }); } catch (e) {} }
}
function btnHtml_(label, url, bg) {
  if (!url) return '';
  return '<a href="' + url + '" style="display:inline-block;padding:14px 20px;margin:6px 8px 6px 0;'
    + 'background:' + bg + ';color:#ffffff;text-decoration:none;border-radius:10px;font-weight:bold;'
    + 'font-size:16px;font-family:Arial,sans-serif">' + label + '</a>';
}
function leadButtonsHtml_(lead) {
  const phone = String(lead['Phone'] || '').replace(/\D/g, '');
  const links = internNudgeLinks_(lead);
  const tel = phone ? ('tel:+' + normalizePhone_(phone)) : '';
  return '<div style="margin:6px 0">' + btnHtml_('📞 Call', tel, '#16a34a')
    + btnHtml_('💬 WhatsApp', links.wa, '#22c55e')
    + btnHtml_('✉️ Email', links.mail, '#475569') + '</div>';
}
function emailShell_(heading, subhtml, bodyhtml) {
  return '<div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;padding:18px;color:#0f172a">'
    + '<div style="font-size:20px;font-weight:bold;color:#4f46e5">Reactivify</div>'
    + '<div style="font-size:18px;font-weight:bold;margin-top:12px">' + heading + '</div>'
    + (subhtml ? ('<div style="color:#64748b;margin:4px 0 14px;font-size:14px">' + subhtml + '</div>') : '')
    + bodyhtml
    + '<div style="color:#94a3b8;font-size:12px;margin-top:22px;border-top:1px solid #e2e8f0;padding-top:10px">'
    + 'Sent by Reactivify.</div></div>';
}

/* --------- Doctor meeting emails (link = add-to-calendar) --------- */
function getRawFlag_(k) { return PropertiesService.getScriptProperties().getProperty(k) === '1'; }
function setRawFlag_(k) { PropertiesService.getScriptProperties().setProperty(k, '1'); }

// Deterministic "add to Google Calendar" link — same link every time for a slot.
function gcalTemplateUrl_(lead) {
  const at = toDateTime_(lead['Meeting Date'], lead['Meeting Time']);
  if (!at) return '';
  const tz = Session.getScriptTimeZone();
  const end = new Date(at.getTime() + CONFIG.MEETING_DURATION_MIN * 60000);
  const stamp = function (d) { return Utilities.formatDate(d, tz, "yyyyMMdd'T'HHmmss"); };
  const text = encodeURIComponent(CONFIG.COMPANY_NAME + ' meeting: ' + (lead['Clinic'] || lead['Name'] || ''));
  const details = encodeURIComponent('Your meeting with ' + CONFIG.COMPANY_NAME + '.\nContact: ' + (lead['Phone'] || ''));
  return 'https://calendar.google.com/calendar/render?action=TEMPLATE&text=' + text
    + '&dates=' + stamp(at) + '/' + stamp(end)
    + '&details=' + details
    + '&ctz=' + encodeURIComponent(tz);
}

// Send the doctor a rich email with the calendar link. kind: 'booking' | offset label.
function sendDoctorEmail_(lead, kind, at) {
  const to = lead['Email'];
  if (!to) return;
  const dateStr = fmtDate_(lead['Meeting Date']);
  const timeStr = fmtTime_(lead['Meeting Time']);
  const link = gcalTemplateUrl_(lead);
  const isBooking = (kind === 'booking');
  const heading = isBooking ? '✅ Your meeting is confirmed' : '⏰ Reminder: meeting in ' + kind;
  const sub = 'with ' + CONFIG.COMPANY_NAME + ' &middot; <b>' + dateStr + '</b> at <b>' + timeStr + '</b>';
  const intro = isBooking
    ? 'Hi Dr. ' + (lead['Name'] || '') + ', thanks for booking a meeting with us. Add it to your calendar so it’s locked in:'
    : 'Hi Dr. ' + (lead['Name'] || '') + ', a quick reminder about your upcoming meeting with us. Add it to your calendar:';
  const html = emailShell_(heading, sub,
    '<div style="font-size:14px;margin-bottom:10px">' + intro + '</div>'
    + btnHtml_('📅 Add to my calendar', link, '#4f46e5'));
  const subject = isBooking
    ? ('Meeting confirmed with ' + CONFIG.COMPANY_NAME + ' — ' + dateStr + ' ' + timeStr)
    : ('Reminder: your ' + CONFIG.COMPANY_NAME + ' meeting — ' + dateStr + ' ' + timeStr);
  htmlEmail_(to, subject, 'Meeting ' + dateStr + ' at ' + timeStr + '. Add to calendar: ' + link, html);
}

/* ================================================================
 *              CALL RECORDINGS + AI COACHING (Groq)
 * ============================================================== */
function last10_(v){ const d = String(v || '').replace(/\D/g, ''); return d.length >= 10 ? d.slice(-10) : ''; }

// Walk a Drive folder AND all of its subfolders, calling cb(file) for every
// file. Interns upload into per-name / per-date subfolders, so we must recurse.
function eachFileDeep_(folder, cb) {
  const files = folder.getFiles();
  while (files.hasNext()) cb(files.next());
  const subs = folder.getFolders();
  while (subs.hasNext()) eachFileDeep_(subs.next(), cb);
}

// Phase 1 (free): scan the Drive folder (recursively, incl. name/date
// subfolders), match each recording to a lead by the phone number in its
// filename, and write the Drive link onto that lead.
// Idempotent — re-running only adds links that aren't already there.
function attachRecordings_(){
  const sh = leadsSheet_();
  const idx = ensureColumns_(sh, [CONFIG.RECORD_COL]);
  if (idx[CONFIG.RECORD_COL] === undefined) return { error: 'Could not create a "Recording" column.' };
  const { rows } = readLeads_();
  const byPhone = {};
  rows.forEach(function (r) { const k = last10_(r['Phone']); if (k) (byPhone[k] = byPhone[k] || []).push(r); });

  let folder;
  try { folder = DriveApp.getFolderById(CONFIG.RECORDINGS_FOLDER_ID); }
  catch (e) { return { error: 'Cannot open the recordings folder — check it is shared with this account.' }; }

  let attached = 0, scanned = 0, unmatched = 0;
  eachFileDeep_(folder, function (f) {
    scanned++;
    const m = f.getName().match(/\+?(\d{10,13})/);
    if (!m) { unmatched++; return; }
    const key = m[1].slice(-10);
    const matches = byPhone[key];
    if (!matches || !matches.length) { unmatched++; return; }
    const url = f.getUrl();
    matches.forEach(function (r) {
      const cell = sh.getRange(r._row, idx[CONFIG.RECORD_COL] + 1);
      const cur = String(cell.getValue() || '');
      if (cur.indexOf(url) < 0) { cell.setValue(cur ? (cur + '\n' + url) : url); attached++; }
    });
  });
  return { attached: attached, scanned: scanned, unmatched: unmatched };
}

// Phase 2 (Groq): transcribe + coach recordings that don't have coaching yet.
// onlyToday=true limits to leads worked today (Last Follow-up Date = today).
function coachCalls_(onlyToday){
  if (!CONFIG.GROQ_API_KEY || CONFIG.GROQ_API_KEY.indexOf('PASTE') === 0) return { error: 'Paste your Groq API key into CONFIG.GROQ_API_KEY first.' };
  const sh = leadsSheet_();
  const idx = headerIndex_(sh);
  if (idx[CONFIG.RECORD_COL] === undefined || idx[CONFIG.COACH_COL] === undefined) return { error: 'Add "Recording" and "Coaching" column headers first.' };
  const { rows } = readLeads_();
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const candidates = rows.filter(function (r) {
    if (!String(r[CONFIG.RECORD_COL] || '').trim()) return false;
    if (String(r[CONFIG.COACH_COL] || '').trim()) return false;
    if (onlyToday && String(r['Last Follow-up Date'] || '').indexOf(todayStr) < 0) return false;
    return true;
  });

  let coached = 0, errors = 0, remaining = 0;
  for (let i = 0; i < candidates.length; i++) {
    if (coached >= CONFIG.COACH_BATCH) { remaining = candidates.length - i; break; }
    const r = candidates[i];
    try {
      const url = String(r[CONFIG.RECORD_COL]).split('\n')[0];
      const fid = (url.match(/[-\w]{25,}/) || [])[0];
      if (!fid) { continue; }
      const blob = DriveApp.getFileById(fid).getBlob();
      const transcript = groqTranscribe_(blob);
      const coaching = groqCoach_(transcript, r);
      if (idx[CONFIG.TRANSCRIPT_COL] !== undefined) sh.getRange(r._row, idx[CONFIG.TRANSCRIPT_COL] + 1).setValue(String(transcript).slice(0, 45000));
      sh.getRange(r._row, idx[CONFIG.COACH_COL] + 1).setValue(String(coaching).slice(0, 45000));
      coached++;
    } catch (e) { errors++; Logger.log('coach fail ' + r.ID + ': ' + e); }
  }
  return { coached: coached, remaining: remaining, errors: errors, candidates: candidates.length };
}

function groqTranscribe_(blob){
  const resp = UrlFetchApp.fetch('https://api.groq.com/openai/v1/audio/transcriptions', {
    method: 'post',
    headers: { Authorization: 'Bearer ' + CONFIG.GROQ_API_KEY },
    payload: { file: blob, model: CONFIG.GROQ_TRANSCRIBE_MODEL, response_format: 'text' },
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  const txt = resp.getContentText();
  if (code >= 300) throw new Error('groq transcribe ' + code + ': ' + txt.slice(0, 200));
  return txt;
}

function groqCoach_(transcript, lead){
  const prompt = 'You are a sales coach for Reactivify (we sell websites to physiotherapy / clinic owners in India). '
    + 'The call may be in Hindi / Hinglish. Review this cold/sales call transcript and coach the caller. Be specific and concise.\n\n'
    + 'Format your answer as:\n1. Score: X/10\n2. What went well (2-3 short bullets)\n3. What to improve (2-3 short bullets)\n'
    + '4. Objections / missed opportunities\n5. One concrete next action\n\n'
    + 'Clinic: ' + (lead['Clinic'] || lead['Name'] || '') + ' | Logged outcome: ' + (lead['Status'] || '') + '\n\nTranscript:\n' + transcript;
  const resp = UrlFetchApp.fetch('https://api.groq.com/openai/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + CONFIG.GROQ_API_KEY },
    payload: JSON.stringify({ model: CONFIG.GROQ_COACH_MODEL, temperature: 0.4, messages: [{ role: 'user', content: prompt }] }),
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  const txt = resp.getContentText();
  if (code >= 300) throw new Error('groq coach ' + code + ': ' + txt.slice(0, 200));
  const j = JSON.parse(txt);
  return (j.choices && j.choices[0] && j.choices[0].message && j.choices[0].message.content) || '(no coaching returned)';
}

// Immediately email the doctor when a meeting is first booked (once per slot).
function maybeSendDoctorBookingEmail_(lead) {
  if (!CONFIG.EMAIL_DOCTOR) return;
  if (!/booked/i.test(String(lead['Status'] || ''))) return;
  if (!lead['Email']) return;
  const at = toDateTime_(lead['Meeting Date'], lead['Meeting Time']);
  if (!at) return;
  const flag = 'DOCBOOK_' + lead.ID + '_' + Utilities.formatDate(at, Session.getScriptTimeZone(), 'yyyyMMddHHmm');
  if (getRawFlag_(flag)) return;
  sendDoctorEmail_(lead, 'booking', at);
  setRawFlag_(flag);
}
