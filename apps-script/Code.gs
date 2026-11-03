/**
 * Oceanside Promo Planner: Google Sheet backend.
 *
 * Setup (once):
 *   1. Make a new Google Sheet called "Promo Planner". Extensions → Apps Script. Paste this whole file over Code.gs and save.
 *   2. Pick `setup` in the function menu and click Run. Approve the permissions it asks for.
 *      It makes the tabs and writes an office code to the Execution log. To change the code later:
 *      Project Settings → Script properties → OFFICE_CODE.
 *   3. Optional: Script properties → add NOTIFY_EMAIL = the address(es) that should hear about new promo requests.
 *   4. Deploy → New deployment → type "Web app". Execute as: Me. Who has access: Anyone. Deploy.
 *   5. Copy the Web app URL into CONFIG.apiUrl in index.html.
 *
 * After editing this file later: Deploy → Manage deployments → edit (pencil) → Version: New version → Deploy.
 * The URL stays the same.
 *
 * Privacy: nothing is readable without the office code. The only thing the public can do is send a request,
 * which lands in the Campaigns tab with status "request".
 */

const CAMPAIGNS = 'Campaigns';
const TICKS = 'Ticks';
const LINES = 'Stories';
const SUNDAYS = 'Sundays';
const TABS = {
  Campaigns: ['id', 'name', 'date', 'level', 'weeks', 'ministry', 'notes', 'status', 'ongoing', 'requester', 'contact', 'created', 'by', 'story'],
  Ticks: ['key', 'by', 'at'],
  Stories: ['key', 'story', 'by', 'at'],
  Sundays: ['date', 'stage', 'booth', 'by', 'at'],
};
const CACHE_KEY = 'all';
const TZ = () => Session.getScriptTimeZone();

function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(TABS).forEach((name, i) => {
    const sh = ss.getSheetByName(name) || (i === 0 && ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() === 0
      ? ss.getSheets()[0].setName(name) : ss.insertSheet(name));
    ensureHeaders_(sh, TABS[name]);
    // Plain text, so Sheets doesn't turn dates and ids into something else.
    sh.getRange(1, 1, sh.getMaxRows(), TABS[name].length).setNumberFormat('@');
    sh.setFrozenRows(1);
  });
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('OFFICE_CODE')) props.setProperty('OFFICE_CODE', Utilities.getUuid().slice(0, 6));
  Logger.log('Office code: ' + props.getProperty('OFFICE_CODE'));
}

function ensureHeaders_(sh, cols) {
  const last = sh.getLastColumn();
  const have = last ? sh.getRange(1, 1, 1, last).getValues()[0].map(String) : [];
  cols.filter(c => have.indexOf(c) < 0).forEach(c => {
    have.push(c);
    sh.getRange(1, have.length).setValue(c);
  });
}

/* ── Read: everything, for people with the office code ── */
function doGet(e) {
  if (!authed_(e && e.parameter && e.parameter.key)) return json_({ ok: false, error: 'code' });
  const cache = CacheService.getScriptCache();
  const hit = cache.get(CACHE_KEY);
  if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  const out = JSON.stringify({
    ok: true,
    campaigns: rows_(CAMPAIGNS),
    ticks: rows_(TICKS),
    stories: rows_(LINES),
    sundays: rows_(SUNDAYS),
  });
  if (out.length < 90000) cache.put(CACHE_KEY, out, 20);
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}

/* ── Write ── */
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad' }); }
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (err) { return json_({ ok: false, error: 'busy' }); }
  try {
    if (body.action === 'request') return json_(request_(body));
    if (!authed_(body.key)) return json_({ ok: false, error: 'code' });
    const act = {
      saveCampaign: () => upsert_(CAMPAIGNS, 'id', campaign_(body.campaign)),
      deleteCampaign: () => removeWhere_(CAMPAIGNS, 'id', body.id),
      tick: () => (body.on ? upsert_(TICKS, 'key', { key: body.key_, by: clean_(body.by, 60), at: stamp_() }) : removeWhere_(TICKS, 'key', body.key_)),
      story: () => (clean_(body.story, 2000)
        ? upsert_(LINES, 'key', { key: body.key_, story: clean_(body.story, 2000), by: clean_(body.by, 60), at: stamp_() })
        : removeWhere_(LINES, 'key', body.key_)),
      sunday: () => upsert_(SUNDAYS, 'date', { date: clean_(body.date, 10), stage: clean_(body.stage, 80), booth: clean_(body.booth, 200), by: clean_(body.by, 60), at: stamp_() }),
    }[body.action];
    if (!act) return json_({ ok: false, error: 'unknown' });
    act();
    CacheService.getScriptCache().remove(CACHE_KEY);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function onEdit() {
  CacheService.getScriptCache().remove(CACHE_KEY);
}

/* ── A ministry leader's request. Public, so it can only add a row with status "request". ── */
function request_(b) {
  if (b.website) return { ok: true }; // honeypot
  const name = clean_(b.name, 120), date = clean_(b.date, 10), ongoing = b.ongoing ? '1' : '';
  if (!name || (!ongoing && !/^\d{4}-\d{2}-\d{2}$/.test(date))) return { ok: false, error: 'missing' };
  const rec = {
    id: 'r' + Utilities.getUuid().slice(0, 8), name, date, level: /^[1-5]$/.test(String(b.level)) ? String(b.level) : '', weeks: /^[234]$|^8$/.test(String(b.weeks)) ? String(b.weeks) : '',
    ministry: clean_(b.ministry, 120), notes: clean_(b.notes, 2000), status: 'request', ongoing, story: clean_(b.story, 10),
    requester: clean_(b.requester, 120), contact: clean_(b.contact, 160), created: stamp_().slice(0, 10), by: '',
  };
  upsert_(CAMPAIGNS, 'id', rec);
  CacheService.getScriptCache().remove(CACHE_KEY);
  const to = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL');
  if (to) {
    try {
      MailApp.sendEmail({
        to: to.split(/[,;\s]+/).filter(Boolean).join(','),
        subject: 'Promo request: ' + name,
        body: [name, date, rec.ministry, rec.requester, rec.contact, '', rec.notes, '', SpreadsheetApp.getActive().getUrl()].join('\n'),
        name: 'Promo Planner',
      });
    } catch (err) { /* mail never blocks a request */ }
  }
  return { ok: true };
}

function campaign_(c) {
  const rec = {};
  TABS.Campaigns.forEach(k => { rec[k] = clean_(c && c[k], k === 'notes' ? 2000 : 160); });
  if (!rec.id) throw new Error('no id');
  return rec;
}

/* ── Sheet helpers ── */
function authed_(key) {
  const code = PropertiesService.getScriptProperties().getProperty('OFFICE_CODE');
  return !!code && String(key || '').trim().toLowerCase() === code.trim().toLowerCase();
}

function rows_(tab) {
  const sh = SpreadsheetApp.getActive().getSheetByName(tab);
  if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const head = data.shift().map(String);
  return data
    .filter(r => String(r[0]).trim())
    .map(r => {
      const o = {};
      head.forEach((h, i) => { o[h] = r[i] instanceof Date ? Utilities.formatDate(r[i], TZ(), 'yyyy-MM-dd') : String(r[i]); });
      return o;
    });
}

function upsert_(tab, idCol, rec) {
  const sh = SpreadsheetApp.getActive().getSheetByName(tab);
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const row = head.map(h => safe_(rec[h]));
  const at = findRow_(sh, head.indexOf(idCol), rec[idCol]);
  if (at) sh.getRange(at, 1, 1, head.length).setValues([row]);
  else sh.appendRow(row);
}

function removeWhere_(tab, idCol, id) {
  const sh = SpreadsheetApp.getActive().getSheetByName(tab);
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const at = findRow_(sh, head.indexOf(idCol), id);
  if (at) sh.deleteRow(at);
}

function findRow_(sh, col, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1 || col < 0 || !id) return 0;
  const vals = sh.getRange(2, col + 1, n, 1).getValues();
  for (let i = 0; i < n; i++) {
    const v = vals[i][0] instanceof Date ? Utilities.formatDate(vals[i][0], TZ(), 'yyyy-MM-dd') : String(vals[i][0]);
    if (v === String(id)) return i + 2;
  }
  return 0;
}

function clean_(v, max) {
  return String(v === undefined || v === null ? '' : v).trim().slice(0, max);
}

// Keeps a typed "=" or "+" from running as a formula.
function safe_(v) {
  const s = v === undefined || v === null ? '' : String(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function stamp_() {
  return Utilities.formatDate(new Date(), TZ(), "yyyy-MM-dd'T'HH:mm");
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
