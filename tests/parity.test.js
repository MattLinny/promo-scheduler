// Run: node tests/parity.test.js
// Checks the plan in index.html and in apps-script/Code.gs agree, and runs Code.gs against a stand-in for Sheets.
const fs = require('fs'), vm = require('vm'), path = require('path');
const root = path.join(__dirname, '..');
const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const gs = fs.readFileSync(path.join(root, 'apps-script/Code.gs'), 'utf8');
let failed = 0;
const assert = (c, m) => { console.log((c ? 'ok   ' : 'FAIL ') + m); if (!c) { failed++; process.exitCode = 1; } };

class Sheet {
  constructor(n) { this.n = n; this.d = []; }
  getLastRow() { return this.d.length; } getLastColumn() { return this.d[0] ? this.d[0].length : 0; } getMaxRows() { return 1000; }
  setName(n) { this.n = n; return this; } setFrozenRows() {}
  getRange(r, c, nr = 1, nc = 1) { const s = this; return {
    getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => (s.d[r - 1 + i] || [])[c - 1 + j] ?? '')),
    setValue: v => { (s.d[r - 1] = s.d[r - 1] || [])[c - 1] = v; },
    setValues: vs => vs.forEach((row, i) => row.forEach((v, j) => { (s.d[r - 1 + i] = s.d[r - 1 + i] || [])[c - 1 + j] = v; })),
    setNumberFormat: () => {} }; }
  appendRow(r) { this.d.push(r.slice()); } deleteRow(i) { this.d.splice(i - 1, 1); }
}
const sheets = [new Sheet('Sheet1')], props = {}, cache = {}, mails = [], triggers = [], copies = [];
const ss = { getSheetByName: n => sheets.find(s => s.n === n) || null, getSheets: () => sheets, insertSheet: n => { const s = new Sheet(n); sheets.push(s); return s; },
  getUrl: () => 'url', getName: () => 'Promo', getId: () => 'id', copy: n => { copies.push(n); return { getId: () => 'copy' }; } };
const builder = { timeBased() { return this; }, onWeekDay() { return this; }, atHour() { return this; }, create() { triggers.push(this.h); } };
let uuid = 0;
const ctx = { SpreadsheetApp: { getActive: () => ss }, PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
  CacheService: { getScriptCache: () => ({ get: k => cache[k] || null, put: (k, v) => { cache[k] = v; }, remove: k => { delete cache[k]; } }) },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) }, ContentService: { createTextOutput: t => ({ t, setMimeType() { return this; } }), MimeType: { JSON: 1 } },
  Utilities: { getUuid: () => 'abc123-def456-789' + (uuid++), formatDate: d => d.toISOString().slice(0, 10) }, Session: { getScriptTimeZone: () => 'America/Vancouver', getEffectiveUser: () => ({ getEmail: () => 'owner@x.ca' }) },
  Logger: { log() {} }, MailApp: { sendEmail: m => mails.push(m) }, ScriptApp: { getProjectTriggers: () => [], deleteTrigger() {}, WeekDay: {}, newTrigger: h => Object.assign(builder, { h }) },
  DriveApp: { getFileById: () => ({ getParents: () => ({ hasNext: () => false }) }) }, console };
vm.createContext(ctx);
vm.runInContext(gs + ';Object.assign(this,{setup,doGet,doPost,tasksOf_,lineup_,digest_,norm_,sendWeeklyDigest,installTriggers,backupSheet});', ctx);
const get = k => JSON.parse(ctx.doGet({ parameter: { key: k } }).t), post = b => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify(b) } }).t);

// the plan in the page, without a browser
const i = html.indexOf('const W=(phase'), j = html.indexOf('// Why each week looks');
const clientTables = html.slice(i, j);
const pick = n => { const a = html.indexOf('const ' + n + '='); return html.slice(a, html.indexOf('\n', a)); };
const blockStart = gs.indexOf('// <plan>'), blockEnd = gs.indexOf('// </plan>');
const block = gs.slice(blockStart, blockEnd);
assert(block.includes(clientTables.trim()), 'playbook tables match index.html');
['WEEKMAP', 'AFTER', 'HIGHLIGHT', 'PREP'].forEach(n => assert(block.includes(pick(n)), n + ' matches index.html'));

ctx.setup();
const code = props.OFFICE_CODE;
assert(['Campaigns', 'Ticks', 'Stories', 'Sundays', 'Owners', 'Log'].every(n => ss.getSheetByName(n)), 'tabs made');
assert(!!props.MC_TOKEN, 'MC token made');
assert(get('wrong').error === 'code', 'wrong code refused');
post({ action: 'saveCampaign', key: code, campaign: { id: 'c1', name: 'Christmas', date: '2026-12-24', level: 1, weeks: 8, status: 'active', created: '2026-11-02', by: 'Sam' } });
post({ action: 'saveCampaign', key: code, campaign: { id: 'c2', name: 'Tea', date: '2026-11-28', level: 5, weeks: 4, status: 'active', created: '2026-11-02' } });
assert(get(code).mcToken === props.MC_TOKEN, 'office data includes the MC token');
assert(ss.getSheetByName('Log').d.length === 3, 'writes go in the log');

const c = ctx.norm_(get(code).campaigns[0]);
const keys = ctx.tasksOf_(c).map(t => t.key + '|' + t.date);
assert(keys.includes('c1~prep~0|2026-10-29') && keys.includes('c1~w1~email2|2026-11-03') && keys.includes('c1~after~1|2026-12-25'), 'task keys and dates as expected');

const mcBad = JSON.parse(ctx.doGet({ parameter: { mc: 'nope' } }).t);
assert(mcBad.ok === false, 'bad MC token refused');
const mc = JSON.parse(ctx.doGet({ parameter: { mc: props.MC_TOKEN } }).t);
assert(mc.ok && Array.isArray(mc.items) && !('campaigns' in mc), 'MC link returns only the lineup');

post({ action: 'tick', key: code, key_: 'c1~w1~email2', on: true, by: 'Sam' });
post({ action: 'story', key: code, key_: 'link~c1~w1~email2', story: 'https://x.ca', by: 'Sam' });
post({ action: 'owner', key: code, channel: 'desk', name: 'Pat', by: 'Sam' });
assert(get(code).owners[0].name === 'Pat', 'owner saved');
post({ action: 'owner', key: code, channel: 'desk', name: '', by: 'Sam' });
assert(get(code).owners.length === 0, 'owner cleared');

const d = ctx.digest_('2026-11-02');
assert(d.byCampaign.Christmas && !d.byCampaign.Christmas.some(t => t.key === 'c1~w1~email2'), 'Monday email leaves out finished tasks');
assert(d.byCampaign.Christmas.some(t => t.key === 'c1~w1~text3'), 'Monday email lists open tasks');
assert(d.launches.some(x => x.startsWith('Christmas')), 'Monday email lists launches');
props.DIGEST_EMAIL = 'office@x.ca';
mails.length = 0;
ctx.sendWeeklyDigest();
assert(mails.length === 1 && mails[0].to === 'office@x.ca' && /to do/.test(mails[0].subject), 'Monday email sends');

mails.length = 0;
assert(post({ action: 'request', name: 'Picnic', date: '2027-06-12', requester: 'R', contact: 'r@x.ca', level: 5, story: 'no' }).ok, 'public request');
assert(mails.some(m => m.to === 'r@x.ca' && /We got your promo request/.test(m.subject)), 'requester gets a confirmation');
mails.length = 0;
assert(post({ action: 'replyEmail', key: code, to: 'r@x.ca', subject: 'Picnic', body: 'Hi', by: 'Sam' }).ok && mails[0].to === 'r@x.ca', 'reply emails the requester');
assert(post({ action: 'replyEmail', to: 'r@x.ca', subject: 'x', body: 'x' }).error === 'code', 'reply needs the office code');
assert(post({ action: 'request', name: 'Ongoing', ongoing: '1', requester: 'R', contact: 'not-an-email' }).ok, 'ongoing request with a bad address still saves');

post({ action: 'deleteCampaign', key: code, id: 'c1' });
const a = get(code);
assert(!a.campaigns.find(x => x.id === 'c1') && !a.ticks.length && !a.stories.length, 'delete clears ticks, stories and links');
ctx.installTriggers();
assert(triggers.includes('sendWeeklyDigest') && triggers.includes('backupSheet'), 'triggers installed');
ctx.backupSheet();
assert(copies.length === 1, 'backup copies the sheet');
console.log(failed ? failed + ' failed' : 'all passed');
