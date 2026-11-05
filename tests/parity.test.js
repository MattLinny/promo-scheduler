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
vm.runInContext(gs + ';Object.assign(this,{RANKWEEKS,add_,sunOf_,setup,doGet,doPost,tasksOf_,weeksOf_,lineup_,digest_,norm_,sendWeeklyDigest,installTriggers,backupSheet});', ctx);
const get = k => JSON.parse(ctx.doGet({ parameter: { key: k } }).t), post = b => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify(b) } }).t);

// the plan in the page, without a browser
const i = html.indexOf('const W=(phase'), j = html.indexOf('// Why each week looks');
const clientTables = html.slice(i, j);
const pick = n => { const a = html.indexOf('const ' + n + '='); return html.slice(a, html.indexOf('\n', a)); };
const blockStart = gs.indexOf('// <plan>'), blockEnd = gs.indexOf('// </plan>');
const block = gs.slice(blockStart, blockEnd);
assert(block.includes(clientTables.trim()), 'playbook tables match index.html');
['WEEKMAP', 'RUNMAP', 'RUNNAME', 'AFTER', 'HIGHLIGHT', 'RANKWEEKS', 'ASSETS', 'ASSETFLAGS', 'ASSETLIST', 'PREP'].forEach(n => assert(block.includes(pick(n)), n + ' matches index.html'));

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

// a fundraiser fills its whole run, and is cut short if it's already under way
const run = (start, created) => ctx.norm_({ id: 'f1', name: 'Fundraiser', level: 4, status: 'active', date: '2026-12-06', start, created });
const says = c => ctx.tasksOf_(c).filter(t => /~wr\d+~/.test(t.key)).map(t => t.date);
const runKeys = ctx.tasksOf_(run('2026-11-03', '2026-11-03')).map(t => t.key);
assert(says(run('2026-11-03', '2026-11-03'))[0] === '2026-11-08' && runKeys.includes('f1~wr3~email0') && runKeys.includes('f1~wr5~email0') && new Set(runKeys).size === runKeys.length, 'a five week run has launch, new story and last week, with unique task keys');
assert(ctx.tasksOf_(run('2026-11-03', '2026-11-03')).find(t => t.key === 'f1~wr5~email0').date === '2026-12-01', 'the final push goes out before a run ends on a Sunday');
assert(!runKeys.some(k => /~wr[24]~/.test(k)), 'rest weeks in a run have no tasks');
assert(says(run('2026-10-01', '2026-11-20'))[0] === '2026-11-22', 'a run already under way starts from when it was added');
assert(ctx.tasksOf_(ctx.norm_({ id: 'f2', level: 1, date: '2026-12-24', start: '2026-12-01', created: '2026-11-01' })).some(t => t.key === 'f2~w1~email2'), 'Level 1 keeps its 8 week plan');
post({ action: 'request', name: 'Run request', date: '2027-03-01', start: '2027-01-10', requester: 'R', contact: 'r@x.ca', level: 5 });
assert(get(code).campaigns.find(x => x.name === 'Run request').start === '2027-01-10', 'a request keeps its start date');

// a campaign's own task changes
const edited = ctx.norm_({ id: 'e1', name: 'Conference', level: 1, status: 'active', date: '2027-03-20', created: '2026-11-03', edits: JSON.stringify({ add: [{ id: 'a', label: 'Save the date', date: '2026-11-15', say: 1 }, { id: 'b', label: 'Book the venue', date: '2026-11-18' }], del: ['e1~prep~0', 'e1~w1~say'] }) });
const editedTasks = ctx.tasksOf_(edited);
assert(editedTasks.some(t => t.key === 'e1~x~b' && t.date === '2026-11-18') && !editedTasks.some(t => t.key === 'e1~prep~0' || t.key === 'e1~w1~say'), 'tasks can be added and removed');
const editedData = { campaigns: [edited], stories: [], ticks: [], sundays: [] };
assert(ctx.lineup_('2026-11-15', editedData).stage.w.angle === 'Save the date', 'an early announcement goes on its Sunday');
const firstSun = ctx.weeksOf_(edited)[0].sun;
assert(!ctx.lineup_(firstSun, editedData).stage, 'a removed announcement comes off its Sunday');

// a plan task can be renamed, and an added task can have a type
const renamed = ctx.norm_({ id: 'e2', name: 'R', level: 1, status: 'active', date: '2027-03-20', created: '2026-11-03', edits: JSON.stringify({ add: [{ id: 'f', label: 'Hand out flyers', date: '2026-10-10', ch: 'print' }], del: [], ren: { 'e2~prep~0': 'Find two stories', 'e2~w1~say': 'Big reveal' } }) });
const renTasks = ctx.tasksOf_(renamed);
assert(renTasks.find(t => t.key === 'e2~prep~0').label === 'Find two stories' && renTasks.find(t => t.key === 'e2~x~f').ch === 'print', 'tasks can be renamed and typed');
assert(ctx.lineup_(ctx.weeksOf_(renamed)[0].sun, { campaigns: [renamed], stories: [], ticks: [], sundays: [] }).stage.w.angle === 'Big reveal', 'a renamed announcement shows on Sunday');

// the plan says which graphics to make, due the Thursday before first use
const artOf = c => ctx.tasksOf_(c).filter(t => t.ch === 'art');
const xmas = ctx.norm_({ id: 'x', level: 1, status: 'active', date: '2026-12-24', created: '2026-10-01' }), xArt = artOf(xmas);
const byKey = k => xArt.find(t => t.key === 'x~art~' + k);
assert(['story', 'feed', 'fb', 'email', 'slide', 'main', 'page', 'poll', 'sprint'].every(byKey), 'a Level 1 plan lists its graphics');
assert(!['launcher', 'info', 'banner', 'post', 'calendar', 'bulletin', 'desk'].some(byKey), 'links and cards need no graphic, and one main graphic covers post, calendar and bulletin');
assert(xArt.filter(t => t.key === 'x~art~main').length === 1, 'the main graphic is listed once');
assert(byKey('story').date === '2026-10-29', 'each graphic is due the Thursday before its first use');
assert(!ctx.tasksOf_(xmas).some(t => t.label === 'Make the graphics') && ctx.tasksOf_(xmas).some(t => t.key === 'x~prep~2' && t.label === 'Build the web page'), 'the list replaces Make the graphics and keeps other task keys');
const small5 = artOf(ctx.norm_({ id: 'y', level: 5, weeks: 4, status: 'active', date: '2026-12-24', created: '2026-10-01' }));
assert(!small5.some(t => t.key === 'y~art~slide') && small5.some(t => t.key === 'y~art~main') && small5.find(t => t.key === 'y~art~feed').opt, 'a small event gets only what its plan uses');
const noFb = artOf(ctx.norm_({ id: 'x', level: 1, status: 'active', date: '2026-12-24', created: '2026-10-01', edits: JSON.stringify({ add: [], del: ['x~w1~fb5'] }) }));
assert(!noFb.some(t => t.key === 'x~art~fb'), 'removing a task removes its graphic');

// every launch writes the website post, adds it to Featured Events and the main calendar, and old keys stay put
const xAll = ctx.tasksOf_(xmas), lab = k => (xAll.find(t => t.key === k) || {}).label;
assert(lab('x~prep~3') === 'Write the website post' && lab('x~w1~web6') === 'Add to Featured Events' && lab('x~w1~web7') === 'Add to the main calendar' && lab('x~w1~fb5') === 'Create the Facebook event', 'a Level 1 launch adds the website post, Featured Events and the main calendar');
const l4 = ctx.tasksOf_(ctx.norm_({ id: 'm', level: 4, weeks: 4, status: 'active', date: '2026-12-24', created: '2026-10-01' })), lab4 = k => (l4.find(t => t.key === k) || {}).label;
assert(lab4('m~prep~1') === 'Write the website post' && lab4('m~w1~web1') === 'Add to the main calendar' && lab4('m~w1~web7') === 'Add to Featured Events' && lab4('m~w1~fb6') === 'Create the Facebook event', 'a Level 4 launch does the same');
assert(lab4('m~w1~web8') === 'Publish the Info Card' && lab4('m~after~2') === 'Remove the Info Card and Next Steps section', 'Level 4 puts up an Info Card and takes it down after');
const l5 = ctx.tasksOf_(ctx.norm_({ id: 'n', level: 5, weeks: 4, status: 'active', date: '2026-12-24', created: '2026-10-01' }));
assert(!l5.some(t => /Info Card/.test(t.label)) && l5.find(t => t.key === 'n~after~2').label === 'Remove the Next Steps section', 'Level 5 has no Info Card');

// four launches in a week: one leads, one rides along, the rest wait for room
const rt = (id, level, end, date = '2026-11-03', extra = {}) => ({ key: id + '~s', id, level, end, date, opt: false, fixed: false, kind: 'social', ...extra });
const four = ctx.RANKWEEKS([rt('xmas', 1, '2026-12-24'), rt('baby', 2, '2026-11-08'), rt('movie', 3, '2026-11-30'), rt('soul', 5, '2026-11-28')], {}, ctx.add_, ctx.sunOf_);
assert(four['xmas~s'].rank === 'lead' && four['baby~s'].rank === 'sprinkle' && four['baby~s'].lead === 'xmas', 'highest level leads, next is the sprinkle');
assert(four['movie~s'].rank === 'lead' && four['movie~s'].date === '2026-11-10' && four['movie~s'].moved && four['soul~s'].rank === 'sprinkle' && four['soul~s'].date === '2026-11-10', 'the rest move to the next week with room');
const tooLate = ctx.RANKWEEKS([rt('a', 1, '2026-12-24'), rt('b', 2, '2026-12-24'), rt('c', 3, '2026-11-08')], {}, ctx.add_, ctx.sunOf_);
assert(tooLate['c~s'].hold === 1 && tooLate['c~s'].date === '2026-11-03', 'a post with no room before its event goes on hold');
const swapped = ctx.RANKWEEKS([rt('a', 1, '2026-12-24'), rt('b', 2, '2026-12-24')], { 'social~2026-11-01': 'b,a' }, ctx.add_, ctx.sunOf_);
assert(swapped['b~s'].rank === 'lead' && swapped['a~s'].rank === 'sprinkle', 'a hand change wins');
const small = ctx.RANKWEEKS([rt('m', 4, '2026-12-24')], {}, ctx.add_, ctx.sunOf_);
assert(small['m~s'].rank === 'sprinkle', 'Level 4 and 5 never lead on their own');

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
