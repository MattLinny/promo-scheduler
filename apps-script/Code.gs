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
 *   6. Optional: run `installTriggers` once for the Monday email and the weekly backup.
 *      The Monday email goes to DIGEST_EMAIL (Script properties), or to NOTIFY_EMAIL if that's not set.
 *
 * After editing this file later: Deploy → Manage deployments → edit (pencil) → Version: New version → Deploy.
 * The URL stays the same.
 *
 * Privacy: nothing is readable without the office code, except the MC link, which needs its own token and only
 * returns next Sunday's lineup. The public can send a request, which lands in the Campaigns tab with status "request".
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
  Owners: ['channel', 'name'],
  Log: ['at', 'by', 'action', 'what'],
};
const SITE = 'https://mattlinny.github.io/promo-scheduler/';
const KEEP_BACKUPS = 8;
const CACHE_KEY = 'all';
const CHURCH = 'Oceanside Church';
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
  if (!props.getProperty('MC_TOKEN')) props.setProperty('MC_TOKEN', Utilities.getUuid().replace(/-/g, '').slice(0, 16));
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
  const mc = e && e.parameter && e.parameter.mc;
  if (mc) {
    const token = PropertiesService.getScriptProperties().getProperty('MC_TOKEN');
    return json_(token && mc === token ? mcData_() : { ok: false, error: 'code' });
  }
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
    owners: rows_('Owners'),
    mcToken: PropertiesService.getScriptProperties().getProperty('MC_TOKEN') || '',
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
      deleteCampaign: () => {
        removeWhere_(CAMPAIGNS, 'id', body.id);
        // A deleted campaign takes its ticks, stories and task links with it.
        removePrefix_(TICKS, body.id + '~');
        removePrefix_(LINES, body.id + '~');
        removePrefix_(LINES, 'link~' + body.id + '~');
      },
      owner: () => (clean_(body.name, 60)
        ? upsert_('Owners', 'channel', { channel: clean_(body.channel, 20), name: clean_(body.name, 60) })
        : removeWhere_('Owners', 'channel', body.channel)),
      replyEmail: () => {
        const to = clean_(body.to, 160);
        if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(to)) throw new Error('bad address');
        MailApp.sendEmail({ to, subject: clean_(body.subject, 120) + ': your promo plan', body: clean_(body.body, 4000), name: CHURCH, replyTo: replyTo_() });
      },
      tick: () => (body.on ? upsert_(TICKS, 'key', { key: body.key_, by: clean_(body.by, 60), at: stamp_() }) : removeWhere_(TICKS, 'key', body.key_)),
      story: () => (clean_(body.story, 2000)
        ? upsert_(LINES, 'key', { key: body.key_, story: clean_(body.story, 2000), by: clean_(body.by, 60), at: stamp_() })
        : removeWhere_(LINES, 'key', body.key_)),
      sunday: () => upsert_(SUNDAYS, 'date', { date: clean_(body.date, 10), stage: clean_(body.stage, 80), booth: clean_(body.booth, 200), by: clean_(body.by, 60), at: stamp_() }),
    }[body.action];
    if (!act) return json_({ ok: false, error: 'unknown' });
    act();
    log_(clean_(body.by || (body.campaign && body.campaign.by), 60), body.action, clean_(body.key_ || body.id || body.date || body.channel || body.to || (body.campaign && body.campaign.name), 120));
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
  log_(rec.requester, 'request', name);
  // Tell the person asking that it arrived.
  if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(rec.contact)) {
    try {
      MailApp.sendEmail({
        to: rec.contact,
        subject: 'We got your promo request: ' + name,
        body: 'Thanks for sending ' + name + '.\n\nWe\'ll look at where it fits and reply with your promo plan.\n\n' + CHURCH,
        name: CHURCH,
        replyTo: replyTo_(),
      });
    } catch (err) { /* mail never blocks a request */ }
  }
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

function replyTo_() {
  const n = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL') || '';
  return n.split(/[,;\s]+/).filter(Boolean)[0] || Session.getEffectiveUser().getEmail();
}

function log_(by, action, what) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Log');
  if (sh) sh.appendRow([stamp_(), safe_(by), action, safe_(what)]);
}

/* ── The plan, for the Monday email and the MC link ──
   The tables below are copied from index.html. tests/parity.test.js fails if they drift apart. */
// <plan>
const W=(phase,angle,say,slide,desk,bull,tasks=[])=>({phase,angle,say,slide,desk,bull,tasks});
const TUE=(email,text,social,opt={})=>[[2,'email',email,opt.email],[2,'text',text,opt.text],[2,'social',social,opt.social]];
const PB={};
PB[1]={1:W('Awareness','Launch','req',1,1,'Featured',[[0,'web','Publish the Launcher action'],[0,'web','Publish the web page and Info Card'],...TUE('Send the launch email','Send the launch text','Post a Story and feed sprinkle'),[2,'fb','Create the Facebook event']]),
  2:W('Awareness','Rest',null,1,1,'Upcoming'),
  3:W('Awareness','The Handoff','req',1,1,'Highlight',TUE('Send a sprinkle email','Send a sprinkle text','Post a poll or something light',{email:1,text:1})),
  4:W('Awareness','The Callback','req',1,1,'Highlight',TUE('Send a sprinkle email','Send a sprinkle text','Post a sprinkle',{email:1,text:1})),
  5:W('Interest','Move to interest','req',1,1,'Highlight'),
  6:W('Interest','Values and voice','req',1,1,'Highlight',TUE('Email: repurpose Sunday','Text: repurpose Sunday','Social: repurpose Sunday')),
  7:W('Interest','Fresh creative','req',1,1,'Highlight',TUE('Email: repurpose Sunday','Text: repurpose Sunday','Social: repurpose Sunday')),
  8:W('Urgency','Sprint','req',1,1,'Featured',[[0,'web','Upgrade the Launcher action to Featured'],[0,'web','Publish the floating banner'],[0,'social','Run the sprint: daily social, email or text, 7 days at most']])};
const MID=l=>{const who=l==2?'churchwide':'ministry';return {
  1:W('Interest','Launch','req',1,1,'Highlight',[[0,'web','Publish the Launcher action'],[0,'web','Publish the page or Next Steps section'],...TUE(`Send the ${who} sprinkle email`,`Send the ${who} sprinkle text`,'Social: repurpose Sunday'),[2,'fb','Create the Facebook event',1]]),
  2:W(l==3?'Awareness':'Interest','Rest',null,1,1,'Highlight'),
  3:W('Interest','Values and voice','opt',1,1,'Highlight',TUE(`Send the ${who} sprinkle email`,`Send the ${who} sprinkle text`,'Social: repurpose Sunday',{email:l==2,text:l==2})),
  4:W('Urgency','Urgency','req',1,1,'Highlight',TUE(`Send the ${who} sprinkle email`,`Send the ${who} sprinkle text`,'Social: repurpose Sunday'))}};
PB[2]=MID(2);PB[3]=MID(3);
PB[4]=PB[5]={1:W('Interest','Launch',null,0,1,'Featured',[[0,'web','Publish the post in the ministry collection'],[0,'web','Add the event to the Calendar page'],[0,'web','Add a Next Steps section',1],...TUE('Send the ministry email','Send the ministry text','Post a sprinkle',{social:1}),[2,'fb','Create the Facebook event',1]]),
  2:W('Interest','Hold',null,0,1,'Featured'),
  3:W('Interest','New creative',null,0,1,'Featured',TUE('Send the ministry email, new creative','Send the ministry text','Post a sprinkle',{social:1})),
  4:W('Urgency','This weekend','smuggle',0,1,'Featured: this weekend',TUE('Send the ministry email: final chance','Send the ministry text','Post a sprinkle',{social:1}))};
const WEEKMAP={8:[1,2,3,4,5,6,7,8],4:[1,2,3,4],3:[1,3,4],2:[1,4]};
const AFTER=l=>[[0,'post','Post live updates and behind-the-scenes'],[1,'post','Publish the recap post and email'],[1,'web',l==1?'Remove the banner, Launcher action and Info Card':'Remove the Launcher action and Next Steps section']];
const HIGHLIGHT=[[2,'social','Post a spotlight sprinkle'],[2,'email','Send a churchwide sprinkle email',1],[2,'text','Send a churchwide sprinkle text',1]];
const PREP=l=>l<=3?['Find the story','Make the graphics','Build the web page']:['Find a story or photo','Write the post'];
// </plan>
const D_ = s => { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); };
const S_ = d => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
const add_ = (s, n) => { const x = D_(s); x.setDate(x.getDate() + n); return S_(x); };
const sunOf_ = s => add_(s, -D_(s).getDay());
const nextSun_ = s => (D_(s).getDay() === 0 ? s : add_(sunOf_(s), 7));
const fmt_ = s => D_(s).toLocaleDateString('en-CA', { weekday: 'short', month: 'short', day: 'numeric' });
const today_ = () => Utilities.formatDate(new Date(), TZ(), 'yyyy-MM-dd');

function norm_(c) {
  return { ...c, level: +c.level || 0, weeks: +c.weeks || 0, ongoing: !!c.ongoing && c.ongoing !== '0' && c.ongoing !== 'false' };
}
const length_ = c => (c.level == 1 ? 8 : (WEEKMAP[c.weeks] ? c.weeks : 4));
function weeksOf_(c) {
  const n = length_(c), map = WEEKMAP[n], last = sunOf_(c.date), pb = PB[c.level] || PB[5];
  return map.map((tw, i) => Object.assign({ i: i + 1, n, tw, sun: add_(last, -(n - 1 - i) * 7) }, pb[tw]));
}
function tasksOf_(c) {
  const out = [], start = weeksOf_(c)[0].sun;
  PREP(c.level).forEach((label, j) => out.push({ key: c.id + '~prep~' + j, date: add_(start, -3), ch: 'prep', label, opt: false, c }));
  weeksOf_(c).forEach(w => {
    if ((w.say === 'req' || w.say === 'opt') && w.sun <= c.date) out.push({ key: c.id + '~w' + w.tw + '~say', date: w.sun, ch: 'say', label: w.angle, opt: w.say === 'opt', c, say: true });
    w.tasks.forEach((t, j) => {
      const date = add_(w.sun, t[0]);
      if (date <= c.date) out.push({ key: c.id + '~w' + w.tw + '~' + t[1] + j, date, ch: t[1], label: t[2], opt: !!t[3], c });
    });
  });
  AFTER(c.level).forEach((t, j) => out.push({ key: c.id + '~after~' + j, date: add_(c.date, t[0]), ch: t[1], label: t[2], opt: false, c }));
  return out;
}
function lineup_(sun, data) {
  const active = data.campaigns.filter(c => c.status === 'active');
  const items = active.filter(c => !c.ongoing && c.date && c.level && c.date >= sun)
    .map(c => ({ c, w: weeksOf_(c).find(w => w.sun === sun) })).filter(x => x.w);
  const req = items.filter(x => x.w.say === 'req').sort((a, b) => a.c.level - b.c.level || a.c.date.localeCompare(b.c.date));
  const list = active.filter(c => c.ongoing && c.level == 2).sort((a, b) => a.name.localeCompare(b.name));
  const every = Math.max(4, Math.min(6, list.length));
  const pick = list.length ? list[Math.round(D_(sun) / 6048e5) % every] : null;
  const hl = pick ? { c: pick, hl: true, w: { angle: 'Highlight week' } } : null;
  const all = req.concat(items.filter(x => x.w.say === 'opt' || x.w.say === 'smuggle'), hl ? [hl] : []);
  const find = id => all.find(x => x.c.id === id) || null;
  const over = data.sundays.find(x => x.date === sun);
  let stage, booth;
  if (over) {
    stage = over.stage ? find(over.stage) : null;
    booth = String(over.booth || '').split(',').filter(Boolean).map(find).filter(Boolean).slice(0, 2);
  } else {
    stage = req.find(x => x.c.level <= 2) || null;
    booth = req.filter(x => x !== stage).slice(0, 2);
    if (hl && (stage ? 1 : 0) + booth.length < 2) { if (!stage) stage = hl; else booth.push(hl); }
  }
  return { stage, booth };
}
function data_() {
  return { campaigns: rows_(CAMPAIGNS).map(norm_), ticks: rows_(TICKS), stories: rows_(LINES), sundays: rows_(SUNDAYS) };
}
function mcData_() {
  const data = data_(), sun = nextSun_(today_()), l = lineup_(sun, data);
  const story = c => { const s = data.stories.find(x => x.key === c.id + '~' + sun); return s ? s.story : ''; };
  const items = [[l.stage, 'Stage']].concat(l.booth.map(x => [x, 'Booth'])).filter(x => x[0])
    .map(([it, where]) => ({ where, name: it.c.name, angle: it.w.angle, story: story(it.c) }));
  return { ok: true, date: sun, items };
}

/* ── Monday email ── */
function digest_(today) {
  const data = data_(), done = {};
  data.ticks.forEach(t => { done[t.key] = true; });
  const end = nextSun_(today), byCampaign = {}, overdue = [], launches = [];
  data.campaigns.filter(c => c.status === 'active' && !c.ongoing && c.date && c.level).forEach(c => {
    const start = weeksOf_(c)[0].sun;
    if (start <= end && add_(start, 6) >= today) launches.push(c.name + ' (' + fmt_(start) + ')');
    tasksOf_(c).filter(t => !t.opt && !t.say && !done[t.key]).forEach(t => {
      if (t.date < today && t.date >= (c.created || '')) overdue.push(t);
      else if (t.date >= today && t.date <= end) (byCampaign[c.name] = byCampaign[c.name] || []).push(t);
    });
  });
  const l = lineup_(end, data);
  const says = [[l.stage, 'Stage']].concat(l.booth.map(x => [x, 'Booth'])).filter(x => x[0]).map(([it, w]) => w + ': ' + it.c.name + ' (' + it.w.angle + ')');
  const count = Object.values(byCampaign).reduce((n, ts) => n + ts.length, 0);
  return { end, count, overdue, launches, says, byCampaign };
}
function sendWeeklyDigest() {
  const props = PropertiesService.getScriptProperties();
  const to = (props.getProperty('DIGEST_EMAIL') || props.getProperty('NOTIFY_EMAIL') || '').split(/[,;\s]+/).filter(Boolean).join(',');
  if (!to) return;
  const d = digest_(today_()), lines = [];
  if (d.overdue.length) { lines.push('OVERDUE (' + d.overdue.length + ')'); d.overdue.forEach(t => lines.push('- ' + t.c.name + ': ' + t.label + ', ' + fmt_(t.date))); lines.push(''); }
  if (d.launches.length) { lines.push('LAUNCHING'); d.launches.forEach(x => lines.push('- ' + x)); lines.push(''); }
  if (d.says.length) { lines.push('SUNDAY, ' + fmt_(d.end)); d.says.forEach(x => lines.push('- ' + x)); lines.push(''); }
  Object.keys(d.byCampaign).sort().forEach(name => { lines.push(name); d.byCampaign[name].sort((a, b) => a.date.localeCompare(b.date)).forEach(t => lines.push('- ' + fmt_(t.date) + ': ' + t.label)); lines.push(''); });
  lines.push(SITE);
  MailApp.sendEmail({ to, subject: 'Promo this week: ' + d.count + ' to do' + (d.overdue.length ? ', ' + d.overdue.length + ' overdue' : ''), body: lines.join('\n'), name: 'Promo Planner' });
}

/* ── Weekly backup: a copy of this spreadsheet beside it, keeping the newest few ── */
function backupSheet() {
  const ss = SpreadsheetApp.getActive(), stamp = today_();
  const copy = ss.copy(ss.getName() + ' backup ' + stamp);
  const parent = DriveApp.getFileById(ss.getId()).getParents();
  if (!parent.hasNext()) return copy.getId();
  const folder = parent.next(), files = folder.getFiles(), old = [];
  while (files.hasNext()) { const f = files.next(); if (f.getName().indexOf(ss.getName() + ' backup ') === 0) old.push(f); }
  old.sort((a, b) => b.getName().localeCompare(a.getName())).slice(KEEP_BACKUPS).forEach(f => f.setTrashed(true));
  return copy.getId();
}
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => { if (['sendWeeklyDigest', 'backupSheet'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sendWeeklyDigest').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  ScriptApp.newTrigger('backupSheet').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(22).create();
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

function removePrefix_(tab, prefix) {
  const sh = SpreadsheetApp.getActive().getSheetByName(tab);
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const vals = sh.getRange(2, 1, n, 1).getValues();
  for (let i = n - 1; i >= 0; i--) if (String(vals[i][0]).indexOf(prefix) === 0) sh.deleteRow(i + 2);
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
