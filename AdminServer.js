/* ==========================================================================
   ADMIN CONSOLE — server side
   Served at the web app's URL (see doGet in Code.js). File is AdminServer.js — Apps Script drops extensions, so it can't share a name with Admin.html. Reads and writes
   the same tabs the member app reads, in the same cell formats, so anything
   saved here shows up in the app on the next refresh.

   Roles (from the Users tab)
     role = Mentor, scope = admin  — admin: everything, including creating and
                                     editing pulse checks; sees every member's
                                     pulse answers with their name.
     role = Mentor, scope empty    — mentor: reads and writes everything, but
                                     Pulse Check is read-only and limited to
                                     their own group's members.
   Everyone else (members, or any other scope value) cannot sign in. Run
   setupAdminConsole() once from the editor to add the scope column and make
   the script owner an admin.

   Every call re-derives the caller from their session token, and pulse data
   is scoped here before it leaves the server — a mentor never receives
   another group's members or answers.
   ========================================================================== */

const ADMIN_SESSION_PREFIX = 'admsess_';
const ADMIN_UPLOAD_FOLDER_ID = '10Z7VWwJt6KuCL15CZ68ETsgex4cnls5R';
const ADMIN_ROLES = ['admin', 'leader_group'];
const EVENT_ALL = 'All groups';
const EVENT_MENTORS = 'Mentors only';

function adminDoGet_() {
  return HtmlService.createTemplateFromFile('Admin')
    .evaluate()
    .setTitle('SSG Internal Tools')
    .setFaviconUrl('https://lh3.googleusercontent.com/d/17-3NVNtoneKapt_mJMzMXEK7EVuRC_RE#.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* One-time setup, run from the Apps Script editor: adds a "scope" column to
   Users and sets the script owner's scope to admin. */
function setupAdminConsole() {
  const sheet = sheetWithHeaders_('Users', ['scope']);
  const owner = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  const row = readRows_('Users').find(r => String(r.email || '').trim().toLowerCase() === owner);
  if (!row) return 'Added the scope column. ' + (owner || 'The owner') + ' has no Users row — set scope = admin by hand.';
  writeByHeader_(sheet, row.__row, { scope: 'admin' });
  return 'Added the scope column — ' + owner + ' is an admin.';
}

/* ---- Sheet helpers ---------------------------------------------------------- */

function isTruthyCell_(v) {
  return v === true || String(v).trim().toUpperCase() === 'TRUE';
}

function headerMap_(sheet) {
  const lastCol = sheet.getLastColumn();
  const map = {};
  if (!lastCol) return map;
  sheet.getRange(1, 1, 1, lastCol).getValues()[0].forEach((h, i) => {
    const k = String(h).trim();
    if (k && !(k in map)) map[k] = i + 1;
  });
  return map;
}

/* Returns the tab, creating it or adding any missing header columns. */
function sheetWithHeaders_(name, headers) {
  const ss = ss_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  const map = headerMap_(sheet);
  let lastHeader = Object.keys(map).reduce((m, k) => Math.max(m, map[k]), 0);
  headers.forEach(h => {
    if (!(h in map)) {
      lastHeader += 1;
      sheet.getRange(1, lastHeader).setValue(h);
      clearSheetMemo_();
      map[h] = lastHeader;
    }
  });
  return sheet;
}

function appendByHeader_(sheet, obj) {
  const map = headerMap_(sheet);
  const width = Math.max(sheet.getLastColumn(), 1);
  const row = new Array(width).fill('');
  Object.keys(obj).forEach(k => { if (map[k]) row[map[k] - 1] = obj[k]; });
  sheet.appendRow(row);
  clearSheetMemo_();
  return sheet.getLastRow();
}

/* Text columns (times like "7:00 PM") are forced to plain text first, or
   Sheets would turn them into 1899 time values. */
function writeByHeader_(sheet, rowNum, obj, textKeys) {
  const map = headerMap_(sheet);
  Object.keys(obj).forEach(k => {
    if (!map[k]) return;
    const cell = sheet.getRange(rowNum, map[k]);
    if (textKeys && textKeys.indexOf(k) !== -1) cell.setNumberFormat('@');
    cell.setValue(obj[k]);
  });
  clearSheetMemo_();
}

/* Rows as objects with their sheet row number; fully blank rows are skipped. */
function readRows_(name) {
  const data = sheetValues_(name);
  if (!data || data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (r.every(v => v === '' || v === null)) continue;
    const obj = { __row: i + 1 };
    headers.forEach((h, j) => { if (h && !(h in obj)) obj[h] = r[j]; });
    out.push(obj);
  }
  return out;
}

function tz_() { return Session.getScriptTimeZone(); }

function asDate_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (v === '' || v === null || v === undefined) return null;
  const d = new Date(String(v));
  return isNaN(d.getTime()) ? null : d;
}

function ymd_(d) { return d ? Utilities.formatDate(d, tz_(), 'yyyy-MM-dd') : ''; }
function hm_(d) { return d ? Utilities.formatDate(d, tz_(), 'HH:mm') : ''; }

/* "yyyy-MM-dd" + "HH:mm" read as wall-clock time in the script's timezone. */
function parseLocal_(date, time) {
  if (!date) return null;
  const d = Utilities.parseDate(date + ' ' + (time || '00:00'), tz_(), 'yyyy-MM-dd HH:mm');
  return isNaN(d.getTime()) ? null : d;
}

/* Event time cells: "4:00 PM" text, "19:00", or a Sheets time value. */
function timeCellTo24_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'HH:mm');
  const m = String(v || '').trim().match(/^(\d{1,2})(?::(\d{2}))?\s*([AaPp][Mm])?$/);
  if (!m) return '';
  let h = parseInt(m[1], 10);
  const min = m[2] ? parseInt(m[2], 10) : 0;
  const mer = (m[3] || '').toLowerCase();
  if (mer === 'pm' && h < 12) h += 12;
  if (mer === 'am' && h === 12) h = 0;
  if (h > 23 || min > 59) return '';
  return ('0' + h).slice(-2) + ':' + ('0' + min).slice(-2);
}

function time24ToCell_(t) {
  const m = String(t || '').match(/^(\d{2}):(\d{2})$/);
  if (!m) return '';
  let h = parseInt(m[1], 10);
  const mer = h >= 12 ? 'PM' : 'AM';
  h = h % 12 || 12;
  return h + ':' + m[2] + ' ' + mer;
}

function str_(v) { return v === null || v === undefined ? '' : String(v); }
function id_(v) { return String(v === null || v === undefined ? '' : v).trim(); }

/* ---- Pulse storage -------------------------------------------------------
   Pulse tab:          pulse_id | windowTitle | open date | closes date |
                       questions | open_questions   (questions "|"-separated)
   PulseResponses tab: userId | pulse id | timestamp | answers | note
   answers is "{4;3;"text"}" — scale answers first, then quoted open answers,
   matching the order the member app asks them in.
   -------------------------------------------------------------------------- */

function splitQuestions_(v) {
  return String(v || '').split('|').map(q => q.trim()).filter(q => q !== '');
}

function readPulses_() {
  return readRows_('Pulse').filter(r => id_(r.pulse_id) || str_(r.windowTitle)).map(r => {
    const scale = splitQuestions_(r.questions).map(text => ({ text: text, type: 'scale' }));
    const open = splitQuestions_(r.open_questions).map(text => ({ text: text, type: 'open' }));
    return {
      row: r.__row,
      id: id_(r.pulse_id),
      raw: r,
      windowTitle: str_(r.windowTitle),
      openAt: asDate_(r['open date']),
      closeAt: asDate_(r['closes date']),
      items: scale.concat(open)
    };
  });
}

function looksLikeTimestamp_(v) {
  return v instanceof Date || /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(String(v || ''));
}

/* Older rows were written without a pulse id, so their columns sit one to the
   left of the header (timestamp under "pulse id", answers under "timestamp").
   Values are recognised by shape, and such rows are credited to the pulse
   whose window contains their timestamp (else the first pulse). */
function readPulseResponses_(pulses) {
  const sheet = ss_().getSheetByName('PulseResponses');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const h = data[0].map(x => String(x).trim().toLowerCase());
  const col = names => { for (const n of names) { const i = h.indexOf(n); if (i !== -1) return i; } return -1; };
  const iUser = col(['userid', 'user id']);
  const iPulse = col(['pulse id', 'pulseid', 'pulse_id']);
  const byKey = {};

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const userId = id_(row[iUser !== -1 ? iUser : 0]);
    if (!userId) continue;
    const answers = row.find(v => /^\s*\{/.test(String(v)));
    const stamp = row.find(v => looksLikeTimestamp_(v));
    if (answers === undefined) continue;
    const at = asDate_(stamp);

    let pulseId = iPulse !== -1 ? row[iPulse] : '';
    if (looksLikeTimestamp_(pulseId) || /^\s*\{/.test(String(pulseId))) pulseId = '';
    pulseId = id_(pulseId);
    if (!pulseId && pulses.length) {
      const hit = at && pulses.find(p => (!p.openAt || p.openAt <= at) && (!p.closeAt || at < p.closeAt));
      pulseId = (hit || pulses[0]).id;
    }

    // A member who submitted twice counts once, with their latest answers.
    const key = pulseId + '|' + userId;
    const prev = byKey[key];
    if (!prev || (at && (!prev.at || at > prev.at))) {
      byKey[key] = { pulseId: pulseId, userId: userId, at: at, answersRaw: String(answers) };
    }
  }
  return Object.keys(byKey).map(k => byKey[k]);
}

function parseAnswers_(raw, items) {
  let s = String(raw || '').trim().replace(/^\{/, '').replace(/\}$/, '');
  const tokens = [];
  let i = 0;
  while (i < s.length) {
    if (s[i] === '"') {
      let j = s.indexOf('";', i + 1);
      if (j === -1) j = s.lastIndexOf('"');
      if (j <= i) j = s.length;
      tokens.push(s.slice(i + 1, j));
      i = j + 2;
    } else {
      let j = s.indexOf(';', i);
      if (j === -1) j = s.length;
      tokens.push(s.slice(i, j).trim());
      i = j + 1;
    }
  }
  return items.map((it, k) => {
    const t = tokens[k];
    if (it.type === 'scale') {
      const n = parseInt(t, 10);
      return n >= 1 && n <= 5 ? n : null;
    }
    return t ? String(t).trim() : '';
  });
}

/* ---- Identity & sessions --------------------------------------------------- */

function findGroup_(groups, key) {
  const k = String(key || '').trim().toLowerCase();
  if (!k) return null;
  return groups.find(g => id_(g.groupId).toLowerCase() === k || String(g.name || '').trim().toLowerCase() === k) || null;
}

function adminResolveUser_(email) {
  const clean = String(email || '').trim().toLowerCase();
  if (!clean) return null;
  const u = readSheetAsMap('Users').find(x => String(x.email || '').trim().toLowerCase() === clean);
  if (!u || String(u.status || '').toUpperCase() === 'INACTIVE') return null;
  const groups = readSheetAsMap('Groups');

  // Only mentors get in. scope = admin makes a mentor an admin; an empty
  // scope keeps them a mentor (internally "leader_group"); anything else
  // shuts the console to them.
  if (String(u.role || '').trim().toLowerCase() !== 'mentor') return null;
  const scope = String(u.scope || '').trim().toLowerCase();
  const role = scope === 'admin' ? 'admin' : scope === '' ? 'leader_group' : '';
  const groupKey = u.groupId;
  if (!role) return null;
  const g = findGroup_(groups, groupKey);
  if (role === 'leader_group' && !g) return null;
  const name = str_(u.name);
  return {
    email: clean,
    userId: id_(u.userId),
    name: name,
    initials: getInitials(name),
    role: role,
    groupId: g ? id_(g.groupId) : '',
    groupName: g ? str_(g.name) : ''
  };
}

function adminCtx_(token) {
  let session = null;
  try {
    const raw = token && PropertiesService.getScriptProperties().getProperty(ADMIN_SESSION_PREFIX + String(token));
    session = raw ? JSON.parse(raw) : null;
  } catch (err) { session = null; }
  const ctx = session && adminResolveUser_(session.email);
  if (!ctx) {
    if (token) PropertiesService.getScriptProperties().deleteProperty(ADMIN_SESSION_PREFIX + String(token));
    throw new Error('SESSION_EXPIRED');
  }
  return ctx;
}

function adminLogin(email, password) {
  const cleanEmail = String(email || '').trim().toLowerCase();
  const u = readSheetAsMap('Users').find(x => String(x.email || '').trim().toLowerCase() === cleanEmail);
  if (!u || String(u.password || '').trim() !== String(password || '').trim()) {
    return { success: false, message: 'Invalid email or password.' };
  }
  const ctx = adminResolveUser_(cleanEmail);
  if (!ctx) return { success: false, message: "This account doesn't have admin access." };
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  PropertiesService.getScriptProperties().setProperty(
    ADMIN_SESSION_PREFIX + token, JSON.stringify({ email: cleanEmail, created: Date.now() }));
  return { success: true, token: token, data: adminDataCached_(ctx) };
}

function adminResume(token) {
  return adminCall_(token, null, () => ({}));
}

function adminLogout(token) {
  if (token) PropertiesService.getScriptProperties().deleteProperty(ADMIN_SESSION_PREFIX + String(token));
  return { success: true };
}

/* Wraps every API call: resolves the caller, checks the permission, takes the
   script lock for writes, and reports errors as data rather than throwing
   (google.script.run turns a throw into an opaque message). */
function adminCall_(token, need, fn) {
  let lock = null;
  try {
    const ctx = adminCtx_(token);
    if (need === 'content' && ADMIN_ROLES.indexOf(ctx.role) === -1) throw new Error('Not allowed.');
    if (need === 'pulse' && ctx.role !== 'admin') throw new Error('Only admins can change pulse checks.');
    if (need) { lock = LockService.getScriptLock(); lock.waitLock(20000); }
    const out = fn(ctx) || {};
    if (need) {
      SpreadsheetApp.flush();
      clearSheetMemo_();
      bumpAdminDataGen_();
    }
    out.success = true;
    if (!out.data) out.data = adminDataCached_(ctx);
    return out;
  } catch (err) {
    const msg = String(err && err.message || err);
    return { success: false, expired: msg === 'SESSION_EXPIRED', message: msg === 'SESSION_EXPIRED' ? 'Your session has ended. Sign in again.' : msg };
  } finally {
    if (lock) lock.releaseLock();
  }
}

/* ---- Read model ------------------------------------------------------------
   The payload is cached for ADMIN_CACHE_SECONDS per role + group, so a reload
   skips the sheet reads. Every save bumps a generation number that is part of
   the key, so the next read after a save is always fresh. Edits made directly
   in the spreadsheet show up once the cached copy expires.
   -------------------------------------------------------------------------- */

const ADMIN_CACHE_SECONDS = 60;
const ADMIN_CACHE_CHUNK = 90000;   // CacheService allows 100 KB per value

function adminCacheKey_(ctx) {
  const gen = PropertiesService.getScriptProperties().getProperty('ADMIN_DATA_GEN') || '0';
  return 'admdata_' + gen + '_' + ctx.role + '_' + (ctx.role === 'admin' ? 'all' : ctx.groupId);
}

function bumpAdminDataGen_() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('ADMIN_DATA_GEN', String((parseInt(props.getProperty('ADMIN_DATA_GEN') || '0', 10) || 0) + 1));
}

function adminDataCached_(ctx) {
  const cache = CacheService.getScriptCache(), key = adminCacheKey_(ctx);
  try {
    const n = parseInt(cache.get(key + '_n') || '0', 10);
    if (n > 0) {
      const keys = []; for (let i = 0; i < n; i++) keys.push(key + '_' + i);
      const parts = cache.getAll(keys);
      if (keys.every(k => k in parts)) {
        const data = JSON.parse(keys.map(k => parts[k]).join(''));
        data.serverNow = Date.now();
        return data;
      }
    }
  } catch (err) { /* fall through to a fresh read */ }

  const data = adminData_(ctx);
  try {
    const json = JSON.stringify(data), put = {};
    const n = Math.ceil(json.length / ADMIN_CACHE_CHUNK);
    if (n <= 9) {
      for (let i = 0; i < n; i++) put[key + '_' + i] = json.slice(i * ADMIN_CACHE_CHUNK, (i + 1) * ADMIN_CACHE_CHUNK);
      put[key + '_n'] = String(n);
      cache.putAll(put, ADMIN_CACHE_SECONDS);
    }
  } catch (err) { /* caching is best effort */ }
  return data;
}


function adminData_(ctx) {
  const tz = tz_();
  const groupsRaw = readSheetAsMap('Groups').filter(g => id_(g.groupId));
  const groups = groupsRaw.map(g => ({ id: id_(g.groupId), name: str_(g.name) || id_(g.groupId) }));
  const groupName = gid => { const g = groups.find(x => x.id === gid); return g ? g.name : gid; };

  const modules = readModules_();
  const vs = readSheetAsMap('Verse')[0] || {};

  const ann = readRows_('Announcements').filter(a => str_(a.title)).map(a => {
    const t = asDate_(a.time);
    return {
      row: a.__row, title: str_(a.title), tag: str_(a.tag),
      time: a.time instanceof Date && t ? Utilities.formatDate(t, tz, "EEE, d MMM · h:mm a") : str_(a.time),
      imgUrl: str_(a.imgUrl || a.imageUrl || a.imageLabel), detail: str_(a.detail)
    };
  });

  const events = readRows_('Events').filter(e => str_(e.title)).map(e => {
    const gid = id_(e['Group ID']);
    const group = gid.toUpperCase() === 'ALL' ? (isTruthyCell_(e.is_mentor_only) ? EVENT_MENTORS : EVENT_ALL) : groupName(gid);
    const d = asDate_(e.date);
    return {
      row: e.__row, title: str_(e.title), group: group,
      date: d ? ymd_(d) : '', start: timeCellTo24_(e.time), end: timeCellTo24_(e.end),
      location: str_(e.location), detail: str_(e.description), link: str_(e.link)
    };
  }).sort((a, b) => (a.date + a.start).localeCompare(b.date + b.start));

  const res = readRows_('Resources').filter(r => str_(r.title) || str_(r.fileUrl)).map(r => ({
    row: r.__row, title: str_(r.title), group: str_(r.module), fileUrl: str_(r.fileUrl)
  }));

  // Members, scoped: a group leader only ever receives their own group.
  const scopedToGroup = ctx.role === 'leader_group';
  const allMembers = readSheetAsMap('Users')
    .filter(u => id_(u.userId) && String(u.status || '').toUpperCase() !== 'INACTIVE')
    .map(u => ({ userId: id_(u.userId), name: str_(u.name), groupId: id_(u.groupId), group: groupName(id_(u.groupId)), email: str_(u.email) }));
  const members = scopedToGroup ? allMembers.filter(m => m.groupId === ctx.groupId) : allMembers;
  const memberGroup = {};
  allMembers.forEach(m => { memberGroup[m.userId] = m.groupId; });

  const pulseRows = readPulses_();
  const responses = readPulseResponses_(pulseRows);
  const pulses = pulseRows.map(p => {
    const all = responses.filter(r => r.pulseId === p.id).map(r => ({
      userId: r.userId, groupId: memberGroup[r.userId] || '', at: r.at ? r.at.getTime() : 0,
      answers: parseAnswers_(r.answersRaw, p.items)
    }));
    const mine = scopedToGroup ? all.filter(r => r.groupId === ctx.groupId) : all;
    const out = {
      id: p.id, row: p.row, windowTitle: p.windowTitle, items: p.items,
      openDate: ymd_(p.openAt), openTime: hm_(p.openAt), closeDate: ymd_(p.closeAt), closeTime: hm_(p.closeAt),
      openMs: p.openAt ? p.openAt.getTime() : 0, closeMs: p.closeAt ? p.closeAt.getTime() : 0,
      submitted: mine.map(r => ({ userId: r.userId, at: r.at, answers: r.answers })),
      responseCount: all.length
    };
    return out;
  }).sort((a, b) => b.openMs - a.openMs);

  return {
    me: { email: ctx.email, name: ctx.name, initials: ctx.initials, role: ctx.role, groupId: ctx.groupId, groupName: ctx.groupName },
    tz: tz,
    serverNow: Date.now(),
    groups: groups,
    modules: modules,
    verse: { label: str_(vs.label), text: str_(vs.text), reference: str_(vs.reference) },
    ann: ann, events: events, res: res, pulses: pulses, members: members
  };
}

/* ---- Writes ---------------------------------------------------------------- */

/* Module and verse live on the first data row, which is what the app reads. */
function firstDataRow_(sheet) {
  return 2;
}

/* ---- Modules --------------------------------------------------------------
   ServiceUpdate tab: enabled | theme | module | note | date_start | date_end |
   prompt ("|"-separated reflection prompts). One row is ever enabled — the
   one on members' Home — and only while today is within its dates. Date
   ranges may not overlap. Date cells are only rewritten when they change, so
   the sheet's chained formulas (=F2+1) survive edits to other fields.
   -------------------------------------------------------------------------- */

function readModules_() {
  return readRows_('ServiceUpdate').filter(r => str_(r.theme).trim() || str_(r.module).trim()).map(r => ({
    row: r.__row,
    enabled: isTruthyCell_(r.enabled),
    theme: str_(r.theme).trim(),
    module: str_(r.module).trim(),
    note: str_(r.note).trim(),
    start: ymd_(asDate_(r.date_start)),
    end: ymd_(asDate_(r.date_end)),
    prompts: String(r.prompt || '').split('|').map(x => x.trim()).filter(x => x)
  }));
}

/* Shared with the client's live validation; returns an error message or ''. */
function moduleProblem_(d, others, today) {
  if (!str_(d.theme).trim() || !str_(d.module).trim()) return 'Theme and module are required.';
  if (!d.start || !d.end) return 'Start and end dates are required.';
  if (d.end < d.start) return 'The end date must be on or after the start date.';
  const clash = others.find(o => o.start && o.end && d.start <= o.end && o.start <= d.end);
  if (clash) return 'Dates overlap with ' + (clash.module || clash.theme) + ' (' + clash.start + ' – ' + clash.end + ').';
  if (d.enabled && !(d.start <= today && today <= d.end)) return 'Show on Home is only possible while today is within the module\'s dates.';
  return '';
}

function adminSaveModule(token, row, expectTheme, d) {
  return adminCall_(token, 'content', () => {
    const sheet = sheetWithHeaders_('ServiceUpdate', ['enabled', 'theme', 'module', 'note', 'date_start', 'date_end', 'prompt']);
    const editing = row !== null && row !== undefined && row !== -1;
    const all = readModules_();
    const current = editing ? all.find(m => m.row === Number(row)) : null;
    if (editing && (!current || current.theme !== String(expectTheme))) throw new Error('This module changed in the sheet. Reload and try again.');
    const clean = {
      enabled: !!d.enabled, theme: str_(d.theme).trim(), module: str_(d.module).trim(), note: str_(d.note).trim(),
      start: str_(d.start), end: str_(d.end)
    };
    const problem = moduleProblem_(clean, all.filter(m => !current || m.row !== current.row), ymd_(new Date()));
    if (problem) throw new Error(problem);

    const values = {
      enabled: clean.enabled, theme: clean.theme, module: clean.module, note: clean.note,
      prompt: (d.prompts || []).map(x => str_(x).replace(/\|/g, '/').trim()).filter(x => x).join('|')
    };
    if (!current || current.start !== clean.start) values.date_start = parseLocal_(clean.start, '00:00');
    if (!current || current.end !== clean.end) values.date_end = parseLocal_(clean.end, '00:00');

    const target = current ? current.row : sheet.getLastRow() + 1;
    writeByHeader_(sheet, target, values);
    // Only one module is ever live on Home.
    if (clean.enabled) all.forEach(m => { if (m.row !== target && m.enabled) writeByHeader_(sheet, m.row, { enabled: false }); });
  });
}

function adminDeleteModule(token, row, expectTheme) {
  return adminCall_(token, 'content', () => {
    const sheet = ss_().getSheetByName('ServiceUpdate');
    const current = readModules_().find(m => m.row === Number(row));
    if (!sheet || !current || current.theme !== String(expectTheme)) throw new Error('This module changed in the sheet. Reload and try again.');
    sheet.deleteRow(current.row);
  });
}

function adminSaveVerse(token, v) {
  return adminCall_(token, 'content', () => {
    const sheet = sheetWithHeaders_('Verse', ['label', 'text', 'reference']);
    writeByHeader_(sheet, firstDataRow_(sheet), {
      label: str_(v.label).trim(), text: str_(v.text).trim(), reference: str_(v.reference).trim()
    });
  });
}

const ADMIN_KINDS = {
  ann: { sheet: 'Announcements', headers: ['tag', 'title', 'time', 'imageLabel', 'detail'] },
  ev: { sheet: 'Events', headers: ['Group ID', 'title', 'date', 'time', 'location', 'description', 'is_mentor_only', 'end', 'link'] },
  res: { sheet: 'Resources', headers: ['title', 'fileUrl', 'module', 'dateAdded', 'type_size'] }
};

/* Rows are addressed by sheet row number; the title the client last saw must
   still be there, so an edit made in the sheet meanwhile is never overwritten. */
function checkRow_(sheet, row, expectTitle) {
  const map = headerMap_(sheet);
  const r = Number(row);
  if (!(r >= 2) || r > sheet.getLastRow() || !map.title ||
      String(sheet.getRange(r, map.title).getValue()) !== String(expectTitle)) {
    throw new Error('This item changed in the sheet. Reload and try again.');
  }
  return r;
}

function detectLinkType_(url) {
  const m = String(url || '').match(/^https?:\/\/([^\/?#]+)([^?#]*)/i);
  if (!m) return '';
  const h = m[1].toLowerCase().replace(/^www\./, ''), p = m[2];
  if (/\.pdf$/i.test(p)) return 'PDF';
  if (/youtube\.com|youtu\.be/.test(h)) return 'YouTube';
  if (h === 'docs.google.com' && p.indexOf('/presentation') === 0) return 'Google Slides';
  if (h === 'docs.google.com' && p.indexOf('/document') === 0) return 'Google Docs';
  if (h === 'docs.google.com' && p.indexOf('/spreadsheets') === 0) return 'Google Sheets';
  if (h === 'drive.google.com') return 'Drive';
  if (/canva\.com/.test(h)) return 'Canva';
  return 'Website';
}

function uploadAnnouncementImage_(dataUrl, title) {
  const m = String(dataUrl).match(/^data:(image\/[\w.+-]+);base64,(.*)$/);
  if (!m) throw new Error('Images must be image files.');
  const ext = (m[1].split('/')[1] || 'png').replace('jpeg', 'jpg');
  const name = 'announcement-' + String(title || 'image').replace(/[^\w\- ]+/g, '').slice(0, 40).trim() + '-' + Date.now() + '.' + ext;
  const file = DriveApp.getFolderById(ADMIN_UPLOAD_FOLDER_ID).createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name));
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) { /* domain policy — URL still works for signed-in viewers */ }
  return 'https://lh3.googleusercontent.com/d/' + file.getId();
}

function eventGroupCells_(label) {
  if (label === EVENT_ALL) return { gid: 'ALL', mentorOnly: false };
  if (label === EVENT_MENTORS) return { gid: 'ALL', mentorOnly: true };
  const g = findGroup_(readSheetAsMap('Groups'), label);
  if (!g) throw new Error('Unknown group: ' + label);
  return { gid: id_(g.groupId), mentorOnly: false };
}

function adminSaveItem(token, kind, row, expectTitle, d) {
  return adminCall_(token, 'content', () => {
    const cfg = ADMIN_KINDS[kind];
    if (!cfg) throw new Error('Unknown item type.');
    const sheet = sheetWithHeaders_(cfg.sheet, cfg.headers);
    const editing = row !== null && row !== undefined && row !== -1;
    const title = str_(d.title).trim();
    if (!title) throw new Error('Title is required.');
    let values, textKeys = [];

    if (kind === 'ann') {
      let img = str_(d.imgUrl);
      if (/^data:/.test(img)) img = uploadAnnouncementImage_(img, title);
      values = { title: title, tag: str_(d.tag).trim(), imageLabel: img, detail: str_(d.detail).trim() };
    } else if (kind === 'ev') {
      if (!d.date || !d.group) throw new Error('Event needs a group and a date.');
      const gc = eventGroupCells_(d.group);
      values = {
        'Group ID': gc.gid, is_mentor_only: gc.mentorOnly, title: title,
        date: parseLocal_(d.date, '00:00'), time: time24ToCell_(d.start), end: time24ToCell_(d.end),
        location: str_(d.location).trim(), description: str_(d.detail).trim(), link: str_(d.link).trim()
      };
      textKeys = ['time', 'end'];
    } else {
      if (!str_(d.fileUrl).trim() || !str_(d.group).trim()) throw new Error('Resource needs a module group and a link.');
      values = { title: title, fileUrl: str_(d.fileUrl).trim(), module: str_(d.group).trim(), type_size: detectLinkType_(d.fileUrl) };
    }

    if (editing) {
      writeByHeader_(sheet, checkRow_(sheet, row, expectTitle), values, textKeys);
    } else if (kind === 'ann') {
      // Newest announcement first, as the app lists them in sheet order.
      values.time = new Date();
      sheet.insertRowBefore(2);
      clearSheetMemo_();
      writeByHeader_(sheet, 2, values);
    } else {
      if (kind === 'res') values.dateAdded = new Date();
      writeByHeader_(sheet, sheet.getLastRow() + 1, values, textKeys);
    }
  });
}

function adminDeleteItem(token, kind, row, expectTitle) {
  return adminCall_(token, 'content', () => {
    const cfg = ADMIN_KINDS[kind];
    if (!cfg) throw new Error('Unknown item type.');
    const sheet = ss_().getSheetByName(cfg.sheet);
    if (!sheet) throw new Error('Sheet not found.');
    sheet.deleteRow(checkRow_(sheet, row, expectTitle));
  });
}

function pulseRowById_(sheet, id) {
  const p = readPulses_().find(x => x.id === String(id));
  if (!p) throw new Error('Pulse check not found. Reload and try again.');
  return p.row;
}

function adminSavePulse(token, id, d) {
  let newId = null;
  return adminCall_(token, 'pulse', () => {
    const title = str_(d.windowTitle).trim();
    const openAt = parseLocal_(d.openDate, d.openTime || '00:00');
    const closeAt = parseLocal_(d.closeDate, d.closeTime || '23:59');
    if (!title || !openAt || !closeAt) throw new Error('Title, open and close are required.');
    if (closeAt <= openAt) throw new Error('Close time must be after the open time.');
    const sheet = sheetWithHeaders_('Pulse', ['pulse_id', 'windowTitle', 'open date', 'closes date', 'questions', 'open_questions']);
    const values = { windowTitle: title, 'open date': openAt, 'closes date': closeAt };
    if (id) {
      writeByHeader_(sheet, pulseRowById_(sheet, id), values);
    } else {
      const maxId = readPulses_().reduce((m, p) => Math.max(m, parseInt(p.id, 10) || 0), 0);
      newId = String(maxId + 1);
      values.pulse_id = maxId + 1;
      values.questions = '';
      values.open_questions = '';
      appendByHeader_(sheet, values);
    }
    return { pulseId: newId || String(id) };
  });
}

function adminSaveQuestions(token, id, items) {
  return adminCall_(token, 'pulse', () => {
    const sheet = sheetWithHeaders_('Pulse', ['pulse_id', 'windowTitle', 'open date', 'closes date', 'questions', 'open_questions']);
    const clean = (items || []).map(it => ({ text: str_(it.text).replace(/\|/g, '/').trim(), type: it.type === 'open' ? 'open' : 'scale' }))
      .filter(it => it.text);
    writeByHeader_(sheet, pulseRowById_(sheet, id), {
      questions: clean.filter(it => it.type === 'scale').map(it => it.text).join('|'),
      open_questions: clean.filter(it => it.type === 'open').map(it => it.text).join('|')
    });
  });
}

/* Removes the pulse's row only; its responses stay in PulseResponses. */
function adminDeletePulse(token, id) {
  return adminCall_(token, 'pulse', () => {
    const sheet = ss_().getSheetByName('Pulse');
    if (!sheet) throw new Error('Sheet not found.');
    sheet.deleteRow(pulseRowById_(sheet, id));
  });
}
