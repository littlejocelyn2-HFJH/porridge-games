/**
 * 閱讀小誌（診斷後鷹架練習）— 通用後端 Google Apps Script
 * 需要兩個檔案：Code.gs（本檔）＋ Answers.gs（老師另外拿到，不在 GitHub 上）
 * 部署：部署 → 新增部署作業 → 網頁應用程式｜執行身分：我｜存取權：所有人
 * 這是通用版：以後改題目、改答案，只要在老師頁「教師區 → 更新答案」貼上代碼，不用再改這個檔案。
 */
// ★ 請把下一行的 change-me 改成老師自己的密碼（10 個字以上、英文夾數字），保留引號
const TEACHER_PASSWORD = 'change-me';

const SHEET_DATA = '資料（程式用）';
const SHEET_LOG = '練習紀錄（老師看）';
const MAX_SESSIONS = 20;
const DATA_HEADERS = ['key', '班級', '座號', '姓名', '前測層級', '練習次數', '最後練習', 'sessionsJSON', '更新時間'];
const LOG_HEADERS = ['開始時間', '班級', '座號', '姓名', '前測層級', '模式', '完成話數',
  '第一話 一次就對', '第一話 嘗試', '第二話 一次就對', '第二話 嘗試', '第三話 一次就對', '第三話 嘗試', '第四話 一次就對', '第四話 嘗試',
  '總用時(秒)', '第一話 文字', '第二話 文字', '第三話 文字', '第四話 文字', 'sessionId'];
const LOG_ID_COL = 21;

function doGet() {
  const A = answers_();
  return out({ ok: true, msg: '閱讀小誌後端運作中。這是後端網址，請貼到 config.js。', answers: !!(A && A.sets), version: A && A.version || '' });
}

function doPost(e) {
  let req = {};
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return out({ ok: false, error: '資料格式錯誤' }); }
  try {
    const A = answers_();
    if (!A || !A.sets) return out({ ok: false, error: '老師還沒設定答案（Answers.gs 或老師頁的「更新答案」）' });
    switch (req.action) {
      case 'ping': return out({ ok: true });
      case 'start': return out(startAction_(req, A));
      case 'check': return out(checkAction_(req, A));
      case 'list': return out(listAction_(req, A));
      case 'save': return out(withLock_(function () { return saveAction_(req, A); }));
      case 'clearClass': return out(withLock_(function () { return clearAction_(req); }));
      case 'deleteSession': return out(withLock_(function () { return deleteSessionAction_(req); }));
      case 'setPretest': return out(setPretestAction_(req));
      case 'setAnswers': return out(setAnswersAction_(req));
      default: return out({ ok: false, error: '未知的動作' });
    }
  } catch (err) { return out({ ok: false, error: String(err && err.message || err) }); }
}

/* ---------- 工具 ---------- */
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function withLock_(fn) { const l = LockService.getScriptLock(); l.waitLock(25000); try { return fn(); } finally { l.releaseLock(); } }
function str_(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u0008\u000b-\u001f]/g, '').trim().slice(0, max || 200); }
function safe_(v) { const s = String(v == null ? '' : v); return /^[=+\-@]/.test(s) ? "'" + s : s; } // 防試算表公式注入
function normSeat_(s) { const n = parseInt(String(s).replace(/[^\d]/g, ''), 10); return isNaN(n) ? '' : String(n); }
function keyOf_(cls, seat) { return str_(cls, 10) + '|' + normSeat_(seat); }
function parseJSON_(s, d) { try { return s ? JSON.parse(s) : d; } catch (e) { return d; } }
function twTime_(iso) { try { return Utilities.formatDate(new Date(iso), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss'); } catch (e) { return iso; } }
function props_() { return PropertiesService.getScriptProperties(); }

/* 答案：優先用老師頁更新過的版本（存在指令碼屬性），否則用 Answers.gs */
function answers_() {
  const p = props_(), n = Number(p.getProperty('ANS_CHUNKS') || 0);
  if (n > 0) {
    let s = ''; for (let i = 0; i < n; i++) s += p.getProperty('ANS_' + i) || '';
    const a = parseJSON_(s, null); if (a && a.sets) return a;
  }
  return (typeof ANSWERS !== 'undefined') ? ANSWERS : null;
}

function validStudent_(st) {
  st = st || {};
  const cls = str_(st.cls, 10), seat = normSeat_(st.seat), name = str_(st.name, 20);
  if (!/^\d{3}$/.test(cls)) return { error: '班級請填三位數字，例如 902' };
  const allow = (typeof ALLOWED_CLASSES !== 'undefined') ? ALLOWED_CLASSES : [];
  if (allow.length && allow.indexOf(cls) < 0) return { error: '這個班級沒有開放，請確認班級' };
  const n = Number(seat);
  if (!(n >= 1 && n <= 60)) return { error: '座號請填 1～60' };
  if (!name) return { error: '請填姓名' };
  return { cls: cls, seat: seat, name: name, key: cls + '|' + seat };
}

/* 密碼：每次錯誤延遲 2 秒；10 分鐘內錯 10 次就暫停登入 10 分鐘 */
function pwCheck_(p) {
  if (!p) return { ok: false };
  const cache = CacheService.getScriptCache();
  if (cache.get('pwlock')) return { ok: false, locked: true };
  if (TEACHER_PASSWORD !== 'change-me' && p === TEACHER_PASSWORD) return { ok: true };
  Utilities.sleep(2000);
  const n = Number(cache.get('pwfail') || 0) + 1;
  cache.put('pwfail', String(n), 600);
  if (n >= 10) { cache.put('pwlock', '1', 600); cache.remove('pwfail'); }
  return { ok: false };
}

function pretestMap_() {
  const base = (typeof PRETEST !== 'undefined') ? PRETEST : {};
  const extra = parseJSON_(props_().getProperty('PRETEST_JSON'), {});
  const m = {}; Object.keys(base).forEach(function (k) { m[k] = base[k]; }); Object.keys(extra).forEach(function (k) { m[k] = extra[k]; });
  return m;
}
function modeOf_(lv) { return lv == null ? 'standard' : lv <= 1 ? 'support' : lv >= 4 ? 'challenge' : 'standard'; }

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#efe4cf');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), headers.length + 4).setNumberFormat('@');
  } else {
    const cur = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    if (cur.join('|') !== headers.join('|')) sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#efe4cf');
  }
  return sh;
}
function readData_() {
  const sh = sheet_(SHEET_DATA, DATA_HEADERS);
  const last = sh.getLastRow();
  if (last < 2) return { sh: sh, rows: [] };
  return { sh: sh, rows: sh.getRange(2, 1, last - 1, DATA_HEADERS.length).getValues().map(function (r, i) {
    return { row: i + 2, key: String(r[0]), cls: String(r[1]), seat: String(r[2]), name: String(r[3]), sessions: parseJSON_(r[7], []) };
  }) };
}

/* ---------- start：回傳學生的鷹架模式 ---------- */
function startAction_(req, A) {
  const st = validStudent_(req.student);
  if (st.error) return { ok: false, error: st.error };
  const lv = pretestMap_()[st.key];
  const mode = modeOf_(lv);
  const res = { ok: true, mode: mode, pretest: lv == null ? null : lv };
  if (mode === 'support' && A.supportHints) { // 支持模式：一開始就給段落提示
    res.paraHints = {};
    Object.keys(A.supportHints).forEach(function (lvNo) { const s = A.sets[A.supportHints[lvNo]]; if (s && s.para) res.paraHints[lvNo] = s.para; });
  }
  return res;
}

/* ---------- check：通用批改（sort / order / fill / ref），不寫入試算表 ---------- */
function checkAction_(req, A) {
  const set = A.sets[str_(req.set, 30)];
  if (!set) return { ok: false, error: '沒有這一題' };
  const tryNo = Math.max(1, Math.min(9, Number(req.tryNo) || 1));
  const res = { ok: true, results: {}, hints: {} };
  if (set.type === 'fill') {
    const f = str_(req.fill, 30).replace(/[、，,\s]/g, '');
    res.fillOk = (set.accept || []).indexOf(f) >= 0 || (!!set.all && set.all.every(function (c) { return f.indexOf(c) >= 0; }) && f.length <= (set.maxLen || 99));
    if (!res.fillOk && tryNo >= 2) res.fillReveal = (set.accept || [''])[0];
    return res;
  }
  if (set.type === 'ref') { // 寫完每一格才給參考筆記
    const w = req.written || {};
    if ((set.need || []).some(function (k) { return !str_(w[k], 300); })) return { ok: false, error: '每一格都寫一點，再看參考筆記' };
    return { ok: true, ref: set.ref || null };
  }
  if (set.type === 'order') {
    const order = (req.order || []).slice(0, 30).map(String);
    order.forEach(function (id, i) { res.results[id] = set.order[i] === id; if (!res.results[id]) res.hints[id] = (set.para || {})[id] || ''; });
    res.allRight = order.length === set.order.length && order.every(function (id, i) { return set.order[i] === id; });
    if (!res.allRight && tryNo >= 3) res.reveal = set.order;
  } else { // sort
    const p = req.placements || {};
    Object.keys(set.key).forEach(function (id) { const ok = p[id] === set.key[id]; res.results[id] = ok; if (!ok) res.hints[id] = (set.para || {})[id] || ''; });
    res.allRight = Object.keys(res.results).every(function (id) { return res.results[id]; });
    if (!res.allRight && tryNo >= 3) { res.reveal = {}; Object.keys(set.key).forEach(function (id) { if (!res.results[id]) res.reveal[id] = set.key[id]; }); }
  }
  if ((res.allRight || res.reveal) && set.ref) res.ref = set.ref;
  return res;
}

/* 伺服器重算一話的成績（不信任前端分數） */
function grade_(A, lv, tries) {
  const set = A.sets[(A.grade || {})[lv]]; if (!set) return null;
  function score(t) {
    let c = 0;
    if (set.type === 'order') { const o = Array.isArray(t) ? t : []; set.order.forEach(function (id, i) { if (o[i] === id) c++; }); }
    else Object.keys(set.key || {}).forEach(function (id) { if (t && t[id] === set.key[id]) c++; });
    return c;
  }
  const total = set.type === 'order' ? set.order.length : Object.keys(set.key || {}).length;
  const first = tries.length ? score(tries[0]) : 0;
  let solvedAt = 0; for (let i = 0; i < tries.length; i++) { if (score(tries[i]) === total) { solvedAt = i + 1; break; } }
  return { first: first, total: total, tries: tries.length, solvedAt: solvedAt };
}

/* ---------- save：同一次練習會一話一話更新 ---------- */
function saveAction_(req, A) {
  const st = validStudent_(req.student);
  if (st.error) return { ok: false, error: st.error };
  const se = req.session || {};
  const sid = str_(se.id, 60);
  if (!sid) return { ok: false, error: '缺少練習編號' };
  const d = readData_();
  let rec = null; d.rows.forEach(function (r) { if (r.key === st.key) rec = r; });
  let sessions = rec ? rec.sessions : [];
  const exist = sessions.filter(function (x) { return x.id === sid; })[0];
  if (!exist) { // 新的一次練習：同一座號 60 秒內只能開一次
    const cache = CacheService.getScriptCache(), ck = 'new_' + st.key;
    if (cache.get(ck)) return { ok: false, error: '太快了，請稍等一分鐘再開始新的練習' };
    cache.put(ck, '1', 60);
  }
  const lvPre = pretestMap_()[st.key];
  const clean = { id: sid, start: str_(se.start, 40), t: new Date().toISOString(), mode: modeOf_(lvPre), pretest: lvPre == null ? null : lvPre, levels: {} };
  Object.keys(se.levels || {}).slice(0, 10).forEach(function (lv) {
    const r = se.levels[lv]; if (!r) return;
    const tries = (Array.isArray(r.tries) ? r.tries : []).slice(0, 9).map(function (t) {
      if (Array.isArray(t)) return t.slice(0, 30).map(function (x) { return str_(x, 12); });
      const o = {}; Object.keys(t || {}).slice(0, 40).forEach(function (k) { o[str_(k, 12)] = str_(t[k], 16); }); return o;
    });
    const texts = {};
    Object.keys(r.texts || {}).slice(0, 30).forEach(function (k) { const v = str_(r.texts[k], 400); if (v) texts[str_(k, 40)] = v; });
    clean.levels[lv] = { tries: tries, dur: Math.max(0, Math.min(36000, Math.round(Number(r.dur) || 0))), done: !!r.done,
      hints: Math.min(99, Number(r.hints) || 0), texts: texts, g: grade_(A, lv, tries) };
  });
  if (exist) sessions = sessions.map(function (x) { return x.id === sid ? clean : x; });
  else sessions.push(clean);
  if (sessions.length > MAX_SESSIONS) sessions = [sessions[0]].concat(sessions.slice(sessions.length - MAX_SESSIONS + 1));
  const row = [st.key, st.cls, st.seat, safe_(st.name), lvPre == null ? '' : String(lvPre), String(sessions.length), clean.t, JSON.stringify(sessions), new Date().toISOString()];
  if (rec) d.sh.getRange(rec.row, 1, 1, row.length).setValues([row]); else d.sh.appendRow(row);
  logUpsert_(st, clean);
  const graded = {}; Object.keys(clean.levels).forEach(function (lv) { graded[lv] = clean.levels[lv].g; });
  return { ok: true, graded: graded, mode: clean.mode, pretest: clean.pretest };
}
function textsOf_(L, lv) { const t = L[lv] && L[lv].texts; if (!t) return ''; return Object.keys(t).map(function (k) { return k + '：' + t[k]; }).join('\n'); }
function logUpsert_(st, s) {
  const sh = sheet_(SHEET_LOG, LOG_HEADERS);
  const L = s.levels, cell = function (lv, k) { return L[lv] && L[lv].g ? String(k === 'first' ? L[lv].g.first + '/' + L[lv].g.total : L[lv].g.tries) : ''; };
  const done = Object.keys(L).filter(function (lv) { return L[lv].done; }).length;
  const dur = Object.keys(L).reduce(function (a, lv) { return a + (L[lv].dur || 0); }, 0);
  const modeName = { support: '支持', standard: '一般', challenge: '挑戰' }[s.mode];
  const row = [twTime_(s.start || s.t), st.cls, st.seat, safe_(st.name), s.pretest == null ? '' : String(s.pretest), modeName, String(done),
    cell(1, 'first'), cell(1, 'tries'), cell(2, 'first'), cell(2, 'tries'), cell(3, 'first'), cell(3, 'tries'), cell(4, 'first'), cell(4, 'tries'),
    String(dur), safe_(textsOf_(L, 1)), safe_(textsOf_(L, 2)), safe_(textsOf_(L, 3)), safe_(textsOf_(L, 4)), s.id];
  const last = sh.getLastRow();
  if (last >= 2) {
    const ids = sh.getRange(2, LOG_ID_COL, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === s.id) { sh.getRange(i + 2, 1, 1, row.length).setValues([row]); return; }
  }
  sh.appendRow(row);
}

/* ---------- list ---------- */
function listAction_(req, A) {
  const d = readData_();
  if (req.password) {
    const pw = pwCheck_(req.password);
    if (pw.ok) return { ok: true, authed: true, rows: d.rows.map(function (r) { return { cls: r.cls, seat: r.seat, name: r.name, sessions: r.sessions }; }), answers: A, pretest: pretestMap_() };
    return { ok: true, authed: false, pwError: true, locked: !!pw.locked, rows: publicRows_(d.rows) };
  }
  return { ok: true, authed: false, rows: publicRows_(d.rows) };
}
function publicRows_(rows) {
  return rows.map(function (r) {
    const best = r.sessions.reduce(function (m, s) { return Math.max(m, Object.keys(s.levels || {}).filter(function (lv) { return s.levels[lv].done; }).length); }, 0);
    return { cls: r.cls, seat: r.seat, name: r.name, chapters: best };
  });
}

/* ---------- 老師：清除班級、刪除單次練習、匯入前測層級、更新答案 ---------- */
function clearAction_(req) {
  if (!pwCheck_(req.password).ok) return { ok: false, error: '密碼錯誤或暫停登入中' };
  const cls = str_(req.cls, 10); if (!cls) return { ok: false, error: '請指定班級' };
  [sheet_(SHEET_DATA, DATA_HEADERS), sheet_(SHEET_LOG, LOG_HEADERS)].forEach(function (sh) {
    const last = sh.getLastRow(); if (last < 2) return;
    const v = sh.getRange(2, 2, last - 1, 1).getValues();
    for (let i = v.length - 1; i >= 0; i--) if (String(v[i][0]) === cls) sh.deleteRow(i + 2);
  });
  return { ok: true };
}
function deleteSessionAction_(req) {
  if (!pwCheck_(req.password).ok) return { ok: false, error: '密碼錯誤或暫停登入中' };
  const key = keyOf_(req.cls, req.seat), sid = str_(req.sessionId, 60);
  const d = readData_(); let rec = null; d.rows.forEach(function (r) { if (r.key === key) rec = r; });
  if (!rec) return { ok: false, error: '找不到這位學生' };
  const left = rec.sessions.filter(function (s) { return s.id !== sid; });
  if (!left.length) d.sh.deleteRow(rec.row);
  else d.sh.getRange(rec.row, 6, 1, 3).setValues([[String(left.length), left[left.length - 1].t, JSON.stringify(left)]]);
  const sh = sheet_(SHEET_LOG, LOG_HEADERS), last = sh.getLastRow();
  if (last >= 2) { const ids = sh.getRange(2, LOG_ID_COL, last - 1, 1).getValues(); for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === sid) sh.deleteRow(i + 2); }
  return { ok: true };
}
function setPretestAction_(req) {
  if (!pwCheck_(req.password).ok) return { ok: false, error: '密碼錯誤或暫停登入中' };
  const cur = parseJSON_(props_().getProperty('PRETEST_JSON'), {});
  let n = 0;
  (req.items || []).slice(0, 300).forEach(function (it) {
    const k = keyOf_(it.cls, it.seat), lv = Number(it.level);
    if (/^\d{3}\|\d+$/.test(k) && lv >= 0 && lv <= 4) { cur[k] = lv; n++; }
  });
  props_().setProperty('PRETEST_JSON', JSON.stringify(cur));
  return { ok: true, count: n };
}
/* 更新答案：老師在老師頁貼上 PZA1: 開頭的代碼（存在指令碼屬性，分段存放） */
function setAnswersAction_(req) {
  if (!pwCheck_(req.password).ok) return { ok: false, error: '密碼錯誤或暫停登入中' };
  const p = props_();
  if (req.reset) { const n = Number(p.getProperty('ANS_CHUNKS') || 0); for (let i = 0; i < n; i++) p.deleteProperty('ANS_' + i); p.deleteProperty('ANS_CHUNKS'); return { ok: true, reset: true }; }
  const code = String(req.code || '').trim();
  if (code.indexOf('PZA1:') !== 0) return { ok: false, error: '代碼格式不對，請確認是 PZA1: 開頭' };
  let json;
  try { json = Utilities.newBlob(Utilities.base64Decode(code.slice(5))).getDataAsString('UTF-8'); } catch (e) { return { ok: false, error: '代碼無法解讀，可能沒有複製完整' }; }
  const a = parseJSON_(json, null);
  if (!a || !a.sets || !a.grade) return { ok: false, error: '代碼內容不完整' };
  const SIZE = 8000, chunks = [];
  for (let i = 0; i < json.length; i += SIZE) chunks.push(json.slice(i, i + SIZE));
  const old = Number(p.getProperty('ANS_CHUNKS') || 0);
  for (let i = 0; i < old; i++) p.deleteProperty('ANS_' + i);
  chunks.forEach(function (c, i) { p.setProperty('ANS_' + i, c); });
  p.setProperty('ANS_CHUNKS', String(chunks.length));
  return { ok: true, version: a.version || '', sets: Object.keys(a.sets).length };
}
