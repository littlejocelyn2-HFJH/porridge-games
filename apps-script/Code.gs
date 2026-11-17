/**
 * 〈以粥養生〉閱讀小誌（診斷後鷹架練習）— 後端 Google Apps Script
 * 需要兩個檔案：Code.gs（本檔）＋ Answers.gs（老師另外拿到，不在 GitHub 上）
 * 部署：部署 → 新增部署作業 → 網頁應用程式｜執行身分：我｜存取權：所有人
 * 改了程式之後：部署 → 管理部署作業 → 編輯 → 版本選「新版本」→ 部署（網址不變）
 */
// ★ 請把下一行的 change-me 改成老師自己的密碼（10 個字以上、英文夾數字），保留引號
const TEACHER_PASSWORD = 'change-me';

const ITEM = 'porridge-games';
const SHEET_DATA = '資料（程式用）';
const SHEET_LOG = '練習紀錄（老師看）';
const MAX_SESSIONS = 20;
const DATA_HEADERS = ['key', '班級', '座號', '姓名', '前測層級', '練習次數', '最後練習', 'sessionsJSON', '更新時間'];
const LOG_HEADERS = ['開始時間', '班級', '座號', '姓名', '前測層級', '模式', '完成話數',
  '第一話 一次就對', '第一話 嘗試', '第二話 一次就對', '第二話 嘗試', '第三話 一次就對', '第三話 嘗試', '第四話 一次就對', '第四話 嘗試',
  '總用時(秒)', '第一話填空', '第二話想法', '第三話一句話', '第四話想法', 'sessionId',
  '任務一：看見生活', '任務二：取名字', '任務三：句型', '任務三：統整詞與理由'];
const LOG_ID_COL = 21; // sessionId 所在欄（固定）

function doGet() { return out({ ok: true, msg: '〈以粥養生〉閱讀小誌後端運作中。這是後端網址，請貼到 config.js。', answers: typeof ANSWERS !== 'undefined' }); }

function doPost(e) {
  let req = {};
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return out({ ok: false, error: '資料格式錯誤' }); }
  try {
    if (typeof ANSWERS === 'undefined') return out({ ok: false, error: '老師還沒貼上 Answers.gs' });
    switch (req.action) {
      case 'ping': return out({ ok: true });
      case 'start': return out(startAction_(req));
      case 'check': return out(checkAction_(req));
      case 'list': return out(listAction_(req));
      case 'save': return out(withLock_(function () { return saveAction_(req); }));
      case 'clearClass': return out(withLock_(function () { return clearAction_(req); }));
      case 'deleteSession': return out(withLock_(function () { return deleteSessionAction_(req); }));
      case 'setPretest': return out(setPretestAction_(req));
      default: return out({ ok: false, error: '未知的動作' });
    }
  } catch (err) { return out({ ok: false, error: String(err && err.message || err) }); }
}

/* ---------- 工具 ---------- */
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function withLock_(fn) { const l = LockService.getScriptLock(); l.waitLock(25000); try { return fn(); } finally { l.releaseLock(); } }
function str_(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u0008\u000b-\u001f]/g, '').trim().slice(0, max || 200); }
// 防試算表公式注入：開頭是 = + - @ 的文字前面加 '
function safe_(v) { const s = String(v == null ? '' : v); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function normSeat_(s) { const n = parseInt(String(s).replace(/[^\d]/g, ''), 10); return isNaN(n) ? '' : String(n); }
function keyOf_(cls, seat) { return str_(cls, 10) + '|' + normSeat_(seat); }
function parseJSON_(s, d) { try { return s ? JSON.parse(s) : d; } catch (e) { return d; } }
function twTime_(iso) { try { return Utilities.formatDate(new Date(iso), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss'); } catch (e) { return iso; } }

function validStudent_(st) {
  st = st || {};
  const cls = str_(st.cls, 10), seat = normSeat_(st.seat), name = str_(st.name, 20);
  if (!/^\d{3}$/.test(cls)) return { error: '班級請填三位數字，例如 902' };
  if (typeof ALLOWED_CLASSES !== 'undefined' && ALLOWED_CLASSES.length && ALLOWED_CLASSES.indexOf(cls) < 0) return { error: '這個班級沒有開放，請確認班級' };
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
  const extra = parseJSON_(PropertiesService.getScriptProperties().getProperty('PRETEST_JSON'), {});
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
    sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
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
function startAction_(req) {
  const st = validStudent_(req.student);
  if (st.error) return { ok: false, error: st.error };
  const lv = pretestMap_()[st.key];
  const mode = modeOf_(lv);
  const res = { ok: true, mode: mode, pretest: lv == null ? null : lv };
  // 支持模式：第一話一開始就標出段落提示
  if (mode === 'support') res.paraHints = { 1: ANSWERS[1].para, 2: ANSWERS[2].para };
  return res;
}

/* ---------- check：批改一次嘗試（不寫入試算表） ---------- */
function checkAction_(req) {
  const lv = Number(req.level), A = ANSWERS[lv];
  if (!A) return { ok: false, error: '沒有這一話' };
  const tryNo = Math.max(1, Math.min(9, Number(req.tryNo) || 1));
  const res = { ok: true, results: {}, hints: {} };
  if (lv === 1 && req.induceRef) { // 歸納任務：學生每一格都寫了，才回傳參考筆記
    const task = Number(req.induceRef), w = req.written || {}, ref = (ANSWERS[1].induce || {})['t' + task];
    const need = { 1: ['dad', 'ama', 'mom'], 2: ['r1', 'r2', 'r3', 'r4'], 3: ['s1', 's2', 's3', 's4', 'word', 'reason'] }[task];
    if (!need || !ref) return { ok: false, error: '沒有這個任務' };
    if (need.some(function (k) { return !str_(w[k], 200); })) return { ok: false, error: '每一格都寫一點，再看參考筆記' };
    return { ok: true, ref: ref };
  }
  if (lv === 1 && req.fillOnly) {
    const f = str_(req.fill, 20).replace(/[、，,\s]/g, '');
    res.fillOk = A.fill.indexOf(f) >= 0 || (f.indexOf('甜') >= 0 && f.indexOf('鹹') >= 0 && f.length <= 4);
    if (!res.fillOk && tryNo >= 2) res.fillReveal = A.fill[0];
    return res;
  }
  if (lv === 2) {
    const order = (req.order || []).slice(0, 10).map(String);
    order.forEach(function (id, i) { res.results[id] = A.order[i] === id; if (!res.results[id]) res.hints[id] = A.para[id] || ''; });
    res.allRight = order.length === A.order.length && order.every(function (id, i) { return A.order[i] === id; });
    if (!res.allRight && tryNo >= 3) res.reveal = A.order;
  } else {
    const p = req.placements || {};
    Object.keys(A.key).forEach(function (id) {
      const ok = p[id] === A.key[id]; res.results[id] = ok;
      if (!ok) res.hints[id] = A.para[id] || '';
    });
    res.allRight = Object.keys(res.results).every(function (id) { return res.results[id]; });
    if (!res.allRight && tryNo >= 3) { res.reveal = {}; Object.keys(A.key).forEach(function (id) { if (!res.results[id]) res.reveal[id] = A.key[id]; }); }
  }
  if (lv === 1 && req.fill != null) {
    const f = str_(req.fill, 20).replace(/[、，,\s]/g, '');
    res.fillOk = A.fill.indexOf(f) >= 0 || (f.indexOf('甜') >= 0 && f.indexOf('鹹') >= 0 && f.length <= 4);
    if (!res.fillOk && tryNo >= 2) res.fillReveal = A.fill[0];
  }
  return res;
}

/* 伺服器重算一話的成績（不信任前端分數） */
function grade_(lv, rec) {
  const A = ANSWERS[lv]; if (!A || !rec) return null;
  const tries = Array.isArray(rec.tries) ? rec.tries.slice(0, 9) : [];
  function score(t) {
    if (lv === 2) { const o = Array.isArray(t) ? t : []; let c = 0; A.order.forEach(function (id, i) { if (o[i] === id) c++; }); return c; }
    let c = 0; Object.keys(A.key).forEach(function (id) { if (t && t[id] === A.key[id]) c++; }); return c;
  }
  const total = lv === 2 ? A.order.length : Object.keys(A.key).length;
  const first = tries.length ? score(tries[0]) : 0;
  let solvedAt = 0; for (let i = 0; i < tries.length; i++) { if (score(tries[i]) === total) { solvedAt = i + 1; break; } }
  return { first: first, total: total, tries: tries.length, solvedAt: solvedAt };
}

/* ---------- save：儲存（同一次練習會一話一話更新） ---------- */
function saveAction_(req) {
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
  [1, 2, 3, 4].forEach(function (lv) {
    const r = se.levels && se.levels[lv]; if (!r) return;
    const tries = (Array.isArray(r.tries) ? r.tries : []).slice(0, 9).map(function (t) {
      if (lv === 2) return (Array.isArray(t) ? t : []).slice(0, 10).map(function (x) { return str_(x, 10); });
      const o = {}; Object.keys(t || {}).slice(0, 20).forEach(function (k) { o[str_(k, 10)] = str_(t[k], 12); }); return o;
    });
    const g = grade_(lv, { tries: tries });
    clean.levels[lv] = { tries: tries, dur: Math.max(0, Math.min(36000, Math.round(Number(r.dur) || 0))), done: !!r.done, hints: Math.min(99, Number(r.hints) || 0),
      fill: str_(r.fill, 20), reflect: str_(r.reflect, 600), g: g };
    if (lv === 1 && r.induce) clean.levels[1].induce = cleanInduce_(r.induce);
  });
  if (exist) sessions = sessions.map(function (x) { return x.id === sid ? clean : x; });
  else sessions.push(clean);
  if (sessions.length > MAX_SESSIONS) sessions = [sessions[0]].concat(sessions.slice(sessions.length - MAX_SESSIONS + 1));
  const now = new Date().toISOString();
  const row = [st.key, st.cls, st.seat, safe_(st.name), lvPre == null ? '' : String(lvPre), String(sessions.length), clean.t, JSON.stringify(sessions), now];
  if (rec) d.sh.getRange(rec.row, 1, 1, row.length).setValues([row]); else d.sh.appendRow(row);
  logUpsert_(st, clean);
  // 回傳伺服器算出的成績，給學生的報告書使用
  const graded = {}; Object.keys(clean.levels).forEach(function (lv) { graded[lv] = clean.levels[lv].g; });
  return { ok: true, graded: graded, mode: clean.mode, pretest: clean.pretest };
}
function cleanInduce_(o) {
  const pick = function (x, keys, max) { const r = {}; keys.forEach(function (k) { r[k] = str_(x && x[k], max); }); return r; };
  return { t1: pick(o.t1, ['dad', 'ama', 'mom'], 200), t2: pick(o.t2, ['r1', 'r2', 'r3', 'r4'], 40),
           t3: pick(o.t3, ['s1', 's2', 's3', 's4', 'word', 'reason'], 200) };
}
function induceText_(i, part) {
  if (!i) return '';
  if (part === 1) return '爸爸：' + i.t1.dad + '／阿媽：' + i.t1.ama + '／媽媽：' + i.t1.mom;
  if (part === 2) return [i.t2.r1, i.t2.r2, i.t2.r3, i.t2.r4].join('／');
  if (part === 3) return [i.t3.s1, i.t3.s2, i.t3.s3, i.t3.s4].join('／');
  return i.t3.word + '：' + i.t3.reason;
}
function logUpsert_(st, s) {
  const sh = sheet_(SHEET_LOG, LOG_HEADERS);
  const L = s.levels, cell = function (lv, k) { return L[lv] && L[lv].g ? String(k === 'first' ? L[lv].g.first + '/' + L[lv].g.total : L[lv].g.tries) : ''; };
  const done = [1, 2, 3, 4].filter(function (lv) { return L[lv] && L[lv].done; }).length;
  const dur = [1, 2, 3, 4].reduce(function (a, lv) { return a + (L[lv] ? L[lv].dur : 0); }, 0);
  const modeName = { support: '支持', standard: '一般', challenge: '挑戰' }[s.mode];
  const row = [twTime_(s.start || s.t), st.cls, st.seat, safe_(st.name), s.pretest == null ? '' : String(s.pretest), modeName, String(done),
    cell(1, 'first'), cell(1, 'tries'), cell(2, 'first'), cell(2, 'tries'), cell(3, 'first'), cell(3, 'tries'), cell(4, 'first'), cell(4, 'tries'),
    String(dur), safe_(L[1] ? L[1].fill : ''), safe_(L[2] ? L[2].reflect : ''), safe_(L[3] ? L[3].reflect : ''), safe_(L[4] ? L[4].reflect : ''), s.id,
    safe_(induceText_(L[1] && L[1].induce, 1)), safe_(induceText_(L[1] && L[1].induce, 2)), safe_(induceText_(L[1] && L[1].induce, 3)), safe_(induceText_(L[1] && L[1].induce, 4))];
  const last = sh.getLastRow();
  if (last >= 2) {
    const ids = sh.getRange(2, LOG_ID_COL, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === s.id) { sh.getRange(i + 2, 1, 1, row.length).setValues([row]); return; }
  }
  sh.appendRow(row);
}

/* ---------- list ---------- */
function listAction_(req) {
  const d = readData_();
  if (req.password) {
    const pw = pwCheck_(req.password);
    if (pw.ok) return { ok: true, authed: true, rows: d.rows.map(function (r) { return { cls: r.cls, seat: r.seat, name: r.name, sessions: r.sessions }; }), answers: ANSWERS, pretest: pretestMap_() };
    return { ok: true, authed: false, pwError: true, locked: !!pw.locked, rows: publicRows_(d.rows) };
  }
  return { ok: true, authed: false, rows: publicRows_(d.rows) };
}
function publicRows_(rows) {
  return rows.map(function (r) {
    const best = r.sessions.reduce(function (m, s) { return Math.max(m, [1, 2, 3, 4].filter(function (lv) { return s.levels && s.levels[lv] && s.levels[lv].done; }).length); }, 0);
    return { cls: r.cls, seat: r.seat, name: r.name, chapters: best };
  });
}

/* ---------- 老師：清除班級、刪除單次練習、匯入前測層級 ---------- */
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
  else { d.sh.getRange(rec.row, 6, 1, 3).setValues([[String(left.length), left[left.length - 1].t, JSON.stringify(left)]]); }
  const sh = sheet_(SHEET_LOG, LOG_HEADERS), last = sh.getLastRow();
  if (last >= 2) { const ids = sh.getRange(2, LOG_ID_COL, last - 1, 1).getValues(); for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === sid) sh.deleteRow(i + 2); }
  return { ok: true };
}
function setPretestAction_(req) {
  if (!pwCheck_(req.password).ok) return { ok: false, error: '密碼錯誤或暫停登入中' };
  const props = PropertiesService.getScriptProperties();
  const cur = parseJSON_(props.getProperty('PRETEST_JSON'), {});
  let n = 0;
  (req.items || []).slice(0, 300).forEach(function (it) {
    const k = keyOf_(it.cls, it.seat), lv = Number(it.level);
    if (/^\d{3}\|\d+$/.test(k) && lv >= 0 && lv <= 4) { cur[k] = lv; n++; }
  });
  props.setProperty('PRETEST_JSON', JSON.stringify(cur));
  return { ok: true, count: n };
}
