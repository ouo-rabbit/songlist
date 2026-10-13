/**
 * 兔子島點歌本 —— 網頁點歌後台（Google Apps Script）
 *
 * 觀眾在網頁歌單按「確定點歌」→ 記到這份試算表的「點歌紀錄」分頁
 * UtaStage 每幾秒來拿新的點歌 → 自動加進待唱清單
 * 網頁歌單會來問「今天點過哪些歌」→ 在歌名旁邊顯示「今天已點過」
 *
 * 安裝方法請看「網頁點歌設定教學.md」。要改的只有下面 CONFIG。
 */
var CONFIG = {
  // UtaStage 拿點歌用的密碼：自己亂打一串英數字，UtaStage 設定裡填一樣的
  SECRET: 'change-me-please',
  SONG_SHEET: '網頁歌單',        // 歌單分頁（欄位：語言 | 歌手 | 歌名 | 備註）
  LOG_SHEET: '點歌紀錄',         // 會自動建立
  TIMEZONE: 'Asia/Taipei',
  DAY_START_HOUR: 5,            // 凌晨 5 點前還算前一天（深夜場不會被切成兩天）
  PER_HOUR_LIMIT: 60,           // 全部人加起來一小時最多幾首（防洗版）
  REQUIRE_LIVE: true            // true = UtaStage 沒開著（或沒勾網頁點歌）時不開放網頁點歌
};

var VERSION = 3;
var HEADER = ['編號', '時間', '日期', '歌名', '歌手', '暱稱', '語言', '觀眾代號'];

/* ------------------------------------------------------------------ 入口 */
function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.action === 'pull') return json(pull(p));
    if (p.action === 'ping') return json({ ok: true, version: VERSION, day: dayInfo().key });
    return json(today());
  } catch (err) {
    return json({ ok: false, err: 'server', msg: String(err) });
  }
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {
    return json({ ok: false, err: 'bad' });
  }
  try {
    if (body.action === 'sync') return json(sync(body));
    if (body.action === 'request') return json(request(body));
    return json({ ok: false, err: 'bad' });
  } catch (err) {
    return json({ ok: false, err: 'server', msg: String(err) });
  }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------------ 小工具 */
function norm(s) {
  return String(s == null ? '' : s).toLowerCase()
    .replace(/\.[a-z0-9]{2,4}$/, '')
    .replace(/[\s\-_()\[\]（）【】·・,，.。!！?？'"「」《》\/\\]/g, '');
}

function dayInfo(ms) {
  var now = ms || Date.now();
  var shifted = new Date(now - CONFIG.DAY_START_HOUR * 3600000);
  var key = Utilities.formatDate(shifted, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  var off = Utilities.formatDate(shifted, CONFIG.TIMEZONE, 'Z');          // +0800
  var hh = ('0' + CONFIG.DAY_START_HOUR).slice(-2);
  var start = Date.parse(key + 'T' + hh + ':00:00' + off.slice(0, 3) + ':' + off.slice(3));
  return { key: key, start: start };
}

function cache() { return CacheService.getScriptCache(); }

function isLive() {
  var t = Number(cache().get('live') || 0);
  return Date.now() - t < 90000;
}

function checkKey(k) {
  if (!CONFIG.SECRET || CONFIG.SECRET === 'change-me-please') return false;
  return String(k || '') === CONFIG.SECRET;
}

function logSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CONFIG.LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.LOG_SHEET);
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('yyyy/MM/dd HH:mm:ss');
  }
  return sh;
}

/* 最近 600 筆紀錄（今天的一定在最底下） */
function recentRows() {
  var sh = logSheet();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var first = Math.max(2, last - 599);
  return sh.getRange(first, 1, last - first + 1, HEADER.length).getValues().map(function (r) {
    return { seq: Number(r[0]) || 0, time: r[1] instanceof Date ? r[1].getTime() : Number(r[1]) || 0,
             day: String(r[2] instanceof Date ? Utilities.formatDate(r[2], CONFIG.TIMEZONE, 'yyyy-MM-dd') : r[2]),
             title: String(r[3]), artist: String(r[4]), name: String(r[5]), lang: String(r[6]), client: String(r[7]) };
  });
}

function songs() {
  var c = cache().get('songs');
  if (c) return JSON.parse(c);
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SONG_SHEET);
  if (!sh) return [];
  var rows = sh.getDataRange().getDisplayValues();
  if (!rows.length) return [];
  var h = rows[0].map(String);
  var ix = function (k) { for (var i = 0; i < h.length; i++) if (h[i].indexOf(k) >= 0) return i; return -1; };
  var li = ix('語言'), ai = ix('歌手'), ti = ix('歌名'), body = rows.slice(1);
  if (ti < 0) { li = 0; ai = 1; ti = 2; body = rows; }
  var list = body.map(function (r) {
    return { lang: String(r[li] || '').trim(), artist: String(r[ai] || '').trim(), title: String(r[ti] || '').trim() };
  }).filter(function (s) { return s.title; });
  try { cache().put('songs', JSON.stringify(list), 300); } catch (e) {}
  return list;
}

/* ------------------------------------------------------------------ 今天點過的歌（網頁歌單用） */
function today() {
  var d = dayInfo();
  var c = cache().get('today:' + d.key);
  var data;
  if (c) data = JSON.parse(c);
  else {
    var rows = recentRows().filter(function (r) { return r.day === d.key; });
    data = { web: rows.map(function (r) { return r.title; }) };
    try { cache().put('today:' + d.key, JSON.stringify(data), 30); } catch (e) {}
  }
  var synced = [];
  try { synced = JSON.parse(cache().get('synced:' + d.key) || '[]'); } catch (e) {}
  var seen = {}, titles = [];
  data.web.concat(synced).forEach(function (t) {
    var k = norm(t);
    if (k && !seen[k]) { seen[k] = true; titles.push(t); }
  });
  /* open = 現在開放網頁點歌（REQUIRE_LIVE 時要 UtaStage 開著而且有勾網頁點歌） */
  return { ok: true, day: d.key, open: !CONFIG.REQUIRE_LIVE || isLive(), titles: titles };
}

/* ------------------------------------------------------------------ UtaStage 來拿新點歌 */
function pull(p) {
  if (!checkKey(p.key)) return { ok: false, err: 'key' };
  cache().put('live', String(Date.now()), 600);
  var d = dayInfo();
  var after = Number(p.after) || 0;
  var last = Number(cache().get('seq') || PropertiesService.getScriptProperties().getProperty('seq') || 0);
  if (after >= last && after > 0) return { ok: true, day: d.key, dayStart: d.start, last: last, items: [] };
  var items = recentRows().filter(function (r) { return r.seq > after && r.day === d.key; }).map(function (r) {
    return { seq: r.seq, time: r.time, title: r.title, artist: r.artist, name: r.name };
  });
  return { ok: true, day: d.key, dayStart: d.start, last: last, items: items };
}

/* UtaStage 送來「今天點過的歌」（聊天室點的也算）→ 網頁歌單會標「今天已點過」 */
function sync(b) {
  if (!checkKey(b.key)) return { ok: false, err: 'key' };
  var d = dayInfo();
  var titles = (Array.isArray(b.titles) ? b.titles : []).slice(0, 400).map(function (t) {
    return String(t).slice(0, 120);
  });
  cache().put('synced:' + d.key, JSON.stringify(titles), 21600);
  cache().put('live', String(Date.now()), 600);
  return { ok: true, day: d.key, dayStart: d.start, n: titles.length };
}

/* ------------------------------------------------------------------ 觀眾點歌 */
function request(b) {
  var title = String(b.title || '').trim().slice(0, 120);
  var artist = String(b.artist || '').trim().slice(0, 80);
  if (!title) return { ok: false, err: 'bad' };
  if (CONFIG.REQUIRE_LIVE && !isLive()) return { ok: false, err: 'closed' };

  /* 只能點歌單上有的歌 */
  var key = norm(title), song = null;
  songs().some(function (s) {
    if (norm(s.title) === key && (!artist || norm(s.artist) === norm(artist))) { song = s; return true; }
    return false;
  });
  if (!song) return { ok: false, err: 'notfound' };

  var c = cache();
  var d = dayInfo();
  var hourKey = 'hour:' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMddHH');
  if (Number(c.get(hourKey) || 0) >= CONFIG.PER_HOUR_LIMIT) return { ok: false, err: 'busy' };

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) return { ok: false, err: 'busy' };
  try {
    var rows = recentRows().filter(function (r) { return r.day === d.key; });
    if (rows.some(function (r) { return norm(r.title) === key; })) {
      return { ok: true, dup: true, titles: today().titles };
    }
    var props = PropertiesService.getScriptProperties();
    var seq = (Number(props.getProperty('seq')) || 0) + 1;
    props.setProperty('seq', String(seq));
    /* 文字前面加 ' ：不讓試算表把歌名當成日期 / 數字 */
    var txt = function (v) { return "'" + v; };
    logSheet().appendRow([seq, new Date(), txt(d.key), txt(song.title), txt(song.artist), '', txt(song.lang), '']);
    SpreadsheetApp.flush();
    c.put('seq', String(seq), 21600);
    c.remove('today:' + d.key);
    c.put(hourKey, String(Number(c.get(hourKey) || 0) + 1), 3700);
  } finally {
    lock.releaseLock();
  }
  var t = today();
  return { ok: true, seq: seq, open: t.open, titles: t.titles };
}

/* 在編輯器裡按「執行」測一下（不會真的點歌）：看看讀不讀得到歌單、今天點了哪些 */
function testSetup() {
  Logger.log('歌單：' + songs().length + ' 首');
  Logger.log('今天（' + dayInfo().key + '）點過：' + JSON.stringify(today().titles));
  Logger.log('SECRET 已設定：' + (CONFIG.SECRET !== 'change-me-please'));
}
