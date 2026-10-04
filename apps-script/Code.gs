/**
 * 八縣市首長賭盤後端(Google Apps Script)
 *
 * 部署步驟:
 * 1. 新建一份 Google 試算表。
 * 2. 選單「擴充功能」→「Apps Script」,把這個檔案整份貼進去,取代原本內容。
 * 3. 把下面 HOST_KEY 改成只有你知道的密碼(莊家用來填開票結果、封盤)。
 * 4. 右上「部署」→「新增部署作業」→ 類型選「網頁應用程式」。
 *    執行身分:我;誰可以存取:所有人。按部署並授權。
 * 5. 複製「網頁應用程式網址」(結尾是 /exec),貼到 bet.html 最上面的 API_URL。
 *
 * 之後改了這份程式,要再「管理部署作業」→ 編輯 → 新版本,網址才會生效。
 */

var HOST_KEY = 'change-me';
var LOCK_AT = new Date('2026-11-28T08:00:00+08:00').getTime();
var COUNTIES = ['臺北市', '新北市', '桃園市', '臺中市', '臺南市', '高雄市', '新竹縣', '新竹市'];

function sheet_(name, header) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    if (header) s.appendRow(header);
  }
  return s;
}

function readConfig_() {
  var v = sheet_('config').getRange('A1').getValue();
  try { return JSON.parse(v) || {}; } catch (e) { return {}; }
}

function isLocked_(cfg) {
  return !!cfg.locked || Date.now() >= LOCK_AT;
}

function state_() {
  var cfg = readConfig_();
  var rows = sheet_('picks', ['name', 'token', 'picks', 'updatedAt']).getDataRange().getValues().slice(1);
  var picks = rows.filter(function (r) { return r[0]; }).map(function (r) {
    var p = {};
    try { p = JSON.parse(r[2]) || {}; } catch (e) {}
    return { name: String(r[0]), picks: p };
  });
  return {
    config: { opts: cfg.opts || {}, results: cfg.results || {}, locked: !!cfg.locked },
    picks: picks,
    now: Date.now()
  };
}

function pick_(p) {
  var cfg = readConfig_();
  if (isLocked_(cfg)) return { error: 'locked' };
  var name = String(p.name || '').trim();
  var token = String(p.token || '');
  if (!name || name.length > 16 || token.length < 8) return { error: 'bad_request' };

  var incoming = {};
  try { incoming = JSON.parse(p.picks || '{}'); } catch (e) { return { error: 'bad_request' }; }
  var clean = {};
  COUNTIES.forEach(function (c) {
    if (typeof incoming[c] === 'string' && incoming[c].length <= 30) clean[c] = incoming[c];
  });

  var s = sheet_('picks', ['name', 'token', 'picks', 'updatedAt']);
  var data = s.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === name) {
      if (String(data[i][1]) !== token) return { error: 'name_taken' };
      s.getRange(i + 1, 3, 1, 2).setValues([[JSON.stringify(clean), Date.now()]]);
      return { ok: true };
    }
  }
  s.appendRow([name, token, JSON.stringify(clean), Date.now()]);
  return { ok: true };
}

function config_(p) {
  if (String(p.key || '') !== HOST_KEY) return { error: 'bad_key' };
  var cfg;
  try { cfg = JSON.parse(p.config); } catch (e) { return { error: 'bad_request' }; }
  sheet_('config').getRange('A1').setValue(JSON.stringify({
    opts: cfg.opts || {},
    results: cfg.results || {},
    locked: !!cfg.locked
  }));
  return { ok: true };
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  var lock = LockService.getScriptLock();
  var out;
  lock.waitLock(10000);
  try {
    if (p.a === 'pick') out = pick_(p);
    else if (p.a === 'config') out = config_(p);
    else out = state_();
  } catch (err) {
    out = { error: String(err) };
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
