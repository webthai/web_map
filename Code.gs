/**
 * ===== OFFLINE MAP APP - Google Apps Script Backend =====
 * Database: Google Sheets (auto-created on first run)
 * Sheets:
 *   Users  : id | username | passwordHash | role | token | createdAt
 *   Places : id | userId | name | lat | lng | note | createdAt | updatedAt
 *
 * Deploy as Web App:
 *   Execute as: Me
 *   Who has access: Anyone
 *   IMPORTANT: every time you edit this code, you must create a NEW
 *   deployment version (Deploy > Manage deployments > Edit > New version)
 *   or your changes will NOT go live.
 */

var TZ = 'Asia/Bangkok';
var CACHE_TTL = 30; // seconds

function getSS_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function ensureSheet_(name, headers) {
  var ss = getSS_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function ensureSchema_() {
  ensureSheet_('Users', ['id', 'username', 'passwordHash', 'role', 'token', 'createdAt']);
  ensureSheet_('Places', ['id', 'userId', 'name', 'lat', 'lng', 'note', 'createdAt', 'updatedAt']);
}

function sheetToObjects_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (!row[0]) continue; // skip blank id rows
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = row[j];
    out.push(obj);
  }
  return out;
}

function findRowIndexById_(sheet, id) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1; // 1-based row number
  }
  return -1;
}

function nowISO_() {
  return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function invalidateCache_() {
  CacheService.getScriptCache().removeAll(['bootstrap_all', 'bootstrap_public']);
}

/**
 * All requests are POST with a text/plain body containing JSON:
 * { action: "...", ...params }
 * (text/plain avoids CORS preflight issues with Apps Script)
 */
function doPost(e) {
  ensureSchema_();
  var body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut_({ ok: false, error: 'ข้อมูลคำขอไม่ถูกต้อง' });
  }
  var action = body.action;
  try {
    switch (action) {
      case 'register': return jsonOut_(handleRegister_(body));
      case 'login': return jsonOut_(handleLogin_(body));
      case 'bootstrap': return jsonOut_(handleBootstrap_(body));
      case 'addPlace': return jsonOut_(handleAddPlace_(body));
      case 'updatePlace': return jsonOut_(handleUpdatePlace_(body));
      case 'deletePlace': return jsonOut_(handleDeletePlace_(body));
      default: return jsonOut_({ ok: false, error: 'ไม่รู้จักคำสั่งนี้' });
    }
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  return jsonOut_({ ok: true, message: 'Offline Map App API is running' });
}

/* ---------- Auth helpers ---------- */

function getUserByToken_(token) {
  if (!token) return null;
  var sheet = ensureSheet_('Users', ['id', 'username', 'passwordHash', 'role', 'token', 'createdAt']);
  var users = sheetToObjects_(sheet);
  for (var i = 0; i < users.length; i++) {
    if (users[i].token === token) return users[i];
  }
  return null;
}

function handleRegister_(body) {
  var username = (body.username || '').trim();
  var passwordHash = body.passwordHash || '';
  if (!username || !passwordHash) return { ok: false, error: 'กรุณากรอกชื่อผู้ใช้และรหัสผ่าน' };

  var sheet = ensureSheet_('Users', ['id', 'username', 'passwordHash', 'role', 'token', 'createdAt']);
  var users = sheetToObjects_(sheet);
  for (var i = 0; i < users.length; i++) {
    if (users[i].username.toLowerCase() === username.toLowerCase()) {
      return { ok: false, error: 'มีชื่อผู้ใช้นี้อยู่แล้ว' };
    }
  }
  var isFirstUser = users.length === 0;
  var id = Utilities.getUuid();
  var token = Utilities.getUuid();
  sheet.appendRow([id, username, passwordHash, isFirstUser ? 'admin' : 'user', token, nowISO_()]);
  invalidateCache_();
  return { ok: true, userId: id, username: username, role: isFirstUser ? 'admin' : 'user', token: token };
}

function handleLogin_(body) {
  var username = (body.username || '').trim();
  var passwordHash = body.passwordHash || '';
  var sheet = ensureSheet_('Users', ['id', 'username', 'passwordHash', 'role', 'token', 'createdAt']);
  var users = sheetToObjects_(sheet);
  for (var i = 0; i < users.length; i++) {
    if (users[i].username.toLowerCase() === username.toLowerCase()) {
      if (users[i].passwordHash !== passwordHash) {
        return { ok: false, error: 'รหัสผ่านไม่ถูกต้อง' };
      }
      // rotate token each login
      var newToken = Utilities.getUuid();
      var rowIndex = findRowIndexById_(sheet, users[i].id);
      sheet.getRange(rowIndex, 5).setValue(newToken); // token column = 5
      return { ok: true, userId: users[i].id, username: users[i].username, role: users[i].role, token: newToken };
    }
  }
  return { ok: false, error: 'ไม่พบชื่อผู้ใช้นี้' };
}

/* ---------- Bootstrap (single combined load) ---------- */

function handleBootstrap_(body) {
  var user = getUserByToken_(body.token);
  if (!user) return { ok: false, error: 'กรุณาเข้าสู่ระบบใหม่' };

  var cacheKey = user.role === 'admin' ? 'bootstrap_all' : ('bootstrap_' + user.id);
  var cache = CacheService.getScriptCache();
  var cached = cache.get(cacheKey);
  if (cached) {
    var parsed = JSON.parse(cached);
    parsed.user = { id: user.id, username: user.username, role: user.role };
    parsed.cached = true;
    return parsed;
  }

  var placesSheet = ensureSheet_('Places', ['id', 'userId', 'name', 'lat', 'lng', 'note', 'createdAt', 'updatedAt']);
  var allPlaces = sheetToObjects_(placesSheet);
  var places = user.role === 'admin' ? allPlaces : allPlaces.filter(function (p) { return p.userId === user.id; });

  var result = { ok: true, places: places };
  cache.put(cacheKey, JSON.stringify(result), CACHE_TTL);

  result.user = { id: user.id, username: user.username, role: user.role };
  result.cached = false;
  return result;
}

/* ---------- Places CRUD (permission enforced server-side) ---------- */

function handleAddPlace_(body) {
  var user = getUserByToken_(body.token);
  if (!user) return { ok: false, error: 'กรุณาเข้าสู่ระบบใหม่' };
  var name = (body.name || '').trim();
  if (!name || typeof body.lat !== 'number' && isNaN(parseFloat(body.lat))) {
    return { ok: false, error: 'ข้อมูลสถานที่ไม่ครบ' };
  }
  var sheet = ensureSheet_('Places', ['id', 'userId', 'name', 'lat', 'lng', 'note', 'createdAt', 'updatedAt']);
  var id = Utilities.getUuid();
  var ts = nowISO_();
  sheet.appendRow([id, user.id, name, parseFloat(body.lat), parseFloat(body.lng), body.note || '', ts, ts]);
  invalidateCache_();
  return { ok: true, id: id };
}

function handleUpdatePlace_(body) {
  var user = getUserByToken_(body.token);
  if (!user) return { ok: false, error: 'กรุณาเข้าสู่ระบบใหม่' };
  var sheet = ensureSheet_('Places', ['id', 'userId', 'name', 'lat', 'lng', 'note', 'createdAt', 'updatedAt']);
  var rowIndex = findRowIndexById_(sheet, body.id);
  if (rowIndex === -1) return { ok: false, error: 'ไม่พบข้อมูล' };
  var ownerId = sheet.getRange(rowIndex, 2).getValue();
  if (user.role !== 'admin' && String(ownerId) !== String(user.id)) {
    return { ok: false, error: 'ไม่มีสิทธิ์แก้ไขข้อมูลนี้' };
  }
  if (body.name !== undefined) sheet.getRange(rowIndex, 3).setValue(body.name);
  if (body.lat !== undefined) sheet.getRange(rowIndex, 4).setValue(parseFloat(body.lat));
  if (body.lng !== undefined) sheet.getRange(rowIndex, 5).setValue(parseFloat(body.lng));
  if (body.note !== undefined) sheet.getRange(rowIndex, 6).setValue(body.note);
  sheet.getRange(rowIndex, 8).setValue(nowISO_());
  invalidateCache_();
  return { ok: true };
}

function handleDeletePlace_(body) {
  var user = getUserByToken_(body.token);
  if (!user) return { ok: false, error: 'กรุณาเข้าสู่ระบบใหม่' };
  var sheet = ensureSheet_('Places', ['id', 'userId', 'name', 'lat', 'lng', 'note', 'createdAt', 'updatedAt']);
  var rowIndex = findRowIndexById_(sheet, body.id);
  if (rowIndex === -1) return { ok: false, error: 'ไม่พบข้อมูล' };
  var ownerId = sheet.getRange(rowIndex, 2).getValue();
  if (user.role !== 'admin' && String(ownerId) !== String(user.id)) {
    return { ok: false, error: 'ไม่มีสิทธิ์ลบข้อมูลนี้' };
  }
  sheet.deleteRow(rowIndex);
  invalidateCache_();
  return { ok: true };
}
