/**
 * ===== OFFLINE MAP APP - Google Apps Script Backend =====
 * บัญชีเดียวคงที่ (username: meen / password: 5340) — ตรวจฝั่งเว็บแอปแล้วส่ง
 * APP_TOKEN มาด้วยทุกครั้ง ฝั่งนี้แค่เช็คว่า token ตรงกันก่อนอนุญาตให้อ่าน/เขียนข้อมูล
 *
 * ⚠️ ค่า APP_TOKEN ด้านล่างต้องเหมือนกับ APP_TOKEN ใน index.html เป๊ะๆ
 *
 * Database: Google Sheets (auto-create sheet/หัวตารางให้เองตอนรันครั้งแรก)
 * Sheets:
 *   Places : id | name | lat | lng | note | createdAt | updatedAt
 *
 * Deploy as Web App:
 *   Execute as: Me
 *   Who has access: Anyone
 *   IMPORTANT: แก้โค้ดทีไร ต้องสร้าง deployment เวอร์ชันใหม่ทุกครั้ง
 *   (Deploy > Manage deployments > แก้ไข > Version: New version > Deploy)
 */

var TZ = 'Asia/Bangkok';
var CACHE_TTL = 30; // seconds
var APP_TOKEN = 'meen5340-a8f3e1c9'; // ต้องตรงกับ APP_TOKEN ใน index.html

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
  ensureSheet_('Places', ['id', 'name', 'lat', 'lng', 'note', 'createdAt', 'updatedAt']);
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
  CacheService.getScriptCache().remove('bootstrap');
}

function checkAuth_(body) {
  return body && body.token === APP_TOKEN;
}

/**
 * All requests are POST with a text/plain body containing JSON:
 * { action: "...", token: "...", ...params }
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

  if (!checkAuth_(body)) {
    return jsonOut_({ ok: false, error: 'ไม่ได้รับอนุญาต (token ไม่ถูกต้อง)' });
  }

  var action = body.action;
  try {
    switch (action) {
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

/* ---------- Bootstrap (single combined load) ---------- */

function handleBootstrap_(body) {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('bootstrap');
  if (cached) {
    var parsed = JSON.parse(cached);
    parsed.cached = true;
    return parsed;
  }

  var placesSheet = ensureSheet_('Places', ['id', 'name', 'lat', 'lng', 'note', 'createdAt', 'updatedAt']);
  var places = sheetToObjects_(placesSheet);

  var result = { ok: true, places: places, cached: false };
  cache.put('bootstrap', JSON.stringify(result), CACHE_TTL);
  return result;
}

/* ---------- Places CRUD ---------- */

function handleAddPlace_(body) {
  var name = (body.name || '').trim();
  if (!name || (typeof body.lat !== 'number' && isNaN(parseFloat(body.lat)))) {
    return { ok: false, error: 'ข้อมูลสถานที่ไม่ครบ' };
  }
  var sheet = ensureSheet_('Places', ['id', 'name', 'lat', 'lng', 'note', 'createdAt', 'updatedAt']);
  var id = Utilities.getUuid();
  var ts = nowISO_();
  sheet.appendRow([id, name, parseFloat(body.lat), parseFloat(body.lng), body.note || '', ts, ts]);
  invalidateCache_();
  return { ok: true, id: id };
}

function handleUpdatePlace_(body) {
  var sheet = ensureSheet_('Places', ['id', 'name', 'lat', 'lng', 'note', 'createdAt', 'updatedAt']);
  var rowIndex = findRowIndexById_(sheet, body.id);
  if (rowIndex === -1) return { ok: false, error: 'ไม่พบข้อมูล' };
  if (body.name !== undefined) sheet.getRange(rowIndex, 2).setValue(body.name);
  if (body.lat !== undefined) sheet.getRange(rowIndex, 3).setValue(parseFloat(body.lat));
  if (body.lng !== undefined) sheet.getRange(rowIndex, 4).setValue(parseFloat(body.lng));
  if (body.note !== undefined) sheet.getRange(rowIndex, 5).setValue(body.note);
  sheet.getRange(rowIndex, 7).setValue(nowISO_());
  invalidateCache_();
  return { ok: true };
}

function handleDeletePlace_(body) {
  var sheet = ensureSheet_('Places', ['id', 'name', 'lat', 'lng', 'note', 'createdAt', 'updatedAt']);
  var rowIndex = findRowIndexById_(sheet, body.id);
  if (rowIndex === -1) return { ok: false, error: 'ไม่พบข้อมูล' };
  sheet.deleteRow(rowIndex);
  invalidateCache_();
  return { ok: true };
}
