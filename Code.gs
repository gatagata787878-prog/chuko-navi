/**
 * アクアリンク 中古品在庫管理 ― バックエンド（Google Apps Script）
 * ※ 写真は複数（最大10枚）対応。fileIdをカンマ区切りで photo 列に保存します。
 * シート「在庫」は初回アクセス時に自動生成されます。
 */

var SHEET_NAME = '在庫';
var PHOTO_FOLDER = 'aqualink-photos';
var HEADERS = ['mg','t','g','name','w','d','h','std','rank','status',
               'price','loc','note','damage','dirt','remark','propTo','propUntil','photo','updatedAt'];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); }
  if (sh.getLastRow() === 0) { sh.appendRow(HEADERS); }
  return sh;
}

function thumbUrls_(photos) {
  return photos.map(function (id) { return 'https://drive.google.com/thumbnail?id=' + id + '&sz=w1000'; });
}

function readAll_() {
  var sh = getSheet_();
  var values = sh.getDataRange().getValues();
  var head = values[0];
  var items = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (!row[0]) continue;
    var o = {};
    head.forEach(function (h, j) { o[h] = row[j]; });
    ['w','d','h','price'].forEach(function (k) { o[k] = Number(o[k]) || 0; });
    o.photos = o.photo ? String(o.photo).split(',').filter(function (s) { return s; }) : [];
    o.photoUrls = thumbUrls_(o.photos);
    items.push(o);
  }
  return items;
}

function doGet(e) {
  return json_({ ok: true, items: readAll_() });
}

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    if (req.action === 'save')      return json_(saveItem_(req.item));
    if (req.action === 'setStatus') return json_(setStatus_(req));
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function rowIndexByMg_(sh, mg) {
  var n = sh.getLastRow() - 1;
  if (n <= 0) return -1;
  var mgs = sh.getRange(2, 1, n, 1).getValues();
  for (var i = 0; i < mgs.length; i++) if (mgs[i][0] === mg) return i + 2;
  return -1;
}

function nextMg_(sh) {
  var n = sh.getLastRow() - 1, max = 0;
  if (n > 0) {
    var mgs = sh.getRange(2, 1, n, 1).getValues();
    mgs.forEach(function (r) {
      var m = String(r[0]).match(/(\d+)/);
      if (m && Number(m[1]) > max) max = Number(m[1]);
    });
  }
  return 'AQ-' + ('000' + (max + 1)).slice(-3);
}

function saveItem_(item) {
  var sh = getSheet_();
  if (!item.mg) item.mg = nextMg_(sh);

  // 既存の写真リスト（fileId配列）
  var photos = item.photos || [];
  if (typeof photos === 'string') photos = photos ? photos.split(',') : [];
  photos = photos.filter(function (s) { return s && String(s).indexOf('data:') !== 0; });

  // 新規アップロード分をDriveに保存して追加
  if (item.photoDataUrls && item.photoDataUrls.length) {
    for (var k = 0; k < item.photoDataUrls.length; k++) {
      photos.push(savePhoto_(item.mg, item.photoDataUrls[k]));
    }
  }
  item.photo = photos.join(',');
  item.updatedAt = new Date();

  var row = HEADERS.map(function (h) { return item[h] !== undefined && item[h] !== null ? item[h] : ''; });
  var idx = rowIndexByMg_(sh, item.mg);
  if (idx > 0) sh.getRange(idx, 1, 1, HEADERS.length).setValues([row]);
  else sh.appendRow(row);

  var saved = {};
  HEADERS.forEach(function (h, j) { saved[h] = row[j]; });
  saved.photos = photos;
  saved.photoUrls = thumbUrls_(photos);
  return { ok: true, item: saved };
}

function setStatus_(req) {
  var sh = getSheet_();
  var idx = rowIndexByMg_(sh, req.mg);
  if (idx < 0) return { ok: false, error: 'not found' };
  sh.getRange(idx, 10).setValue(req.status || '');
  sh.getRange(idx, 17).setValue(req.propTo || '');
  sh.getRange(idx, 18).setValue(req.propUntil || '');
  sh.getRange(idx, 20).setValue(new Date());
  return { ok: true };
}

function savePhoto_(mg, dataUrl) {
  var folder = getFolder_();
  var m = dataUrl.match(/^data:(.+?);base64,(.*)$/);
  var contentType = m[1];
  var bytes = Utilities.base64Decode(m[2]);
  var blob = Utilities.newBlob(bytes, contentType, mg + '_' + Date.now() + '.jpg');
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getId();
}

function getFolder_() {
  var it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
