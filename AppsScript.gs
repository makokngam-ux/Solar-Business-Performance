/**
 * Pile Line 2026 — Google Apps Script Backend
 * ใช้คู่กับ index.html (Sales Pipeline Tracker)
 *
 * วิธีติดตั้ง:
 *  1) สร้าง Google Sheet ใหม่ 1 ไฟล์ (ปล่อยว่างได้)
 *  2) เมนู ส่วนขยาย (Extensions) › Apps Script
 *  3) ลบโค้ดเดิมทั้งหมด แล้ววางโค้ดนี้ทั้งไฟล์ › บันทึก (Ctrl+S)
 *  3.1) ตั้งรหัสซิงก์ที่ตัวแปร TOKEN ด้านล่าง (เช่น คำยาว ๆ ที่เดายาก) แล้วบันทึก
 *       จากนั้นใส่รหัสเดียวกันในแอป หน้าตั้งค่า ช่อง "รหัสซิงก์"
 *  4) กด Deploy › New deployment › เลือก type = Web app
 *  5) Execute as: Me   |   Who has access: Anyone
 *  6) Deploy › อนุญาตสิทธิ์ › คัดลอก Web app URL (.../exec)
 *  7) นำ URL ไปวางในแอป (ปุ่ม "เชื่อมต่อ Sheet")
 *
 * เมื่อแก้โค้ดภายหลัง ให้ Deploy › Manage deployments › แก้ไข version เป็น New
 *
 * การป้องกันข้อมูล:
 *  - TOKEN: ถ้าตั้งไว้ ทุกคำขออ่าน/เขียนต้องส่งรหัสนี้มาด้วย ไม่งั้นถูกปฏิเสธ
 *  - ก่อนเขียนทับครั้งแรกของแต่ละวัน จะคัดลอกชีตเดิมเก็บเป็นชีตซ่อนชื่อ Backup_<ชื่อชีต>_<วันที่>
 *    เก็บย้อนหลัง BACKUP_KEEP วัน (กู้คืนได้โดย คลิกขวาที่แถบชีต › Show hidden sheets)
 *  - ไม่ยอมให้ส่งรายการว่างมาล้างชีตที่มีข้อมูล เว้นแต่แอปยืนยันมา (allowEmpty)
 *  - ใช้ LockService กันคำขอเขียนซ้อนกัน
 */

var TOKEN = '';            // รหัสซิงก์ — เว้นว่าง '' = ไม่ตรวจรหัส (ไม่แนะนำ)
var BACKUP_PREFIX = 'Backup_';
var BACKUP_KEEP = 7;       // จำนวนวันที่เก็บชีตสำรองของแต่ละชีต

var SHEET_NAME = 'PileLine';
var APPDATA_SHEET_NAME = 'AppData';
var HEADERS = ['id','name','projName','customer','contact','type','stage','kwp','kwh',
               'value','prob','product','close','followUp','next','notes','locked','review',
               'piRef','piFull','piStatus','piDate','piRev','piUpdated',
               'custTel','custEmail','custTaxId','custAddress','custType','custSale',
               'lineItems','created','updated','activities'];

// คอลัมน์ที่เก็บเป็น JSON (array) ต้อง parse/stringify เวลาอ่าน-เขียน
function isJsonCol_(key){ return key === 'activities' || key === 'lineItems'; }

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readAll_() {
  var sh = getSheet_();
  var rng = sh.getDataRange().getValues();
  if (rng.length < 2) return [];
  var head = rng[0];
  var out = [];
  for (var i = 1; i < rng.length; i++) {
    var row = rng[i];
    if (!row[0] && !row[1]) continue; // skip blank
    var obj = {};
    for (var c = 0; c < head.length; c++) {
      var key = head[c];
      var val = row[c];
      if (isJsonCol_(key)) {
        try { val = val ? JSON.parse(val) : []; } catch (e) { val = []; }
      }
      obj[key] = val;
    }
    out.push(obj);
  }
  return out;
}

function writeAll_(data) {
  var sh = getSheet_();
  var rows = [];
  if (data && data.length) {
    rows = data.map(function (p) {
      return HEADERS.map(function (h) {
        if (isJsonCol_(h)) return JSON.stringify(p[h] || []);
        return p[h] != null ? p[h] : '';
      });
    });
  }
  backupOncePerDay_(sh);
  // เขียนข้อมูลใหม่ทับก่อน แล้วค่อยล้างส่วนเกิน — ถ้าเขียนล้มเหลว ข้อมูลเดิมยังอยู่ ไม่ได้ถูกล้างไปก่อน
  var oldLastRow = sh.getLastRow(), oldLastCol = sh.getLastColumn();
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  if (rows.length) {
    var rng = sh.getRange(2, 1, rows.length, HEADERS.length);
    rng.setNumberFormat('@');   // เก็บทุกคอลัมน์เป็นข้อความ กัน Sheets แปลง "2026-06-18" เป็นวันที่/ISO อัตโนมัติ
    rng.setValues(rows);
  }
  clearLeftover_(sh, rows.length + 1, HEADERS.length, oldLastRow, oldLastCol);
  return data ? data.length : 0;
}

// ล้างแถว/คอลัมน์เก่าที่เกินจากข้อมูลชุดใหม่ (usedRows นับรวมแถวหัวตาราง)
function clearLeftover_(sh, usedRows, usedCols, oldLastRow, oldLastCol) {
  if (oldLastRow > usedRows) {
    sh.getRange(usedRows + 1, 1, oldLastRow - usedRows, Math.max(oldLastCol, usedCols)).clearContent();
  }
  if (oldLastCol > usedCols) {
    sh.getRange(1, usedCols + 1, Math.max(oldLastRow, usedRows), oldLastCol - usedCols).clearContent();
  }
}

// สำรองชีตก่อนเขียนทับ วันละ 1 ครั้ง (ครั้งแรกของวัน = สภาพก่อนการแก้ไขของวันนั้น)
function backupOncePerDay_(sh) {
  try {
    if (sh.getLastRow() < 2) return;   // ไม่มีข้อมูลให้สำรอง
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var base = BACKUP_PREFIX + sh.getName() + '_';
    var name = base + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    if (ss.getSheetByName(name)) return;
    var copy = sh.copyTo(ss);
    copy.setName(name);
    copy.hideSheet();
    // ลบชีตสำรองที่เก่าเกิน BACKUP_KEEP วัน (เฉพาะชีตที่สคริปต์นี้สร้างเอง ดูจากชื่อ)
    var olds = ss.getSheets().filter(function (s) { return s.getName().indexOf(base) === 0; })
      .sort(function (a, b) { return a.getName() < b.getName() ? -1 : 1; });
    for (var i = 0; i < olds.length - BACKUP_KEEP; i++) ss.deleteSheet(olds[i]);
  } catch (err) {
    // สำรองไม่สำเร็จไม่ควรทำให้การบันทึกหลักล้ม
  }
}

function authorized_(token) {
  return !TOKEN || String(token || '') === TOKEN;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getAppDataSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(APPDATA_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(APPDATA_SHEET_NAME);
    sh.getRange(1, 1, 1, 3).setValues([['key','json','updatedAt']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readAppData_() {
  var sh = getAppDataSheet_();
  var rng = sh.getDataRange().getValues();
  var out = {};
  for (var i = 1; i < rng.length; i++) {
    var key = rng[i][0];
    if (!key) continue;
    try { out[key] = rng[i][1] ? JSON.parse(rng[i][1]) : null; }
    catch (e) { out[key] = null; }
  }
  return out;
}

function writeAppData_(data) {
  var sh = getAppDataSheet_();
  var keys = ['arData','arProjects','arReceiveDates','arNotes','arPromiseDates','arCollectStatus','claims'];
  var rows = keys.map(function (k) {
    return [k, JSON.stringify(data && data[k] != null ? data[k] : (k === 'arProjects' || k === 'arReceiveDates' || k === 'arNotes' || k === 'arPromiseDates' || k === 'arCollectStatus' ? {} : null)), new Date().toISOString()];
  });
  backupOncePerDay_(sh);
  var oldLastRow = sh.getLastRow(), oldLastCol = sh.getLastColumn();
  sh.getRange(1, 1, 1, 3).setValues([['key','json','updatedAt']]);
  sh.getRange(2, 1, rows.length, 3).setNumberFormat('@').setValues(rows);
  clearLeftover_(sh, rows.length + 1, 3, oldLastRow, oldLastCol);
  return rows.length;
}

// GET  ?action=load  -> ส่งข้อมูลทั้งหมดกลับ
function doGet(e) {
  try {
    if (!authorized_(e && e.parameter && e.parameter.token)) return json_({ ok: false, error: 'unauthorized' });
    if (e && e.parameter && e.parameter.action === 'loadAppData') {
      return json_({ ok: true, data: readAppData_() });
    }
    return json_(readAll_());
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// POST {action:'save', data:[...]}  -> เขียนทับทั้งชีต
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    if (!authorized_(body.token)) return json_({ ok: false, error: 'unauthorized' });
    if (body.action !== 'save' && body.action !== 'saveAppData') return json_({ ok: false, error: 'unknown action' });
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);   // รอคิวเขียน กันสองเครื่องเขียนซ้อนกันจนข้อมูลปนกัน
    try {
      if (body.action === 'save') {
        if (!Array.isArray(body.data)) return json_({ ok: false, error: 'bad data' });
        // กันเครื่องที่ข้อมูลว่าง (เช่น เพิ่งติดตั้ง) ส่งมาล้างชีตที่มีข้อมูลอยู่
        if (!body.data.length && !body.allowEmpty && getSheet_().getLastRow() > 1) {
          return json_({ ok: false, error: 'refuse_empty' });
        }
        return json_({ ok: true, saved: writeAll_(body.data) });
      }
      if (!body.data || typeof body.data !== 'object') return json_({ ok: false, error: 'bad data' });
      return json_({ ok: true, saved: writeAppData_(body.data) });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
