/**
 * Pile Line 2026 — Google Apps Script Backend
 * ใช้คู่กับ index.html (Sales Pipeline Tracker)
 *
 * วิธีติดตั้ง:
 *  1) สร้าง Google Sheet ใหม่ 1 ไฟล์ (ปล่อยว่างได้)
 *  2) เมนู ส่วนขยาย (Extensions) › Apps Script
 *  3) ลบโค้ดเดิมทั้งหมด แล้ววางโค้ดนี้ทั้งไฟล์ › บันทึก (Ctrl+S)
 *  4) กด Deploy › New deployment › เลือก type = Web app
 *  5) Execute as: Me   |   Who has access: Anyone
 *  6) Deploy › อนุญาตสิทธิ์ › คัดลอก Web app URL (.../exec)
 *  7) นำ URL ไปวางในแอป (ปุ่ม "เชื่อมต่อ Sheet")
 *
 * เมื่อแก้โค้ดภายหลัง ให้ Deploy › Manage deployments › แก้ไข version เป็น New
 */

var SHEET_NAME = 'PileLine';
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
  sh.clearContents();
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  if (rows.length) {
    var rng = sh.getRange(2, 1, rows.length, HEADERS.length);
    rng.setNumberFormat('@');   // เก็บทุกคอลัมน์เป็นข้อความ กัน Sheets แปลง "2026-06-18" เป็นวันที่/ISO อัตโนมัติ
    rng.setValues(rows);
  }
  return data ? data.length : 0;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// GET  ?action=load  -> ส่งข้อมูลทั้งหมดกลับ
function doGet(e) {
  try {
    return json_(readAll_());
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// POST {action:'save', data:[...]}  -> เขียนทับทั้งชีต
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    if (body.action === 'save') {
      var n = writeAll_(body.data || []);
      return json_({ ok: true, saved: n });
    }
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
