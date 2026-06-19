/**
 * CULAC Invoice System - Google Apps Script Backend
 * ตรงกับ index.html (เวอร์ชันแก้เลขรัน + ปรับ UI)
 *
 * วิธีใช้:
 * 1) เปิด Google Sheet ที่ใช้เก็บข้อมูล (หรือสร้างใหม่)
 * 2) Extensions > Apps Script > วางโค้ดนี้ทับของเดิม
 * 3) Deploy > New deployment > Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4) คัดลอก URL (/exec) ไปวางในหน้า "ตั้งค่า" ของระบบ
 *
 * หมายเหตุ: ระบบจะสร้างชีต "Invoices" พร้อมหัวคอลัมน์ให้อัตโนมัติถ้ายังไม่มี
 *           อ่าน/เขียนข้อมูลตาม "ชื่อหัวคอลัมน์" ไม่ใช่ตำแหน่ง จึงทนต่อการสลับคอลัมน์
 */

var SHEET_NAME = 'Invoices';
var HEADERS = [
  'ID', 'InvoiceNo', 'Date', 'Customer', 'Items', 'Total', 'Note',
  'BankName', 'BankBranch', 'BankAccount', 'BankUser',
  'PaymentStatus', 'PaymentDate', 'PaymentDocNo', 'CreatedAt'
];

// ---------- Routing ----------
function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var action = p.action || 'test';
  try {
    switch (action) {
      case 'test':          return json({ success: true, message: 'Connection successful!', version: '3.0' });
      case 'getNextNumber': return json({ success: true, nextNumber: getNextNumber_(p.month) });
      case 'list':          return json({ success: true, invoices: listInvoices_() });
      case 'get':           return json(getInvoice_(p.id));
      default:              return json({ success: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return json({ success: false, error: String(err) });
  }
}

function doPost(e) {
  var data = {};
  try {
    if (e && e.postData && e.postData.contents) data = JSON.parse(e.postData.contents);
  } catch (err) {
    return json({ success: false, error: 'Invalid JSON: ' + String(err) });
  }
  var action = data.action || '';
  try {
    switch (action) {
      case 'create': return json(createInvoice_(data));
      case 'delete': return json(deleteInvoice_(data.id));
      case 'pay':    return json(payInvoice_(data));
      default:       return json({ success: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return json({ success: false, error: String(err) });
  }
}

// ---------- Sheet helpers ----------
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
  }
  ensureHeaders_(sh);
  return sh;
}

// เพิ่มคอลัมน์ที่จำเป็นต่อท้าย ถ้าชีตเดิมยังไม่มี (เช่น PaymentStatus/PaymentDate/PaymentDocNo)
function ensureHeaders_(sh) {
  var headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  var existing = {};
  for (var i = 0; i < headers.length; i++) existing[String(headers[i]).trim()] = true;
  var missing = [];
  for (var j = 0; j < HEADERS.length; j++) {
    if (!existing[HEADERS[j]]) missing.push(HEADERS[j]);
  }
  if (missing.length > 0) {
    sh.getRange(1, sh.getLastColumn() + 1, 1, missing.length).setValues([missing]);
  }
}

function getHeaderMap_(sh) {
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) map[String(headers[i]).trim()] = i;
  return map;
}

// แปลงทุกแถวเป็น object ที่ฝั่ง HTML ใช้ (key เป็น camelCase)
function listInvoices_() {
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var hmap = getHeaderMap_(sh);
  var values = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  var out = [];
  for (var r = 0; r < values.length; r++) {
    out.push(rowToObject_(values[r], hmap));
  }
  out.reverse(); // ใบล่าสุดอยู่บนสุด (invoices[0])
  return out;
}

function rowToObject_(row, hmap) {
  function v(name) { return hmap[name] != null ? row[hmap[name]] : ''; }
  return {
    id:            v('ID'),
    invoiceNo:     v('InvoiceNo'),
    date:          v('Date'),
    customer:      v('Customer'),
    items:         v('Items'),         // เก็บเป็น JSON string
    total:         v('Total'),
    note:          v('Note'),
    bankName:      v('BankName'),
    bankBranch:    v('BankBranch'),
    bankAccount:   v('BankAccount'),
    bankUser:      v('BankUser'),
    paymentStatus: v('PaymentStatus'),
    paymentDate:   v('PaymentDate'),
    paymentDocNo:  v('PaymentDocNo')
  };
}

// ---------- Actions ----------
function getInvoice_(id) {
  if (!id) return { success: false, error: 'Missing id' };
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 2) return { success: false, error: 'Not found' };
  var hmap = getHeaderMap_(sh);
  var values = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  for (var r = 0; r < values.length; r++) {
    if (String(values[r][hmap['ID']]) === String(id)) {
      return { success: true, invoice: rowToObject_(values[r], hmap) };
    }
  }
  return { success: false, error: 'Not found' };
}

function createInvoice_(d) {
  var sh = getSheet_();
  var hmap = getHeaderMap_(sh);
  var id = 'INV' + new Date().getTime();
  var itemsStr = (typeof d.items === 'string') ? d.items : JSON.stringify(d.items || []);

  var record = {
    'ID': id,
    'InvoiceNo': d.invoiceNo || '',
    'Date': d.date || '',
    'Customer': d.customer || '',
    'Items': itemsStr,
    'Total': d.total || 0,
    'Note': d.note || '',
    'BankName': d.bankName || '',
    'BankBranch': d.bankBranch || '',
    // ฝั่ง HTML ส่งมาเป็น bankAcc แต่ตอนอ่านใช้ bankAccount -> รองรับทั้งคู่
    'BankAccount': d.bankAccount || d.bankAcc || '',
    'BankUser': d.bankUser || '',
    'PaymentStatus': 'unpaid',
    'PaymentDate': '',
    'PaymentDocNo': '',
    'CreatedAt': new Date()
  };

  var width = sh.getLastColumn();
  var row = new Array(width).fill('');
  for (var key in record) {
    if (hmap[key] != null) row[hmap[key]] = record[key];
  }
  sh.appendRow(row);

  // บังคับให้ช่องเลขที่เป็น "ข้อความ" เพื่อรักษาศูนย์นำหน้า เช่น 030669 ไม่ให้กลายเป็น 30669
  if (hmap['InvoiceNo'] != null) {
    sh.getRange(sh.getLastRow(), hmap['InvoiceNo'] + 1)
      .setNumberFormat('@')
      .setValue(record['InvoiceNo']);
  }
  return { success: true, id: id, invoiceNo: record['InvoiceNo'] };
}

function deleteInvoice_(id) {
  if (!id) return { success: false, error: 'Missing id' };
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 2) return { success: false, error: 'Not found' };
  var hmap = getHeaderMap_(sh);
  var ids = sh.getRange(2, hmap['ID'] + 1, last - 1, 1).getValues();
  for (var r = 0; r < ids.length; r++) {
    if (String(ids[r][0]) === String(id)) {
      sh.deleteRow(r + 2);
      return { success: true };
    }
  }
  return { success: false, error: 'Not found' };
}

function payInvoice_(d) {
  if (!d.id) return { success: false, error: 'Missing id' };
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 2) return { success: false, error: 'Not found' };
  var hmap = getHeaderMap_(sh);
  var ids = sh.getRange(2, hmap['ID'] + 1, last - 1, 1).getValues();
  for (var r = 0; r < ids.length; r++) {
    if (String(ids[r][0]) === String(d.id)) {
      var rowNum = r + 2;
      if (hmap['PaymentStatus'] != null) sh.getRange(rowNum, hmap['PaymentStatus'] + 1).setValue('paid');
      if (hmap['PaymentDate'] != null)   sh.getRange(rowNum, hmap['PaymentDate'] + 1).setValue(d.paymentDate || '');
      if (hmap['PaymentDocNo'] != null)  sh.getRange(rowNum, hmap['PaymentDocNo'] + 1).setValue(d.paymentDocNo || '');
      return { success: true };
    }
  }
  return { success: false, error: 'Not found' };
}

/**
 * นับเลขรันถัดไปของ "เดือนที่ระบุ" จากข้อมูลจริงในชีต (รันแยกอิสระแต่ละเดือน)
 * รูปแบบเลข: SSMMYY  (SS=ลำดับ, MM=เดือน, YY=ปี พ.ศ. 2 หลักท้าย)
 * @param {string} monthParam  MMYY ของเดือนที่ต้องการ เช่น "0669" (ถ้าไม่ส่งมา ใช้เดือนปัจจุบัน)
 * คืนค่าลำดับถัดไป (จำนวนเต็ม) เช่น เดือนนั้นมีสูงสุด 03 -> คืน 4
 */
function getNextNumber_(monthParam) {
  var suffix;
  if (monthParam && /^\d{4}$/.test(String(monthParam))) {
    suffix = String(monthParam);          // ใช้เดือนที่ฝั่งหน้าเว็บส่งมา (อิงวันที่ในเอกสาร)
  } else {
    var now = new Date();
    var mm = ('0' + (now.getMonth() + 1)).slice(-2);
    var yy = String(now.getFullYear() + 543).slice(-2);
    suffix = mm + yy;                      // เดือนปัจจุบัน
  }

  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 2) return 1;

  var hmap = getHeaderMap_(sh);
  var nos = sh.getRange(2, hmap['InvoiceNo'] + 1, last - 1, 1).getValues();
  var maxSeq = 0;
  for (var r = 0; r < nos.length; r++) {
    var no = String(nos[r][0]).trim();
    // เทียบเฉพาะใบของเดือนนั้น (4 หลักท้ายตรงกัน) รองรับทั้งเลขที่มี/ไม่มีศูนย์นำหน้า
    if (no.length >= 5 && no.slice(-4) === suffix) {
      var seq = parseInt(no.slice(0, no.length - 4), 10);
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    }
  }
  return maxSeq + 1;
}

// ---------- Output ----------
function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ===== เครื่องมือแปลงเลขเก่า (รันครั้งเดียว) =====
 * วิธีใช้: ในหน้า Apps Script เลือกฟังก์ชัน "fixOldInvoiceNumbers" แล้วกด ▶ Run
 * จะแปลงเลขที่ในคอลัมน์ InvoiceNo ให้เป็น 6 หลักเสมอ (เติมศูนย์นำหน้า) และเก็บเป็นข้อความ
 * เช่น 30669 -> 030669, 10669 -> 010669, 20669 -> 020669 (100669 ที่ครบ 6 หลักอยู่แล้วไม่เปลี่ยน)
 */
function fixOldInvoiceNumbers() {
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 2) {
    Logger.log('ไม่มีข้อมูลให้แปลง');
    return;
  }
  var hmap = getHeaderMap_(sh);
  var col = hmap['InvoiceNo'] + 1;            // คอลัมน์ InvoiceNo (1-indexed)
  var range = sh.getRange(2, col, last - 1, 1);
  var values = range.getValues();
  var fixed = 0;

  for (var i = 0; i < values.length; i++) {
    var raw = String(values[i][0]).trim();
    if (raw === '') continue;                 // ข้ามช่องว่าง
    var digits = raw.replace(/\D/g, '');       // เอาเฉพาะตัวเลข
    if (digits === '') continue;
    var padded = digits.padStart(6, '0');     // เติมศูนย์ให้ครบ 6 หลัก
    if (padded !== raw) fixed++;
    values[i][0] = padded;
  }

  range.setNumberFormat('@');                 // บังคับคอลัมน์เป็นข้อความ กันศูนย์หายอีก
  range.setValues(values);
  Logger.log('แปลงเลขเสร็จแล้ว: ปรับ ' + fixed + ' รายการ จากทั้งหมด ' + values.length + ' รายการ');
}
