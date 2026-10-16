const SPREADSHEET_ID = '1KIda5u-xHsYASi-9o5uWo20utckkQkZIJx26dIZ2F7k';
const SHEET_NAME = 'Sheet1';
const LOGO_URL = 'https://drive.google.com/uc?export=view&id=1RYqjqw5bW-nqpgvcMn3Cdhly5QTwnJyK';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ตารางกำหนดการปฏิบัติงาน กอ.รมน.จังหวัด น.ศ.')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSheet_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('ไม่พบแท็บ ' + SHEET_NAME);
  return sheet;
}

function getAppData() {
  const sheet = getSheet_();
  const n = sheet.getLastRow();
  const rows = n > 1 ? sheet.getRange(2, 1, n - 1, 8).getDisplayValues() : [];
  const items = rows.map((r, i) => ({
    rowNumber: i + 2, date: normalizeDate_(r[0]), time: normalizeTime_(r[1]),
    mission: r[2] || '', chair: r[3] || '', location: r[4] || '',
    unit: r[5] || '', operators: r[6] || '', attire: r[7] || ''
  })).filter(x => x.date || x.mission || x.time);
  const dates = [...new Set(items.map(x => x.date).filter(Boolean))].sort();
  let logoDataUrl = '';
  try {
    const blob = DriveApp.getFileById('1RYqjqw5bW-nqpgvcMn3Cdhly5QTwnJyK').getBlob();
    const mime = blob.getContentType() || 'image/png';
    logoDataUrl = 'data:' + mime + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (error) {
    logoDataUrl = '';
  }
  return {items, dates, latestDate: dates.length ? dates[dates.length - 1] : today_(), logoDataUrl};
}

function saveScheduleItem(item) {
  validate_(item);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    const values = [normalizeDate_(item.date), normalizeTime_(item.time), clean_(item.mission),
      clean_(item.chair), clean_(item.location), clean_(item.unit), clean_(item.operators), clean_(item.attire)];
    const row = Number(item.rowNumber || 0);
    if (row >= 2 && row <= sheet.getLastRow()) {
      const current = sheet.getRange(row, 1, 1, 8).getDisplayValues()[0];
      if (item.expected && !sameRow_(current, item.expected)) throw new Error('ข้อมูลรายการเปลี่ยนแปลงแล้ว กรุณาโหลดใหม่ก่อนบันทึก');
      sheet.getRange(row, 1, 1, 8).setNumberFormat('@');
      sheet.getRange(row, 1, 1, 8).setValues([values]);
      return {success:true, message:'แก้ไขภารกิจเรียบร้อยแล้ว'};
    }
    const target = Math.max(2, sheet.getLastRow() + 1);
    sheet.getRange(target, 1, 1, 8).setNumberFormat('@');
    sheet.getRange(target, 1, 1, 8).setValues([values]);
    return {success:true, message:'เพิ่มภารกิจเรียบร้อยแล้ว'};
  } finally { lock.releaseLock(); }
}

function deleteScheduleItem(rowNumber, expected) {
  const row = Number(rowNumber);
  if (!Number.isInteger(row) || row < 2) throw new Error('หมายเลขแถวไม่ถูกต้อง');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    if (row > sheet.getLastRow()) throw new Error('ไม่พบรายการนี้');
    const current = sheet.getRange(row, 1, 1, 8).getDisplayValues()[0];
    if (expected && !sameRow_(current, expected)) throw new Error('รายการเปลี่ยนแปลงแล้ว กรุณาโหลดข้อมูลใหม่ก่อนลบ');
    sheet.deleteRow(row);
    return {success:true, message:'ลบภารกิจเรียบร้อยแล้ว'};
  } finally { lock.releaseLock(); }
}

function validate_(x) {
  if (!x || !/^\d{4}-\d{2}-\d{2}$/.test(String(x.date || ''))) throw new Error('กรุณาระบุวันที่ให้ถูกต้อง');
  if (x.time && !/^\d{4}$/.test(String(x.time))) throw new Error('กรุณากรอกเวลา 4 หลัก เช่น 0830');
  if (!clean_(x.mission)) throw new Error('กรุณาระบุภารกิจ');
}
function sameRow_(r, x) {
  const a = r.map((v,i) => i===0 ? normalizeDate_(v) : i===1 ? normalizeTime_(v) : clean_(v));
  const b = [normalizeDate_(x.date), normalizeTime_(x.time), clean_(x.mission), clean_(x.chair),
    clean_(x.location), clean_(x.unit), clean_(x.operators), clean_(x.attire)];
  return a.every((v,i) => v===b[i]);
}
function clean_(v) { return String(v == null ? '' : v).trim(); }
function normalizeDate_(v) {
  const s = clean_(v); if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) { let y=+m[3]; if(y>2400)y-=543; return y+'-'+String(m[2]).padStart(2,'0')+'-'+String(m[1]).padStart(2,'0'); }
  return s;
}
function normalizeTime_(v) {
  let s=clean_(v); if (!s) return '';
  if (/^\d{1,2}:\d{2}$/.test(s)) s=s.replace(':','');
  return /^\d{1,4}$/.test(s) ? s.padStart(4,'0') : s;
}
function today_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
