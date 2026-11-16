/**
 * ปฏิทินลงเวรส่วนตัว — Google Apps Script Web App
 *
 * ลงเวรเองรายวัน: บด 1–4, ชบด 1–4, ช 1–4, E 1–4, M 1–4, ออกหน่วย หรือ "ลา" (ไม่ทำงาน)
 * (ตัวเลข 1–4 = จุดที่อยู่เวร ไม่มีผลกับเงิน)
 *
 * การคิดเงิน (ค่าเริ่มต้น แก้ได้ในหน้าเว็บ):
 *   วันทำการ จ–ศ (ไม่ใช่วันหยุดราชการ และไม่ได้กด "ลา")  = 1,200 / วัน
 *   + ค่าเวร: บด = 1,200 × 2 · ชบด = 1,200 × 3 · ช 1,200 · E 800 · M 400 · ออกหน่วย 1,200
 *   หักภาษี 10% และประกันสังคม 850 บาท/เดือน
 *
 * วันหยุดราชการ: ดึงจากปฏิทินวันหยุดไทยของ Google (ปีนี้ + ปีหน้า)
 *   - รัน setupMonthlyHolidaySync() ครั้งเดียวจาก editor → อัปเดตอัตโนมัติทุกวันที่ 1 ของเดือน
 *   - หรือกดปุ่ม "อัปเดตวันหยุด" บนหน้าเว็บ
 */

// ถ้าสคริปต์ผูกกับ Sheet (ส่วนขยาย > Apps Script) ปล่อยว่างไว้
const SS_ID = '';
const TZ = 'Asia/Bangkok';

const HOLIDAY_ICS_TH = 'https://calendar.google.com/calendar/ical/th.th%23holiday%40group.v.calendar.google.com/public/basic.ics';
const HOLIDAY_ICS_EN = 'https://calendar.google.com/calendar/ical/en.th%23holiday%40group.v.calendar.google.com/public/basic.ics';

const SHEETS = {
  Shifts:   ['date', 'code', 'note', 'updated'],
  Holidays: ['date', 'name', 'source'],
  Log:      ['timestamp', 'date', 'action', 'detail', 'by'],
  Settings: ['key', 'value', 'label']
};

// key ในหน้าเว็บ → [key ในชีต, ค่าเริ่มต้น, คำอธิบาย]
const RATES = {
  workday: ['rate_workday', 1200, 'วันทำการ จ–ศ (บาท/วัน)'],
  shift:   ['rate_shift',   1200, 'ค่าเวรต่อ 1 เวร (บาท)'],
  bdN:     ['bd_shifts',       2, 'บด = กี่เวร'],
  cbdN:    ['cbd_shifts',      3, 'ชบด = กี่เวร'],
  ch:      ['rate_ch',      1200, 'ช (บาท)'],
  unit:    ['rate_unit',    1200, 'ออกหน่วย (บาท)'],
  e:       ['rate_e',        800, 'E (บาท)'],
  m:       ['rate_m',        400, 'M (บาท)'],
  tax:     ['tax_pct',        10, 'หักภาษี (%)'],
  sso:     ['sso',           850, 'ประกันสังคม (บาท/เดือน)']
};

const CODE_RE = /^(บด|ชบด|ช|E|M)[1-4]$|^ลา$|^ออกหน่วย$/;

/* ---------- Web entry ---------- */

function doGet() {
  setup_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ปฏิทินเวรของฉัน')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- Setup ---------- */

function ss_() { return SS_ID ? SpreadsheetApp.openById(SS_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
function sh_(name) { return ss_().getSheetByName(name); }

function setup_() {
  const ss = ss_();
  Object.keys(SHEETS).forEach(function (n) {
    const h = SHEETS[n];
    let sh = ss.getSheetByName(n);
    // หัวตารางไม่ตรง (เวอร์ชันเก่า) → เก็บไว้เป็น _old แล้วสร้างใหม่
    if (sh && sh.getLastColumn() > 0) {
      const cur = sh.getRange(1, 1, 1, h.length).getValues()[0];
      if (cur.join('|') !== h.join('|')) {
        sh.setName(n + '_old_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmm'));
        sh = null;
      }
    }
    if (sh) return;
    sh = ss.insertSheet(n);
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
    sh.setFrozenRows(1);
    if (n !== 'Settings') sh.getRange('A:B').setNumberFormat('@');
  });
}

/** รันครั้งเดียวจาก editor: ดึงวันหยุดอัตโนมัติทุกวันที่ 1 ของเดือน ตี 3 */
function setupMonthlyHolidaySync() {
  setup_();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncHolidays') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncHolidays').timeBased().onMonthDay(1).atHour(3).inTimezone(TZ).create();
  Logger.log('ตั้งอัปเดตรายเดือนแล้ว · ' + syncHolidays().message);
}

/* ---------- Helpers ---------- */

function toKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v).trim();
}

function rows_(name) {
  const sh = sh_(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, SHEETS[name].length).getValues();
}

function readShifts_() {
  const o = {};
  rows_('Shifts').forEach(function (r) {
    if (r[0] && r[1]) o[toKey_(r[0])] = { code: String(r[1]), note: String(r[2] || '') };
  });
  return o;
}

function writeShifts_(map) {
  const sh = sh_('Shifts');
  const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
  const data = Object.keys(map).sort().map(function (k) {
    return [k, map[k].code, map[k].note || '', map[k].updated || now];
  });
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, SHEETS.Shifts.length).clearContent();
  sh.getRange('A:B').setNumberFormat('@');
  if (data.length) sh.getRange(2, 1, data.length, SHEETS.Shifts.length).setValues(data);
}

function readHolidays_() {
  const o = {};
  rows_('Holidays').forEach(function (r) {
    if (r[0]) o[toKey_(r[0])] = { name: String(r[1] || 'วันหยุด'), source: String(r[2] || 'manual') };
  });
  return o;
}

function writeHolidays_(map) {
  const sh = sh_('Holidays');
  const data = Object.keys(map).sort().map(function (k) { return [k, map[k].name, map[k].source]; });
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 3).clearContent();
  sh.getRange('A:A').setNumberFormat('@');
  if (data.length) sh.getRange(2, 1, data.length, 3).setValues(data);
}

function getSetting_(key) {
  const row = rows_('Settings').filter(function (r) { return r[0] === key; })[0];
  return row ? row[1] : '';
}

function setSetting_(key, value, label) {
  const sh = sh_('Settings');
  const rows = rows_('Settings');
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === key) { sh.getRange(i + 2, 2).setValue(value); return; }
  }
  sh.appendRow([key, value, label || key]);
}

function readRates_() {
  const r = {};
  Object.keys(RATES).forEach(function (k) {
    const v = getSetting_(RATES[k][0]);
    r[k] = (v === '' || v === null) ? RATES[k][1] : Number(v) || 0;
  });
  return r;
}

function log_(date, action, detail) {
  let by = '';
  try { by = Session.getActiveUser().getEmail(); } catch (e) {}
  sh_('Log').appendRow([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), date, action, detail || '', by]);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------- ดึงวันหยุดราชการ ---------- */

function ymdToDate_(s) { return new Date(+s.slice(0, 4), +s.slice(4, 6) - 1, +s.slice(6, 8)); }

/** แยก event จากไฟล์ .ics → [{date, name, desc}] */
function parseIcs_(text) {
  const unfolded = text.replace(/\r?\n[ \t]/g, '');
  const out = [];
  unfolded.split('BEGIN:VEVENT').slice(1).forEach(function (blk) {
    const ds = blk.match(/DTSTART(?:;[^:\n]*)?:(\d{8})/);
    const de = blk.match(/DTEND(?:;[^:\n]*)?:(\d{8})/);
    const su = blk.match(/\nSUMMARY(?:;[^:\n]*)?:(.*)/);
    const dc = blk.match(/\nDESCRIPTION(?:;[^:\n]*)?:(.*)/);
    if (!ds) return;
    const name = su ? su[1].replace(/\\,/g, ',').replace(/\\n/g, ' ').trim() : 'วันหยุด';
    const desc = dc ? dc[1].replace(/\\n/g, ' ') : '';
    let d = ymdToDate_(ds[1]);
    const end = de ? ymdToDate_(de[1]) : new Date(d.getTime() + 86400000);
    while (d < end) {
      out.push({ date: Utilities.formatDate(d, TZ, 'yyyy-MM-dd'), name: name, desc: desc });
      d = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1);
    }
  });
  return out;
}

function fetchIcs_(url) {
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('ดึงปฏิทินวันหยุดไม่ได้ (HTTP ' + res.getResponseCode() + ')');
  return parseIcs_(res.getContentText('UTF-8'));
}

/**
 * ดึงวันหยุดราชการของปีนี้และปีหน้า
 * - ปฏิทินภาษาอังกฤษใช้แยก "Public holiday" ออกจาก "Observance" (วันสำคัญที่ไม่ได้หยุด)
 * - ชื่อภาษาไทยเอามาจากปฏิทินภาษาไทย
 * - วันหยุดที่ตั้งเอง (manual) และที่ยกเลิกเอง (removed) จะไม่ถูกทับ
 */
function syncHolidays() {
  setup_();
  return withLock_(function () {
    const y = new Date().getFullYear();
    const years = [String(y), String(y + 1)];
    const inYears = function (k) { return years.indexOf(k.slice(0, 4)) >= 0; };

    const en = fetchIcs_(HOLIDAY_ICS_EN).filter(function (e) { return inYears(e.date); });
    let th = [];
    try { th = fetchIcs_(HOLIDAY_ICS_TH).filter(function (e) { return inYears(e.date); }); } catch (e) {}

    let pub = en.filter(function (e) { return /public holiday/i.test(e.desc); });
    if (!pub.length) pub = en;

    const thName = {};
    th.forEach(function (e) { thName[e.date] = thName[e.date] ? thName[e.date] + ' / ' + e.name : e.name; });

    const map = readHolidays_();
    Object.keys(map).forEach(function (k) { if (map[k].source === 'auto' && inYears(k)) delete map[k]; });

    let added = 0;
    pub.forEach(function (e) {
      if (map[e.date] && map[e.date].source !== 'auto') return;
      map[e.date] = { name: thName[e.date] || e.name, source: 'auto' };
      added++;
    });
    writeHolidays_(map);

    setSetting_('holiday_synced', Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), 'อัปเดตวันหยุดล่าสุด');
    log_(years.join('-'), 'อัปเดตวันหยุด', added + ' วัน');
    return { message: 'อัปเดตวันหยุดแล้ว ' + added + ' วัน (' + years.map(function (x) { return +x + 543; }).join(', ') + ')' };
  });
}

/* ---------- API (เรียกจากหน้าเว็บ) ---------- */

function getAll() {
  if (!getSetting_('holiday_synced')) { try { syncHolidays(); } catch (e) {} }
  const shifts = readShifts_();
  const hol = readHolidays_();
  const holOut = {};
  Object.keys(hol).forEach(function (k) { if (hol[k].source !== 'removed') holOut[k] = hol[k]; });
  return {
    shifts: shifts,
    holidays: holOut,
    rates: readRates_(),
    synced: String(getSetting_('holiday_synced') || '')
  };
}

function syncHolidaysNow() {
  const r = syncHolidays();
  const out = getAll();
  out.message = r.message;
  return out;
}

/**
 * ลงเวรวันนั้น
 * @param {string} date yyyy-MM-dd
 * @param {string} code เช่น 'บด1', 'บด3,M1', 'ชบด3', 'ช2', 'ออกหน่วย', 'ลา' หรือ '' = ล้าง
 */
function setShift(date, code, note) {
  return withLock_(function () {
    // หลายเวรในวันเดียวคั่นด้วย , เช่น 'บด3,M1' (แถวละ 1 อย่าง, 'ลา' ใช้เดี่ยว)
    const parts = String(code || '').split(',').map(function (x) { return x.trim(); }).filter(String);
    parts.forEach(function (p) { if (!CODE_RE.test(p)) throw new Error('รหัสเวรไม่ถูกต้อง: ' + p); });
    if (parts.indexOf('ลา') >= 0 && parts.length > 1) throw new Error('วันที่ลา ลงเวรอื่นไม่ได้');
    const kinds = parts.map(function (p) { return p.replace(/[1-4]$/, ''); });
    if (new Set(kinds).size !== kinds.length) throw new Error('เลือกเวรประเภทเดียวกันซ้ำในวันเดียว');
    code = parts.join(',');
    const map = readShifts_();
    const before = map[date] ? map[date].code : '';
    if (code) map[date] = { code: code, note: note || '' };
    else delete map[date];
    writeShifts_(map);
    log_(date, code ? (before ? 'แก้ไข' : 'ลงเวร') : 'ล้าง', (before || '-') + ' → ' + (code || '-'));
    return getAll();
  });
}

/** ตั้ง/ยกเลิกวันหยุดราชการเอง */
function toggleHoliday(date, name) {
  return withLock_(function () {
    const map = readHolidays_();
    const h = map[date];
    if (h && h.source !== 'removed') {
      if (h.source === 'manual') delete map[date];
      else map[date] = { name: h.name, source: 'removed' };
      log_(date, 'ยกเลิกวันหยุด', h.name);
    } else {
      map[date] = { name: name || 'วันหยุด', source: 'manual' };
      log_(date, 'เพิ่มวันหยุด', name || '');
    }
    writeHolidays_(map);
    return getAll();
  });
}

/** บันทึกอัตรา {workday, bd, cbd, e, m, tax, sso} */
function saveRates(rates) {
  return withLock_(function () {
    Object.keys(RATES).forEach(function (k) {
      setSetting_(RATES[k][0], Number(rates[k]) || 0, RATES[k][2]);
    });
    return getAll();
  });
}
