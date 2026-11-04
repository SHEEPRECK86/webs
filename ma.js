//==============================================================
//                      ประกาศตัวแปรสำคัญ
//==============================================================
const SPREADSHEET_ID = '1mQxa6V0uFmm1qHM5HmsMUW--ZFGXa38oT_1XYeaUmNM';
const LOG_SHEET_NAME = 'meeting';
const DATA_SHEET_NAME = 'data';

const BOOKING_MINUTES = 90;
const BOOKING_START_TIME = '16:30';
const BOOKING_END_TIME = '20:00';
const ALLOWED_START_TIMES = ['16:30', '17:00', '17:30', '18:00', '18:30'];
const ADMIN_EMAIL = 'torfuninlao@gmail.com';
const TIMEZONE = 'Asia/Bangkok';
const STATUS_DEFAULT = 'จองแล้ว';
const STATUS_CANCELLED = 'ยกเลิก';

//==============================================================
//                   ฟังก์ชันหลักของ Web App
//==============================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ระบบจองเวลาซ้อม')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

//==============================================================
//                  ดึงข้อมูลจาก Google Sheets
//==============================================================
function getColumnData(sheetName, column) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID)
    .getSheetByName(sheetName);

  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(`${column}2:${column}${lastRow}`)
    .getValues()
    .map(r => String(r[0]).trim())
    .filter(v => v !== '');
}

//==============================================================
//             ดึงรายการจองสำหรับแสดงในปฏิทิน
//==============================================================
/**
 * รับปี ค.ศ. และเดือน 1-12 แล้วคืนรายการจองในเดือนนั้น
 * ส่งเฉพาะข้อมูลที่จำเป็นต่อการแสดงปฏิทิน
 */
function getBookingsForCalendar(year, month) {
  const y = Number(year);
  const m = Number(month);

  if (!Number.isInteger(y) || !Number.isInteger(m) ||
      m < 1 || m > 12 || y < 2000 || y > 2100) {
    throw new Error('เดือนหรือปีไม่ถูกต้อง');
  }

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID)
    .getSheetByName(LOG_SHEET_NAME);

  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
  const monthKey = `${y}-${String(m).padStart(2, '0')}`;

  return values.reduce((result, row) => {
    if (!(row[6] instanceof Date) ||
        !(row[7] instanceof Date) ||
        !(row[9] instanceof Date)) return result;

    const date = Utilities.formatDate(row[6], TIMEZONE, 'yyyy-MM-dd');
    if (!date.startsWith(monthKey)) return result;

    result.push({
      date: date,
      name: String(row[1] || ''),
      department: String(row[5] || ''),
      startTime: Utilities.formatDate(row[7], TIMEZONE, 'HH:mm'),
      endTime: Utilities.formatDate(row[9], TIMEZONE, 'HH:mm'),
      status: String(row[12] || STATUS_DEFAULT)
    });
    return result;
  }, []).sort((a, b) =>
    (a.date + a.startTime).localeCompare(b.date + b.startTime)
  );
}

//==============================================================
//                         ระบบจอง
//==============================================================
/**
 * form ประกอบด้วย name, department, contact, email, date, startTime
 */
function saveBooking(form) {
  try {
    if (!form || !form.date || !form.startTime) {
      return { ok: false, message: 'กรุณากรอกข้อมูลให้ครบทุกช่อง' };
    }

    form.name = String(form.name || '').trim().slice(0, 100);
    form.department = String(form.department || '').trim().slice(0, 100);
    form.contact = String(form.contact || '').trim().slice(0, 100);
    form.email = String(form.email || '').trim().slice(0, 150);

    if (!form.name || !form.department || !form.contact || !form.email) {
      return {
        ok: false,
        message: 'กรุณากรอกชื่อ ฝ่าย ช่องทางการติดต่อกลับ และ Email ให้ครบ'
      };
    }

    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(form.email)) {
      return { ok: false, message: 'รูปแบบ Email ไม่ถูกต้อง' };
    }

    ['name', 'department', 'contact', 'email'].forEach(key => {
      if (/^[=+\-@]/.test(form[key])) form[key] = "'" + form[key];
    });

    if (!/^\d{4}-\d{2}-\d{2}$/.test(form.date) ||
        !/^([01]\d|2[0-3]):[0-5]\d$/.test(form.startTime)) {
      return { ok: false, message: 'รูปแบบวันที่หรือเวลาไม่ถูกต้อง' };
    }

    // อนุญาตเฉพาะเวลาเริ่มที่กำหนด และต้องสิ้นสุดไม่เกิน 20:00 น.
    if (!ALLOWED_START_TIMES.includes(form.startTime)) {
      return {
        ok: false,
        message: 'กรุณาเลือกเวลาเริ่ม 16:30, 17:00, 17:30, 18:00 หรือ 18:30 น. เท่านั้น'
      };
    }

    const [y, m, d] = form.date.split('-').map(Number);
    const [hh, mm] = form.startTime.split(':').map(Number);
    const start = new Date(y, m - 1, d, hh, mm);
    const end = new Date(start.getTime() + BOOKING_MINUTES * 60000);
    const endMinutes = end.getHours() * 60 + end.getMinutes();
    const [limitHour, limitMinute] = BOOKING_END_TIME.split(':').map(Number);
    if (endMinutes > limitHour * 60 + limitMinute) {
      return {
        ok: false,
        message: 'เวลาจองต้องสิ้นสุดไม่เกิน 20:00 น.'
      };
    }

    if (start.getFullYear() !== y ||
        start.getMonth() !== m - 1 ||
        start.getDate() !== d) {
      return { ok: false, message: 'วันที่ไม่ถูกต้อง' };
    }

    if (Utilities.formatDate(start, TIMEZONE, 'yyyy-MM-dd') !==
        Utilities.formatDate(end, TIMEZONE, 'yyyy-MM-dd')) {
      return {
        ok: false,
        message: 'เวลาสิ้นสุดข้ามเที่ยงคืน กรุณาเลือกเวลาเริ่มให้เร็วขึ้น'
      };
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      const sheet = SpreadsheetApp.openById(SPREADSHEET_ID)
        .getSheetByName(LOG_SHEET_NAME);

      if (!sheet) throw new Error('ไม่พบชีต ' + LOG_SHEET_NAME);

      const clash = findClash_(sheet, start, end);
      if (clash) {
        return {
          ok: false,
          message: `เวลาชนกับการจองของ ${clash.name} (${clash.from}-${clash.to})`
        };
      }

      const row = new Array(13).fill('');
      row[0] = new Date();
      row[1] = form.name;
      row[2] = form.contact;
      row[3] = form.email;
      row[5] = form.department;
      row[6] = start;
      row[7] = start;
      row[8] = end;
      row[9] = end;
      row[12] = STATUS_DEFAULT;

      const r = sheet.getLastRow() + 1;
      sheet.getRange(r, 1, 1, 13).setValues([row]);

      sheet.getRange(r, 7).setNumberFormat('dd/MM/yyyy');
      sheet.getRange(r, 9).setNumberFormat('dd/MM/yyyy');
      sheet.getRange(r, 8).setNumberFormat('HH:mm');
      sheet.getRange(r, 10).setNumberFormat('HH:mm');
    } finally {
      lock.releaseLock();
    }

    const fmt = t => Utilities.formatDate(t, TIMEZONE, 'HH:mm');
    const booking = {
      name: form.name.replace(/^'/, ''),
      department: form.department.replace(/^'/, ''),
      contact: form.contact.replace(/^'/, ''),
      email: form.email.replace(/^'/, ''),
      date: Utilities.formatDate(start, TIMEZONE, 'yyyy-MM-dd'),
      startTime: fmt(start),
      endTime: fmt(end),
      status: STATUS_DEFAULT
    };

    // ส่งอีเมลแจ้งเฉพาะผู้ดูแลระบบ โดยไม่ให้ปัญหาอีเมลทำให้การจองหายไป
    let emailSent = false;
    try {
      sendBookingEmail_(booking);
      emailSent = true;
    } catch (emailError) {
      console.error('ส่งอีเมลแจ้ง Admin ไม่สำเร็จ: ' + emailError.message);
    }

    return {
      ok: true,
      emailSent: emailSent,
      message: `จองสำเร็จ ${fmt(start)}-${fmt(end)} น.`,
      booking: booking
    };

  } catch (err) {
    return { ok: false, message: 'เกิดข้อผิดพลาด: ' + err.message };
  }
}


//==============================================================
//                 ส่งอีเมลแจ้งผู้ดูแลระบบ
//==============================================================
function sendBookingEmail_(booking) {
  const dateParts = booking.date.split('-');
  const dateText = dateParts[2] + '/' + dateParts[1] + '/' + dateParts[0];
  const subject = 'แจ้งเตือนการจองเวลาซ้อมใหม่';
  const body = [
    'มีรายการจองเวลาซ้อมใหม่',
    '',
    'ชื่อผู้จอง: ' + booking.name,
    'ฝ่าย: ' + booking.department,
    'วันที่: ' + dateText,
    'เวลา: ' + booking.startTime + ' - ' + booking.endTime + ' น.',
    'ช่องทางติดต่อกลับ: ' + booking.contact,
    'Email ผู้จอง: ' + booking.email,
    'สถานะ: ' + booking.status
  ].join('\n');

  MailApp.sendEmail({
    to: ADMIN_EMAIL,
    subject: subject,
    body: body
  });
}

//==============================================================
//                 ตรวจสอบเวลาจองซ้ำ
//==============================================================
function findClash_(sheet, start, end) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const day = Utilities.formatDate(start, TIMEZONE, 'yyyy-MM-dd');
  const toMin = t => {
    const [h, m] = Utilities.formatDate(t, TIMEZONE, 'HH:mm')
      .split(':').map(Number);
    return h * 60 + m;
  };

  const newFrom = toMin(start);
  const newTo = toMin(end);
  const rows = sheet.getRange(2, 1, lastRow - 1, 13).getValues();

  for (const row of rows) {
    if (!(row[6] instanceof Date) ||
        !(row[7] instanceof Date) ||
        !(row[9] instanceof Date)) continue;

    if (row[12] === STATUS_CANCELLED) continue;
    if (Utilities.formatDate(row[6], TIMEZONE, 'yyyy-MM-dd') !== day) continue;

    const from = toMin(row[7]);
    const to = toMin(row[9]);

    if (newFrom < to && newTo > from) {
      return {
        name: row[1],
        from: Utilities.formatDate(row[7], TIMEZONE, 'HH:mm'),
        to: Utilities.formatDate(row[9], TIMEZONE, 'HH:mm')
      };
    }
  }
  return null;
}

//==============================================================
//                  ดูข้อมูลการจองทั้งหมด
//==============================================================
function showAllBookings() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID)
    .getSheetByName(LOG_SHEET_NAME);

  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('ยังไม่มีข้อมูลการจอง');
    return;
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
  const f = (v, p) => v instanceof Date
    ? Utilities.formatDate(v, TIMEZONE, p)
    : v;

  Logger.log('=== ข้อมูลการจองทั้งหมด ===');
  data.forEach((row, index) => {
    Logger.log(`แถวที่ ${index + 2}:`);
    Logger.log(`ผู้จอง: ${row[1]}`);
    Logger.log(`ช่องทางติดต่อกลับ: ${row[2]}`);
    Logger.log(`Email: ${row[3]}`);
    Logger.log(`ฝ่าย: ${row[5]}`);
    Logger.log(`วันที่: ${f(row[6], 'dd/MM/yyyy')}`);
    Logger.log(`เวลาเริ่ม: ${f(row[7], 'HH:mm')}`);
    Logger.log(`เวลาสิ้นสุด: ${f(row[9], 'HH:mm')}`);
    Logger.log(`สถานะ: ${row[12]}`);
    Logger.log('--------------------------');
  });
}
