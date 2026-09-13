// GOOGLE APPS SCRIPT BACKEND - QUAN LY HOC VIEN

const SHEET_ID = '1JI8U96SwwQXn5ZLqWNlLtGaxkRzArpljjdW9QdIOr7A';
const STUDENT_SHEET = 'BẢNG QLTTHV 2026';
const LEAVE_SHEET = 'VE_QUE';
const ATT_SHEET = 'DIEM_DANH';

function sh(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}

function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function clean(v) {
  return String(v ?? '').trim();
}

function today() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function normalizeDate(v) {
  if (v instanceof Date && !isNaN(v)) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = clean(v);
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return m[3] + '-' + String(m[2]).padStart(2,'0') + '-' + String(m[1]).padStart(2,'0');
  const d = new Date(s);
  return isNaN(d) ? s : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// ==================== GET ====================

function doGet(e) {
  try {
    const action = clean(e && e.parameter && e.parameter.action);
    if (action === 'getAll' || !action) return json(getAll());
    return json({ok:false, error:'Unknown action: ' + action});
  } catch (err) {
    return json({ok:false, error:String(err)});
  }
}

// ==================== POST ====================

function doPost(e) {
  try {
    const p = JSON.parse(e && e.postData && e.postData.contents ? e.postData.contents : '{}');
    const action = clean(p.action);

    if (action === 'getAll') return json(getAll());

    if (action === 'attendance' || action === 'saveAttendance') {
      return json(saveAttendance(p.date, p.class, p.records || p.rows || []));
    }

    if (action === 'addLeave') {
      addLeave(p.row || {});
      return json({ok:true});
    }

    if (action === 'updateLeave') {
      updateLeave(p.id, p.status);
      return json({ok:true});
    }

    return json({ok:false, error:'Unknown action: ' + action});
  } catch (err) {
    return json({ok:false, error:String(err)});
  }
}

// ==================== SHEET OBJECT ====================

function rowsToObjects(sheet) {
  if (!sheet) return [];
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];

  const headers = values[0].map(clean);

  return values.slice(1)
    .filter(r => r.some(x => clean(x) !== ''))
    .map(r => {
      const obj = {};
      headers.forEach((h,i) => {
        if (h) obj[h] = clean(r[i]);
      });
      return obj;
    });
}

// ==================== HOC VIEN ====================

function getStudents() {
  const sheet = sh(STUDENT_SHEET);

  if (!sheet) {
    throw new Error('Không tìm thấy sheet: ' + STUDENT_SHEET);
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 3 || lastCol < 3) return [];

  // Hàng 2 là tiêu đề tiếng Việt.
  const headers = sheet.getRange(2,1,1,lastCol)
    .getDisplayValues()[0].map(clean);

  // Dữ liệu bắt đầu từ hàng 3.
  const values = sheet.getRange(3,1,lastRow-2,lastCol)
    .getDisplayValues();

  const students = [];
  const seen = new Set();

  values.forEach(r => {
    // C = MÃ SỐ HV
    const maHV = clean(r[2]);

    // Chỉ lấy mã thật, loại bỏ các dòng tiêu đề/rác.
    if (!/^TH\d{6,}$/.test(maHV)) return;
    if (seen.has(maHV)) return;
    seen.add(maHV);

    // Toàn bộ dữ liệu gốc theo tiêu đề.
    const detail = {};
    headers.forEach((h,i) => {
      if (h) detail[h] = clean(r[i]);
    });

    students.push({
      maHV: maHV,
      code: maHV,

      hoTen: clean(r[3]),
      name: clean(r[3]),

      katakana: clean(r[4]),
      doiNgoai: clean(r[5]),

      trungTam: clean(r[6]),
      center: clean(r[6]),

      lop: clean(r[7]),
      class: clean(r[7]),

      gvcn: clean(r[8]),
      teacher: clean(r[8]),

      level: clean(r[9]),
      gioiTinh: clean(r[10]),
      ngaySinh: clean(r[11]),
      birth: clean(r[11]),

      nghiepDoan: clean(r[12]),
      nghiepDoanKana: clean(r[13]),

      congTy: clean(r[14]),
      congTyKana: clean(r[15]),

      nganhNghe: clean(r[16]),
      ngayThi: clean(r[17]),
      ngayNhapHoc: clean(r[18]),

      trangThai: clean(r[1]),
      status: clean(r[1]),

      // Giữ toàn bộ cột của sheet để dùng cho thẻ chi tiết.
      detail: detail
    });
  });

  return students;
}

// ==================== GET ALL ====================

function getAll() {
  const students = getStudents();
  const leaves = rowsToObjects(sh(LEAVE_SHEET));
  const attendance = [];

  const sheet = sh(ATT_SHEET);

  if (sheet && sheet.getLastRow() >= 2) {
    const rows = sheet.getRange(
      2, 1,
      sheet.getLastRow()-1,
      Math.max(5, sheet.getLastColumn())
    ).getDisplayValues();

    const currentDate = today();

    rows.forEach(r => {
      const code = clean(r[0]);
      if (!code) return;

      const date = normalizeDate(r[3]);
      if (date !== currentDate) return;

      attendance.push({
        code: code,
        name: clean(r[1]),
        status: clean(r[2]),
        date: date,
        savedAt: clean(r[4])
      });
    });
  }

  return {
    ok: true,
    students: students,
    leaves: leaves,
    attendance: attendance,
    total: students.length
  };
}

// ==================== TAO SHEET ====================

function ensureSheet(name, headers) {
  let sheet = sh(name);

  if (!sheet) {
    sheet = SpreadsheetApp.openById(SHEET_ID).insertSheet(name);
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
  }

  return sheet;
}

// ==================== VE QUE ====================

function addLeave(r) {
  const sheet = ensureSheet(LEAVE_SHEET, [
    'id','code','name','class','from','to',
    'reason','phone','contact','status','createdAt'
  ]);

  sheet.appendRow([
    clean(r.id),
    clean(r.code),
    clean(r.name),
    clean(r.class),
    clean(r.from),
    clean(r.to),
    clean(r.reason),
    clean(r.phone),
    clean(r.contact),
    clean(r.status),
    new Date()
  ]);
}

function updateLeave(id,status) {
  const sheet = sh(LEAVE_SHEET);
  if (!sheet) return;

  const values = sheet.getDataRange().getValues();

  for (let i=1; i<values.length; i++) {
    if (clean(values[i][0]) === clean(id)) {
      sheet.getRange(i+1,10).setValue(status);
      break;
    }
  }
}

// ==================== DIEM DANH ====================

function saveAttendance(selectedDate, classValue, records) {
  const sheet = ensureSheet(ATT_SHEET, [
    'code','name','status','date','savedAt'
  ]);

  const date = normalizeDate(selectedDate) || today();
  const students = getStudents();
  const studentMap = new Map();

  students.forEach(s => studentMap.set(clean(s.maHV), s));

  const codeSet = new Set();

  records.forEach(r => {
    const code = clean(r.code || r.id);
    if (code) codeSet.add(code);
  });

  // Xóa bản ghi cũ cùng ngày và cùng mã.
  if (sheet.getLastRow() >= 2) {
    const old = sheet.getRange(
      2,1,sheet.getLastRow()-1,5
    ).getDisplayValues();

    for (let i=old.length-1; i>=0; i--) {
      const oldCode = clean(old[i][0]);
      const oldDate = normalizeDate(old[i][3]);

      if (oldDate === date && codeSet.has(oldCode)) {
        sheet.deleteRow(i+2);
      }
    }
  }

  const output = [];

  records.forEach(r => {
    const code = clean(r.code || r.id);
    if (!code) return;

    const student = studentMap.get(code);

    output.push([
      code,
      student ? clean(student.hoTen) : clean(r.name),
      clean(r.status) || 'Có mặt',
      date,
      new Date()
    ]);
  });

  if (output.length) {
    sheet.getRange(
      sheet.getLastRow()+1,
      1,
      output.length,
      5
    ).setValues(output);
  }

  return {
    ok:true,
    message:'Lưu điểm danh thành công',
    date:date,
    class:clean(classValue),
    count:output.length
  };
}

// ==================== KIEM TRA NGUON ====================

function testSource() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(STUDENT_SHEET);

  Logger.log('FILE: ' + ss.getName());
  Logger.log('FILE ID: ' + ss.getId());
  Logger.log('SHEET: ' + (sheet ? sheet.getName() : 'KHÔNG CÓ'));

  if (!sheet) return;

  Logger.log('SỐ DÒNG: ' + sheet.getLastRow());
  Logger.log('SỐ CỘT: ' + sheet.getLastColumn());

  const maxRows = Math.min(sheet.getLastRow(),50);

  const values = sheet.getRange(
    1,3,maxRows,1
  ).getDisplayValues();

  Logger.log('MSHV CỘT C:');

  values.forEach((row,index) => {
    Logger.log((index+1) + ': ' + clean(row[0]));
  });

  Logger.log(
    'TỔNG HỌC VIÊN HỢP LỆ: ' +
    getStudents().length
  );
}
