// ============================================================
// BACKEND - QUẢN LÝ HỌC VIÊN
// ============================================================

const SHEET_ID = '1JI8U96SwwQXn5ZLqWNlLtGaxkRzArpljjdW9QdIOr7A';

const STUDENT_SHEET = 'BẢNG QLTHV 2026';
const LEAVE_SHEET = 'VE_QUE';
const ATT_SHEET = 'ĐIỂM DANH';


// ============================================================
// HÀM CƠ BẢN
// ============================================================

function sh(name) {
  return SpreadsheetApp
    .openById(SHEET_ID)
    .getSheetByName(name);
}


function json(o) {
  return ContentService
    .createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}


function today() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}


// ============================================================
// GET
// ============================================================

function doGet(e) {
  try {

    const action = String(
      e && e.parameter && e.parameter.action
        ? e.parameter.action
        : ''
    );

    if (action === 'getAll') {
      return json(getAll());
    }

    return json({
      ok: true,
      message: 'API Quản lý học viên đang hoạt động'
    });

  } catch (err) {

    return json({
      ok: false,
      error: String(err)
    });

  }
}


// ============================================================
// POST
// ============================================================

function doPost(e) {
  try {

    if (!e || !e.postData || !e.postData.contents) {
      return json({
        ok: false,
        error: 'Không có dữ liệu POST'
      });
    }

    const p = JSON.parse(e.postData.contents || '{}');

    // Lấy toàn bộ học viên
    if (p.action === 'getAll') {
      return json(getAll());
    }

    // Thêm về quê
    if (p.action === 'addLeave') {
      addLeave(p.row || {});
      return json({
        ok: true
      });
    }

    // Cập nhật về quê
    if (p.action === 'updateLeave') {
      updateLeave(p.id, p.status);
      return json({
        ok: true
      });
    }

    // Lưu điểm danh
    if (p.action === 'saveAttendance') {

      const rows = Array.isArray(p.rows)
        ? p.rows
        : [];

      const result = saveAttendance(rows);

      return json(result);
    }

    return json({
      ok: false,
      error: 'Unknown action: ' + String(p.action || '')
    });

  } catch (err) {

    return json({
      ok: false,
      error: String(err && err.stack ? err.stack : err)
    });

  }
}


// ============================================================
// ĐỌC HỌC VIÊN
// ============================================================

function getAll() {

  const studentSheet = sh(STUDENT_SHEET);

  const students = [];

  if (studentSheet) {

    const lastRow = studentSheet.getLastRow();

    if (lastRow >= 3) {

      // Sheet BẢNG QLTHV 2026:
      // B = Trạng thái
      // C = Mã số HV
      // D = Tên Romaji
      // H = Lớp
      // I = GVCN

      const values = studentSheet
        .getRange(3, 2, lastRow - 2, 8)
        .getDisplayValues();

      const seen = new Set();

      values.forEach(function(r) {

        const trangThai = String(r[0] || '').trim();
        const maHV = String(r[1] || '').trim();
        const hoTen = String(r[2] || '').trim();
        const lop = String(r[6] || '').trim();
        const gvcn = String(r[7] || '').trim();

        // Chỉ nhận mã học viên dạng TH202601...
        if (!/^TH\d{6,}$/.test(maHV)) {
          return;
        }

        if (seen.has(maHV)) {
          return;
        }

        seen.add(maHV);

        students.push({
          code: maHV,
          maHV: maHV,

          name: hoTen,
          hoTen: hoTen,

          class: lop,
          lop: lop,

          teacher: gvcn,
          gvcn: gvcn,

          status: trangThai,
          trangThai: trangThai
        });

      });
    }
  }


  // ==========================================================
  // ĐỌC ĐIỂM DANH HÔM NAY
  // ==========================================================

  const attendanceSheet = sh(ATT_SHEET);

  const attendance = [];

  if (attendanceSheet) {

    const lastRow = attendanceSheet.getLastRow();

    if (lastRow >= 2) {

      const values = attendanceSheet
        .getRange(2, 1, lastRow - 1, 5)
        .getDisplayValues();

      const currentDate = today();

      values.forEach(function(r) {

        const code = String(r[0] || '').trim();
        const name = String(r[1] || '').trim();
        const status = String(r[2] || '').trim();
        const date = normalizeDate(r[3]);

        if (!code) {
          return;
        }

        if (date !== currentDate) {
          return;
        }

        attendance.push({
          code: code,
          name: name,
          status: status,
          date: date,
          savedAt: String(r[4] || '')
        });

      });
    }
  }


  // ==========================================================
  // ĐỌC ĐƠN VỀ QUÊ
  // ==========================================================

  const leaves = rowsToObjects(sh(LEAVE_SHEET));


  return {
    ok: true,
    students: students,
    leaves: leaves,
    attendance: attendance,
    total: students.length
  };
}


// ============================================================
// ĐỌC SHEET THÀNH OBJECT
// ============================================================

function rowsToObjects(sheet) {

  if (!sheet) {
    return [];
  }

  const values = sheet.getDataRange().getDisplayValues();

  if (values.length < 2) {
    return [];
  }

  const headers = values[0].map(function(x) {
    return String(x || '').trim();
  });

  return values
    .slice(1)
    .filter(function(row) {
      return row.some(function(x) {
        return String(x || '').trim() !== '';
      });
    })
    .map(function(row) {

      const obj = {};

      headers.forEach(function(key, i) {

        if (key) {
          obj[key] = String(row[i] || '').trim();
        }

      });

      return obj;

    });
}


// ============================================================
// TẠO SHEET NẾU CHƯA CÓ
// ============================================================

function ensureSheet(name, headers) {

  const ss = SpreadsheetApp.openById(SHEET_ID);

  let sheet = ss.getSheetByName(name);

  if (!sheet) {

    sheet = ss.insertSheet(name);

    sheet
      .getRange(1, 1, 1, headers.length)
      .setValues([headers]);
  }

  return sheet;
}


// ============================================================
// VỀ QUÊ
// ============================================================

function addLeave(r) {

  const s = ensureSheet(
    LEAVE_SHEET,
    [
      'id',
      'code',
      'name',
      'class',
      'from',
      'to',
      'reason',
      'phone',
      'contact',
      'status',
      'createdAt'
    ]
  );

  s.appendRow([
    r.id || '',
    r.code || '',
    r.name || '',
    r.class || '',
    r.from || '',
    r.to || '',
    r.reason || '',
    r.phone || '',
    r.contact || '',
    r.status || '',
    new Date()
  ]);
}


// ============================================================
// CẬP NHẬT VỀ QUÊ
// ============================================================

function updateLeave(id, status) {

  const s = sh(LEAVE_SHEET);

  if (!s) {
    return;
  }

  const values = s.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {

    if (
      String(values[i][0] || '').trim() ===
      String(id || '').trim()
    ) {

      // Cột J = status
      s.getRange(i + 1, 10).setValue(status || '');

      break;
    }
  }
}


// ============================================================
// CHUẨN HÓA NGÀY
// ============================================================

function normalizeDate(value) {

  if (
    value instanceof Date &&
    !isNaN(value.getTime())
  ) {

    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );
  }


  const text = String(value || '').trim();

  if (!text) {
    return '';
  }


  // yyyy-MM-dd
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    return text;
  }


  // dd/MM/yyyy
  let m = text.match(
    /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/
  );

  if (m) {

    return (
      m[3] + '-' +
      ('0' + m[2]).slice(-2) + '-' +
      ('0' + m[1]).slice(-2)
    );
  }


  // dd-MM-yyyy
  m = text.match(
    /^(\d{1,2})-(\d{1,2})-(\d{4})$/
  );

  if (m) {

    return (
      m[3] + '-' +
      ('0' + m[2]).slice(-2) + '-' +
      ('0' + m[1]).slice(-2)
    );
  }


  return text;
}


// ============================================================
// LƯU ĐIỂM DANH
// ============================================================

function saveAttendance(rows) {

  if (!Array.isArray(rows)) {

    return {
      ok: false,
      error: 'Dữ liệu điểm danh không hợp lệ'
    };
  }


  const s = ensureSheet(
    ATT_SHEET,
    [
      'code',
      'name',
      'status',
      'date',
      'savedAt'
    ]
  );


  const currentDate = today();


  // ----------------------------------------------------------
  // Đọc dữ liệu hiện có
  // ----------------------------------------------------------

  const lastRow = s.getLastRow();

  const oldValues =
    lastRow >= 2
      ? s.getRange(2, 1, lastRow - 1, 5).getValues()
      : [];


  // ----------------------------------------------------------
  // Tạo map bản ghi hiện tại
  // key = code + date
  // ----------------------------------------------------------

  const existing = new Map();

  oldValues.forEach(function(r, index) {

    const code = String(r[0] || '').trim();

    const date = normalizeDate(r[3]);

    if (!code || !date) {
      return;
    }

    existing.set(
      code + '|' + date,
      {
        row: index + 2
      }
    );

  });


  // ----------------------------------------------------------
  // Cập nhật hoặc thêm
  // ----------------------------------------------------------

  rows.forEach(function(r) {

    const code = String(
      r.code ||
      r.maHV ||
      r.maHv ||
      ''
    ).trim();

    const name = String(
      r.name ||
      r.hoTen ||
      ''
    ).trim();

    const status = String(
      r.status ||
      r.trangThai ||
      ''
    ).trim();


    if (!code) {
      return;
    }


    const date =
      normalizeDate(r.date) ||
      currentDate;


    const key = code + '|' + date;


    const data = [
      code,
      name,
      status,
      date,
      new Date()
    ];


    // Đã có → cập nhật
    if (existing.has(key)) {

      const rowNumber =
        existing.get(key).row;

      s.getRange(
        rowNumber,
        1,
        1,
        5
      ).setValues([data]);

    }

    // Chưa có → thêm
    else {

      s.appendRow(data);

      existing.set(key, {
        row: s.getLastRow()
      });
    }

  });


  return {
    ok: true,
    message: 'Đã lưu điểm danh',
    count: rows.length,
    date: currentDate
  };
}
