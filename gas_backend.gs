const SHEET_ID = "1JI8U96SwwQXn5ZLqWNlLtGaxkRzArpljjdW9QdIOr7A";
const SHEET_GID = 62168460;

// ================= CHUẨN HÓA =================
function normalize(str) {
  return String(str || "")
    .trim()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/\s+/g, "");
}

// ================= TÌM CỘT =================
function getColumnIndex(headers, possibleNames) {
  const h = headers.map(x => normalize(x));

  for (const name of possibleNames) {
    const n = normalize(name);
    const idx = h.findIndex(x => x === n || x.includes(n));
    if (idx !== -1) return idx;
  }

  return -1;
}

// ================= LẤY SHEET THEO GID =================
function getMainSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  for (const sheet of ss.getSheets()) {
    if (sheet.getSheetId() === SHEET_GID) {
      return sheet;
    }
  }

  throw new Error("Không tìm thấy sheet có GID: " + SHEET_GID);
}

// ================= LẤY GIÁ TRỊ Ô =================
function cell(row, index) {
  if (index < 0 || index >= row.length) return "";
  return String(row[index] ?? "").trim();
}

// ================= TẠO CHI TIẾT TOÀN BỘ CỘT =================
function taoChiTiet(headers, row) {
  const detail = {};

  headers.forEach((header, index) => {
    const tenCot = String(header || "").trim();

    if (tenCot !== "") {
      detail[tenCot] = row[index] ?? "";
    }
  });

  // BỔ SUNG RIÊNG CỘT AC
  detail["Nội dung cột AC"] = cell(row, 28);

  return detail;
}

// ================= LƯU ĐIỂM DANH =================
function luuDuLieu(params) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName("DiemDanh");

    if (!sheet) {
      sheet = ss.insertSheet("DiemDanh");
    }

    sheet.appendRow([
      params.lop || "",
      params.from || "",
      params.to || "",
      params.data || "",
      new Date().toLocaleString("vi-VN")
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({
        ok: true,
        message: "✅ Lưu điểm danh thành công!"
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        ok: false,
        message: "Lỗi: " + err.message
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ================= API CHÍNH =================
function doGet(e) {
  const action = e.parameter.action || "";

  if (action === "saveGrid") {
    return luuDuLieu(e.parameter);
  }

  try {
    const sheet = getMainSheet();

    // Dùng getDisplayValues để giữ nguyên mã HV, ngày tháng,
    // số điện thoại, nội dung tiền và dữ liệu hiển thị trong Sheet.
    const allData = sheet.getDataRange().getDisplayValues();

    if (allData.length < 2) {
      throw new Error("Google Sheet chưa có dữ liệu");
    }

    const headers = allData[0];
    const rows = allData.slice(1);

    // ================= XÁC ĐỊNH CỘT =================
    const idxMaHV = getColumnIndex(headers, [
      "Mã HV",
      "Mã học viên",
      "Mã số HV",
      "Mã số",
      "Mã"
    ]);

    const idxTen = getColumnIndex(headers, [
      "Họ tên",
      "Họ và tên",
      "Tên học viên",
      "Tên Romaji",
      "TenRomaji",
      "Romaji",
      "Tên"
    ]);

    const idxLop = getColumnIndex(headers, [
      "Lớp",
      "Lớp học",
      "Khóa",
      "Lop"
    ]);

    const idxGVCN = getColumnIndex(headers, [
      "GVCN",
      "Giáo viên chủ nhiệm",
      "Chủ nhiệm",
      "Giáo viên"
    ]);

    const idxTrangThai = getColumnIndex(headers, [
      "Trạng thái",
      "Tình trạng",
      "Trạng thái học tập",
      "TrangThai"
    ]);

    // ================= LẤY TÊN =================
    function layTen(row) {
      let ten = cell(row, idxTen);

      // Nếu chưa tìm thấy cột tên, dò thêm các tên thường gặp
      if (!ten) {
        const danhSachTen = [
          "hoten",
          "hovaten",
          "tenhocvien",
          "tenromaji",
          "romaji",
          "ten"
        ];

        for (let i = 0; i < headers.length; i++) {
          if (danhSachTen.includes(normalize(headers[i]))) {
            ten = cell(row, i);
            if (ten) break;
          }
        }
      }

      return ten;
    }

    // ================= GET ALL =================
    if (action === "getAll") {
      const students = rows
        .map(row => ({
          maHV: cell(row, idxMaHV),
          hoTen: layTen(row),
          lop: cell(row, idxLop) || "Chưa có",
          gvcn: cell(row, idxGVCN) || "Chưa có",
          trangThai: cell(row, idxTrangThai) || "Đang học",

          // Gửi thêm cột AC
          cotAC: cell(row, 28)
        }))
        .filter(s => s.maHV || s.hoTen);

      let coMat = 0;
      let vang = 0;
      let veQue = 0;

      students.forEach(s => {
        const tt = normalize(s.trangThai);

        if (
          tt === "" ||
          tt === "danghoc" ||
          tt === "comat" ||
          tt === "binhthuong"
        ) {
          coMat++;
        } else if (
          tt === "vang" ||
          tt === "vangmat" ||
          tt === "nghi" ||
          tt === "nghihoc"
        ) {
          vang++;
        } else if (
          tt === "veque" ||
          tt === "nghiphep"
        ) {
          veQue++;
        } else {
          coMat++;
        }
      });

      return ContentService
        .createTextOutput(JSON.stringify({
          ok: true,
          students: students,
          stats: {
            tong: students.length,
            coMat: coMat,
            vang: vang,
            veQue: veQue
          }
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ================= CHI TIẾT HỌC VIÊN =================
    if (action === "getStudentDetail") {
      const maCanTim = String(
        e.parameter.maHV || e.parameter.mshv || ""
      ).trim();

      if (!maCanTim) {
        throw new Error("Thiếu mã học viên");
      }

      const studentRow = rows.find(row =>
        cell(row, idxMaHV) === maCanTim
      );

      if (!studentRow) {
        throw new Error("Không tìm thấy học viên: " + maCanTim);
      }

      return ContentService
        .createTextOutput(JSON.stringify({
          ok: true,
          mshv: maCanTim,
          detail: taoChiTiet(headers, studentRow)
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    throw new Error("Hành động không hợp lệ");

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        ok: false,
        message: err.message
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
