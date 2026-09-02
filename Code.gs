/**
 * Backend Google Apps Script cho web quản lý trung tâm.
 * Script Property bắt buộc: ADMIN_KEY
 */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var callback = String(p.callback || "");
  try {
    var mode = String(p.mode || "all");
    if (mode === "setup") {
      requireAdmin(p.adminKey);
      setupSheets();
      return output({ ok: true, message: "Đã tạo cấu trúc dữ liệu." }, callback);
    }
    if (mode === "all") {
      requireAdmin(p.adminKey);
      return output({ ok: true, data: readAllData() }, callback);
    }
    if (mode === "parent") {
      var pin = digits(p.pin);
      if (!/^\d{5}$/.test(pin)) throw new Error("Mã tra cứu không hợp lệ.");
      var profile = readParentProfile(pin);
      if (!profile) throw new Error("Mã tra cứu không chính xác.");
      return output({ ok: true, data: profile }, callback);
    }
    if (mode === "save") {
      requireAdmin(p.adminKey);
      var payload = JSON.parse(String(p.payload || "{}"));
      saveRecord(payload.type, payload.record);
      return output({ ok: true, data: readAllData() }, callback);
    }
    throw new Error("Chế độ không hợp lệ.");
  } catch (err) {
    return output({ ok: false, message: err.message || "Có lỗi xảy ra." }, callback);
  }
}

function requireAdmin(value) {
  var saved = PropertiesService.getScriptProperties().getProperty("ADMIN_KEY");
  if (!saved || String(value || "") !== saved) {
    throw new Error("Mã quản trị không chính xác.");
  }
}

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet(ss, "HOC_SINH", [
    "ID", "MaHS", "HoTen", "LopTruong", "SDT_PHHS", "MaTraCuu"
  ]);
  ensureSheet(ss, "LOP_HOC", [
    "ID", "TenLop", "MonHoc", "LichHoc", "HocPhiMoiBuoi"
  ]);
  ensureSheet(ss, "DANG_KY_LOP", [
    "ID", "HocSinhID", "LopHocID"
  ]);
  ensureSheet(ss, "DIEM_DANH", [
    "ID", "Ngay", "LopHocID", "HocSinhID", "TrangThai"
  ]);
  ensureSheet(ss, "BANG_DIEM", [
    "ID", "Ngay", "LopHocID", "HocSinhID", "TenBai", "Diem"
  ]);
  ensureSheet(ss, "NHAN_XET", [
    "ID", "Ngay", "LopHocID", "HocSinhID", "NoiDung"
  ]);
  ensureSheet(ss, "HOC_PHI", [
    "ID", "HocSinhID", "Thang", "SoTien", "TrangThai", "NgayDong",
    "LopHocID", "SoBuoi", "DonGia"
  ]);
  ensureSheet(ss, "CAI_DAT_TRUNG_TAM", ["Khoa", "GiaTri"]);
  ensureColumns(ss.getSheetByName("HOC_PHI"), [
    "LopHocID", "SoBuoi", "DonGia"
  ]);
  ss.getSheetByName("LOP_HOC").getRange(1, 5).setValue("HocPhiMoiBuoi");
}

function ensureSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#173b5d")
      .setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }
}

function ensureColumns(sheet, names) {
  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  names.forEach(function(name) {
    if (headers.indexOf(name) >= 0) return;
    sheet.getRange(1, headers.length + 1).setValue(name);
    headers.push(name);
  });
}

function readAllData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSheets();
  return {
    students: rows(ss, "HOC_SINH").map(function(r) {
      return {
        id: r.ID, code: r.MaHS, name: r.HoTen,
        schoolClass: r.LopTruong, phone: r.SDT_PHHS, pin: r.MaTraCuu
      };
    }),
    classes: rows(ss, "LOP_HOC").map(function(r) {
      return {
        id: r.ID, name: r.TenLop, subject: r.MonHoc,
        schedule: r.LichHoc, tuition: r.HocPhiMoiBuoi || r.HocPhiThang
      };
    }),
    enrollments: rows(ss, "DANG_KY_LOP").map(function(r) {
      return { id: r.ID, studentId: r.HocSinhID, classId: r.LopHocID };
    }),
    attendance: rows(ss, "DIEM_DANH").map(function(r) {
      return {
        id: r.ID, date: r.Ngay, classId: r.LopHocID,
        studentId: r.HocSinhID, status: r.TrangThai
      };
    }),
    grades: rows(ss, "BANG_DIEM").map(function(r) {
      return {
        id: r.ID, date: r.Ngay, classId: r.LopHocID,
        studentId: r.HocSinhID, title: r.TenBai,
        score: Number(String(r.Diem || "").replace(",", "."))
      };
    }),
    comments: rows(ss, "NHAN_XET").map(function(r) {
      return {
        id: r.ID, date: r.Ngay, classId: r.LopHocID,
        studentId: r.HocSinhID, text: r.NoiDung
      };
    }),
    fees: rows(ss, "HOC_PHI").map(function(r) {
      return {
        id: r.ID, studentId: r.HocSinhID, month: r.Thang,
        amount: Number(r.SoTien || 0), status: r.TrangThai, paidAt: r.NgayDong,
        classId: r.LopHocID, sessions: Number(r.SoBuoi || 0),
        unitPrice: Number(r.DonGia || 0)
      };
    })
  };
}

function rows(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getDataRange().getDisplayValues();
  var headers = values.shift();
  return values.filter(function(row) {
    return row.some(function(value) { return String(value).trim(); });
  }).map(function(row) {
    var item = {};
    headers.forEach(function(header, i) { item[header] = row[i]; });
    return item;
  });
}

function saveRecord(type, record) {
  if (!record || !record.id) throw new Error("Bản ghi thiếu ID.");
  var map = {
    student: {
      sheet: "HOC_SINH",
      values: [
        record.id, record.code, record.name, record.schoolClass,
        digits(record.phone), record.pin
      ]
    },
    class: {
      sheet: "LOP_HOC",
      values: [
        record.id, record.name, record.subject, record.schedule, record.tuition
      ]
    },
    enrollment: {
      sheet: "DANG_KY_LOP",
      values: [record.id, record.studentId, record.classId]
    },
    attendance: {
      sheet: "DIEM_DANH",
      values: [
        record.id, record.date, record.classId, record.studentId, record.status
      ]
    },
    grade: {
      sheet: "BANG_DIEM",
      values: [
        record.id, record.date, record.classId, record.studentId,
        record.title, record.score
      ]
    },
    comment: {
      sheet: "NHAN_XET",
      values: [
        record.id, record.date, record.classId, record.studentId, record.text
      ]
    },
    fee: {
      sheet: "HOC_PHI",
      values: [
        record.id, record.studentId, record.month, record.amount,
        record.status, record.paidAt || "", record.classId || "",
        Number(record.sessions || 0), Number(record.unitPrice || 0)
      ]
    }
  };
  var config = map[type];
  if (!config) throw new Error("Loại dữ liệu không hợp lệ.");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSheets();
  var sheet = ss.getSheetByName(config.sheet);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var last = sheet.getLastRow();
    var target = last + 1;
    if (last >= 2) {
      var ids = sheet.getRange(2, 1, last - 1, 1).getDisplayValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(record.id)) {
          target = i + 2;
          break;
        }
      }
    }
    sheet.getRange(target, 1, 1, config.values.length).setValues([config.values]);
    sheet.getRange(target, 1).setNumberFormat("@");
    if (type === "student") {
      sheet.getRange(target, 2).setNumberFormat("@");
      sheet.getRange(target, 5, 1, 2).setNumberFormat("@");
    }
  } finally {
    lock.releaseLock();
  }
}

function readParentProfile(pin) {
  var data = readAllData();
  var student = data.students.filter(function(s) {
    return digits(s.pin) === pin;
  })[0];
  if (!student) return null;
  return {
    student: student,
    enrollments: data.enrollments.filter(function(x) {
      return x.studentId === student.id;
    }),
    attendance: data.attendance.filter(function(x) {
      return x.studentId === student.id;
    }),
    fees: data.fees.filter(function(x) {
      return x.studentId === student.id;
    }),
    grades: data.grades.filter(function(x) {
      return x.studentId === student.id;
    }),
    comments: data.comments.filter(function(x) {
      return x.studentId === student.id;
    })
  };
}

function digits(value) {
  return String(value || "").replace(/\D/g, "");
}

function output(data, callback) {
  var json = JSON.stringify(data);
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

