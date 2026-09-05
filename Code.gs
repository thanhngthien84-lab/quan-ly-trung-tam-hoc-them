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
      var cache = CacheService.getScriptCache();
      var cacheKey = "parent-profile-" + pin;
      var cachedProfile = cache.get(cacheKey);
      var profile = cachedProfile ? JSON.parse(cachedProfile) : readParentProfile(pin);
      if (!profile) throw new Error("Mã tra cứu không chính xác.");
      if (!cachedProfile) cache.put(cacheKey, JSON.stringify(profile), 120);
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
  ensureSheet(ss, "KET_QUA_HOC_TAP", [
    "Ngay", "TenLop", "MaHS", "HoTen", "TenBai", "Diem", "NhanXet"
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
  configureLearningInput(ss);
  syncClassScoreSheets(ss);
}

function configureLearningInput(ss) {
  var sheet = ss.getSheetByName("KET_QUA_HOC_TAP");
  var classSheet = ss.getSheetByName("LOP_HOC");
  var studentSheet = ss.getSheetByName("HOC_SINH");
  if (!sheet || !classSheet || !studentSheet) return;
  var classRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(classSheet.getRange("B2:B1000"), true)
    .setAllowInvalid(true).build();
  var studentRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(studentSheet.getRange("B2:B1000"), true)
    .setAllowInvalid(true).build();
  sheet.getRange("B2:B1000").setDataValidation(classRule);
  sheet.getRange("C2:C1000").setDataValidation(studentRule);
  sheet.getRange("A2:A1000").setNumberFormat("dd/MM/yyyy");
  sheet.getRange("F2:F1000").setNumberFormat("0.00");
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

function scoreSheetName(classItem) {
  var safe = String(classItem.TenLop || classItem.ID || "LOP")
    .replace(/[\\\/?*\[\]:]/g, "-").trim().slice(0, 80);
  return "DIEM_" + safe;
}

function findScoreSheet(ss, classId) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    if (sheet.getName().indexOf("DIEM_") !== 0 || sheet.getLastRow() < 1) continue;
    if (String(sheet.getRange("B1").getDisplayValue()) === String(classId)) return sheet;
  }
  return null;
}

function uniqueSheetName(ss, base) {
  var name = base.slice(0, 95);
  var candidate = name;
  var index = 2;
  while (ss.getSheetByName(candidate)) {
    candidate = (name.slice(0, 90) + "_" + index).slice(0, 99);
    index++;
  }
  return candidate;
}

function syncClassScoreSheets(ss) {
  var classRows = rows(ss, "LOP_HOC");
  var students = rows(ss, "HOC_SINH");
  var enrollments = rows(ss, "DANG_KY_LOP");
  classRows.forEach(function(classItem) {
    var sheet = findScoreSheet(ss, classItem.ID);
    if (!sheet) {
      sheet = ss.insertSheet(uniqueSheetName(ss, scoreSheetName(classItem)));
    }
    sheet.getRange("A1:B2").setValues([
      ["LOP_HOC_ID", classItem.ID],
      ["TEN_LOP", classItem.TenLop]
    ]);
    if (sheet.getLastRow() < 5) {
      sheet.getRange("A3:B5").setValues([
        ["MaHS", "HoTen"],
        ["", ""],
        ["", ""]
      ]);
      sheet.getRange("A3:B3").setFontWeight("bold")
        .setBackground("#173b5d").setFontColor("#ffffff");
      sheet.setFrozenRows(5);
      sheet.setFrozenColumns(2);
      sheet.setColumnWidth(1, 120);
      sheet.setColumnWidth(2, 220);
    }
    var existing = {};
    if (sheet.getLastRow() >= 6) {
      sheet.getRange(6, 1, sheet.getLastRow() - 5, 1)
        .getDisplayValues().forEach(function(row) {
          existing[String(row[0]).trim()] = true;
        });
    }
    var memberIds = enrollments.filter(function(item) {
      return String(item.LopHocID) === String(classItem.ID);
    }).map(function(item) { return String(item.HocSinhID); });
    var missing = students.filter(function(student) {
      return memberIds.indexOf(String(student.ID)) >= 0 &&
        !existing[String(student.MaHS).trim()];
    }).map(function(student) {
      return [String(student.MaHS || ""), student.HoTen || ""];
    });
    if (missing.length) {
      var startRow = Math.max(sheet.getLastRow() + 1, 6);
      sheet.getRange(startRow, 1, missing.length, 2).setValues(missing);
      sheet.getRange(startRow, 1, missing.length, 1).setNumberFormat("@");
    }
  });
}

function readClassScores(ss, classes, students) {
  var result = [];
  classes.forEach(function(classItem) {
    var sheet = findScoreSheet(ss, classItem.id);
    if (!sheet || sheet.getLastRow() < 6 || sheet.getLastColumn() < 3) return;
    var values = sheet.getDataRange().getDisplayValues();
    for (var column = 2; column < values[0].length; column++) {
      var testCode = String((values[2] || [])[column] || "").trim();
      var title = String((values[3] || [])[column] || "").trim();
      var date = String((values[4] || [])[column] || "").trim();
      if (!title && !testCode) continue;
      for (var row = 5; row < values.length; row++) {
        var studentCode = String(values[row][0] || "").trim();
        var rawScore = String(values[row][column] || "").trim().replace(",", ".");
        if (!studentCode || rawScore === "" || !isFinite(Number(rawScore))) continue;
        var student = findStudent(students, studentCode, values[row][1]);
        if (!student) continue;
        result.push({
          id: "bang-" + sheet.getSheetId() + "-" + column + "-" + student.id,
          date: date, classId: classItem.id, studentId: student.id,
          title: title || testCode || "Bài kiểm tra", score: Number(rawScore)
        });
      }
    }
  });
  return result;
}

function readAllData(skipSetup) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!skipSetup) setupSheets();

  var students = rows(ss, "HOC_SINH").map(function(r) {
    return {
      id: r.ID, code: r.MaHS, name: r.HoTen,
      schoolClass: r.LopTruong, phone: r.SDT_PHHS, pin: r.MaTraCuu
    };
  });
  var classes = rows(ss, "LOP_HOC").map(function(r) {
    return {
      id: r.ID, name: r.TenLop, subject: r.MonHoc,
      schedule: r.LichHoc, tuition: r.HocPhiMoiBuoi || r.HocPhiThang,
      scoreSheetUrl: findScoreSheet(ss, r.ID)
        ? ss.getUrl() + "#gid=" + findScoreSheet(ss, r.ID).getSheetId()
        : ss.getUrl()
    };
  });
  var enrollments = rows(ss, "DANG_KY_LOP").map(function(r) {
    return { id: r.ID, studentId: r.HocSinhID, classId: r.LopHocID };
  });
  var attendance = rows(ss, "DIEM_DANH").map(function(r) {
    return {
      id: r.ID, date: r.Ngay, classId: r.LopHocID,
      studentId: r.HocSinhID, status: r.TrangThai
    };
  });
  var fees = rows(ss, "HOC_PHI").map(function(r) {
    return {
      id: r.ID, studentId: r.HocSinhID, month: r.Thang,
      amount: Number(r.SoTien || 0), status: r.TrangThai, paidAt: r.NgayDong,
      classId: r.LopHocID, sessions: Number(r.SoBuoi || 0),
      unitPrice: Number(r.DonGia || 0)
    };
  });
  var grades = rows(ss, "BANG_DIEM").map(function(r) {
    return {
      id: r.ID, date: r.Ngay, classId: r.LopHocID,
      studentId: r.HocSinhID, title: r.TenBai,
      score: Number(String(r.Diem || "").replace(",", "."))
    };
  });
  var comments = rows(ss, "NHAN_XET").map(function(r) {
    return {
      id: r.ID, date: r.Ngay, classId: r.LopHocID,
      studentId: r.HocSinhID, text: r.NoiDung
    };
  });

  grades = grades.concat(readClassScores(ss, classes, students));

  var learningRows = rows(ss, "KET_QUA_HOC_TAP");
  learningRows.forEach(function(r, index) {
    var classItem = findClass(classes, r.TenLop);
    var student = findStudent(students, r.MaHS, r.HoTen);
    if (!classItem || !student) return;
    var baseId = "sheet-" + (index + 2) + "-" + classItem.id + "-" + student.id;
    var rawScore = String(r.Diem || "").trim().replace(",", ".");
    if (rawScore !== "" && isFinite(Number(rawScore))) {
      grades.push({
        id: baseId + "-diem", date: r.Ngay, classId: classItem.id,
        studentId: student.id, title: r.TenBai || "Bài kiểm tra",
        score: Number(rawScore)
      });
    }
    if (String(r.NhanXet || "").trim()) {
      comments.push({
        id: baseId + "-nx", date: r.Ngay, classId: classItem.id,
        studentId: student.id, text: String(r.NhanXet).trim()
      });
    }
  });

  return {
    students: students, classes: classes, enrollments: enrollments,
    attendance: attendance, fees: fees, grades: grades, comments: comments,
    spreadsheetUrl: ss.getUrl()
  };
}

function normalizedText(value) {
  return String(value || "").trim().toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function findClass(classes, name) {
  var key = normalizedText(name);
  return classes.filter(function(item) {
    return normalizedText(item.name) === key || String(item.id) === String(name);
  })[0];
}

function findStudent(students, code, name) {
  var codeKey = String(code || "").trim();
  if (codeKey) {
    var byCode = students.filter(function(item) {
      return String(item.code || "").trim() === codeKey;
    })[0];
    if (byCode) return byCode;
  }
  var nameKey = normalizedText(name);
  if (!nameKey) return null;
  var matches = students.filter(function(item) {
    return normalizedText(item.name) === nameKey;
  });
  return matches.length === 1 ? matches[0] : null;
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
  if (type === "enrollment" || type === "student" || type === "class") {
    syncClassScoreSheets(ss);
  }
}

function readParentProfile(pin) {
  // Cổng PHHS chỉ đọc dữ liệu; không chạy setup/đồng bộ sheet ở mỗi lượt tra cứu.
  var data = readAllData(true);
  var student = data.students.filter(function(s) {
    return digits(s.pin) === pin;
  })[0];
  if (!student) return null;
  var studentGrades = data.grades.filter(function(x) {
    return x.studentId === student.id;
  }).map(function(grade) {
    var sameTest = data.grades.filter(function(item) {
      return item.classId === grade.classId &&
        String(item.title || "") === String(grade.title || "") &&
        String(item.date || "") === String(grade.date || "");
    });
    var average = sameTest.length
      ? sameTest.reduce(function(sum, item) {
          return sum + Number(item.score || 0);
        }, 0) / sameTest.length
      : null;
    var result = {};
    Object.keys(grade).forEach(function(key) { result[key] = grade[key]; });
    result.classAverage = average === null ? null : Number(average.toFixed(2));
    return result;
  });
  return {
    student: student,
    updatedAt: new Date().toISOString(),
    enrollments: data.enrollments.filter(function(x) {
      return x.studentId === student.id;
    }),
    attendance: data.attendance.filter(function(x) {
      return x.studentId === student.id;
    }),
    fees: data.fees.filter(function(x) {
      return x.studentId === student.id;
    }),
    grades: studentGrades,
    comments: data.comments.filter(function(x) {
      return x.studentId === student.id;
    }),
    classes: data.classes.filter(function(c) {
      return data.enrollments.some(function(x) {
        return x.studentId === student.id && x.classId === c.id;
      });
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

