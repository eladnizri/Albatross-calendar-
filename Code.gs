/**
 * ============================================================
 *  ALBATROSS — Backend (Google Apps Script)
 *  ניהול בקשות לו"ז · אחסון ב-Google Sheet
 * ============================================================
 *
 *  מה זה עושה:
 *    - doPost  : יצירת בקשה חדשה (action=create)
 *                עדכון סטטוס בקשה (action=updateStatus) — דורש טוקן
 *    - doGet   : שליפת כל הבקשות (action=list) — דורש טוקן
 *
 *  הבקשות נשמרות בגליון בשם "Requests" בקובץ ה-Google Sheet
 *  שאליו מקושר הסקריפט (או שמזוהה לפי SPREADSHEET_ID למטה).
 *
 *  הוראות פריסה מלאות נמצאות בקובץ README.md
 * ============================================================
 */

/* ---------- הגדרות ---------- */
// חייב להיות זהה לסיסמה שב-index.html (CONFIG.ADMIN_PASSWORD)
const ADMIN_TOKEN = "albatross";

// אם הסקריפט משויך לגליון (Container-bound) השאר ריק.
// אחרת הדבק כאן את מזהה ה-Spreadsheet (מתוך כתובת ה-URL שלו).
const SPREADSHEET_ID = "";

const SHEET_NAME = "Requests";
const HEADERS = [
  "id", "createdAt", "requesterName", "scheduleName",
  "date", "startTime", "endTime", "type", "status"
];

/* ---------- Sheet helpers ---------- */
function getSheet_() {
  const ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("לא נמצא Spreadsheet. הגדר SPREADSHEET_ID.");
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  return sh;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- GET: list ---------- */
function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    if (p.action === "list") {
      if (p.token !== ADMIN_TOKEN) return json_({ ok: false, error: "אין הרשאה" });
      return json_({ ok: true, requests: readAll_() });
    }
    return json_({ ok: true, message: "Albatross backend פעיל" });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/* ---------- POST: create / updateStatus ---------- */
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");

    if (body.action === "create") {
      return json_(createRequest_(body));
    }
    if (body.action === "updateStatus") {
      if (body.token !== ADMIN_TOKEN) return json_({ ok: false, error: "אין הרשאה" });
      return json_(updateStatus_(body.id, body.status));
    }
    return json_({ ok: false, error: "פעולה לא מוכרת" });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/* ---------- Core logic ---------- */
function createRequest_(body) {
  // ולידציה בסיסית
  const required = ["requesterName", "scheduleName", "date", "startTime", "endTime", "type"];
  for (const f of required) {
    if (!body[f]) return { ok: false, error: "שדה חסר: " + f };
  }
  const sh = getSheet_();
  const id = "req_" + Date.now() + "_" + Math.floor(Math.random() * 1000);
  const row = {
    id: id,
    createdAt: new Date().toISOString(),
    requesterName: body.requesterName,
    scheduleName: body.scheduleName,
    date: body.date,
    startTime: body.startTime,
    endTime: body.endTime,
    type: body.type,
    status: "pending"
  };
  sh.appendRow(HEADERS.map(h => row[h]));
  return { ok: true, id: id };
}

function readAll_() {
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const head = values[0];
  const rows = values.slice(1).map(r => {
    const o = {};
    head.forEach((h, i) => { o[h] = r[i]; });
    // תאריך/שעה עשויים לחזור כאובייקט Date — נמיר למחרוזת אחידה
    o.date = normalizeDate_(o.date);
    o.startTime = normalizeTime_(o.startTime);
    o.endTime = normalizeTime_(o.endTime);
    return o;
  });
  // חדש -> ישן
  return rows.reverse();
}

function updateStatus_(id, status) {
  if (!id || !["approved", "rejected", "pending"].includes(status)) {
    return { ok: false, error: "פרמטרים לא תקינים" };
  }
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  const head = values[0];
  const idCol = head.indexOf("id");
  const statusCol = head.indexOf("status");
  for (let i = 1; i < values.length; i++) {
    if (values[i][idCol] === id) {
      sh.getRange(i + 1, statusCol + 1).setValue(status);
      return { ok: true };
    }
  }
  return { ok: false, error: "בקשה לא נמצאה" };
}

/* ---------- Normalizers ---------- */
function normalizeDate_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(v || "");
}
function normalizeTime_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "HH:mm");
  }
  return String(v || "");
}
