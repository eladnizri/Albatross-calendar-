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

// כתובת המייל שתקבל התראה על כל בקשה חדשה (רק אתה). השאר ריק כדי לכבות.
const NOTIFY_EMAIL = "eladn2006@gmail.com";

// שם היומן שממנו נמשך הלו"ז השבועי במסך "צפייה ביומן".
// חייב להיות זהה לשם היומן ב-Google Calendar, ומשותף/בבעלות החשבון שמריץ סקריפט זה.
const CALENDAR_NAME = "יחידת אלבטרוס";

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
    if (p.action === "weekEvents") {
      // צפייה בלבד בלו"ז השבועי — פתוח לכל חברי היחידה (ללא טוקן)
      return json_(weekEvents_());
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
  notifyAdmin_(row); // התראת מייל אליך בלבד — לא מפילה את הבקשה אם נכשלת
  return { ok: true, id: id };
}

/* ---------- התראת מייל למנהל על בקשה חדשה ---------- */
function notifyAdmin_(row) {
  if (!NOTIFY_EMAIL) return;
  try {
    const typeLine = row.type ? "\nסוג: " + row.type : "";
    const body =
      "התקבלה בקשת לו\"ז חדשה ליחידת אלבטרוס:\n\n" +
      "שם הלו\"ז: " + row.scheduleName + "\n" +
      "מבקש: " + row.requesterName + "\n" +
      "תאריך: " + row.date + "\n" +
      "שעות: " + row.startTime + "–" + row.endTime +
      typeLine + "\n\n" +
      "היכנס למסך הבקרה כדי לאשר או לדחות.";
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: "🕊️ בקשת לו\"ז חדשה: " + row.scheduleName,
      body: body
    });
  } catch (err) {
    // לא עוצרים את יצירת הבקשה גם אם שליחת המייל נכשלה
  }
}

/* ---------- צפייה ביומן: אירועי השבוע הנוכחי ---------- */
function weekEvents_() {
  const cals = CalendarApp.getCalendarsByName(CALENDAR_NAME);
  if (!cals || !cals.length) {
    return { ok: false, error: 'לא נמצא יומן בשם "' + CALENDAR_NAME + '". ודא שהוא משותף עם החשבון שמריץ את הסקריפט.' };
  }
  const cal = cals[0];
  const tz = Session.getScriptTimeZone();

  // תחילת השבוע = יום ראשון האחרון בשעה 00:00 · סוף = יום ראשון הבא (טווח חצי-פתוח)
  const weekStart = new Date();
  weekStart.setHours(0, 0, 0, 0);
  weekStart.setDate(weekStart.getDate() - weekStart.getDay()); // getDay(): 0 = ראשון
  const weekEnd = new Date(weekStart);
  weekEnd.setDate(weekEnd.getDate() + 7);

  const events = cal.getEvents(weekStart, weekEnd).map(function (ev) {
    const allDay = ev.isAllDayEvent();
    const s = ev.getStartTime();
    const e = ev.getEndTime();
    return {
      title: ev.getTitle(),
      allDay: allDay,
      startLabel: allDay ? "" : Utilities.formatDate(s, tz, "HH:mm"),
      endLabel:   allDay ? "" : Utilities.formatDate(e, tz, "HH:mm"),
      location: ev.getLocation() || "",
      // הימים (yyyy-MM-dd) שהאירוע משתרע עליהם בתוך השבוע — כדי לשבץ אותו בכל יום רלוונטי
      days: eventDays_(ev, weekStart, weekEnd, tz)
    };
  });

  return {
    ok: true,
    calendarName: CALENDAR_NAME,
    weekStart: Utilities.formatDate(weekStart, tz, "yyyy-MM-dd"),
    events: events
  };
}

// מחזיר את רשימת התאריכים (yyyy-MM-dd) שאירוע נוגע בהם, מקוצץ לגבולות השבוע
function eventDays_(ev, weekStart, weekEnd, tz) {
  const allDay = ev.isAllDayEvent();
  const s = ev.getStartTime();
  const e = ev.getEndTime();
  // באירוע "כל היום" זמן הסיום הוא חצות של היום שאחרי — לכן היום האחרון הוא סוף פחות מילישנייה
  const lastMs = allDay ? e.getTime() - 1 : e.getTime();
  let cur = new Date(Math.max(s.getTime(), weekStart.getTime()));
  cur.setHours(0, 0, 0, 0);
  const last = Math.min(lastMs, weekEnd.getTime() - 1);
  const days = [];
  let guard = 0;
  while (cur.getTime() <= last && guard++ < 8) {
    days.push(Utilities.formatDate(cur, tz, "yyyy-MM-dd"));
    cur.setDate(cur.getDate() + 1);
  }
  return days;
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
