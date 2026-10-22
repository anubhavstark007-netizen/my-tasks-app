const SHEET_NAME = "Tasks";
const HEADERS = ["id", "type", "section", "text", "createdAt", "createdBy", "completedAt", "completedBy"];

function doGet(e) {
  let result;
  try {
    result = handleRequest_(e && e.parameter ? e.parameter : {});
  } catch (err) {
    result = { ok: false, error: String(err && err.message ? err.message : err) };
  }

  const callback = e && e.parameter ? String(e.parameter.callback || "") : "";
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + "(" + JSON.stringify(result) + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleRequest_(p) {
  const action = String(p.action || "list");
  const sheet = getTasksSheet_();

  if (action === "list") {
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return { ok: true, tasks: [] };
    const headers = values[0].map(String);
    const tasks = values.slice(1).filter(row => row[0]).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i] instanceof Date ? row[i].toISOString() : row[i]);
      return obj;
    });
    return { ok: true, tasks: tasks };
  }

  if (action === "add") {
    const id = clean_(p.id, 100);
    const type = clean_(p.type, 20);
    const section = clean_(p.section, 80);
    const taskText = clean_(p.text, 180);
    const by = clean_(p.by, 40);
    if (!id || !taskText || !section || !by) throw new Error("Task, section aur member name required hain.");
    if (!["member", "followup"].includes(type)) throw new Error("Invalid task type.");
    const allowedMembers = ["Akash", "Annuj", "Rishabh", "Anubhav", "Surbhi", "Ashok"];
    const allowedCategories = ["Payment Follow-up", "Customer Follow-up", "Quotation Follow-up"];
    if (type === "member" && !allowedMembers.includes(section)) throw new Error("Invalid team member.");
    if (type === "followup" && !allowedCategories.includes(section)) throw new Error("Invalid follow-up category.");
    if (findRowById_(sheet, id) > 0) return { ok: true, message: "Already saved", id: id };
    sheet.appendRow([id, type, section, taskText, new Date().toISOString(), by, "", ""]);
    return { ok: true, id: id };
  }

  if (action === "complete") {
    const id = clean_(p.id, 100);
    const row = findRowById_(sheet, id);
    if (row < 2) throw new Error("Task nahi mila. List refresh karke dobara try karein.");
    if (!sheet.getRange(row, 7).getValue()) {
      sheet.getRange(row, 7).setValue(new Date().toISOString());
      sheet.getRange(row, 8).setValue(clean_(p.by, 40) || "Unknown");
    }
    return { ok: true, id: id };
  }

  if (action === "delete") {
    const id = clean_(p.id, 100);
    const row = findRowById_(sheet, id);
    if (row < 2) throw new Error("Task pehle hi delete ho chuka hai ya nahi mila.");
    sheet.deleteRow(row);
    return { ok: true, id: id };
  }

  throw new Error("Unknown action.");
}

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Is Apps Script ko Google Sheet ke Extensions > Apps Script se open karein.");
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  sheet.clear();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
  sheet.autoResizeColumns(1, HEADERS.length);
  return "Database ready. Ab Deploy > New deployment > Web app karein.";
}

function getTasksSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Script ko Google Sheet ke saath attached hona chahiye.");
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
  const existingHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (String(existingHeaders[0]) !== "id") {
    sheet.clear();
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findRowById_(sheet, id) {
  if (!id || sheet.getLastRow() < 2) return -1;
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) return i + 2;
  }
  return -1;
}

function clean_(value, maxLength) {
  return String(value == null ? "" : value).trim().slice(0, maxLength);
}
