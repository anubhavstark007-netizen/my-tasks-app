const SHEET_NAME = "Tasks";
const SHIPMENT_SHEET_NAME = "Shipments";

const HEADERS = [
  "id", "type", "section", "text",
  "createdAt", "createdBy", "completedAt", "completedBy"
];

const SHIPMENT_HEADERS = [
  "id", "referenceNo", "shipper", "consignee",
  "origin", "destination", "vesselFlight",
  "containerNo", "blMawbNo", "atd", "eta", "ata",
  "status", "remarks", "createdAt", "createdBy",
  "updatedAt", "updatedBy"
];

const MEMBERS = ["Akash", "Annuj", "Rishabh", "Anubhav", "Surbhi", "Ashok"];
const CATEGORIES = ["Payment Follow-up", "Customer Follow-up", "Quotation Follow-up"];

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

  // Shipment requests use their own sheet.
  if (action.indexOf("shipment") === 0) {
    return handleShipmentRequest_(action, p);
  }

  const sheet = getTasksSheet_();

  if (action === "list") {
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return { ok: true, tasks: [] };

    const headers = values[0].map(String);
    const tasks = values.slice(1).filter(row => row[0]).map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = row[i] instanceof Date ? row[i].toISOString() : row[i];
      });
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

    if (!id || !taskText || !section || !by) {
      throw new Error("Task, section aur member name required hain.");
    }
    if (!["member", "followup"].includes(type)) {
      throw new Error("Invalid task type.");
    }
    if (type === "member" && !MEMBERS.includes(section)) {
      throw new Error("Invalid team member.");
    }
    if (type === "followup" && !CATEGORIES.includes(section)) {
      throw new Error("Invalid follow-up category.");
    }

    if (findRowById_(sheet, id) > 0) {
      return { ok: true, message: "Already saved", id: id };
    }

    sheet.appendRow([
      id, type, section, taskText,
      new Date().toISOString(), by, "", ""
    ]);
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

  throw new Error("Unknown action: " + action);
}

// ---------- SHIPMENT TRACK ----------

function handleShipmentRequest_(action, p) {
  const sheet = getShipmentsSheet_();

  if (action === "shipmentList") {
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return { ok: true, shipments: [] };

    const headers = values[0].map(String);
    const shipments = values.slice(1).filter(row => row[0]).map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        let value = row[i];
        if (value instanceof Date) {
          value = h === "atd" || h === "eta" || h === "ata"
            ? Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd")
            : value.toISOString();
        }
        obj[h] = value == null ? "" : value;
      });
      return obj;
    });

    return { ok: true, shipments: shipments };
  }

  if (action === "shipmentAdd" || action === "shipmentUpdate") {
    const id = clean_(p.id, 100) || Utilities.getUuid();
    const by = clean_(p.by, 40) || "Unknown";
    const now = new Date().toISOString();
    const existingRow = findShipmentRowById_(sheet, id);

    if (action === "shipmentUpdate" && existingRow < 2) {
      throw new Error("Shipment nahi mila. List refresh karke dobara try karein.");
    }

    const shipment = {
      id: id,
      referenceNo: clean_(p.referenceNo, 100),
      shipper: clean_(p.shipper, 180),
      consignee: clean_(p.consignee, 180),
      origin: clean_(p.origin, 120),
      destination: clean_(p.destination, 120),
      vesselFlight: clean_(p.vesselFlight, 120),
      containerNo: clean_(p.containerNo, 180),
      blMawbNo: clean_(p.blMawbNo, 120),
      atd: clean_(p.atd, 20),
      eta: clean_(p.eta, 20),
      ata: clean_(p.ata, 20),
      status: clean_(p.status, 80),
      remarks: clean_(p.remarks, 1000)
    };

    if (!shipment.referenceNo) throw new Error("Reference No. required hai.");
    if (!shipment.origin || !shipment.destination) {
      throw new Error("Origin aur Destination required hain.");
    }

    const createdAt = existingRow >= 2
      ? sheet.getRange(existingRow, 15).getValue()
      : now;
    const createdBy = existingRow >= 2
      ? sheet.getRange(existingRow, 16).getValue()
      : by;

    const rowData = [
      shipment.id, shipment.referenceNo, shipment.shipper,
      shipment.consignee, shipment.origin, shipment.destination,
      shipment.vesselFlight, shipment.containerNo, shipment.blMawbNo,
      shipment.atd, shipment.eta, shipment.ata, shipment.status,
      shipment.remarks, createdAt, createdBy, now, by
    ];

    if (existingRow >= 2) {
      sheet.getRange(existingRow, 1, 1, SHIPMENT_HEADERS.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    return { ok: true, id: id, message: "Shipment saved successfully." };
  }

  if (action === "shipmentDelete") {
    const id = clean_(p.id, 100);
    const row = findShipmentRowById_(sheet, id);
    if (row < 2) throw new Error("Shipment nahi mila ya pehle hi delete ho chuka hai.");

    sheet.deleteRow(row);
    return { ok: true, id: id };
  }

  throw new Error("Unknown shipment action: " + action);
}

// ---------- SAFE SHEET SETUP ----------

// Is function ko existing project par run karna safe hai.
// Ye Tasks sheet ko clear nahi karta.
function setupShipmentDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("Apps Script ko Google Sheet ke Extensions > Apps Script se open karein.");
  }

  let sheet = ss.getSheetByName(SHIPMENT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHIPMENT_SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, SHIPMENT_HEADERS.length).setValues([SHIPMENT_HEADERS]);
  } else {
    const firstCell = String(sheet.getRange(1, 1).getValue() || "");
    if (firstCell !== "id") {
      throw new Error(
        "Shipments tab mein pehle se data hai aur header alag hai. Data safety ke liye sheet ko change nahi kiya."
      );
    }
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, SHIPMENT_HEADERS.length).setFontWeight("bold");
  sheet.autoResizeColumns(1, SHIPMENT_HEADERS.length);

  return "Shipments tab ready. Existing Tasks data ko touch nahi kiya gaya.";
}

// ---------- SHARED HELPERS ----------

function getTasksSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Script ko Google Sheet ke saath attached hona chahiye.");

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  } else {
    const firstCell = String(sheet.getRange(1, 1).getValue() || "");
    if (firstCell !== "id") {
      throw new Error("Tasks sheet ka header expected format mein nahi hai. Data safety ke liye sheet clear nahi ki gayi.");
    }
  }

  return sheet;
}

function getShipmentsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Script ko Google Sheet ke saath attached hona chahiye.");

  let sheet = ss.getSheetByName(SHIPMENT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHIPMENT_SHEET_NAME);
    sheet.getRange(1, 1, 1, SHIPMENT_HEADERS.length).setValues([SHIPMENT_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, SHIPMENT_HEADERS.length).setFontWeight("bold");
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, SHIPMENT_HEADERS.length).setValues([SHIPMENT_HEADERS]);
    sheet.setFrozenRows(1);
  } else {
    const headers = sheet.getRange(1, 1, 1, SHIPMENT_HEADERS.length).getValues()[0];
    if (String(headers[0]) !== "id") {
      throw new Error("Shipments sheet ka header expected format mein nahi hai.");
    }
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

function findShipmentRowById_(sheet, id) {
  return findRowById_(sheet, id);
}

function clean_(value, maxLength) {
  return String(value == null ? "" : value).trim().slice(0, maxLength);
}
