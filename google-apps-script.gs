// Baby registry backend. Paste this into Extensions > Apps Script in your Google Sheet,
// then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
// Claims appear in the "Claims" tab. Delete a row there to clear someone's claim.

const SHEET = "Claims";

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(["Gift id", "Name", "Guest id", "Claimed at"]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function rows_(sh) {
  const n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, 4).getValues();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Stops names like "=SUM(...)" being treated as spreadsheet formulas.
function plain_(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function doGet(e) {
  const guest = String(e.parameter.guest || "");
  const claims = {};
  rows_(sheet_()).forEach(function (r) {
    if (r[0]) claims[r[0]] = { name: String(r[1]), mine: !!guest && r[2] === guest };
  });
  return json_({ claims: claims });
}

function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents); } catch (err) {}
  const id = String(b.id || "");
  const guest = String(b.guest || "");
  const name = String(b.name || "").trim().slice(0, 60);
  if (!/^[a-z0-9-]{1,64}$/.test(id) || !/^[A-Za-z0-9-]{8,64}$/.test(guest)) return json_({ error: "bad_request" });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const rows = rows_(sh);
    const i = rows.findIndex(function (r) { return r[0] === id; });

    if (b.action === "claim") {
      if (!name) return json_({ error: "name_required" });
      if (i >= 0) return rows[i][2] === guest ? json_({ ok: true }) : json_({ error: "taken" });
      if (rows.length >= 200) return json_({ error: "full" });
      sh.appendRow([id, plain_(name), guest, new Date()]);
      return json_({ ok: true });
    }

    if (b.action === "unclaim") {
      if (i < 0) return json_({ ok: true });
      if (rows[i][2] !== guest) return json_({ error: "forbidden" });
      sh.deleteRow(i + 2);
      return json_({ ok: true });
    }

    return json_({ error: "bad_request" });
  } finally {
    lock.releaseLock();
  }
}
