// ══════════════════════════════════════════════════════════════════════════════
// FamilyVault – Google Apps Script  (Fixed: handles CORS + text/plain POST)
//
// HOW TO DEPLOY:
//   1. Open your Google Sheet → Extensions → Apps Script
//   2. Delete all existing code, paste this entire file
//   3. Click Save (💾)
//   4. Click Deploy → New Deployment
//      - Type: Web App
//      - Execute as: Me
//      - Who has access: Anyone
//   5. Click Deploy → Authorize → Allow
//   6. Copy the Web App URL → paste into index.html
//
// IMPORTANT — after ANY code change you must redeploy:
//   Deploy → Manage Deployments → Edit (pencil) → New Version → Deploy
// ══════════════════════════════════════════════════════════════════════════════

const SHEET_NAME = "Investments";

const COLUMNS = [
  "id", "member", "type", "detail", "amount", "currentValue",
  "interestRate", "maturityDate", "grams", "purchasePrice",
  "currentPrice", "units", "nav", "avgPrice", "quantity"
];

// ── READ ──────────────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    const sheet = getSheet();
    const data  = sheet.getDataRange().getValues();

    if (data.length < 2) return ok({ rows: [] });

    const header = data[0].map(h => String(h).trim());
    const rows   = data.slice(1)
      .filter(r => r.some(c => c !== ""))
      .map(r => {
        const obj = {};
        header.forEach((h, i) => { obj[h] = String(r[i] || ""); });
        return obj;
      });

    return ok({ rows });
  } catch (err) {
    return ok({ rows: [], error: err.message });
  }
}

// ── WRITE ─────────────────────────────────────────────────────────────────────
// Receives POST with Content-Type: text/plain (avoids CORS preflight 405 error)
function doPost(e) {
  try {
    const body    = JSON.parse(e.postData.contents);
    const action  = body.action;
    const payload = body.payload || {};
    const sheet   = getSheet();

    if (action === "append") {
      sheet.appendRow(COLUMNS.map(c => payload[c] || ""));
      return ok({ ok: true });
    }

    if (action === "update") {
      const { rowIndex } = findById(sheet, String(payload.id));
      if (rowIndex < 0) return ok({ ok: false, error: "ID not found: " + payload.id });
      sheet.getRange(rowIndex, 1, 1, COLUMNS.length)
           .setValues([COLUMNS.map(c => payload[c] || "")]);
      return ok({ ok: true });
    }

    if (action === "delete") {
      const { rowIndex } = findById(sheet, String(payload.id));
      if (rowIndex < 0) return ok({ ok: false, error: "ID not found: " + payload.id });
      sheet.deleteRow(rowIndex);
      return ok({ ok: true });
    }

    return ok({ ok: false, error: "Unknown action: " + action });

  } catch (err) {
    return ok({ ok: false, error: err.message });
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function getSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findById(sheet, id) {
  const data  = sheet.getDataRange().getValues();
  const idCol = data[0].map(h => String(h).trim()).indexOf("id");
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]).trim() === id) return { rowIndex: i + 1 };
  }
  return { rowIndex: -1 };
}

function ok(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
