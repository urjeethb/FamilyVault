// ══════════════════════════════════════════════════════════════════════════════
// FamilyVault – Google Apps Script
// Handles ALL data reads and writes. No API key needed. Completely free.
//
// HOW TO DEPLOY (takes ~5 minutes):
//   1. Open your Google Sheet
//   2. Extensions → Apps Script
//   3. Delete existing code, paste this entire file
//   4. Click Save (💾 icon)
//   5. Click "Deploy" → "New Deployment"
//      - Type: Web App
//      - Execute as: Me
//      - Who has access: Anyone
//   6. Click Deploy → Authorize → Allow
//   7. Copy the Web App URL → paste into index.html
// ══════════════════════════════════════════════════════════════════════════════

const SHEET_NAME = "Investments";

// Must match the header row you created in the Google Sheet
const COLUMNS = [
  "id", "member", "type", "detail", "amount", "currentValue",
  "interestRate", "maturityDate", "grams", "purchasePrice",
  "currentPrice", "units", "nav", "avgPrice", "quantity"
];

// ── READ ─────────────────────────────────────────────────────────────────────
// Called when the React app loads (GET request)
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === "read" || !action) {
    try {
      const sheet = getSheet();
      const data  = sheet.getDataRange().getValues();

      if (data.length < 2) {
        // Sheet exists but has no data rows yet
        return jsonResponse({ rows: [] });
      }

      const header = data[0].map(h => String(h).trim());
      const rows   = data.slice(1)
        .filter(r => r.some(cell => cell !== ""))   // skip blank rows
        .map(r => {
          const obj = {};
          header.forEach((h, i) => { obj[h] = String(r[i] || ""); });
          return obj;
        });

      return jsonResponse({ rows });
    } catch (err) {
      return jsonResponse({ rows: [], error: err.message });
    }
  }

  return jsonResponse({ status: "FamilyVault Apps Script is running ✓" });
}

// ── WRITE ────────────────────────────────────────────────────────────────────
// Called when adding, editing, or deleting (POST request)
function doPost(e) {
  try {
    const body    = JSON.parse(e.postData.contents);
    const action  = body.action;
    const payload = body.payload || {};
    const sheet   = getSheet();

    if (action === "append") {
      // Add a new row
      const row = COLUMNS.map(c => payload[c] || "");
      sheet.appendRow(row);
      return jsonResponse({ ok: true });
    }

    if (action === "update") {
      // Find the row by ID and overwrite it
      const { rowIndex } = findRowById(sheet, String(payload.id));
      if (rowIndex === -1) {
        return jsonResponse({ ok: false, error: "Row not found for id: " + payload.id });
      }
      const row = COLUMNS.map(c => payload[c] || "");
      sheet.getRange(rowIndex, 1, 1, COLUMNS.length).setValues([row]);
      return jsonResponse({ ok: true });
    }

    if (action === "delete") {
      // Find the row by ID and delete it
      const { rowIndex } = findRowById(sheet, String(payload.id));
      if (rowIndex === -1) {
        return jsonResponse({ ok: false, error: "Row not found for id: " + payload.id });
      }
      sheet.deleteRow(rowIndex);
      return jsonResponse({ ok: true });
    }

    return jsonResponse({ ok: false, error: "Unknown action: " + action });

  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function getSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    // Auto-create the sheet and headers if missing
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function findRowById(sheet, id) {
  const data   = sheet.getDataRange().getValues();
  const header = data[0].map(h => String(h).trim());
  const idCol  = header.indexOf("id");

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]).trim() === id) {
      return { rowIndex: i + 1, data: data[i] }; // 1-based row index for Sheets
    }
  }
  return { rowIndex: -1 };
}

function jsonResponse(obj) {
  const output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
