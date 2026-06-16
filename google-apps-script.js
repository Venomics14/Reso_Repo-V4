// ─────────────────────────────────────────────────────────────────────────────
// BRANDING AGENCY · Resource Library — Google Apps Script
// Paste this entire file into your Google Sheet's Apps Script editor.
// Extensions → Apps Script → paste → Deploy → New Deployment → Web App
// Set "Who has access" to "Anyone" → copy the Web App URL into resource-form.html
// ─────────────────────────────────────────────────────────────────────────────

const SHEET_NAME = 'Resource Library'; // Change if your sheet tab has a different name

// ── Handle form submissions (POST) ───────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet();

    // Final duplicate check server-side (safety net)
    if (data.url) {
      const duplicate = findDuplicateUrl(sheet, data.url);
      if (duplicate) {
        return jsonResponse({
          status: 'duplicate',
          message: 'This URL already exists in the library',
          name: duplicate.name,
          addedBy: duplicate.addedBy,
          date: duplicate.date
        });
      }
    }

    // Append row
    sheet.appendRow([
      sheet.getLastRow(),            // Auto row number
      data.name        || '',
      data.category    || '',
      data.subcategory || '',
      data.description || '',
      data.url         || '',
      data.tag1        || '',
      data.tag2        || '',
      data.tag3        || '',
      data.tag4        || '',
      data.tag5        || '',
      data.contributor || '',
      data.date        || new Date().toISOString().split('T')[0],
      data.rating      || '',
      data.notes       || '',
    ]);

    return jsonResponse({ status: 'success' });

  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

// ── Handle URL duplicate check (GET) ─────────────────────────────────────────
function doGet(e) {
  try {
    const action = e.parameter.action;

    // Return the full library (used by the Browse / Search section of the page)
    if (action === 'getResources') {
      const sheet = getReadSheet();
      return jsonResponse({
        status: 'success',
        sheet: sheet ? sheet.getName() : null,
        resources: sheet ? getAllResources(sheet) : []
      });
    }

    // Diagnostics — open this URL in your browser to see what the script sees
    if (action === 'debug') {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const tabs = ss.getSheets().map((s) => ({
        name: s.getName(),
        rows: s.getLastRow(),
        cols: s.getLastColumn(),
        firstRow: s.getLastRow() ? s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0] : []
      }));
      const read = getReadSheet();
      return jsonResponse({
        status: 'debug',
        spreadsheet: ss.getName(),
        expectedTab: SHEET_NAME,
        readingFrom: read ? read.getName() : null,
        resourceCount: read ? getAllResources(read).length : 0,
        tabs
      });
    }

    if (action === 'checkUrl') {
      const url = (e.parameter.url || '').trim().replace(/\/$/, '').toLowerCase();
      if (!url) return jsonResponse({ duplicate: false });

      const sheet = getOrCreateSheet();
      const duplicate = findDuplicateUrl(sheet, url);

      if (duplicate) {
        return jsonResponse({
          duplicate: true,
          name: duplicate.name,
          addedBy: duplicate.addedBy,
          date: duplicate.date
        });
      } else {
        return jsonResponse({ duplicate: false });
      }
    }

    // Default GET — health check
    return jsonResponse({ status: 'ok' });

  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

// ── Check if a URL already exists in the sheet ───────────────────────────────
function findDuplicateUrl(sheet, incomingUrl) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null; // No data rows yet

  // URL is column 6, Name is column 2, Added By is column 12, Date is column 13
  const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();

  const normalise = (url) => (url || '').toString().trim().replace(/\/$/, '').toLowerCase();
  const incoming = normalise(incomingUrl);

  for (const row of data) {
    const existingUrl = normalise(row[5]); // Column F (index 5)
    if (existingUrl && existingUrl === incoming) {
      return {
        name:    row[1]  || '',   // Column B
        addedBy: row[11] || '',   // Column L
        date:    row[12] || ''    // Column M
      };
    }
  }
  return null;
}

// ── Pick the sheet that actually holds the library data ──────────────────────
// Prefers the "Resource Library" tab, but falls back to whatever tab has data,
// so the page works even if the data lives in "Sheet1" or another tab.
function getReadSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const preferred = ss.getSheetByName(SHEET_NAME);
  if (preferred && preferred.getLastRow() >= 2) return preferred;

  const sheets = ss.getSheets();
  for (const s of sheets) if (s.getLastRow() >= 2) return s; // first tab with rows
  for (const s of sheets) if (s.getLastRow() >= 1) return s;
  return preferred || sheets[0] || null;
}

// ── Read every resource row → array of objects (newest first) ────────────────
// Adapts to the sheet's actual layout:
//   • detects whether row 1 is a header (and maps columns by header name), and
//   • when there's no header, detects a leading "#" number column and shifts.
function getAllResources(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 15);
  if (lastRow < 1) return [];

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  const fmtDate = (v) => {
    if (!v) return '';
    if (Object.prototype.toString.call(v) === '[object Date]') {
      const y = v.getFullYear();
      const m = String(v.getMonth() + 1).padStart(2, '0');
      const d = String(v.getDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }
    return v.toString();
  };
  const clean = (v) => (v || '').toString().trim();

  // — Is row 1 a header row? —
  const firstRow = values[0].map((c) => clean(c).toLowerCase());
  const headerHints = ['resource name', 'name', 'category', 'description', 'url', 'link / url'];
  const hasHeader = firstRow.some((c) => headerHints.includes(c));

  let startIdx, map, tagCols;

  if (hasHeader) {
    const find = (names) => {
      for (let i = 0; i < firstRow.length; i++) if (names.includes(firstRow[i])) return i;
      return -1;
    };
    map = {
      name: find(['resource name', 'name']),
      category: find(['category']),
      subcategory: find(['sub-category', 'subcategory', 'sub category']),
      description: find(['description']),
      url: find(['url', 'link', 'link / url']),
      addedBy: find(['added by', 'contributor', 'your name']),
      date: find(['date added', 'date']),
      rating: find(['rating']),
      notes: find(['notes', 'notes / tips', 'notes/tips']),
    };
    // Tag columns: any header containing "hashtag" or "tag"
    tagCols = [];
    firstRow.forEach((h, i) => { if (/hashtag|tag/.test(h)) tagCols.push(i); });
    if (!tagCols.length && map.url >= 0) tagCols = [map.url + 1, map.url + 2, map.url + 3, map.url + 4, map.url + 5];
    startIdx = 1;
  } else {
    // No header. Detect a leading number/"#" column (added by the form's doPost).
    const dataRows = values; // all rows are data
    const col0Numeric = dataRows.every((r) => {
      const v = clean(r[0]);
      return v === '' || /^\d+$/.test(v);
    });
    const off = col0Numeric ? 1 : 0;
    map = {
      name: off + 0, category: off + 1, subcategory: off + 2, description: off + 3,
      url: off + 4, addedBy: off + 10, date: off + 11, rating: off + 12, notes: off + 13,
    };
    tagCols = [off + 5, off + 6, off + 7, off + 8, off + 9];
    startIdx = 0;
  }

  const at = (row, idx) => (idx != null && idx >= 0 ? clean(row[idx]) : '');

  const resources = [];
  for (let i = startIdx; i < values.length; i++) {
    const row = values[i];
    const name = at(row, map.name);
    if (!name) continue; // skip blank rows
    resources.push({
      name,
      category:    at(row, map.category),
      subcategory: at(row, map.subcategory),
      description: at(row, map.description),
      url:         at(row, map.url),
      tags: tagCols.map((c) => clean(row[c])).filter((t) => t.length > 0),
      addedBy:     at(row, map.addedBy),
      date:        map.date >= 0 ? fmtDate(row[map.date]) : '',
      rating:      at(row, map.rating),
      notes:       at(row, map.notes),
    });
  }

  return resources.reverse(); // newest first
}

// ── Helper: return JSON response ──────────────────────────────────────────────
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Creates the sheet with headers if it doesn't exist yet ───────────────────
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = [
      '#', 'Resource Name', 'Category', 'Sub-Category', 'Description',
      'URL', 'Hashtag 1', 'Hashtag 2', 'Hashtag 3', 'Hashtag 4', 'Hashtag 5',
      'Added By', 'Date Added', 'Rating', 'Notes'
    ];
    sheet.appendRow(headers);

    // Style the header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1a1a2e');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setFontFamily('Arial');
    headerRange.setFontSize(10);
    headerRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    // Column widths
    sheet.setColumnWidth(1, 40);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 130);
    sheet.setColumnWidth(4, 130);
    sheet.setColumnWidth(5, 320);
    sheet.setColumnWidth(6, 220);
    sheet.setColumnWidths(7, 5, 120);
    sheet.setColumnWidth(12, 130);
    sheet.setColumnWidth(13, 110);
    sheet.setColumnWidth(14, 70);
    sheet.setColumnWidth(15, 250);
  }

  return sheet;
}

// Run this once manually to initialise the sheet and grant permissions
function setup() {
  getOrCreateSheet();
}
