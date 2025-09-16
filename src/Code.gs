/** ===== Param Parser – universal URL param extractor ===== */

const APP = {
  defaultParams: [
    'locale',
    'surface_type',
    'surface_detail',
    'surface_inter_position',
    'surface_intra_position',
  ],
  defaultCarry: ['Event name', 'Event count'], // columns to keep as metadata
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Param Parser')
    .addItem('Open parser', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar').evaluate()
    .setTitle('Param Parser');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** === API for sidebar === */
function listSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets().map(s => s.getName());
}

/**
 * Build a new sheet with parsed URL parameters.
 * cfg: { sheetName, headerRow, urlHeader, carryHeadersCSV, paramListCSV, outputName }
 */
function buildParsedTable(cfg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(cfg.sheetName);
  if (!sh) throw new Error(`Sheet "${cfg.sheetName}" not found.`);

  const headerRow = Number(cfg.headerRow || 1);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow <= headerRow) throw new Error('No data under header.');

  // headers
  const headers = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(v => (v || '').toString().trim());
  const headersLower = headers.map(h => h.toLowerCase());

  const urlHeader = (cfg.urlHeader || 'Landing page + query string').trim();
  const urlIdx = headersLower.indexOf(urlHeader.toLowerCase());
  if (urlIdx === -1) throw new Error(`URL column "${urlHeader}" not found in header row ${headerRow}.`);

  // carry columns (metadata to keep)
  const carryHeaders = (cfg.carryHeadersCSV || APP.defaultCarry.join(','))
    .split(',')
    .map(s => s.trim())
    .filter(Boolean);

  const carryIdx = carryHeaders.map(h => {
    const i = headersLower.indexOf(h.toLowerCase());
    if (i === -1) throw new Error(`Carry column "${h}" not found in header row ${headerRow}.`);
    return i; // 0-based
  });

  // params
  const params = (cfg.paramListCSV || APP.defaultParams.join(','))
    .split(',')
    .map(s => s.trim())
    .filter(Boolean);

  // read data
  const numRows = lastRow - headerRow;
  const data = sh.getRange(headerRow + 1, 1, numRows, lastCol).getValues();

  // output headers
  const outHeaders = [...carryHeaders, 'Landing path', ...params, 'Full URL'];

  // process rows
  const outRows = [];
  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    const fullUrl = toStr_(row[urlIdx]);
    if (!fullUrl) continue;

    const landingPath = fullUrl.split('?')[0];
    const qp = parseQueryParams_(fullUrl);

    const carryVals = carryIdx.map(i => toStr_(row[i]));
    const paramVals = params.map(p => (p in qp ? normalizeNumericMaybe_(qp[p]) : ''));
    outRows.push([...carryVals, landingPath, ...paramVals, fullUrl]);
  }

  // output sheet
  const outName = (cfg.outputName || (cfg.sheetName + ' - parsed')).substring(0, 99);
  let out = ss.getSheetByName(outName);
  if (!out) out = ss.insertSheet(outName);
  out.clear({ contentsOnly: true });

  // write
  if (outRows.length) {
    out.getRange(1, 1, 1, outHeaders.length).setValues([outHeaders]);
    out.getRange(2, 1, outRows.length, outHeaders.length).setValues(outRows);
  } else {
    out.getRange(1, 1, 1, outHeaders.length).setValues([outHeaders]);
  }

  // formatting
  prettyFormat_(out, outHeaders);

  return { sheetName: outName, rows: outRows.length, cols: outHeaders.length };
}

/** ===== Helpers ===== */

function parseQueryParams_(url) {
  let qs = (url.split('?')[1] || '').split('#')[0] || '';
  const params = {};
  if (!qs) return params;
  qs.split('&').forEach(pair => {
    if (!pair) return;
    const eq = pair.indexOf('=');
    let k, v;
    if (eq === -1) { k = pair; v = ''; }
    else { k = pair.slice(0, eq); v = pair.slice(eq + 1); }
    const key = safeDecode_(k);
    const val = safeDecode_(v);
    if (!(key in params)) params[key] = val; // keep first occurrence
  });
  return params;
}

function safeDecode_(s) {
  try { return decodeURIComponent((s || '').replace(/\+/g, ' ')); }
  catch(e) { return s || ''; }
}

function toStr_(v) { return v == null ? '' : v.toString(); }

function normalizeNumericMaybe_(v) {
  if (v === '' || v == null) return '';
  const n = Number(v);
  return Number.isFinite(n) ? n : v;
}

function prettyFormat_(sheet, headers) {
  const totalCols = headers.length;
  const lastRow = Math.max(sheet.getLastRow(), 2);

  // header
  const headerRange = sheet.getRange(1, 1, 1, totalCols);
  headerRange.setFontWeight('bold');
  sheet.setFrozenRows(1);

  // filter + banding
  sheet.getRange(1, 1, lastRow, totalCols).createFilter();
  sheet.getRange(1, 1, lastRow, totalCols)
       .applyRowBanding(SpreadsheetApp.BandingTheme.BLUE, true, false);

  // auto resize
  sheet.autoResizeColumns(1, totalCols);

  // try to format "Event count" as number
  const idxCount = headers.indexOf('Event count') + 1; // 1-based
  if (idxCount > 0 && lastRow > 1) {
    sheet.getRange(2, idxCount, lastRow - 1, 1).setNumberFormat('#,##0');
  }
}
