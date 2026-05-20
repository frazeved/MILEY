// GABRIEL — Map & Finance Specialist routes.
// Task 1: Map Sync — collects new Styles + their POs from the Tradestone
// Database and Production & PO Database tabs and appends them to Anthro Map 2026.
// Full spec: github.com/frazeved/GABRIEL/blob/main/SPEC.md

const { Router } = require('express');
const { google } = require('googleapis');

const router = Router();

// ---- Sources ----------------------------------------------------------------
const PROD_DB_SPREADSHEET_ID    = '1y0iL7PJldbVQmPIAnJi9wvA2hvjB8_aK2bU2kxvUf5Q';
const TRADESTONE_DB_TAB         = 'Tradestone Database';
const PROD_DB_TAB               = 'Production & PO Database';

// ---- Target -----------------------------------------------------------------
const ANTHRO_MAP_SPREADSHEET_ID = '1W88MKYr-q9g3F2fLFu2jjxvXzigK12PWohSVMsOQst4';
const ANTHRO_MAP_TAB            = 'Anthro Map 2026';

// ---- Filters / keys ---------------------------------------------------------
const CANCEL_DATE_YEAR          = 2026;
const EXCLUDE_STATUS_PATTERN    = /cancel/i;
const STYLE_KEY_TARGET          = 'STYLE #';
const STYLE_KEY_TRADESTONE      = 'Vendor Style #';
const STYLE_KEY_PROD_DB         = 'ORIGINAL STYLE#';
const CANCEL_DATE_COL           = 'Cancel Date';
const STATUS_COL                = 'Status';

// ---- Field map: Anthro Map column -> source rule ---------------------------
const TARGET_FIELD_MAP = {
  'STYLE #':           { tradestone: 'Vendor Style #' },
  'Purchase Order':    { tradestone: 'PO#' },
  'Style Description': { tradestone: 'Style Description' },
  'Total Qty':         { tradestone: 'Total Qty' },
  'Ship Date':         { tradestone: 'Ship Date' },
  'Cancel Date':       { tradestone: 'Cancel Date' },
  'STATUS':            { constant:   'PRODUCTION' },
  'SUPPLIER':          { prod_db:    'SUPPLIER' },
  'CATEGORY':          { prod_db:    'CATEGORY' },
  'SUB-CATEGORY':      { prod_db:    'SUB-CATEGORY' },
  'HTS CODE':          { prod_db:    'HTS CODE' },
  'DUTY':              { prod_db:    'DUTY' },
  'COST':              { prod_db:    'COST' },
  'FREIGHT':           { prod_db:    'FREIGHT' },
  'YEAR':              { prod_db:    'YEAR' },
};

// ---- Helpers ----------------------------------------------------------------
function sheetsClient(readonly = true) {
  const sa   = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const auth = new google.auth.GoogleAuth({
    credentials: sa,
    scopes: [readonly
      ? 'https://www.googleapis.com/auth/spreadsheets.readonly'
      : 'https://www.googleapis.com/auth/spreadsheets'],
  });
  return google.sheets({ version: 'v4', auth });
}

function requireCreds(res) {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    res.status(500).json({ error: 'Google credentials not configured' });
    return false;
  }
  return true;
}

const norm = h => String(h || '').trim();

function indexHeaders(header) {
  const m = new Map();
  header.forEach((h, i) => m.set(norm(h), i));
  return m;
}

function getCell(row, header, name) {
  const i = indexHeaders(header).get(norm(name));
  return i == null ? '' : (row[i] ?? '');
}

function parseYear(value) {
  const s = String(value || '').trim();
  if (!s) return null;
  let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (m) { let y = parseInt(m[3], 10); if (y < 100) y += 2000; return y; }
  m = s.match(/^(\d{4})[\/\-]\d{1,2}[\/\-]\d{1,2}/);
  if (m) return parseInt(m[1], 10);
  m = s.match(/(\d{4})/);
  return m ? parseInt(m[1], 10) : null;
}

function cleanValue(target, value) {
  const s = String(value);
  if (target === 'Purchase Order') return s.replace(/^0+(\d)/, '$1');
  return s;
}

async function readTab(sheets, spreadsheetId, tab) {
  const res = await sheets.spreadsheets.values.get({ spreadsheetId, range: tab });
  const values = res.data.values || [];
  if (values.length === 0) return { header: [], rows: [] };
  const [header, ...rows] = values;
  return { header: header.map(norm), rows };
}

async function appendRows(sheets, spreadsheetId, tab, rows) {
  if (rows.length === 0) return;
  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: `${tab}!A1`,
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: rows },
  });
}

function buildRow(mapHeader, tsHeader, tsRow, prodHeader, prodRow) {
  return mapHeader.map(target => {
    const rule = TARGET_FIELD_MAP[target];
    if (!rule) return '';
    if (rule.constant != null) return rule.constant;
    if (rule.tradestone) {
      const v = getCell(tsRow, tsHeader, rule.tradestone);
      if (v !== '' && v != null) return cleanValue(target, v);
    }
    if (rule.prod_db && prodRow) {
      const v = getCell(prodRow, prodHeader, rule.prod_db);
      if (v !== '' && v != null) return cleanValue(target, v);
    }
    return '';
  });
}

// ---- The map-sync run -------------------------------------------------------
async function runMapSync(opts = {}) {
  const sheetsRO = sheetsClient(true);
  const sheetsRW = sheetsClient(false);
  const logs = [];
  const log = s => { logs.push(s); console.log(`[gabriel] ${s}`); };

  log(`Reading ${TRADESTONE_DB_TAB}...`);
  const tradestone = await readTab(sheetsRO, PROD_DB_SPREADSHEET_ID, TRADESTONE_DB_TAB);

  log(`Reading ${PROD_DB_TAB}...`);
  const prodDb = await readTab(sheetsRO, PROD_DB_SPREADSHEET_ID, PROD_DB_TAB);

  log(`Reading ${ANTHRO_MAP_TAB}...`);
  const anthroMap = await readTab(sheetsRO, ANTHRO_MAP_SPREADSHEET_ID, ANTHRO_MAP_TAB);

  const prodStyleCol = indexHeaders(prodDb.header).get(norm(STYLE_KEY_PROD_DB));
  const prodByStyle = new Map();
  if (prodStyleCol != null) {
    for (const r of prodDb.rows) {
      const s = String(r[prodStyleCol] || '').trim();
      if (s && !prodByStyle.has(s)) prodByStyle.set(s, r);
    }
  }

  const existingStyleCol = indexHeaders(anthroMap.header).get(norm(STYLE_KEY_TARGET));
  const existing = new Set();
  if (existingStyleCol != null) {
    for (const r of anthroMap.rows) {
      const s = String(r[existingStyleCol] || '').trim();
      if (s) existing.add(s);
    }
  }

  const tsStyleCol  = indexHeaders(tradestone.header).get(norm(STYLE_KEY_TRADESTONE));
  const tsCancelCol = indexHeaders(tradestone.header).get(norm(CANCEL_DATE_COL));
  const tsStatusCol = indexHeaders(tradestone.header).get(norm(STATUS_COL));
  if (tsStyleCol == null)  throw new Error(`"${STYLE_KEY_TRADESTONE}" not found in ${TRADESTONE_DB_TAB}`);
  if (tsCancelCol == null) throw new Error(`"${CANCEL_DATE_COL}" not found in ${TRADESTONE_DB_TAB}`);

  const newRows = [];
  const newStyles = new Set();
  let skipYear = 0, skipCancel = 0, skipBlank = 0, skipExist = 0;

  for (const row of tradestone.rows) {
    if (parseYear(row[tsCancelCol]) !== CANCEL_DATE_YEAR) { skipYear++; continue; }
    if (tsStatusCol != null && EXCLUDE_STATUS_PATTERN.test(String(row[tsStatusCol] || ''))) {
      skipCancel++; continue;
    }
    const style = String(row[tsStyleCol] || '').trim();
    if (!style) { skipBlank++; continue; }
    if (existing.has(style)) { skipExist++; continue; }
    newStyles.add(style);
    newRows.push(buildRow(anthroMap.header, tradestone.header, row, prodDb.header, prodByStyle.get(style)));
  }

  const summary = {
    tradestoneRows:   tradestone.rows.length,
    skippedNotYear:   skipYear,
    skippedCanceled:  skipCancel,
    skippedBlank:     skipBlank,
    skippedExisting:  skipExist,
    newStyles:        newStyles.size,
    newRows:          newRows.length,
  };
  log(`Summary: ${JSON.stringify(summary)}`);

  if (newRows.length > 0 && !opts.dryRun) {
    await appendRows(sheetsRW, ANTHRO_MAP_SPREADSHEET_ID, ANTHRO_MAP_TAB, newRows);
    log(`Appended ${newRows.length} rows to ${ANTHRO_MAP_TAB}.`);
  } else if (opts.dryRun) {
    log(`Dry-run mode — no rows appended.`);
  } else {
    log(`Nothing to append.`);
  }

  return { ok: true, summary, logs };
}

// ---- Routes -----------------------------------------------------------------
router.get('/', (_req, res) => res.json({ ok: true, agent: 'gabriel', tasks: ['map-sync'] }));

router.post('/map-sync', async (req, res) => {
  if (!requireCreds(res)) return;
  try {
    const dryRun = String(req.query.dryRun || '').toLowerCase() === 'true';
    const result = await runMapSync({ dryRun });
    res.json(result);
  } catch (err) {
    console.error('[gabriel/map-sync]', err);
    res.status(500).json({ ok: false, error: err.message });
  }
});

// Convenience GET for quick browser preview (always dry-run).
router.get('/map-sync/preview', async (_req, res) => {
  if (!requireCreds(res)) return;
  try {
    const result = await runMapSync({ dryRun: true });
    res.json(result);
  } catch (err) {
    console.error('[gabriel/map-sync/preview]', err);
    res.status(500).json({ ok: false, error: err.message });
  }
});

module.exports = router;
