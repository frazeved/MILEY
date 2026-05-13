const { Router } = require('express');
const { google } = require('googleapis');

const router   = Router();
const SHEET_ID = '1y0iL7PJldbVQmPIAnJi9wvA2hvjB8_aK2bU2kxvUf5Q';
const TAB      = 'SAMPLE DATABASE NEW';

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

function nowFormatted() {
  const d  = new Date();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${mm}/${dd}/${d.getFullYear()}`;
}

// GET /api/paul/database — all rows from SAMPLE DATABASE NEW
router.get('/database', async (req, res) => {
  if (!requireCreds(res)) return;
  try {
    const sheets = sheetsClient(true);
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `'${TAB}'`,
      valueRenderOption: 'FORMATTED_VALUE',
      dateTimeRenderOption: 'FORMATTED_STRING',
    });
    const rows    = r.data.values || [];
    const headers = (rows[0] || []).map(h => String(h).trim());
    const data    = rows.slice(1).map(row =>
      Object.fromEntries(headers.map((h, i) => [h, row[i] ?? '']))
    );
    res.json({ headers, data });
  } catch (e) {
    console.error('[paul/database]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// GET /api/paul/search?tracking=XXX — find row by tracking number
router.get('/search', async (req, res) => {
  const tracking = String(req.query.tracking || '').trim().toUpperCase();
  if (!tracking) return res.status(400).json({ error: 'tracking required' });
  if (!requireCreds(res)) return;
  try {
    const sheets = sheetsClient(true);
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `'${TAB}'`,
      valueRenderOption: 'FORMATTED_VALUE',
      dateTimeRenderOption: 'FORMATTED_STRING',
    });
    const rows        = r.data.values || [];
    const headers     = (rows[0] || []).map(h => String(h).trim());
    const trackingCol = headers.findIndex(h => h.toUpperCase() === 'TRACKING');
    if (trackingCol < 0) return res.status(500).json({ error: 'TRACKING column not found' });

    const rowIdx = rows.slice(1).findIndex(r =>
      String(r[trackingCol] || '').trim().toUpperCase() === tracking
    );
    if (rowIdx < 0) return res.status(404).json({ error: `Tracking "${tracking}" not found` });

    const row    = rows.slice(1)[rowIdx];
    const result = Object.fromEntries(headers.map((h, i) => [h, row[i] ?? '']));
    res.json({ sample: result });
  } catch (e) {
    console.error('[paul/search]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// POST /api/paul/save — upsert by TRACKING number
router.post('/save', async (req, res) => {
  if (!requireCreds(res)) return;
  const body = req.body || {};
  const tracking = String(body['TRACKING'] || '').trim();
  if (!tracking) return res.status(400).json({ error: 'TRACKING is required' });

  try {
    const sheets = sheetsClient(false);
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `'${TAB}'`,
      valueRenderOption: 'FORMATTED_VALUE',
    });
    const rows        = r.data.values || [];
    const headers     = (rows[0] || []).map(h => String(h).trim());
    const trackingCol = headers.findIndex(h => h.toUpperCase() === 'TRACKING');
    const createdCol  = headers.findIndex(h => h.toUpperCase() === 'CREATED DATE');
    const updatedCol  = headers.findIndex(h => h.toUpperCase() === 'LAST UPDATED');

    const rowIdx = trackingCol >= 0
      ? rows.slice(1).findIndex(r => String(r[trackingCol] || '').trim().toUpperCase() === tracking.toUpperCase())
      : -1;

    const now = nowFormatted();

    const buildRow = (isNew) => headers.map(h => {
      const k = h.toUpperCase();
      if (k === 'CREATED DATE') return isNew ? now : (rows.slice(1)[rowIdx]?.[createdCol] ?? now);
      if (k === 'LAST UPDATED') return now;
      const match = Object.keys(body).find(bk => bk.trim().toUpperCase() === h.toUpperCase());
      return match !== undefined ? String(body[match] ?? '') : '';
    });

    if (rowIdx >= 0) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `'${TAB}'!A${rowIdx + 2}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [buildRow(false)] },
      });
      return res.json({ ok: true, action: 'updated' });
    }

    const newRow = buildRow(true);
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: `'${TAB}'!A1`,
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [newRow] },
    });
    res.json({ ok: true, action: 'created' });
  } catch (e) {
    console.error('[paul/save]', e.message);
    res.status(500).json({ error: e.message });
  }
});

module.exports = router;
