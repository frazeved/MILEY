require('dotenv').config();
const express  = require('express');
const path     = require('path');
const fs       = require('fs');
const nodemailer = require('nodemailer');
const XLSX     = require('xlsx');
const ExcelJS  = require('exceljs');
const { google } = require('googleapis');
const session  = require('express-session');

const suppliers   = require('./contacts/suppliers');
const team305     = require('./contacts/team305');
const TEAM_USERS  = require('./contacts/users');
const AUTH_CONFIG = require('./contacts/authUsers');
const buyers      = require('./contacts/buyers');

const app = express();
app.use(express.json({ limit: '25mb' }));
app.use(session({
  secret: process.env.SESSION_SECRET || '305secret',
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 24 * 60 * 60 * 1000 },
}));

// ─── Auth users & password store (Google Sheets persistence) ─────────────────
const SHEET_ID       = '1y0iL7PJldbVQmPIAnJi9wvA2hvjB8_aK2bU2kxvUf5Q';
const AUTH_SHEET_TAB   = 'WORKSPACE AUTH';
const TOKENS_SHEET_TAB = 'GMAIL TOKENS';
let passwordMap = {}; // email → changed password (overrides env var default)

async function loadPasswords() {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON) return;
  try {
    const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'] });
    const sheets = google.sheets({ version: 'v4', auth });
    const r = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `'${AUTH_SHEET_TAB}'!A:B` });
    for (const [email, pwd] of (r.data.values || []).slice(1)) {
      if (email && pwd) passwordMap[email.toLowerCase()] = pwd;
    }
  } catch (_) {} // tab may not exist yet — that's fine
}

async function savePassword(email, newPwd) {
  const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
  const sheets = google.sheets({ version: 'v4', auth });
  // Ensure tab exists
  try {
    const info = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
    if (!info.data.sheets?.some(s => s.properties?.title === AUTH_SHEET_TAB)) {
      await sheets.spreadsheets.batchUpdate({ spreadsheetId: SHEET_ID, requestBody: { requests: [{ addSheet: { properties: { title: AUTH_SHEET_TAB } } }] } });
      await sheets.spreadsheets.values.update({ spreadsheetId: SHEET_ID, range: `'${AUTH_SHEET_TAB}'!A1`, valueInputOption: 'RAW', requestBody: { values: [['EMAIL', 'PASSWORD']] } });
    }
  } catch (_) {}
  // Upsert row
  const r = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `'${AUTH_SHEET_TAB}'!A:B` });
  const rows = r.data.values || [];
  const idx  = rows.slice(1).findIndex(row => (row[0] || '').toLowerCase() === email.toLowerCase());
  if (idx >= 0) {
    await sheets.spreadsheets.values.update({ spreadsheetId: SHEET_ID, range: `'${AUTH_SHEET_TAB}'!B${idx + 2}`, valueInputOption: 'RAW', requestBody: { values: [[newPwd]] } });
  } else {
    await sheets.spreadsheets.values.append({ spreadsheetId: SHEET_ID, range: `'${AUTH_SHEET_TAB}'!A:B`, valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS', requestBody: { values: [[email, newPwd]] } });
  }
  passwordMap[email.toLowerCase()] = newPwd;
}

async function loadTokensFromSheets() {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON) return;
  try {
    const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'] });
    const sheets = google.sheets({ version: 'v4', auth });
    const r = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `'${TOKENS_SHEET_TAB}'!A:B` });
    for (const [userId, token] of (r.data.values || []).slice(1)) {
      if (userId && token) userTokens[userId] = { ...(userTokens[userId] || {}), refreshToken: token };
    }
  } catch (_) {}
}

async function saveTokenToSheets(userId, refreshToken) {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON) return;
  try {
    const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
    const sheets = google.sheets({ version: 'v4', auth });
    const info   = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
    if (!info.data.sheets?.some(s => s.properties?.title === TOKENS_SHEET_TAB)) {
      await sheets.spreadsheets.batchUpdate({ spreadsheetId: SHEET_ID, requestBody: { requests: [{ addSheet: { properties: { title: TOKENS_SHEET_TAB } } }] } });
      await sheets.spreadsheets.values.update({ spreadsheetId: SHEET_ID, range: `'${TOKENS_SHEET_TAB}'!A1`, valueInputOption: 'RAW', requestBody: { values: [['USER_ID', 'REFRESH_TOKEN']] } });
    }
    const r    = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `'${TOKENS_SHEET_TAB}'!A:B` });
    const rows = r.data.values || [];
    const idx  = rows.slice(1).findIndex(row => (row[0] || '') === userId);
    if (idx >= 0) {
      await sheets.spreadsheets.values.update({ spreadsheetId: SHEET_ID, range: `'${TOKENS_SHEET_TAB}'!B${idx + 2}`, valueInputOption: 'RAW', requestBody: { values: [[refreshToken]] } });
    } else {
      await sheets.spreadsheets.values.append({ spreadsheetId: SHEET_ID, range: `'${TOKENS_SHEET_TAB}'!A:B`, valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS', requestBody: { values: [[userId, refreshToken]] } });
    }
  } catch (e) { console.error('saveTokenToSheets error:', e.message); }
}

loadPasswords();
loadTokensFromSheets();

const { DEFAULT_PASSWORD, users: AUTH_USERS_CONFIG } = AUTH_CONFIG;
const AUTH_USERS = AUTH_USERS_CONFIG.map(u => ({
  email:    u.email,
  password: process.env[u.envVar] || DEFAULT_PASSWORD,
  name:     u.name,
  role:     u.role,
}));

// ─── Auth middleware ──────────────────────────────────────────────────────────
const OPEN_PATHS = ['/login', '/api/login', '/api/logout'];
app.use((req, res, next) => {
  if (OPEN_PATHS.some(p => req.path === p)) return next();
  if (req.session?.user) return next();
  if (req.path.startsWith('/api/')) return res.status(401).json({ error: 'Not authenticated' });
  return res.redirect('/login');
});

// ─── Login / logout routes ────────────────────────────────────────────────────
app.get('/login', (req, res) => {
  if (req.session?.user) return res.redirect('/');
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

app.post('/api/login', (req, res) => {
  const { email, password } = req.body || {};
  const user = AUTH_USERS.find(u => u.email.toLowerCase() === (email || '').toLowerCase().trim());
  if (!user) return res.status(401).json({ error: 'Email not found' });
  const stored = passwordMap[user.email.toLowerCase()] || user.password;
  if (stored !== password) return res.status(401).json({ error: 'Incorrect password' });
  const mustChangePassword = stored === DEFAULT_PASSWORD;
  req.session.user = { email: user.email, name: user.name, role: user.role };
  res.json({ ok: true, name: user.name, mustChangePassword });
});

app.get('/api/logout', (req, res) => {
  req.session.destroy(() => res.redirect('/login'));
});

app.get('/api/me', (req, res) => {
  if (!req.session?.user) return res.status(401).json({ error: 'Not authenticated' });
  res.json(req.session.user);
});

app.post('/api/change-password', async (req, res) => {
  if (!req.session?.user) return res.status(401).json({ error: 'Not authenticated' });
  const { currentPassword, newPassword } = req.body || {};
  if (!currentPassword || !newPassword) return res.status(400).json({ error: 'Current and new password are required' });
  if (newPassword.length < 6) return res.status(400).json({ error: 'New password must be at least 6 characters' });

  const email  = req.session.user.email;
  const user   = AUTH_USERS.find(u => u.email.toLowerCase() === email.toLowerCase());
  if (!user) return res.status(401).json({ error: 'User not found' });

  const stored = passwordMap[email.toLowerCase()] || user.password;
  if (stored !== currentPassword) return res.status(401).json({ error: 'Current password is incorrect' });

  try {
    await savePassword(email, newPassword);
    res.json({ ok: true });
  } catch (e) {
    console.error('[change-password]', e.message);
    res.status(500).json({ error: 'Failed to save new password. Try again.' });
  }
});

app.use(express.static(path.join(__dirname, 'public'), {
  setHeaders: (res, filePath) => {
    if (filePath.endsWith('.html')) {
      res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
    }
  },
}));

// ─── Constants ────────────────────────────────────────────────────────────────
const GITHUB_TOKEN = process.env.GITHUB_TOKEN;
const REPO         = 'frazeved/SAMANTHA';
const MAP_SHEET_ID = '1W88MKYr-q9g3F2fLFu2jjxvXzigK12PWohSVMsOQst4';
const SHEET_BASE   = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit`;
const csvUrl       = (gid) => `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${gid}`;
const REDIRECT_URI = 'https://workspace305team.onrender.com/auth/callback';

const SIZE_MAP = {
  "3700":"XXS","3740":"XXS PETITE","4000":"XS","3701":"XS PETITE",
  "5000":"S","3702":"S PETITE","6000":"M","3703":"M PETITE",
  "7000":"L","3704":"L PETITE","8000":"XL","3705":"XL PETITE","9000":"XXL",
};

const FEDEX_API_URL      = 'https://apis.fedex.com';
const FEDEX_DRIVE_FOLDER = '1ufkdrO23m2C-MrmhR1iKN3QFSJFwuQPY';
const FEDEX_SERVICE_NAMES = {
  FIRST_OVERNIGHT:         'FedEx First Overnight',
  FEDEX_2_DAY_AM:          'FedEx 2Day AM',
  FEDEX_3_DAY_FREIGHT:     'FedEx 3Day Freight',
  FEDEX_EXPRESS_SAVER:     'FedEx Express Saver',
  PRIORITY_OVERNIGHT:      'FedEx Priority Overnight',
  FEDEX_2_DAY:             'FedEx 2Day',
  FEDEX_1_DAY_FREIGHT:     'FedEx 1Day Freight',
  FEDEX_2_DAY_FREIGHT:     'FedEx 2Day Freight',
  STANDARD_OVERNIGHT:      'FedEx Standard Overnight',
  FEDEX_GROUND:            'FedEx Ground',
  FIRST_OVERNIGHT_FREIGHT: 'FedEx First Overnight Freight',
};
const FEDEX_ADDRESS_BOOK = {
  SHIP_FROM: {
    company: '305 Consulting and Production Inc',
    contact: '305 Consulting and Production Inc',
    phone:   '9174992103',
    street:  ['1800 NW 15TH AVENUE, STE 110', 'UNIT 2 GROUND'],
    city:    'POMPANO BEACH', state: 'FL', zip: '33069', country: 'US',
  },
  BRO: { company: 'URBAN OUTFITTERS INC',   contact: 'BRISTOL RENTAL DC',           phone: '', street: ['2401 GREEN LN'],                             city: 'LEVITTOWN',    state: 'PA', zip: '19057', country: 'US' },
  GAP: { company: 'ANTHROPOLOGIE GAP',      contact: 'URBN GAP DC',                 phone: '', street: ['755 BRACKBILL ROAD'],                        city: 'GAP',          state: 'PA', zip: '17527', country: 'US' },
  GFC: { company: 'ANTHROPOLOGIE',          contact: 'URBN GAP FULFILLMENT CENTER', phone: '', street: ['766 BRACKBILL ROAD'],                        city: 'GAP',          state: 'PA', zip: '17527', country: 'US' },
  KC1: { company: 'URBAN OUTFITTERS INC',   contact: 'KANSAS CITY KANSAS FC',       phone: '', street: ['11681 STATE AVE'],                           city: 'KANSAS CITY',  state: 'KS', zip: '66111', country: 'US' },
  KC3: { company: 'URBAN OUTFITTERS INC',   contact: 'NUULY RAYMORE RENTAL DC',     phone: '', street: ['1300 S. DEAN AVE', 'BUILDING 3, SUITE 100'], city: 'RAYMORE',      state: 'MO', zip: '64083', country: 'US' },
  REN: { company: 'URBN RENO DC',           contact: 'URBN RENO DC',                phone: '', street: ['6640 ECHO AVE'],                             city: 'RENO',         state: 'NV', zip: '89506', country: 'US' },
  RNO: { company: 'ANTHROPOLOGIE',          contact: 'URBN WEST COAST FULFILLMENT', phone: '', street: ['12055 MOYA BLVD'],                           city: 'RENO',         state: 'NV', zip: '89506', country: 'US' },
  YRD: { company: 'NUULY NAVY YARD',        contact: 'NUULY NAVY YARD',             phone: '', street: ['5000 SOUTH BROAD ST'],                       city: 'PHILADELPHIA', state: 'PA', zip: '19112', country: 'US' },
};
const FEDEX_PACKAGING = {
  'ECICO':      { l: 20, w: 12, h: 12 },
  'GAIA':       { l: 20, w: 12, h: 12 },
  'H&F':        { l: 23, w: 12, h: 12 },
  'HS FASHION': { l: 20, w: 15, h: 11 },
  'KONCEPTION': { l: 20, w: 12, h: 12 },
  'MINI BOX':   { l: 12, w: 12, h: 12 },
  'PQSWIM':     { l: 20, w: 12, h: 12 },
  'S&S':        { l: 23, w: 16, h: 11 },
};

const WORKFLOWS = [
  { id: 'po-detail.yml',   name: 'PO DETAIL',  tab: 'PO DETAIL',  sheetUrl: `${SHEET_BASE}?gid=2017761959#gid=2017761959` },
  { id: 'all-pos.yml',     name: 'ALL POs',     tab: 'PO TRADE',   sheetUrl: `${SHEET_BASE}?gid=890202899#gid=890202899`   },
  { id: 'po-headers.yml',  name: 'PO HEADERS',  tab: 'PO NEW',     sheetUrl: `${SHEET_BASE}?gid=406184613#gid=406184613`   },
  { id: 'po-invoiced.yml', name: 'PO INVOICED', tab: 'PO INVOICE', sheetUrl: `${SHEET_BASE}?gid=1379989532#gid=1379989532` },
];

// ─── Token store ──────────────────────────────────────────────────────────────
const TOKEN_FILE = path.join(__dirname, 'tokens.json');
let userTokens = {};

(function loadTokens() {
  try { userTokens = JSON.parse(fs.readFileSync(TOKEN_FILE, 'utf8')); } catch (_) {}
  // Env vars override file — useful for persisting across Render deploys
  for (const [k, v] of Object.entries(process.env)) {
    if (k.startsWith('GMAIL_TOKEN_') && v) {
      const id = k.slice(12).toLowerCase();
      userTokens[id] = { ...(userTokens[id] || {}), refreshToken: v };
    }
  }
})();

function saveTokens() {
  try { fs.writeFileSync(TOKEN_FILE, JSON.stringify(userTokens, null, 2)); } catch (_) {}
}

function makeOAuth2Client() {
  return new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    REDIRECT_URI
  );
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
async function ghFetch(url, opts = {}) {
  return fetch(url, {
    ...opts,
    headers: {
      Authorization: `Bearer ${GITHUB_TOKEN}`,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
      'Content-Type': 'application/json',
      ...opts.headers,
    },
  });
}

function parseCSV(text) {
  if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
  const rows = []; let row = [], field = '', inQ = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQ) {
      if (c === '"' && text[i+1] === '"') { field += '"'; i++; }
      else if (c === '"') inQ = false;
      else field += c;
    } else {
      if (c === '"') inQ = true;
      else if (c === ',') { row.push(field); field = ''; }
      else if (c === '\n' || (c === '\r' && text[i+1] === '\n')) {
        if (c === '\r') i++;
        row.push(field); field = '';
        if (row.some(f => f.trim())) rows.push(row);
        row = [];
      } else field += c;
    }
  }
  if (row.length) { row.push(field); if (row.some(f => f.trim())) rows.push(row); }
  return rows;
}

// Build a raw MIME message buffer using nodemailer (no actual send)
function buildRawMime(mailOptions) {
  return new Promise((resolve, reject) => {
    const t = nodemailer.createTransport({ streamTransport: true, newline: 'unix', buffer: true });
    t.sendMail(mailOptions, (err, info) => err ? reject(err) : resolve(info.message));
  });
}

// ─── Auth routes ──────────────────────────────────────────────────────────────
app.get('/setup', (req, res) => res.sendFile(path.join(__dirname, 'public', 'setup.html')));

app.get('/api/users', (req, res) => {
  res.json(TEAM_USERS.map(u => ({
    ...u,
    connected: !!(userTokens[u.id]?.refreshToken),
  })));
});

app.get('/auth/google', (req, res) => {
  const { userId, return: returnPath } = req.query;
  const user = TEAM_USERS.find(u => u.id === userId);
  if (!user) return res.status(400).send('Unknown user');
  req.session.pendingUserId = userId;
  req.session.oauthReturn = returnPath || null;
  const url = makeOAuth2Client().generateAuthUrl({
    access_type: 'offline',
    prompt: 'consent',
    scope: ['https://www.googleapis.com/auth/gmail.compose'],
    login_hint: user.email,
  });
  res.redirect(url);
});

app.get('/auth/callback', async (req, res) => {
  const { code, error } = req.query;
  if (error) return res.redirect('/setup?error=denied');
  const userId = req.session.pendingUserId;
  if (!code || !userId) return res.redirect('/setup?error=session');
  try {
    const client = makeOAuth2Client();
    const { tokens } = await client.getToken(code);
    if (!tokens.refresh_token) return res.redirect('/setup?error=norefresh');
    userTokens[userId] = { refreshToken: tokens.refresh_token };
    saveTokens();
    saveTokenToSheets(userId, tokens.refresh_token);
    const returnPath = req.session.oauthReturn;
    req.session.oauthReturn = null;
    if (returnPath === 'popup') {
      return res.send(`<!DOCTYPE html><html><body><script>
        try { window.opener.postMessage({ type:'gmailConnected', userId:'${userId}' }, '*'); } catch(e){}
        window.close();
      </script></body></html>`);
    }
    if (returnPath) return res.redirect(`${returnPath}?gmailConnected=1`);
    res.redirect(`/setup?success=${userId}`);
  } catch (e) {
    console.error('auth callback error:', e.message);
    res.redirect('/setup?error=failed');
  }
});

// ─── Existing routes ──────────────────────────────────────────────────────────
app.get('/api/status', async (req, res) => {
  try {
    const r = await ghFetch(`https://api.github.com/repos/${REPO}/actions/workflows/all-extractions.yml/runs?per_page=1`);
    const data = await r.json();
    const run = data.workflow_runs?.[0] || null;
    const duration = run && run.status === 'completed' && run.run_started_at && run.updated_at
      ? Math.round((new Date(run.updated_at) - new Date(run.run_started_at)) / 1000) : null;
    res.json(WORKFLOWS.map(wf => ({
      id: wf.id, name: wf.name, tab: wf.tab, sheetUrl: wf.sheetUrl,
      status: run?.status || 'unknown', conclusion: run?.conclusion || null,
      started_at: run?.run_started_at || null, updated_at: run?.updated_at || null, duration,
    })));
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/run-all', async (req, res) => {
  try {
    const r = await ghFetch(`https://api.github.com/repos/${REPO}/actions/workflows/all-extractions.yml/dispatches`, {
      method: 'POST', body: JSON.stringify({ ref: 'main' }),
    });
    if (r.status !== 204) { const b = await r.text(); return res.status(500).json({ error: `GitHub returned ${r.status}: ${b}` }); }
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

const ALLOWED_WORKFLOWS = new Set(['po-detail.yml', 'all-pos.yml', 'po-headers.yml', 'po-invoiced.yml']);
app.post('/api/run-workflow', async (req, res) => {
  const { id } = req.body || {};
  if (!id || !ALLOWED_WORKFLOWS.has(id)) return res.status(400).json({ error: 'Invalid workflow id' });
  try {
    const r = await ghFetch(`https://api.github.com/repos/${REPO}/actions/workflows/${id}/dispatches`, {
      method: 'POST', body: JSON.stringify({ ref: 'main' }),
    });
    if (r.status !== 204) { const b = await r.text(); return res.status(500).json({ error: `GitHub returned ${r.status}: ${b}` }); }
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

async function triggerJhonnyUpdate(req, res) {
  try {
    const r = await ghFetch(`https://api.github.com/repos/frazeved/JHONNY/actions/workflows/update-fedex-status.yml/dispatches`, {
      method: 'POST', body: JSON.stringify({ ref: 'main' }),
    });
    if (r.status !== 204) { const b = await r.text(); return res.status(500).json({ error: `GitHub returned ${r.status}: ${b}` }); }
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
}
app.post('/api/run-jhonny', triggerJhonnyUpdate);
app.post('/api/fedex/update', triggerJhonnyUpdate);

// ─── Status Summary Dashboard ─────────────────────────────────────────────────
const TRACKED_STATUSES = [
  'NEW', 'WAITING SMS', 'WAITING PRICE FROM SUPPLIER', 'CHECKING PRICE NDC',
  'WAITING PO', 'WAITING REVISED SMS', 'WAITING REVISED AW', 'WAITING PROTO APPROVAL', 'MISSING AW',
];
app.get('/api/po/status-summary', async (req, res) => {
  try {
    const response = await fetch(csvUrl(0));
    if (!response.ok) return res.status(500).json({ error: 'Failed to fetch sheet' });
    const rows = parseCSV(await response.text());
    if (rows.length < 2) return res.json({ counts: {}, styles: {} });
    const H = rows[0].map(h => h.trim().toLowerCase());
    const col = (...kws) => { for (const kw of kws) { const i = H.findIndex(h => h.includes(kw)); if (i >= 0) return i; } return -1; };
    const C = {
      style:    col('style #', 'style#', 'style'),
      status:   col('status'),
      supplier: col('supplier'),
      category: col('category'),
      ndc:      col('ndc month/year', 'ndc month'),
      smsSent:  col('sms sent from supplier', 'sms sent'),
      poIssued:   col('po issued by anthro', 'po issued'),
      printSent:  col('print sent to supplier', 'print sent'),
    };
    if (C.style < 0 || C.status < 0) return res.status(500).json({ error: 'Required columns not found' });
    const get = (r, i) => i >= 0 ? (r[i] || '').trim() : '';
    const MONTH_NAMES = ['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE','JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER'];
    const MONTH_SHORT = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
    const fmtNDC = raw => {
      if (!raw) return '';
      const s = raw.trim().toUpperCase();
      for (let i = 0; i < MONTH_NAMES.length; i++) {
        if (s.startsWith(MONTH_NAMES[i]) || s.startsWith(MONTH_SHORT[i])) return MONTH_NAMES[i];
      }
      const d = new Date(raw);
      if (!isNaN(d)) return MONTH_NAMES[d.getMonth()];
      return '';
    };
    const counts = {}, styles = {};
    TRACKED_STATUSES.forEach(s => { counts[s] = 0; styles[s] = []; });
    for (let i = 1; i < rows.length; i++) {
      const status = get(rows[i], C.status).toUpperCase();
      const style  = get(rows[i], C.style);
      if (!style) continue;
      const match = TRACKED_STATUSES.find(s => s === status);
      if (match) {
        counts[match]++;
        styles[match].push({
          style,
          supplier: get(rows[i], C.supplier),
          category: get(rows[i], C.category),
          ndc:      fmtNDC(get(rows[i], C.ndc)),
          smsSent:   get(rows[i], C.smsSent),
          poIssued:  get(rows[i], C.poIssued),
          printSent: get(rows[i], C.printSent),
          rowIndex: i + 1,
        });
      }
    }
    res.json({ counts, styles });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/po/save-sms-sent', async (req, res) => {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON)
    return res.status(500).json({ error: 'Google credentials not configured' });
  const { updates } = req.body;
  if (!Array.isArray(updates) || updates.length === 0)
    return res.status(400).json({ error: 'updates array required' });
  try {
    const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
    const sheets = google.sheets({ version: 'v4', auth });

    const meta    = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID, fields: 'sheets.properties' });
    const tabName = meta.data.sheets.find(s => s.properties.sheetId === 0)?.properties.title;
    if (!tabName) return res.status(500).json({ error: 'Main PO tab not found' });

    const hRes   = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `'${tabName}'!1:1` });
    const headers = (hRes.data.values?.[0] || []).map(h => (h || '').trim().toLowerCase());
    const styleCol = headers.findIndex(h => h.includes('style #') || h.includes('style#') || h === 'style');
    const smsCol   = headers.findIndex(h => h.includes('sms sent from supplier') || h.includes('sms sent'));
    if (smsCol < 0)   return res.status(404).json({ error: '"SMS SENT FROM SUPPLIER" column not found in main PO sheet' });
    if (styleCol < 0) return res.status(500).json({ error: 'Style column not found in main PO sheet' });

    const dataRes = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `'${tabName}'` });
    const allRows = dataRes.data.values || [];
    const styleToRow = new Map();
    for (let i = 1; i < allRows.length; i++) {
      const s = (allRows[i][styleCol] || '').trim();
      if (s && !styleToRow.has(s)) styleToRow.set(s, i + 1);
    }

    const colLetter = colToLetter(smsCol);
    const data = [], notFound = [];
    for (const { style, smsSent } of updates) {
      const rowNum = styleToRow.get(style);
      if (!rowNum) { notFound.push(style); continue; }
      data.push({ range: `'${tabName}'!${colLetter}${rowNum}`, values: [[smsSent || '']] });
    }
    if (data.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SHEET_ID,
        requestBody: { valueInputOption: 'USER_ENTERED', data },
      });
    }
    res.json({ ok: true, updated: data.length, notFound });
  } catch (e) {
    console.error('[save-sms-sent]', e);
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/po/save-po-issued', async (req, res) => {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON)
    return res.status(500).json({ error: 'Google credentials not configured' });
  const { updates } = req.body;
  if (!Array.isArray(updates) || updates.length === 0)
    return res.status(400).json({ error: 'updates array required' });
  try {
    const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
    const sheets = google.sheets({ version: 'v4', auth });

    const meta    = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID, fields: 'sheets.properties' });
    const tabName = meta.data.sheets.find(s => s.properties.sheetId === 0)?.properties.title;
    if (!tabName) return res.status(500).json({ error: 'Main PO tab not found' });

    const hRes   = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `'${tabName}'!1:1` });
    const headers = (hRes.data.values?.[0] || []).map(h => (h || '').trim().toLowerCase());
    const styleCol = headers.findIndex(h => h.includes('style #') || h.includes('style#') || h === 'style');
    const poCol    = headers.findIndex(h => h.includes('po issued by anthro') || h.includes('po issued'));
    if (poCol < 0)    return res.status(404).json({ error: '"PO ISSUED BY ANTHRO" column not found in main PO sheet' });
    if (styleCol < 0) return res.status(500).json({ error: 'Style column not found in main PO sheet' });

    const dataRes = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `'${tabName}'` });
    const allRows = dataRes.data.values || [];
    const styleToRow = new Map();
    for (let i = 1; i < allRows.length; i++) {
      const s = (allRows[i][styleCol] || '').trim();
      if (s && !styleToRow.has(s)) styleToRow.set(s, i + 1);
    }

    const colLetter = colToLetter(poCol);
    const data = [], notFound = [];
    for (const { style, poIssued } of updates) {
      const rowNum = styleToRow.get(style);
      if (!rowNum) { notFound.push(style); continue; }
      data.push({ range: `'${tabName}'!${colLetter}${rowNum}`, values: [[poIssued || '']] });
    }
    if (data.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SHEET_ID,
        requestBody: { valueInputOption: 'USER_ENTERED', data },
      });
    }
    res.json({ ok: true, updated: data.length, notFound });
  } catch (e) {
    console.error('[save-po-issued]', e);
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/po/save-print-sent', async (req, res) => {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON)
    return res.status(500).json({ error: 'Google credentials not configured' });
  const { updates } = req.body;
  if (!Array.isArray(updates) || updates.length === 0)
    return res.status(400).json({ error: 'updates array required' });
  try {
    const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
    const sheets = google.sheets({ version: 'v4', auth });

    const meta    = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID, fields: 'sheets.properties' });
    const tabName = meta.data.sheets.find(s => s.properties.sheetId === 0)?.properties.title;
    if (!tabName) return res.status(500).json({ error: 'Main PO tab not found' });

    const hRes    = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `'${tabName}'!1:1` });
    const headers = (hRes.data.values?.[0] || []).map(h => (h || '').trim().toLowerCase());
    const styleCol = headers.findIndex(h => h.includes('style #') || h.includes('style#') || h === 'style');
    const printCol = headers.findIndex(h => h.includes('print sent to supplier') || h.includes('print sent'));
    if (printCol < 0)  return res.status(404).json({ error: '"PRINT SENT TO SUPPLIER" column not found in main PO sheet' });
    if (styleCol < 0)  return res.status(500).json({ error: 'Style column not found in main PO sheet' });

    const dataRes = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `'${tabName}'` });
    const allRows = dataRes.data.values || [];
    const styleToRow = new Map();
    for (let i = 1; i < allRows.length; i++) {
      const s = (allRows[i][styleCol] || '').trim();
      if (s && !styleToRow.has(s)) styleToRow.set(s, i + 1);
    }

    const colLetter = colToLetter(printCol);
    const data = [], notFound = [];
    for (const { style, printSent } of updates) {
      const rowNum = styleToRow.get(style);
      if (!rowNum) { notFound.push(style); continue; }
      data.push({ range: `'${tabName}'!${colLetter}${rowNum}`, values: [[printSent || '']] });
    }
    if (data.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SHEET_ID,
        requestBody: { valueInputOption: 'USER_ENTERED', data },
      });
    }
    res.json({ ok: true, updated: data.length, notFound });
  } catch (e) {
    console.error('[save-print-sent]', e);
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/po/save-status', async (req, res) => {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON)
    return res.status(500).json({ error: 'Google credentials not configured' });
  const { updates } = req.body;
  if (!Array.isArray(updates) || updates.length === 0)
    return res.status(400).json({ error: 'updates array required' });
  try {
    const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
    const sheets = google.sheets({ version: 'v4', auth });

    const meta    = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID, fields: 'sheets.properties' });
    const tabName = meta.data.sheets.find(s => s.properties.sheetId === 0)?.properties.title;
    if (!tabName) return res.status(500).json({ error: 'Main PO tab not found' });

    const hRes    = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `'${tabName}'!1:1` });
    const headers = (hRes.data.values?.[0] || []).map(h => (h || '').trim().toLowerCase());
    const styleCol  = headers.findIndex(h => h.includes('style #') || h.includes('style#') || h === 'style');
    const statusCol = headers.findIndex(h => h === 'status');
    if (statusCol < 0) return res.status(404).json({ error: '"STATUS" column not found in main PO sheet' });
    if (styleCol < 0)  return res.status(500).json({ error: 'Style column not found in main PO sheet' });

    const dataRes = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `'${tabName}'` });
    const allRows = dataRes.data.values || [];
    const styleToRow = new Map();
    for (let i = 1; i < allRows.length; i++) {
      const s = (allRows[i][styleCol] || '').trim();
      if (s && !styleToRow.has(s)) styleToRow.set(s, i + 1);
    }

    const colLetter = colToLetter(statusCol);
    const data = [], notFound = [];
    for (const { style, status } of updates) {
      if (!status) continue;
      const rowNum = styleToRow.get(style);
      if (!rowNum) { notFound.push(style); continue; }
      data.push({ range: `'${tabName}'!${colLetter}${rowNum}`, values: [[status]] });
    }
    if (data.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SHEET_ID,
        requestBody: { valueInputOption: 'USER_ENTERED', data },
      });
    }
    res.json({ ok: true, updated: data.length, notFound });
  } catch (e) {
    console.error('[save-status]', e);
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/po/search', async (req, res) => {
  try {
    const { style } = req.body;
    if (!style?.trim()) return res.status(400).json({ error: 'Style # is required' });
    const styleToSearch = style.toString().trim().toLowerCase();
    const response = await fetch(csvUrl(0));
    if (!response.ok) return res.status(500).json({ error: 'Failed to fetch sheet data' });
    const rows = parseCSV(await response.text());
    if (rows.length < 2) return res.status(500).json({ error: 'No data in sheet' });
    const headers = rows[0].map(h => h.trim().toLowerCase());
    const col = (...kws) => {
      for (const kw of kws) { const i = headers.findIndex(h => h === kw.toLowerCase()); if (i >= 0) return i; }
      for (const kw of kws) { const i = headers.findIndex(h => h.includes(kw.toLowerCase())); if (i >= 0) return i; }
      return -1;
    };
    const C = {
      style: col('style #','style#','style'), status: col('status'), supplier: col('supplier'),
      category: col('category'), subcategory: col('sub-category','subcategory'),
      piReceived: col('buyer presentation','pi received','pi date'),
      topSent: col('sms sent to anthro','top sent to anthro','sms sent'),
      topSampleStatus: col('top approval from anthro','top sample status','top approval'),
      basePo: col('po info anthro','base po','base po import farm'),
      finalNDC: col('final ndc','ndc'), topDeadline: col('sms deadline','top deadline'),
      exFactory: col('ex factory / flight date','ex factory','flight date'),
      cost: col('cost'), freight: col('freight'), duty: col('duty'), hts: col('hts code','hts'),
    };
    const get = (r, i) => i >= 0 ? (r[i] || '').trim() : '';
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (get(r, C.style).toLowerCase() !== styleToSearch) continue;
      if (!get(r, C.status).toLowerCase().includes("po'd")) continue;
      const finalNDC = get(r, C.finalNDC), topDeadline = get(r, C.topDeadline);
      let topDeadlineDays = '';
      if (finalNDC && topDeadline) { const a = new Date(finalNDC), b = new Date(topDeadline); if (!isNaN(a)&&!isNaN(b)) topDeadlineDays = Math.round((a-b)/86400000).toString(); }
      return res.json({ style: get(r,C.style), supplier: get(r,C.supplier), category: get(r,C.category), subcategory: get(r,C.subcategory), piReceived: get(r,C.piReceived), topSent: get(r,C.topSent), topSampleStatus: get(r,C.topSampleStatus), basePo: get(r,C.basePo), finalNDC, topDeadline, topDeadlineDays, exFactory: get(r,C.exFactory), cost: get(r,C.cost), freight: get(r,C.freight), duty: get(r,C.duty), hts: get(r,C.hts) });
    }
    res.status(404).json({ error: "Style not found or not in PO'd status" });
  } catch (e) { console.error('PO search error:', e.message); res.status(500).json({ error: e.message }); }
});

// ─── Breakdown Email — creates a Gmail draft ──────────────────────────────────
// ─── Breakdown Report — download Excel only (no email) ───────────────────────
app.post('/api/po/breakdown-report', async (req, res) => {
  try {
    const { style } = req.body;
    if (!style?.trim()) return res.status(400).json({ error: 'Style # is required' });

    const rawStyle     = style.trim();
    const displayStyle = /^[A-Za-z]-/.test(rawStyle) ? rawStyle.split('-').slice(1).join('-').trim() : rawStyle;
    const cleanStyle   = displayStyle.toUpperCase().replace(/\s+/g, '');

    const [pdRes, ptRes] = await Promise.all([fetch(csvUrl(2017761959)), fetch(csvUrl(890202899))]);
    if (!pdRes.ok) throw new Error('Could not fetch PO DETAIL sheet');
    if (!ptRes.ok) throw new Error('Could not fetch PO TRADE sheet');
    const pdRows = parseCSV(await pdRes.text());
    const ptRows = parseCSV(await ptRes.text());

    const channelMap = {};
    for (let i = 1; i < ptRows.length; i++) { const po = (ptRows[i][0]||'').trim(); if (po) channelMap[po] = (ptRows[i][1]||'').trim(); }

    const H = pdRows[0] || [];
    const findCol = (...kws) => {
      for (const kw of kws) { const i = H.findIndex(h=>h.trim().toLowerCase()===kw); if(i>=0) return i; }
      for (const kw of kws) { const i = H.findIndex(h=>h.trim().toLowerCase().includes(kw)); if(i>=0) return i; }
      return -1;
    };
    const C = {
      po: findCol('po#','purchase order','po'), style: findCol('vendor style','style#','style #','style'),
      shortSku: findCol('short sku'), size: findCol('size desc','size'),
      packType: findCol('ship pack','pack type','pack'), sizeCode: findCol('size code'),
      qty: findCol('total qty','qty'), units: findCol('allocated','prepack','total units'),
      itemNum: findCol('long sku','item number','sku'), color: findCol('vendor color','color'),
    };
    const get = (r,i) => i>=0 ? (r[i]||'').trim() : '';

    const groupedByPO = {};
    for (let i=1; i<pdRows.length; i++) {
      const r = pdRows[i];
      if (get(r,C.style).toUpperCase().replace(/\s+/g,'') !== cleanStyle) continue;
      const po = get(r,C.po); if (!po) continue;
      (groupedByPO[po] = groupedByPO[po]||[]).push(r);
    }

    if (!Object.keys(groupedByPO).length) return res.status(404).json({ error: `No PO data found for style ${displayStyle}` });

    const gcd=(a,b)=>{a=Math.abs(a);b=Math.abs(b);while(b){const t=b;b=a%b;a=t;}return a;};
    const gcdArr=ns=>{let g=0;for(const n of ns){if(n)g=g?gcd(g,n):Math.abs(n);}return g||1;};

    const HEADERS = ['PO#','Style#','Type','Pack Type','Item Number','Vendor Color','Size code','Size','Short SKU','Qty','RATIO','Total Units'];
    const allRows = []; let grandQty=0, grandPack=0, grandUnits=0;
    for (const po in groupedByPO) {
      const rows=groupedByPO[po];
      const ppkUnits=rows.filter(r=>get(r,C.packType).toUpperCase()==='PPK').map(r=>Number(get(r,C.units))||0).filter(v=>v>0);
      const gcdUnits=gcdArr(ppkUnits);
      let poQty=0, poPack=0, poTotal=0;
      for (const r of rows) {
        const pack=get(r,C.packType).toUpperCase(), qty=Number(get(r,C.qty))||0, units=Number(get(r,C.units))||0;
        const ratio=pack==='PPK'?Math.round((units/(gcdUnits||1))*1e6)/1e6:1, sc=get(r,C.sizeCode);
        allRows.push({ type:'data', vals:[get(r,C.po),get(r,C.style),channelMap[po]||'',pack,get(r,C.itemNum),get(r,C.color),sc,SIZE_MAP[sc]||get(r,C.size),get(r,C.shortSku),qty,ratio,units] });
        poQty+=qty; poPack+=ratio; poTotal+=units;
      }
      allRows.push({ type:'total', vals:['','','','','','','','TOTAL','',poQty,poPack,poTotal] });
      grandQty+=poQty; grandPack+=poPack; grandUnits+=poTotal;
    }
    allRows.push({ type:'blank', vals: Array(12).fill('') });
    allRows.push({ type:'grand', vals:['','','','','','','','GRAND TOTAL','',grandQty,grandPack,grandUnits] });

    // Build formatted workbook with ExcelJS
    const wbx = new ExcelJS.Workbook();
    const wsx = wbx.addWorksheet('Breakdown');
    wsx.views = [{ state: 'frozen', ySplit: 1 }];

    const headerRow = wsx.addRow(HEADERS);
    headerRow.eachCell(cell => {
      cell.fill   = { type:'pattern', pattern:'solid', fgColor:{ argb:'FF2E75B6' } };
      cell.font   = { bold:true, color:{ argb:'FFFFFFFF' } };
      cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
      cell.alignment = { vertical:'middle' };
    });

    let rowIdx = 0;
    for (const r of allRows) {
      const xlRow = wsx.addRow(r.vals);
      if (r.type === 'total') {
        xlRow.eachCell(cell => {
          cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFDEEAF1' } };
          cell.font = { bold:true };
          cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
        });
      } else if (r.type === 'grand') {
        xlRow.eachCell(cell => {
          cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FF2E75B6' } };
          cell.font = { bold:true, color:{ argb:'FFFFFFFF' } };
          cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
        });
      } else if (r.type === 'data') {
        const bg = rowIdx % 2 === 0 ? 'FFDEEAF1' : 'FFFFFFFF';
        xlRow.eachCell(cell => {
          cell.fill   = { type:'pattern', pattern:'solid', fgColor:{ argb:bg } };
          cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
        });
        rowIdx++;
      }
    }

    wsx.columns.forEach(col => { col.width = 14; });

    const excelBuf = await wbx.xlsx.writeBuffer();
    const filename = `BREAKDOWN STYLE# ${cleanStyle}.xlsx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(excelBuf);
  } catch (e) {
    console.error('breakdown-report error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/po/breakdown-email', async (req, res) => {
  try {
    const { style, supplier, cost, hts, freight, exFactory, message, sendingAs } = req.body;
    if (!style?.trim())    return res.status(400).json({ error: 'Style # is required' });
    if (!supplier?.trim()) return res.status(400).json({ error: 'Supplier is required' });
    if (!sendingAs)        return res.status(400).json({ error: 'Please select who is sending this' });

    const token = userTokens[sendingAs];
    if (!token?.refreshToken) {
      const user = TEAM_USERS.find(u => u.id === sendingAs);
      return res.status(401).json({ error: `${user?.name || sendingAs} has not connected their Gmail yet. Please visit /setup first.` });
    }

    const sender = TEAM_USERS.find(u => u.id === sendingAs);

    // Normalize style
    const rawStyle    = style.trim();
    const displayStyle = /^[A-Za-z]-/.test(rawStyle) ? rawStyle.split('-').slice(1).join('-').trim() : rawStyle;
    const cleanStyle  = displayStyle.toUpperCase().replace(/\s+/g, '');
    const supplierKey = supplier.trim().toUpperCase();

    // Format dates
    const fmtDate = iso => { if (!iso) return 'xxxxxx'; const d = new Date(iso); return isNaN(d) ? 'xxxxxx' : d.toLocaleDateString('en-US',{month:'short',day:'numeric'}); };
    const handoverFormatted = fmtDate(exFactory);
    const invoiceFormatted  = exFactory ? fmtDate(new Date(new Date(exFactory).getTime()-2*86400000).toISOString().split('T')[0]) : 'xxxxxx';

    // Fetch sheets
    const [pdRes, ptRes] = await Promise.all([fetch(csvUrl(2017761959)), fetch(csvUrl(890202899))]);
    if (!pdRes.ok) throw new Error('Could not fetch PO DETAIL sheet');
    if (!ptRes.ok) throw new Error('Could not fetch PO TRADE sheet');
    const pdRows = parseCSV(await pdRes.text());
    const ptRows = parseCSV(await ptRes.text());

    const channelMap = {};
    for (let i = 1; i < ptRows.length; i++) { const po = (ptRows[i][0]||'').trim(); if (po) channelMap[po] = (ptRows[i][1]||'').trim(); }

    const H = pdRows[0] || [];
    const findCol = (...kws) => {
      for (const kw of kws) { const i = H.findIndex(h=>h.trim().toLowerCase()===kw); if(i>=0) return i; }
      for (const kw of kws) { const i = H.findIndex(h=>h.trim().toLowerCase().includes(kw)); if(i>=0) return i; }
      return -1;
    };
    const C = {
      po: findCol('po#','purchase order','po'), style: findCol('vendor style','style#','style #','style'),
      shortSku: findCol('short sku'), size: findCol('size desc','size'),
      packType: findCol('ship pack','pack type','pack'), sizeCode: findCol('size code'),
      qty: findCol('total qty','qty'), units: findCol('allocated','prepack','total units'),
      itemNum: findCol('long sku','item number','sku'), color: findCol('vendor color','color'),
    };
    const get = (r,i) => i>=0 ? (r[i]||'').trim() : '';

    const groupedByPO = {};
    for (let i=1; i<pdRows.length; i++) {
      const r = pdRows[i];
      if (get(r,C.style).toUpperCase().replace(/\s+/g,'') !== cleanStyle) continue;
      const po = get(r,C.po); if (!po) continue;
      (groupedByPO[po] = groupedByPO[po]||[]).push(r);
    }

    const poLines=[], poSubject=[];
    for (const po in groupedByPO) { const ch=channelMap[po]||''; poLines.push(`<b>PO# ${po} - ${ch}</b>`); poSubject.push(`#${po}`); }
    if (!poLines.length) { poLines.push('<span style="color:blue;"><b>PO# xxxxxx - xxxxx</b></span>'); poSubject.push('#xxxxxx'); }

    const gcd=(a,b)=>{a=Math.abs(a);b=Math.abs(b);while(b){const t=b;b=a%b;a=t;}return a;};
    const gcdArr=ns=>{let g=0;for(const n of ns){if(n)g=g?gcd(g,n):Math.abs(n);}return g||1;};

    const EXCEL_HEADERS=['PO#','Style#','Type','Pack Type','Item Number','Vendor Color','Size code','Size','Short SKU','Qty','RATIO','Total Units'];
    const allRows=[]; let grandQty=0,grandPack=0,grandUnits=0;
    for (const po in groupedByPO) {
      const rows=groupedByPO[po];
      const ppkUnits=rows.filter(r=>get(r,C.packType).toUpperCase()==='PPK').map(r=>Number(get(r,C.units))||0).filter(v=>v>0);
      const gcdUnits=gcdArr(ppkUnits);
      let poQty=0,poPack=0,poTotal=0;
      for (const r of rows) {
        const pack=get(r,C.packType).toUpperCase(), qty=Number(get(r,C.qty))||0, units=Number(get(r,C.units))||0;
        const ratio=pack==='PPK'?Math.round((units/(gcdUnits||1))*1e6)/1e6:1, sc=get(r,C.sizeCode);
        allRows.push([get(r,C.po),get(r,C.style),channelMap[po]||'',pack,get(r,C.itemNum),get(r,C.color),sc,SIZE_MAP[sc]||get(r,C.size),get(r,C.shortSku),qty,ratio,units]);
        poQty+=qty; poPack+=ratio; poTotal+=units;
      }
      allRows.push(['','','','','','','','TOTAL','',poQty,poPack,poTotal]);
      grandQty+=poQty; grandPack+=poPack; grandUnits+=poTotal;
    }
    allRows.push(Array(12).fill(''));
    allRows.push(['','','','','','','','GRAND TOTAL','',grandQty,grandPack,grandUnits]);

    const wb=XLSX.utils.book_new(), ws=XLSX.utils.aoa_to_sheet([EXCEL_HEADERS,...allRows]);
    XLSX.utils.book_append_sheet(wb,ws,'Breakdown');
    const excelBuf=XLSX.write(wb,{type:'buffer',bookType:'xlsx'});

    const toEmails  = suppliers.emails[supplierKey]      || ['logistics@creativetwotwelve.com'];
    const greetName = suppliers.mainContact[supplierKey] || supplierKey;
    const subject   = `BREAKDOWN STYLE# ${displayStyle} - PO ${poSubject.join(' ')} - ${supplierKey}`;

    const htmlBody = `
<p><b><span style="font-size:12pt;">Hi ${greetName} and ${supplierKey} team,</span></b></p>
<p>Please find attached the breakdown of<br><b>STYLE# ${displayStyle}</b></p>
<p><b>HTS# ${hts||'xxxxxx'}</b> | <b>${freight||'xxxxxx'} $${cost||'xxxxxx'}</b> | <b>INVOICE/PACKING LIST WITH FLAVIO: ${invoiceFormatted}</b> | <b>AGREED HANDOVER DATE: ${handoverFormatted}</b></p>
<p>${poLines.join('<br>')}</p>
<p>Could you please confirm this style fabric composition?</p>
<p><b><span style="color:red;">IMPORTANT:</span></b><br>Please, send the Invoice and Packing list before shipping for validation, also the custom description, HTS#, TAX ID on it. AWB when available.</p>
<p>Any delay or new agreed ship date on this Style#, please answer this email chain immediately!</p>
${message?`<p>${message}</p>`:''}
<p>Best,<br>${sender?.name||sendingAs}<br>Production Team<br>305 CONSULTING AND PRODUCTION</p>`;

    // Build raw MIME (nodemailer composes, does NOT send)
    const rawMime = await buildRawMime({
      from: `"${sender?.name||sendingAs}" <${sender?.email||''}>`,
      to:   toEmails.join(','),
      cc:   team305.breakdownCC.join(','),
      subject,
      html: htmlBody,
      attachments: [{ filename:`BREAKDOWN STYLE# ${cleanStyle}.xlsx`, content:excelBuf, contentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }],
    });
    const encoded = rawMime.toString('base64').replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');

    // Create draft in the sender's Gmail
    const authClient = makeOAuth2Client();
    authClient.setCredentials({ refresh_token: token.refreshToken });
    const gmail = google.gmail({ version:'v1', auth:authClient });
    await gmail.users.drafts.create({ userId:'me', requestBody:{ message:{ raw:encoded } } });

    res.json({ ok:true, message:'Draft created in your Gmail Drafts folder' });
  } catch (e) {
    console.error('breakdown-email error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── PO Official Email — creates a Gmail draft ───────────────────────────────
app.post('/api/po/official-email', async (req, res) => {
  try {
    const { style, supplier, finalNDC, exFactory, cost, freight, sendingAs } = req.body;
    if (!style?.trim())    return res.status(400).json({ error: 'Style # is required' });
    if (!supplier?.trim()) return res.status(400).json({ error: 'Supplier is required' });
    if (!sendingAs)        return res.status(400).json({ error: 'Please select who is sending this' });

    const token = userTokens[sendingAs];
    if (!token?.refreshToken) {
      const user = TEAM_USERS.find(u => u.id === sendingAs);
      return res.status(401).json({ error: `${user?.name || sendingAs} has not connected their Gmail yet. Please visit /setup first.` });
    }

    const sender = TEAM_USERS.find(u => u.id === sendingAs);

    // Normalize style — strip leading "X-" prefix
    const rawStyle   = style.trim();
    const cleanStyle = rawStyle.toUpperCase().replace(/^[A-Za-z]+-/, '').replace(/\s+/g, '');
    // Digits-only version for subject line (matches GAS create_po_request_draft)
    const styleDigits = cleanStyle.replace(/[^0-9]/g, '');

    // Format NDC date  MM/DD/YYYY
    const fmtNDC = iso => {
      if (!iso) return 'xxxxxx';
      const d = new Date(iso);
      return isNaN(d) ? 'xxxxxx' : d.toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: 'numeric' });
    };
    const anthroNDC = fmtNDC(finalNDC);

    // Fetch PO DETAIL + RLM in parallel
    const [pdRes, rlmRes] = await Promise.all([fetch(csvUrl(2017761959)), fetch(csvUrl(1284509953))]);
    if (!pdRes.ok)  throw new Error('Could not fetch PO DETAIL sheet');
    if (!rlmRes.ok) throw new Error('Could not fetch RLM sheet');
    const pdRows  = parseCSV(await pdRes.text());
    const rlmRows = parseCSV(await rlmRes.text());

    const H = pdRows[0] || [];
    const findCol = (...kws) => {
      for (const kw of kws) { const i = H.findIndex(h => h.trim().toLowerCase() === kw); if (i >= 0) return i; }
      for (const kw of kws) { const i = H.findIndex(h => h.trim().toLowerCase().includes(kw)); if (i >= 0) return i; }
      return -1;
    };
    const C = {
      style:    findCol('vendor style', 'style#', 'style # ', 'style'),
      sizeCode: findCol('size code'),
      size:     findCol('size desc', 'size'),
      qty:      findCol('total qty', 'qty'),
      color:    findCol('vendor color', 'color'),
    };
    const get = (r, i) => i >= 0 ? (r[i] || '').trim() : '';

    // Look up color code from RLM sheet (col F=index 5 = style, col P=index 15 = color code)
    const normalizeStyle = s => s.toUpperCase().replace(/^[A-Za-z]+-/, '').replace(/ /g, '').replace(/\s+/g, '').replace(/[^0-9A-Z-]/g, '');
    let colorCode = null;
    for (let i = 1; i < rlmRows.length; i++) {
      if (normalizeStyle(rlmRows[i][5] || '') === normalizeStyle(cleanStyle)) {
        colorCode = (rlmRows[i][15] || '').trim();
        break;
      }
    }
    if (colorCode === null) {
      return res.status(404).json({ error: `Style ${cleanStyle} was not found in the RLM sheet.` });
    }

    // Aggregate qty by size across all POs for this style
    const sizeTotals = {};
    for (let i = 1; i < pdRows.length; i++) {
      const r = pdRows[i];
      const rowStyle = normalizeStyle(get(r, C.style));
      if (rowStyle !== normalizeStyle(cleanStyle)) continue;
      const sc  = get(r, C.sizeCode);
      const sz  = SIZE_MAP[sc] || get(r, C.size);
      const qty = Number(get(r, C.qty)) || 0;
      if (sz) sizeTotals[sz] = (sizeTotals[sz] || 0) + qty;
    }

    if (Object.keys(sizeTotals).length === 0) {
      return res.status(404).json({ error: `Style ${cleanStyle} was not found in the PO Detail sheet.` });
    }

    // Build Excel — 26-column PO import format with light blue formatting
    const userPO       = cleanStyle + 'A26';
    const etaDate      = exFactory ? new Date(new Date(exFactory).getTime() + 7 * 86400000).toLocaleDateString('en-US') : '';
    const exfFormatted = exFactory ? new Date(exFactory).toLocaleDateString('en-US') : '';
    const poTerms      = freight || '';

    const OFFICIAL_HEADERS = [
      'COMPANY','DIVISION','USER PO#','Season','Year','PO EXF Date','PO Vendor','Warehouse',
      'Style #','Fabric Code (SKU200)','Length Code (SKU300)','Color Code (RLM)','Size',
      'Cost (FOB/DDP)','Quantity','PO PIW/ETA Date','PO Notes','PO Product Notes','PO Cancel Date',
      'Ship Mode','Selling Prd','Selling Prd Year','PO Type','PO Terms (FOB ou DDP)',
      'Special Instructions','Comments/Special 01',
    ];

    // Sort sizes by SIZE_MAP order
    const sizeOrder = Object.values(SIZE_MAP);
    const sortedSizes = Object.entries(sizeTotals).sort((a, b) => {
      const ai = sizeOrder.indexOf(a[0]), bi = sizeOrder.indexOf(b[0]);
      return (ai === -1 ? 999 : ai) - (bi === -1 ? 999 : bi);
    });

    const wb2 = new ExcelJS.Workbook();
    const ws2 = wb2.addWorksheet('PO Import');

    // Header style — medium blue bg, white bold text
    const hdrFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E75B6' } };
    const hdrFont = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    const thinBorder = {
      top: { style: 'thin', color: { argb: 'FFB8CCE4' } },
      left: { style: 'thin', color: { argb: 'FFB8CCE4' } },
      bottom: { style: 'thin', color: { argb: 'FFB8CCE4' } },
      right: { style: 'thin', color: { argb: 'FFB8CCE4' } },
    };

    // Add header row
    const headerRow = ws2.addRow(OFFICIAL_HEADERS);
    headerRow.eachCell(cell => {
      cell.fill   = hdrFill;
      cell.font   = hdrFont;
      cell.border = thinBorder;
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    headerRow.height = 18;

    // Data rows — alternating light blue / white
    sortedSizes.forEach(([sz, qty], idx) => {
      const rowData = [
        '1','1',userPO,'A','26',exfFormatted,supplier,'305',cleanStyle,
        '','',colorCode,sz,Number(cost)||0,qty,etaDate,
        '','',exfFormatted,
        'A','A','26','',poTerms,'','',
      ];
      const row = ws2.addRow(rowData);
      const rowFill = idx % 2 === 0
        ? { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDEEAF1' } }
        : { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
      row.eachCell({ includeEmpty: true }, (cell, colNum) => {
        cell.fill   = rowFill;
        cell.border = thinBorder;
        if (colNum === 14) cell.numFmt = '$#,##0.00';   // Cost
        if (colNum === 15) cell.numFmt = '0';            // Quantity
      });
    });

    // Column widths based on header length
    ws2.columns.forEach((col, i) => {
      col.width = Math.min(30, Math.max(10, OFFICIAL_HEADERS[i].length * 1.1));
    });
    ws2.views = [{ state: 'frozen', ySplit: 1 }];

    const excelBuf = await wb2.xlsx.writeBuffer();

    // Email
    const subject = `[ANTHRO X FARM] official PO request - style # ${styleDigits}`;
    const toEmails = ['inbound@farmrio.com', 'danielle.gouvea@farmrio.com', 'anacarolina.azevedo@farmrio.com'];
    const ccEmails = ['paula@creativetwotwelve.com', 'rafaela@showroom212.com', 'ozan.guruscu@creativetwotwelve.com', 'business@creativetwotwelve.com', 'kamilla@creativetwotwelve.com'];

    const textBody =
      `Hi Ana,\n\n` +
      `Please find attached the base PO import of the following style:\n` +
      `Style#  ${styleDigits}\n` +
      `Supplier: ${supplier}\n` +
      `Anthro NDC ${anthroNDC}:\n\n` +
      `This style has also already been issued in RLM.\n\n` +
      `Thank you,\n\n` +
      `${sender?.name || sendingAs}`;

    const rawMime = await buildRawMime({
      from:        `"${sender?.name || sendingAs}" <${sender?.email || ''}>`,
      to:          toEmails.join(','),
      cc:          ccEmails.join(','),
      subject,
      text:        textBody,
      attachments: [{ filename: `PO_${cleanStyle}.xlsx`, content: excelBuf, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }],
    });
    const encoded = rawMime.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

    const authClient = makeOAuth2Client();
    authClient.setCredentials({ refresh_token: token.refreshToken });
    const gmail = google.gmail({ version: 'v1', auth: authClient });
    await gmail.users.drafts.create({ userId: 'me', requestBody: { message: { raw: encoded } } });

    res.json({ ok: true, message: 'Draft created in your Gmail Drafts folder' });
  } catch (e) {
    console.error('official-email error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── TOP STATUS Email — creates Gmail drafts for all suppliers ────────────────
app.post('/api/po/top-status-email', async (req, res) => {
  try {
    const sendingAs = req.body.sendingAs || 'kamilla';
    const token = userTokens[sendingAs];
    if (!token?.refreshToken) {
      const user = TEAM_USERS.find(u => u.id === sendingAs);
      return res.status(401).json({ error: `${user?.name || sendingAs} has not connected their Gmail. Please visit /setup first.` });
    }
    const sender = TEAM_USERS.find(u => u.id === sendingAs);

    const response = await fetch(csvUrl(0));
    if (!response.ok) throw new Error('Could not fetch production sheet');
    const rows = parseCSV(await response.text());
    if (rows.length < 2) return res.json({ ok: true, draftsCreated: [], skipped: [] });

    const H = rows[0].map(h => (h || '').trim().toLowerCase());
    const findCol = (...kws) => {
      for (const kw of kws) { const i = H.findIndex(h => h.includes(kw.toLowerCase())); if (i >= 0) return i; }
      return -1;
    };
    const findColOr = (fallback, ...kws) => { const i = findCol(...kws); return i >= 0 ? i : fallback; };

    const C = {
      style:       findCol('style #', 'style#', 'style'),
      status:      findCol('status'),
      supplier:    findCol('supplier'),
      category:    findCol('category'),
      subcategory: findCol('sub-category', 'subcategory'),
      responded:   findColOr(57, 'top responded', 'top - responded'),
      sent:        findColOr(58, 'top - sent', 'top sample sent'),
      deadline:    findColOr(59, 'supplier top deadline', 'top deadline'),
    };

    const get = (r, i) => i >= 0 ? (r[i] || '').trim() : '';
    const parseDate = val => { if (!val) return null; const d = new Date(val); return isNaN(d) ? null : d; };

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const formattedDate = `${today.getMonth()+1}/${today.getDate()}/${String(today.getFullYear()).slice(2)}`;

    const CC = [
      'paula@creativetwotwelve.com',
      'Production@showroom212.com',
      'rafaela@showroom212.com',
      'samples@creativetwotwelve.com',
      'kamilla@creativetwotwelve.com',
    ];

    const authClient = makeOAuth2Client();
    authClient.setCredentials({ refresh_token: token.refreshToken });
    const gmail = google.gmail({ version: 'v1', auth: authClient });

    const draftsCreated = [];
    const skipped = [];

    for (const [supplierKey, toEmails] of Object.entries(suppliers.emails)) {
      const contactName = suppliers.mainContact[supplierKey] || supplierKey;
      const entries = [];

      for (let i = 1; i < rows.length; i++) {
        const rawStyle = get(rows[i], C.style);
        if (!rawStyle) continue;
        const m = rawStyle.match(/[0-9][0-9A-Z\-]*$/);
        const styleNumber = m ? m[0] : rawStyle;

        const status      = get(rows[i], C.status);
        const rowSupplier = get(rows[i], C.supplier).toUpperCase();
        const responded   = get(rows[i], C.responded);
        const sent        = get(rows[i], C.sent);
        const deadlineRaw = get(rows[i], C.deadline);

        if (rowSupplier !== supplierKey) continue;
        if (!styleNumber || status !== "PO'd + production ok") continue;
        if (responded !== '') continue;
        if (sent && sent.toLowerCase().includes('sent')) continue;

        const supplierDeadline = parseDate(deadlineRaw);
        if (!supplierDeadline) continue;

        const diffInDays = Math.floor((supplierDeadline - today) / 86400000);

        entries.push({
          style:        styleNumber,
          deadlineDate: supplierDeadline,
          diffInDays,
          category:     get(rows[i], C.category),
          subcategory:  get(rows[i], C.subcategory),
        });
      }

      if (entries.length === 0) { skipped.push(supplierKey); continue; }

      // Fetch CAD images (CID inline attachments — render in all email clients)
      const fmtDate = d => `${d.getMonth()+1}/${d.getDate()}/${String(d.getFullYear()).slice(2)}`;
      const attachments = [];
      const entriesWithCad = await Promise.all(entries.map(async (e, idx) => {
        let cad = await getCadImage(e.style);
        if (!cad.found) {
          const norm = e.style.replace(/^[A-Za-z]+-?/, '').trim();
          if (norm && norm !== e.style) cad = await getCadImage(norm);
        }
        const cid = `cad-${idx}`;
        if (cad.found) {
          attachments.push({
            filename:    `${e.style}.jpg`,
            content:     Buffer.from(cad.imageData, 'base64'),
            encoding:    'base64',
            cid,
            contentType: cad.mimeType || 'image/jpeg',
          });
        }
        return { ...e, hasCad: cad.found, cid };
      }));

      const subject = `TOP Sample ${formattedDate} Status Request - ${supplierKey}`;

      const tableRows = entriesWithCad.map(e => {
        const isOverdue = e.diffInDays < 0;
        const deadlineCell = isOverdue
          ? `<span style="color:red;font-weight:bold;">${fmtDate(e.deadlineDate)} — OVERDUE</span>`
          : fmtDate(e.deadlineDate);
        const cadCell = e.hasCad
          ? `<img src="cid:${e.cid}" width="70" style="display:block;border:0;">`
          : '';
        return `<tr>
          <td style="padding:4px;text-align:center;">${cadCell}</td>
          <td style="padding:6px;">${e.style}</td>
          <td style="padding:6px;">PENDING</td>
          <td style="padding:6px;">${deadlineCell}</td>
          <td style="padding:6px;">${e.category}</td>
          <td style="padding:6px;">${e.subcategory}</td>
          <td style="padding:6px;"></td>
        </tr>`;
      }).join('');

      const htmlBody = `Hi ${contactName} and ${supplierKey} team,<br><br>
Could you please provide an update on the TOP samples listed below?<br><br>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:11pt;">
  <thead>
    <tr style="background-color:#d9edf7;text-align:left;">
      <th style="padding:6px;">CAD</th>
      <th style="padding:6px;">Style #</th>
      <th style="padding:6px;">Status</th>
      <th style="padding:6px;">Deadline</th>
      <th style="padding:6px;">Category</th>
      <th style="padding:6px;">Sub-category</th>
      <th style="padding:6px;">Update</th>
    </tr>
  </thead>
  <tbody>${tableRows}</tbody>
</table>
<br><br>Thank you,<br>${sender?.email || ''}`;

      const rawMime = await buildRawMime({
        from:        `"${sender?.name || sendingAs}" <${sender?.email || ''}>`,
        to:          toEmails.join(','),
        cc:          CC.join(','),
        subject,
        html:        htmlBody,
        attachments,
      });
      const encoded = rawMime.toString('base64').replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');
      await gmail.users.drafts.create({ userId:'me', requestBody:{ message:{ raw:encoded } } });
      draftsCreated.push({ supplier: supplierKey, styles: entries.length });
    }

    res.json({ ok: true, draftsCreated, skipped });
  } catch (e) {
    console.error('top-status-email error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/contact', async (req, res) => {
  const { name, email, message } = req.body;
  if (!name || !message) return res.status(400).json({ error: 'Name and message are required.' });
  try {
    const transporter = nodemailer.createTransport({ service:'gmail', auth:{ user:process.env.EMAIL_USER, pass:process.env.EMAIL_PASS } });
    await transporter.sendMail({
      from: `"305 Workspace" <${process.env.EMAIL_USER}>`,
      to: 'support@creativetwotwelve.com',
      replyTo: email || process.env.EMAIL_USER,
      subject: `Message from ${name} — 305 Workspace`,
      text: `From: ${name}\nEmail: ${email||'not provided'}\n\n${message}`,
      html: `<p><strong>From:</strong> ${name}</p><p><strong>Email:</strong> ${email||'not provided'}</p><br/><p>${message.replace(/\n/g,'<br/>')}</p>`,
    });
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ─── FedEx Validation Email Draft ────────────────────────────────────────────
app.post('/api/fedex/validation-draft', async (req, res) => {
  try {
    const token = userTokens['flavio'];
    if (!token?.refreshToken) return res.status(401).json({ error: 'flavio Gmail not connected. Visit /setup first.' });

    const htmlBody = `
<p>Hi FedEx Label Analysis Team,</p>
<p>Please find attached our test label for Ship API production validation.</p>
<table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:13px;">
  <tr><td style="padding:4px 12px 4px 0;font-weight:bold;color:#4D148C;">Company</td><td>305 Consulting and Production Inc</td></tr>
  <tr><td style="padding:4px 12px 4px 0;font-weight:bold;color:#4D148C;">FedEx Account Number</td><td>740561073</td></tr>
  <tr><td style="padding:4px 12px 4px 0;font-weight:bold;color:#4D148C;">Production API Key</td><td>l766161b6e947e4d08b5284266e7afcee8</td></tr>
  <tr><td style="padding:4px 12px 4px 0;font-weight:bold;color:#4D148C;">Contact</td><td>support@creativetwotwelve.com | 917-499-2103</td></tr>
  <tr><td style="padding:4px 12px 4px 0;font-weight:bold;color:#4D148C;">Services Requested</td><td>FedEx Ground (domestic US)</td></tr>
  <tr><td style="padding:4px 12px 4px 0;font-weight:bold;color:#4D148C;">Application</td><td>Internal warehouse shipping system — 305 Workspace</td></tr>
</table>
<br>
<p>Attached: scanned test label printed at 600 DPI.</p>
<p>Please advise on approval.</p>
<p>Thank you,<br>Flavio Azevedo<br>305 Consulting and Production Inc</p>`;

    const rawMime = await buildRawMime({
      from:    '"Flavio Azevedo" <support@creativetwotwelve.com>',
      to:      'label@fedex.com',
      subject: 'FedEx Ship API Label Validation — 305 Consulting and Production Inc',
      html:    htmlBody,
    });
    const encoded = rawMime.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

    const authClient = makeOAuth2Client();
    authClient.setCredentials({ refresh_token: token.refreshToken });
    const gmail = google.gmail({ version: 'v1', auth: authClient });
    await gmail.users.drafts.create({ userId: 'me', requestBody: { message: { raw: encoded } } });

    res.json({ ok: true, message: 'Draft saved to support@creativetwotwelve.com — check Gmail Drafts' });
  } catch (e) {
    console.error('validation-draft error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── FedEx Label Creation ─────────────────────────────────────────────────────
function colToLetter(col) {
  let letter = '', n = col + 1;
  while (n > 0) { const r = (n - 1) % 26; letter = String.fromCharCode(65 + r) + letter; n = Math.floor((n - 1) / 26); }
  return letter;
}

app.post('/api/fedex/create-label', async (req, res) => {
  try {
    const { po, boxQty, weight, serviceCode, deliverToCode, supplier } = req.body;
    if (!po || !boxQty || !serviceCode || !deliverToCode)
      return res.status(400).json({ error: 'Missing required fields: po, boxQty, serviceCode, deliverToCode' });

    const to = FEDEX_ADDRESS_BOOK[deliverToCode];
    if (!to) return res.status(400).json({ error: `Unknown deliver-to code: ${deliverToCode}` });
    const from = FEDEX_ADDRESS_BOOK.SHIP_FROM;

    const nk = s => (s || '').toLowerCase().trim();
    const dimsEntry = Object.entries(FEDEX_PACKAGING).find(([k]) =>
      nk(k) === nk(supplier) || nk(supplier).includes(nk(k)) || nk(k).includes(nk(supplier))
    );
    const dims = dimsEntry ? dimsEntry[1] : { l: 20, w: 12, h: 12 };
    const qty  = Math.max(1, parseInt(boxQty) || 1);
    const wt   = parseFloat(weight) || 20;

    // 1. FedEx OAuth token
    const tokenRes = await fetch(`${FEDEX_API_URL}/oauth/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: `grant_type=client_credentials&client_id=${encodeURIComponent(process.env.FEDEX_SHIP_API_KEY)}&client_secret=${encodeURIComponent(process.env.FEDEX_SHIP_API_SECRET)}`,
    });
    if (!tokenRes.ok) {
      const t = await tokenRes.text();
      return res.status(500).json({ error: `FedEx auth failed (${tokenRes.status}): ${t}` });
    }
    const { access_token } = await tokenRes.json();

    // 2. Build payload
    const packages = Array.from({ length: qty }, (_, i) => ({
      sequenceNumber: i + 1,
      weight:     { units: 'LB', value: wt },
      dimensions: { length: dims.l, width: dims.w, height: dims.h, units: 'IN' },
    }));

    const payload = {
      labelResponseOptions: 'URL_ONLY',
      requestedShipment: {
        shipper: {
          contact: { personName: from.contact, companyName: from.company, phoneNumber: from.phone },
          address: { streetLines: from.street, city: from.city, stateOrProvinceCode: from.state, postalCode: from.zip, countryCode: from.country },
        },
        recipients: [{
          contact: { personName: to.contact, companyName: to.company, phoneNumber: to.phone || '0000000000' },
          address: { streetLines: to.street, city: to.city, stateOrProvinceCode: to.state, postalCode: to.zip, countryCode: to.country },
        }],
        serviceType: serviceCode,
        packagingType: 'YOUR_PACKAGING',
        pickupType: 'USE_SCHEDULED_PICKUP',
        shippingChargesPayment: {
          paymentType: 'SENDER',
          payor: { responsibleParty: { accountNumber: { value: process.env.FEDEX_ACCOUNT_NUMBER } } },
        },
        labelSpecification: { labelFormatType: 'COMMON2D', imageType: 'PDF', labelStockType: 'PAPER_4X6' },
        customerReferences: [{ customerReferenceType: 'CUSTOMER_REFERENCE', value: po }],
        requestedPackageLineItems: packages,
      },
      accountNumber: { value: process.env.FEDEX_ACCOUNT_NUMBER },
    };

    // 3. Create shipment
    const shipRes = await fetch(`${FEDEX_API_URL}/ship/v1/shipments`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${access_token}`, 'Content-Type': 'application/json', 'x-locale': 'en_US' },
      body: JSON.stringify(payload),
    });
    const shipData = await shipRes.json();
    if (!shipRes.ok) {
      const errMsg = shipData?.errors?.[0]?.message || JSON.stringify(shipData).slice(0, 300);
      return res.status(500).json({ error: `FedEx Ship error: ${errMsg}`, raw: shipData });
    }

    const shipment    = shipData.output?.transactionShipments?.[0];
    if (!shipment) return res.status(500).json({ error: 'No shipment in FedEx response', raw: shipData });

    const tracking    = shipment.masterTrackingNumber || shipment.pieceResponses?.[0]?.trackingNumber || '';
    const allLabels   = (shipment.pieceResponses || []).flatMap(p => (p.packageDocuments || []).map(d => d.url).filter(Boolean));
    const rateDetail  = shipment.completedShipmentDetail?.shipmentRating?.shipmentRateDetails;
    const totalCharge = rateDetail?.[0]?.totalNetCharge?.amount ?? rateDetail?.[0]?.totalNetCharge ?? '';
    const serviceName = FEDEX_SERVICE_NAMES[serviceCode] || serviceCode;

    // 4. Upload labels to Google Drive (best-effort)
    let driveWebUrl = '';
    if (allLabels.length && process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
      try {
        const sa   = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
        const auth = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/drive'] });
        const drive = google.drive({ version: 'v3', auth });
        const MONTHS = ['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE','JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER'];
        let folderId = FEDEX_DRIVE_FOLDER;
        try {
          const listRes = await drive.files.list({
            q: `'${FEDEX_DRIVE_FOLDER}' in parents and mimeType='application/vnd.google-apps.folder' and name='${MONTHS[new Date().getMonth()]}' and trashed=false`,
            fields: 'files(id)', pageSize: 1,
          });
          if (listRes.data.files?.length) folderId = listRes.data.files[0].id;
        } catch (_) {}

        for (let i = 0; i < allLabels.length; i++) {
          const pdfRes = await fetch(allLabels[i]);
          if (!pdfRes.ok) continue;
          const buf = Buffer.from(await pdfRes.arrayBuffer());
          const { Readable } = require('stream');
          const fileName = allLabels.length === 1 ? `FEDEX ${po}` : `FEDEX ${po} - ${i + 1}`;
          const up = await drive.files.create({
            requestBody: { name: fileName, mimeType: 'application/pdf', parents: [folderId] },
            media: { mimeType: 'application/pdf', body: Readable.from(buf) },
            fields: 'id,webViewLink',
          });
          if (i === 0) driveWebUrl = up.data.webViewLink || '';
        }
      } catch (driveErr) { console.error('[Drive upload]', driveErr.message); }
    }

    // 5. Write back to Warehouse Now sheet (best-effort)
    if (tracking && process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
      try {
        const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
        const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
        const sheets = google.sheets({ version: 'v4', auth });
        const TAB    = 'Warehouse Now Database';

        const hdRes = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `'${TAB}'!1:1` });
        const hdrs  = (hdRes.data.values?.[0] || []).map(h => h.trim().toLowerCase());
        const fc = (...kws) => {
          for (const kw of kws) { const i = hdrs.findIndex(h => h === kw.toLowerCase()); if (i >= 0) return i; }
          for (const kw of kws) { const i = hdrs.findIndex(h => h.includes(kw.toLowerCase())); if (i >= 0) return i; }
          return -1;
        };
        const trackCol   = fc('tracking number', 'tracking #', 'tracking');
        const carrierCol = fc('carrier');
        const serviceCol = fc('shipping type', 'service type', 'ship type');
        const costCol    = fc('shipping cost', 'freight cost', 'ship cost');
        const poColIdx   = fc('po#', 'po number', 'po');

        if (poColIdx >= 0 && trackCol >= 0) {
          const colL   = colToLetter(poColIdx);
          const poVals = (await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `'${TAB}'!${colL}:${colL}` })).data.values || [];
          let sheetRow = -1;
          for (let i = 1; i < poVals.length; i++) {
            if ((poVals[i]?.[0] || '').toString().trim() === po.toString().trim()) { sheetRow = i + 1; break; }
          }
          if (sheetRow > 0) {
            const updates = [
              ...(trackCol   >= 0 ? [{ range: `'${TAB}'!${colToLetter(trackCol)}${sheetRow}`,   values: [[tracking]]          }] : []),
              ...(carrierCol >= 0 ? [{ range: `'${TAB}'!${colToLetter(carrierCol)}${sheetRow}`,  values: [['FEDEX']]            }] : []),
              ...(serviceCol >= 0 ? [{ range: `'${TAB}'!${colToLetter(serviceCol)}${sheetRow}`,  values: [[serviceName]]        }] : []),
              ...(costCol    >= 0 ? [{ range: `'${TAB}'!${colToLetter(costCol)}${sheetRow}`,     values: [[String(totalCharge)]]}] : []),
            ];
            if (updates.length) await sheets.spreadsheets.values.batchUpdate({
              spreadsheetId: SHEET_ID,
              requestBody: { valueInputOption: 'RAW', data: updates },
            });
          }
        }
      } catch (sheetErr) { console.error('[Sheet write]', sheetErr.message); }
    }

    res.json({ ok: true, tracking, cost: totalCharge ? `$${totalCharge}` : '', driveUrl: driveWebUrl });
  } catch (e) {
    console.error('[create-label]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── Rebeca: Design / Style routes ───────────────────────────────────────────
app.use('/api/rebeca', require('./routes/rebeca'));

// ─── Paul: Samples Tracking routes ───────────────────────────────────────────
app.use('/api/paul', require('./routes/paul'));

// ATTN/BUYER dropdown value → { brand, category } for buyer lookup
const PAUL_BUYER_MAP = {
  'JULIE/DRESSES':              { brand: 'anthropologie', category: 'DRESS & ROMPER' },
  'IVY/SWIM':                   { brand: 'anthropologie', category: 'SWIMWEAR' },
  'ALY/BLOUSES':                { brand: 'anthropologie', category: 'BLOUSES & SHIRTS' },
  'CAROLINE/BOTTOM':            { brand: 'anthropologie', category: 'PANTS & JUMPSUIT' },
  'ABBY/SKIRTS':                { brand: 'anthropologie', category: 'SKIRTS & SHORTS' },
  'Jacqueline Mazzi/ KNITS':    { brand: 'anthropologie', category: 'LOUNGE' },
  'MC Miskuly/Maternity/Nuuly': { brand: 'nuuly',         category: 'PANTS, JUMPSUIT, SKIRTS & SHORTS' },
};

// 305 team always CC'd on sample emails
const PAUL_CC = [
  'kamilla@creativetwotwelve.com',
  'paula@creativetwotwelve.com',
  'business@creativetwotwelve.com',
  'rafaela.neves@farmrio.com',
  'ozan.guruscu@creativetwotwelve.com',
];

app.post('/api/paul/draft-email', async (req, res) => {
  try {
    const { attnBuyer, samples } = req.body || {};
    if (!attnBuyer)                                    return res.status(400).json({ error: 'Attn/Buyer is required' });
    if (!Array.isArray(samples) || !samples.length)    return res.status(400).json({ error: 'samples array is required' });

    // Use the logged-in user's Gmail account
    const sessionUser = req.session?.user;
    if (!sessionUser) return res.status(401).json({ error: 'Not authenticated' });
    const teamUser = TEAM_USERS.find(u => u.email.toLowerCase() === sessionUser.email.toLowerCase());
    if (!teamUser) return res.status(400).json({ error: `${sessionUser.name} is not configured for Gmail drafts` });
    const token = userTokens[teamUser.id];
    if (!token?.refreshToken) {
      return res.status(401).json({ error: `${sessionUser.name}'s Gmail is not connected` });
    }

    // Resolve buyer list from ATTN/BUYER value
    const mapping   = PAUL_BUYER_MAP[attnBuyer];
    if (!mapping) return res.status(400).json({ error: `No buyer mapping found for: ${attnBuyer}` });
    const buyerList = (buyers[mapping.brand]?.[mapping.category]) || [];
    if (!buyerList.length) return res.status(400).json({ error: `No buyer emails found for category: ${mapping.category}` });

    const toEmails  = buyerList.map(b => b.email);
    const leadFirst = buyerList[0].name.split(' ')[0];

    // Friendly category label for greeting: "DRESS & ROMPER" → "dress"
    const CAT_LABEL = {
      'DRESS & ROMPER':                     'dress',
      'SWIMWEAR':                           'swim',
      'BLOUSES & SHIRTS':                   'blouses',
      'PANTS & JUMPSUIT':                   'pants',
      'SKIRTS & SHORTS':                    'skirts',
      'LOUNGE':                             'lounge',
      'PANTS, JUMPSUIT, SKIRTS & SHORTS':   'bottoms',
    };
    const catLabel = CAT_LABEL[mapping.category] || mapping.category.toLowerCase();

    let subject, htmlBody;

    if (samples.length === 1) {
      const s = samples[0];
      subject = `FARM RIO // Style# ${s.style || ''} - Tracking ${s.tracking} - ${s.sampleType}`;
      htmlBody =
        `<p>Hello ${leadFirst} and ${catLabel} team,</p>` +
        `<p>We sent you the style below by tracking <strong>${s.tracking}</strong><br>` +
        `<strong>${s.style || ''}</strong>${s.style ? ' - ' : ''}${s.sampleType}</p>` +
        (s.disclaimer ? `<p>${s.disclaimer}</p>` : '') +
        `<p>Best regards,<br>${sessionUser.name}<br>305 Consulting and Production</p>`;
    } else {
      subject = `FARM RIO // ${samples.length} Samples - ${catLabel} team`;
      const tableRows = samples.map(s =>
        `<tr><td style="padding:4px 12px 4px 0">${s.style || '—'}</td>` +
        `<td style="padding:4px 12px 4px 0">${s.tracking}</td>` +
        `<td style="padding:4px 12px 4px 0">${s.sampleType}</td>` +
        `<td style="padding:4px 0">${s.disclaimer || ''}</td></tr>`
      ).join('');
      htmlBody =
        `<p>Hello ${leadFirst} and ${catLabel} team,</p>` +
        `<p>We are sending you the following samples:</p>` +
        `<table style="border-collapse:collapse;font-family:sans-serif;font-size:13px">` +
        `<tr style="font-weight:700;border-bottom:2px solid #000">` +
        `<td style="padding:4px 12px 4px 0">Style #</td>` +
        `<td style="padding:4px 12px 4px 0">Tracking</td>` +
        `<td style="padding:4px 12px 4px 0">Sample Type</td>` +
        `<td style="padding:4px 0">Notes</td></tr>` +
        tableRows + `</table>` +
        `<p>Best regards,<br>${sessionUser.name}<br>305 Consulting and Production</p>`;
    }

    const rawMime = await buildRawMime({
      from:    `"${sessionUser.name}" <${sessionUser.email}>`,
      to:      toEmails.join(', '),
      cc:      PAUL_CC.join(', '),
      subject,
      html:    htmlBody,
    });
    const encoded = rawMime.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

    const authClient = makeOAuth2Client();
    authClient.setCredentials({ refresh_token: token.refreshToken });
    const gmail = google.gmail({ version: 'v1', auth: authClient });
    await gmail.users.drafts.create({ userId: 'me', requestBody: { message: { raw: encoded } } });

    res.json({ ok: true });
  } catch (e) {
    console.error('[paul/draft-email]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── Gabriel: MAP DATA (chart feed) ──────────────────────────────────────────
app.get('/api/gabriel/map-data', async (req, res) => {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    return res.status(500).json({ error: 'Google credentials not configured' });
  }
  try {
    const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'] });
    const sheets = google.sheets({ version: 'v4', auth });
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: MAP_SHEET_ID,
      range: `'ANTHRO MAP 2026'`,
      valueRenderOption: 'FORMATTED_VALUE',
      dateTimeRenderOption: 'FORMATTED_STRING',
    });
    const values = r.data.values || [];
    res.setHeader('Cache-Control', 'no-store');
    res.json({ values });
  } catch (e) {
    console.error('[map-data]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── Samantha: PowerBI Database Sync ─────────────────────────────────────────
app.post('/api/samantha/powerbi-sync', async (req, res) => {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    return res.status(500).json({ error: 'Google credentials not configured' });
  }
  try {
    const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
    const sheets = google.sheets({ version: 'v4', auth });

    const readTab = async (tab, formula = false) => {
      const r = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID, range: `'${tab}'`,
        valueRenderOption: formula ? 'FORMULA' : 'UNFORMATTED_VALUE',
        dateTimeRenderOption: 'FORMATTED_STRING',
      });
      return r.data.values || [];
    };
    const isFormulaPB = v => typeof v === 'string' && v.startsWith('=');
    const colLetter = idx => {
      let s = '', n = idx;
      do { s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) - 1; } while (n >= 0);
      return s;
    };
    const fmtMonDayYear = v => {
      if (!v) return '';
      const d = new Date(String(v));
      if (isNaN(d.getTime())) return '';
      const M = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
      return M[d.getMonth()] + '/' + d.getDate() + '/' + d.getFullYear();
    };

    const [metaRes, tsData, poNewData, poInvData] = await Promise.all([
      sheets.spreadsheets.get({ spreadsheetId: SHEET_ID, fields: 'sheets.properties' }),
      readTab('TRADESTONE DATABASE', true), // FORMULA mode — preserves existing formulas
      readTab('PO NEW'), readTab('PO INVOICE'),
    ]);
    const tsSheetId = (metaRes.data.sheets.find(s => s.properties.title === 'TRADESTONE DATABASE')?.properties.sheetId) ?? 0;

    const tsH = tsData[0] || [], poNewH = poNewData[0] || [], poInvH = poInvData[0] || [];
    const hMap = H => { const m = {}; H.forEach((h,i) => { const k = String(h).trim().toLowerCase(); if (k) m[k]=i; }); return m; };
    const tsHM = hMap(tsH), pnHM = hMap(poNewH), piHM = hMap(poInvH);

    const tsPO  = tsHM['po#'] ?? -1;
    const pnPO  = pnHM['po#'] ?? -1;
    const piPO  = piHM['po#'] ?? -1;
    const piDate  = piHM['invoice date']  ?? -1;
    const piValue = piHM['invoice value'] ?? piHM['invoice amount'] ?? -1;
    const tsShip  = tsHM['ship date'] ?? -1;
    const pnShip  = pnHM['ship date'] ?? -1;
    const pnCancel = pnHM['cancel date'] ?? -1;
    const tsIPcol  = tsHM['ip class'] ?? tsHM['ipclass'] ?? -1;
    const realCancelIdx = tsHM['real cancel date'] ?? -1;

    // Columns that must not be overwritten from PO NEW (V=21, AL=37, AM=38)
    const SKIP = new Set([21, 37, 38]);
    const COL_AJ = 35, COL_AK = 36, COL_H = 7, COL_V = 21;

    const IP_CAT = {
      1501:'Blouses',1502:'SWTRS & SWTSHRTS',1504:'Pants',1505:'Skirts',1506:'SHORTS',
      1508:'Dresses',1509:'Dresses',1515:'Dresses',4110:'Blouses',4114:'Sweaters',
      4120:'Skirts',4123:'Pants',4124:'Jumpsuit',4125:'Shorts',4130:'Dresses',
      4134:'Dresses',4141:'LOUNGE',4142:'Swimwear',4148:'Blouses',4152:'Hats',
      4153:'Accessories',4157:'Wraps',4428:'Dresses',4915:'Waters Edge',
      4920:'Blouses',4925:'Skirts',4927:'Pants',4928:'Dresses',4942:'Swimwear',
    };

    // Column mapping: TRADESTONE col → PO NEW col (matched by header name, skip protected)
    const colMap = [];
    for (let ti = 0; ti < tsH.length; ti++) {
      if (SKIP.has(ti)) continue;
      const k = String(tsH[ti]).trim().toLowerCase();
      if (k && pnHM[k] !== undefined) colMap.push({ tsIdx: ti, poIdx: pnHM[k] });
    }

    // Existing TRADESTONE rows by PO
    let lastDataRowIdx = 0;
    const existByPO = new Map();
    for (let i = 1; i < tsData.length; i++) {
      const po = String(tsData[i][tsPO] || '').trim().toUpperCase();
      if (!po) continue;
      if (!existByPO.has(po)) existByPO.set(po, i);
      lastDataRowIdx = i;
    }

    // Invoice data by PO
    const invByPO = new Map();
    for (let i = 1; i < poInvData.length; i++) {
      const po = String(poInvData[i][piPO] || '').trim();
      if (po) invByPO.set(po, {
        AJ: piValue >= 0 ? poInvData[i][piValue] : '',
        AK: piDate  >= 0 ? poInvData[i][piDate]  : '',
      });
    }

    const colH = row => { const ak = row[COL_AK]; return (ak != null && ak !== '') ? ak : (row[6] || ''); };

    // Step 1 — refresh AJ/AK invoice data on ALL existing rows
    for (const [po, inv] of invByPO) {
      const ei = existByPO.get(po.toUpperCase());
      if (ei === undefined) continue;
      const r = tsData[ei];
      if (inv.AJ != null && inv.AJ !== '') r[COL_AJ] = inv.AJ;
      if (inv.AK != null && inv.AK !== '') r[COL_AK] = inv.AK;
      r[COL_H] = colH(r);
      if (realCancelIdx >= 0) r[realCancelIdx] = fmtMonDayYear(r[COL_H]);
    }

    // Step 2 — add only NEW POs (not already in TRADESTONE DATABASE)
    // Find formula column indices by header name
    const cat2Idx    = tsHM['category 2']    ?? tsHM['category2']     ?? -1;
    const cat3Idx    = tsHM['category 3']    ?? tsHM['category3']     ?? -1;
    const boxesIdx   = tsHM['boxes']         ?? -1;
    const totalQtyIdx = tsHM['total qty']    ?? tsHM['total quantity'] ?? tsHM['totalqty'] ?? tsHM['total units'] ?? -1;
    const yr2Idx     = tsHM['year 2']        ?? tsHM['year2']          ?? -1;
    const mo2Idx     = tsHM['month 2']       ?? tsHM['month2']         ?? -1;

    // firstNewSheetRow: the sheet row where the first appended row will land.
    // tsData[lastDataRowIdx] is at sheet row lastDataRowIdx+1, so next row is lastDataRowIdx+2.
    // Using lastDataRowIdx instead of tsData.length avoids inflation from stale formula-only rows
    // that the formula-mode read returns beyond the actual data.
    const firstNewSheetRow = lastDataRowIdx + 2;

    const newRows = [];
    for (let si = 1; si < poNewData.length; si++) {
      const pnRow = poNewData[si];
      const po  = String(pnRow[pnPO]    || '').trim();
      const st  = String(pnRow[pnHM['vendor style #'] ?? -1] || '').trim();
      if (!po || !st) continue;
      if (existByPO.has(po.toUpperCase())) continue; // already in TRADESTONE — skip

      const nr = Array(tsH.length).fill('');
      for (const {tsIdx, poIdx} of colMap) nr[tsIdx] = pnRow[poIdx];
      if (tsShip >= 0 && pnShip >= 0) nr[tsShip] = pnRow[pnShip];
      if (pnCancel >= 0) nr[COL_H] = pnRow[pnCancel];
      const inv = invByPO.get(po);
      if (inv) { nr[COL_AJ] = inv.AJ; nr[COL_AK] = inv.AK; }
      nr[COL_H] = colH(nr);
      if (tsIPcol >= 0) nr[COL_V] = IP_CAT[parseInt(nr[tsIPcol])] || 'OTHER';
      if (realCancelIdx >= 0) nr[realCancelIdx] = fmtMonDayYear(nr[COL_H]);

      newRows.push(nr);
    }

    // Deduplicate newRows by PO# (in case PO NEW has duplicate entries)
    const seenPOs = new Set();
    const dedupedRows = newRows.filter(row => {
      const po = String(row[tsPO] || '').trim().toUpperCase();
      if (!po || seenPOs.has(po)) return false;
      seenPOs.add(po);
      return true;
    });

    // Write AJ/AK/H/RealCancelDate updates — only for actual data rows (up to lastDataRowIdx)
    const ajVals = [], akVals = [], hVals = [], rcVals = [];
    for (let i = 1; i <= lastDataRowIdx; i++) {
      const orig = tsData[i];
      ajVals.push([isFormulaPB(orig[COL_AJ]) ? orig[COL_AJ] : (orig[COL_AJ] ?? '')]);
      akVals.push([isFormulaPB(orig[COL_AK]) ? orig[COL_AK] : (orig[COL_AK] ?? '')]);
      hVals.push([isFormulaPB(orig[COL_H])  ? orig[COL_H]  : (orig[COL_H]  ?? '')]);
      if (realCancelIdx >= 0) rcVals.push([isFormulaPB(orig[realCancelIdx]) ? orig[realCancelIdx] : (orig[realCancelIdx] ?? '')]);
    }

    if (tsData.length > 1) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SHEET_ID,
        requestBody: {
          valueInputOption: 'USER_ENTERED',
          data: [
            { range: `'TRADESTONE DATABASE'!AJ2`, values: ajVals },
            { range: `'TRADESTONE DATABASE'!AK2`, values: akVals },
            { range: `'TRADESTONE DATABASE'!H2`,  values: hVals  },
            ...(realCancelIdx >= 0 && rcVals.length ? [{ range: `'TRADESTONE DATABASE'!${colLetter(realCancelIdx)}2`, values: rcVals }] : []),
          ],
        },
      });
    }

    // Append new rows first (no formula columns), then read the exact rows they landed on
    // from the API response — that's the only reliable way to know the actual row numbers.
    if (dedupedRows.length) {
      const appendRes = await sheets.spreadsheets.values.append({
        spreadsheetId: SHEET_ID, range: `'TRADESTONE DATABASE'!A1`,
        valueInputOption: 'USER_ENTERED', insertDataOption: 'INSERT_ROWS',
        requestBody: { values: dedupedRows },
      });

      // updatedRange looks like "'TRADESTONE DATABASE'!A2920:BQ2925" — extract starting row
      const rangeMatch = (appendRes.data.updates?.updatedRange || '').match(/!A(\d+)/);
      const actualFirstRow = rangeMatch ? parseInt(rangeMatch[1]) : null;

      if (actualFirstRow !== null) {
        const tqCol = totalQtyIdx >= 0 ? colLetter(totalQtyIdx) : null;
        const formulaData = [];
        for (let i = 0; i < dedupedRows.length; i++) {
          const r = actualFirstRow + i;
          if (cat2Idx  >= 0) formulaData.push({ range: `'TRADESTONE DATABASE'!${colLetter(cat2Idx)}${r}`,  values: [[`=IFNA(VLOOKUP(V${r},{"Dresses","Dresses";"Rompers","Dresses";"JUMPERS & ROMPERS","Dresses";"Blouses","Blouses";"BLOUSES & SHIRTS","Blouses";"SLEEP","Lounge";"Fine Gauge","Sweaters";"Sweaters","Sweaters";"SWTRS & SWTSHRTS","Sweaters";"Heavyweight","Knit";"Knit","Knit";"Pants","Bottoms";"PANTS & LEGGINGS","Bottoms";"Jumpsuit","Bottoms";"Swimwear","Swimwear";"Water''s Edge","Swimwear";"Wraps","Accessories";"Shorts","Shorts";"Skirts","Skirts"},2,0),"No Match")`]] });
          if (cat3Idx  >= 0) formulaData.push({ range: `'TRADESTONE DATABASE'!${colLetter(cat3Idx)}${r}`,  values: [[`=IFNA(VLOOKUP(V${r},{"Sleep","Lounge";"Blouses","Blouses";"BLOUSES & SHIRTS","Blouses";"Dresses","Dresses";"Fine Gauge","Sweaters";"Heavyweight","Knit";"JUMPERS & ROMPERS","Bottoms";"Jumpsuit","Bottoms";"Pants","Bottoms";"PANTS & LEGGINGS","Bottoms";"Rompers","Dresses";"Shorts","Skirts";"Skirts","Skirts";"Sweaters","Sweaters";"SWTRS & SWTSHRTS","Sweaters";"Swimwear","Swimwear";"Water''s Edge","Swimwear";"Wraps","Accessories"},2,0),"No Match")`]] });
          if (boxesIdx >= 0 && tqCol) formulaData.push({ range: `'TRADESTONE DATABASE'!${colLetter(boxesIdx)}${r}`, values: [[`=${tqCol}${r}/30`]] });
          if (yr2Idx   >= 0) formulaData.push({ range: `'TRADESTONE DATABASE'!${colLetter(yr2Idx)}${r}`,   values: [[`=YEAR(H${r})`]] });
          if (mo2Idx   >= 0) formulaData.push({ range: `'TRADESTONE DATABASE'!${colLetter(mo2Idx)}${r}`,   values: [[`=TEXT(H${r},"MM") & " - " & TEXT(H${r},"MMM")`]] });
        }
        if (formulaData.length) {
          await sheets.spreadsheets.values.batchUpdate({
            spreadsheetId: SHEET_ID,
            requestBody: { valueInputOption: 'USER_ENTERED', data: formulaData },
          });
        }
      }
    }

    // Dedup: read column C fresh (plain values), find duplicate PO#s, delete those rows
    const colCRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID, range: `'TRADESTONE DATABASE'!C:C`,
      valueRenderOption: 'UNFORMATTED_VALUE',
    });
    const colCVals = colCRes.data.values || [];
    const seenPO = new Set();
    const dupRows = [];
    for (let i = 1; i < colCVals.length; i++) {
      const po = String(colCVals[i]?.[0] || '').trim();
      if (!po) continue;
      if (seenPO.has(po)) dupRows.push(i + 1); else seenPO.add(po);
    }
    if (dupRows.length) {
      const asc = dupRows.sort((a, b) => a - b);
      const ranges = [];
      let s = asc[0], e = s;
      for (let i = 1; i < asc.length; i++) {
        if (asc[i] === e + 1) { e = asc[i]; } else { ranges.push([s, e]); s = e = asc[i]; }
      }
      ranges.push([s, e]);
      ranges.sort((a, b) => b[0] - a[0]);
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SHEET_ID,
        requestBody: {
          requests: ranges.map(([s, e]) => ({
            deleteDimension: { range: { sheetId: tsSheetId, dimension: 'ROWS', startIndex: s - 1, endIndex: e } }
          })),
        },
      });
    }

    res.json({ ok: true, newRows: dedupedRows.length, invoiceRefreshed: ajVals.length, duplicatesRemoved: dupRows.length });
  } catch (e) {
    console.error('[powerbi-sync]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── Samantha: URBN Invoice Generator ────────────────────────────────────────

// Preview: return SHIPPED POs with no URBN INVOICE DATE for the modal checklist
app.get('/api/samantha/invoice-preview', async (req, res) => {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    return res.status(500).json({ error: 'Google credentials not configured' });
  }
  try {
    const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'] });
    const sheets = google.sheets({ version: 'v4', auth });

    const r    = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Warehouse Now Database' });
    const rows = r.data.values || [];
    if (rows.length < 2) return res.json([]);

    const headers = rows[0].map(h => (h || '').trim().toLowerCase());
    const idx = keys => headers.findIndex(h => keys.some(k => h.includes(k)));

    const poCol       = idx(['po#', 'po number']);
    const statusCol   = idx(['status']);
    const styleCol    = idx(['style#', 'style number', 'style']);
    const trackingCol = idx(['tracking number', 'tracking']);
    const invDateCol  = idx(['urbn invoice date']);

    if (poCol < 0 || statusCol < 0) return res.status(500).json({ error: 'Required columns not found' });

    const result = [];
    for (let i = 1; i < rows.length; i++) {
      const row      = rows[i];
      const poNumber = (row[poCol]     || '').trim();
      const status   = (row[statusCol] || '').trim().toUpperCase();
      const invDate  = invDateCol >= 0 ? (row[invDateCol] || '').trim() : null;

      if (!poNumber) continue;
      if (status !== 'SHIPPED') continue;
      if (invDate !== null && invDate !== '') continue;

      result.push({
        poNumber,
        style:    styleCol    >= 0 ? (row[styleCol]    || '').trim() : '',
        tracking: trackingCol >= 0 ? (row[trackingCol] || '').trim() : '',
      });
    }

    res.json(result);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/samantha/invoice-status', async (req, res) => {
  try {
    const r    = await ghFetch(`https://api.github.com/repos/${REPO}/actions/workflows/urbn-invoice-generator.yml/runs?per_page=1`);
    const data = await r.json();
    const run  = data.workflow_runs?.[0] || null;
    const duration = run && run.status === 'completed' && run.run_started_at && run.updated_at
      ? Math.round((new Date(run.updated_at) - new Date(run.run_started_at)) / 1000) : null;
    res.json({
      status:     run?.status     || 'unknown',
      conclusion: run?.conclusion || null,
      updated_at: run?.updated_at || null,
      started_at: run?.run_started_at || null,
      duration,
    });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/samantha/run-invoices', async (req, res) => {
  try {
    const { pos } = req.body || {};
    const inputs  = {};
    if (Array.isArray(pos) && pos.length) inputs.po_numbers = pos.join(',');

    const r = await ghFetch(`https://api.github.com/repos/${REPO}/actions/workflows/urbn-invoice-generator.yml/dispatches`, {
      method: 'POST', body: JSON.stringify({ ref: 'main', inputs }),
    });
    if (r.status !== 204) { const b = await r.text(); return res.status(500).json({ error: `GitHub returned ${r.status}: ${b}` }); }
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ─── Samantha: Invoice PDF Extract ────────────────────────────────────────────

// Preview: SHIPPED POs with invoice date but no PDF link
app.get('/api/samantha/pdf-preview', async (req, res) => {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    return res.status(500).json({ error: 'Google credentials not configured' });
  }
  try {
    const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'] });
    const sheets = google.sheets({ version: 'v4', auth });

    const r    = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Warehouse Now Database' });
    const rows = r.data.values || [];
    if (rows.length < 2) return res.json([]);

    const headers  = rows[0].map(h => (h || '').trim().toLowerCase());
    const idx      = keys => headers.findIndex(h => keys.some(k => h.includes(k)));
    const poCol    = idx(['po#', 'po number']);
    const statCol  = idx(['status']);
    const dateCol  = idx(['urbn invoice date']);
    const linkCol  = idx(['invoice link']);

    if (poCol < 0 || statCol < 0) return res.status(500).json({ error: 'Required columns not found' });

    const result = [];
    for (let i = 1; i < rows.length; i++) {
      const row    = rows[i];
      const po     = (row[poCol]   || '').trim();
      const status = (row[statCol] || '').trim().toUpperCase();
      const date   = dateCol >= 0 ? (row[dateCol] || '').trim() : '';
      const link   = linkCol >= 0 ? (row[linkCol] || '').trim() : '';
      if (!po || status !== 'SHIPPED' || !date || link) continue;
      result.push({ poNumber: po });
    }
    res.json(result);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/samantha/pdf-status', async (req, res) => {
  try {
    const r    = await ghFetch(`https://api.github.com/repos/${REPO}/actions/workflows/invoice-pdf.yml/runs?per_page=1`);
    const data = await r.json();
    const run  = data.workflow_runs?.[0] || null;
    const duration = run && run.status === 'completed' && run.run_started_at && run.updated_at
      ? Math.round((new Date(run.updated_at) - new Date(run.run_started_at)) / 1000) : null;
    res.json({
      status:     run?.status     || 'unknown',
      conclusion: run?.conclusion || null,
      updated_at: run?.updated_at || null,
      started_at: run?.run_started_at || null,
      duration,
    });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/samantha/run-pdf-extract', async (req, res) => {
  try {
    const { pos } = req.body || {};
    const inputs  = {};
    if (Array.isArray(pos) && pos.length) inputs.po_numbers = pos.join(',');
    const r = await ghFetch(`https://api.github.com/repos/${REPO}/actions/workflows/invoice-pdf.yml/dispatches`, {
      method: 'POST', body: JSON.stringify({ ref: 'main', inputs }),
    });
    if (r.status !== 204) { const b = await r.text(); return res.status(500).json({ error: `GitHub returned ${r.status}: ${b}` }); }
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/samantha/run-invoice-for-po', async (req, res) => {
  try {
    const { poNumber } = req.body || {};
    if (!poNumber) return res.status(400).json({ error: 'poNumber required' });
    const inputs = { po_numbers: String(poNumber) };
    const r1 = await ghFetch(`https://api.github.com/repos/${REPO}/actions/workflows/urbn-invoice-generator.yml/dispatches`, {
      method: 'POST', body: JSON.stringify({ ref: 'main', inputs }),
    });
    if (r1.status !== 204) { const b = await r1.text(); return res.status(500).json({ error: `invoice workflow: ${b}` }); }
    const r2 = await ghFetch(`https://api.github.com/repos/${REPO}/actions/workflows/invoice-pdf.yml/dispatches`, {
      method: 'POST', body: JSON.stringify({ ref: 'main', inputs }),
    });
    if (r2.status !== 204) { const b = await r2.text(); return res.status(500).json({ error: `pdf workflow: ${b}` }); }
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ─── Jhonny: trigger FedEx label creation (API) ──────────────────────────────
app.post('/api/jhonny/run-fedex-labels', async (req, res) => {
  try {
    const { pos } = req.body || {};
    if (!Array.isArray(pos) || !pos.length) return res.status(400).json({ error: 'pos array required' });
    const r = await ghFetch(`https://api.github.com/repos/frazeved/JHONNY/actions/workflows/create-labels.yml/dispatches`, {
      method: 'POST',
      body: JSON.stringify({ ref: 'main', inputs: { po_numbers: pos.join(',') } }),
    });
    if (r.status !== 204) { const b = await r.text(); return res.status(500).json({ error: `GitHub: ${b}` }); }
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ─── Jhonny: trigger FedEx label creation (Web) ──────────────────────────────
app.post('/api/jhonny/run-fedex-labels-web', async (req, res) => {
  try {
    const { pos } = req.body || {};
    if (!Array.isArray(pos) || !pos.length) return res.status(400).json({ error: 'pos array required' });
    const r = await ghFetch(`https://api.github.com/repos/frazeved/JHONNY/actions/workflows/create-labels-web.yml/dispatches`, {
      method: 'POST',
      body: JSON.stringify({ ref: 'main', inputs: { po_numbers: pos.join(',') } }),
    });
    if (r.status !== 204) { const b = await r.text(); return res.status(500).json({ error: `GitHub: ${b}` }); }
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ─── Jhonny: PO list for generators ──────────────────────────────────────────
app.get('/api/jhonny/po-list', async (req, res) => {
  const type = req.query.type;
  const validTypes = ['invoice','packing-list','fedex','al-print','pl-print','fedex-print','send-invoices'];
  if (!validTypes.includes(type))
    return res.status(400).json({ error: 'invalid type' });

  try {
    const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'] });
    const sheets = google.sheets({ version: 'v4', auth });
    const r      = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Warehouse Now Database' });
    const rows   = r.data.values || [];
    if (rows.length < 2) return res.json([]);

    const H   = rows[0].map(h => (h || '').trim().toLowerCase());
    const idx = (...keys) => H.findIndex(h => keys.some(k => h.includes(k.toLowerCase())));

    const C = {
      style:      idx('style#', 'style #', 'style'),
      po:         idx('po#', 'po number'),
      status:     idx('status'),
      shipDate:   idx('ship date tradestone', 'ship date', 'shipped date'),
      cancel:     idx('cancel date', 'cancel by', 'cancel'),
      invDate:    idx('urbn invoice date'),
      invEmailSent: idx('invoice email sent'),
      pl:         idx('packing list', 'pack list'),
      alCheck:    idx('anthro label 🏷️', 'anthro label 🏷', 'anthro label'),
      fx:         idx('fedex label'),
      alLink:     idx('po link', 'al link', 'anthro label link'),
      plLink:     idx('pl link', 'packing list link'),
      fxLink:     idx('fedex label link', 'fedex link'),
      alPrinted:  idx('al printed'),
      plPrinted:  idx('pl printed'),
      fxPrinted:  idx('fedeex printed', 'fedex printed'),
      supplier:   idx('supplier'),
      boxQty:     idx('box qty', 'boxes', 'cartons'),
      exFactory:  idx('ex factory', 'ex-factory', 'flight date'),
      arrival:    idx('arrival date', 'eta'),
      supInvoice: idx('sup invoice', 'supplier invoice'),
      mawb:       idx('mawb', 'hawb'),
      poQty:      idx('po qty', 'po quantity'),
      awbFolder:  idx('awb folder'),
      brand:      idx('brand'),
    };
    // category/sub-category: find carefully to avoid cross-matching
    C.subCategory = H.findIndex(h => h.includes('sub') && h.includes('categor'));
    C.category    = H.findIndex(h => h.includes('categor') && !(h.includes('sub')));

    const get = (row, i) => i >= 0 ? (row[i] || '').trim() : '';

    const inTransitWarehouseDelayed = s => {
      const v = s.toLowerCase();
      return v.includes('transit') || v.includes('route') ||
             v.includes('warehouse') || v.includes('received') || v.includes('arrived') ||
             v.includes('delay') || v.includes('late') || v.includes('hold');
    };

    const result = [];
    let lastStyle = '';
    for (let i = 1; i < rows.length; i++) {
      const row   = rows[i];
      const style = get(row, C.style) || lastStyle;
      if (get(row, C.style)) lastStyle = style;
      const po     = get(row, C.po);
      const status = get(row, C.status);
      if (!po) continue;

      let match = false;
      let link  = '';
      if (type === 'invoice')       { match = status.toUpperCase() === 'SHIPPED' && !get(row, C.invDate); }
      if (type === 'send-invoices') { match = status.toUpperCase() === 'SHIPPED' && !get(row, C.invEmailSent); }
      if (type === 'packing-list')  { match = inTransitWarehouseDelayed(status) && !get(row, C.pl) && !get(row, C.alCheck); }
      if (type === 'fedex')        { match = inTransitWarehouseDelayed(status) && !get(row, C.fx); }
      const s = status.toLowerCase();
      const printable = s.includes('transit') || s.includes('warehouse');
      if (type === 'al-print')    { link = get(row, C.alLink);  match = !!link && printable && !get(row, C.alPrinted); }
      if (type === 'pl-print')    { link = get(row, C.plLink);  match = !!link && printable && !get(row, C.plPrinted); }
      if (type === 'fedex-print') { link = get(row, C.fxLink);  match = !!link && printable && !get(row, C.fxPrinted); }

      if (match) {
        const entry = { style, po, status, shipDate: get(row, C.shipDate), cancelDate: get(row, C.cancel), rowIndex: i + 1 };
        if (link) entry.link = link;
        if (type === 'packing-list') Object.assign(entry, {
          supplier:    get(row, C.supplier),
          boxQty:      get(row, C.boxQty),
          exFactory:   get(row, C.exFactory),
          arrivalDate: get(row, C.arrival),
          supInvoice:  get(row, C.supInvoice),
          mawb:        get(row, C.mawb),
          category:    get(row, C.category),
          subCategory: get(row, C.subCategory),
          poQty:       get(row, C.poQty),
          awbFolder:   get(row, C.awbFolder),
          brand:       get(row, C.brand),
        });
        result.push(entry);
      }
    }
    res.json(result);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/jhonny/mark-printed', async (req, res) => {
  const { type, pos } = req.body || {};
  const colKeys = { 'al-print': ['al printed'], 'pl-print': ['pl printed'], 'fedex-print': ['fedeex printed', 'fedex printed'] };
  if (!colKeys[type] || !Array.isArray(pos) || !pos.length)
    return res.status(400).json({ error: 'type and pos required' });

  try {
    const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
    const sheets = google.sheets({ version: 'v4', auth });
    const r      = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Warehouse Now Database' });
    const rows   = r.data.values || [];

    const H          = rows[0].map(h => (h || '').trim().toLowerCase());
    const idx        = (...keys) => H.findIndex(h => keys.some(k => h.includes(k.toLowerCase())));
    const poIdx      = idx('po#', 'po number');
    const printedIdx = idx(...colKeys[type]);
    if (printedIdx < 0) return res.status(400).json({ error: 'PRINTED column not found in sheet' });

    const colLetter = i => { let col = '', n = i; while (n >= 0) { col = String.fromCharCode(65 + (n % 26)) + col; n = Math.floor(n / 26) - 1; } return col; };
    const posSet = new Set(pos);
    const data = [];
    for (let i = 1; i < rows.length; i++) {
      const po = (rows[i][poIdx] || '').trim();
      if (posSet.has(po)) data.push({ range: `Warehouse Now Database!${colLetter(printedIdx)}${i+1}`, values: [['✅']] });
    }
    if (data.length) await sheets.spreadsheets.values.batchUpdate({ spreadsheetId: SHEET_ID, requestBody: { valueInputOption: 'RAW', data } });
    res.json({ ok: true, updated: data.length });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ─── Jhonny: trigger fill-links (scan Drive, fill link + check columns) ──────
app.post('/api/jhonny/run-fill-links', async (req, res) => {
  try {
    const r = await ghFetch(`https://api.github.com/repos/frazeved/JHONNY/actions/workflows/fill-links.yml/dispatches`, {
      method: 'POST', body: JSON.stringify({ ref: 'main' }),
    });
    if (r.status !== 204) { const b = await r.text(); return res.status(500).json({ error: `GitHub: ${b}` }); }
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ─── Jhonny: save packing list form (update sheet + upload to Drive) ─────────
const AL_PL_FOLDER_ID = '1k4k8EpLdhw4EyUvQMn35ZwiRgvqsniJq';
const FEDEX_FOLDER_ID = '1ufkdrO23m2C-MrmhR1iKN3QFSJFwuQPY';
const MONTH_NAMES = ['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE','JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER'];

async function getOrCreateMonthFolder(drive, parentId, month) {
  const r = await drive.files.list({
    q: `name='${month}' and '${parentId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`,
    fields: 'files(id)', supportsAllDrives: true, includeItemsFromAllDrives: true,
  });
  if (r.data.files.length > 0) return r.data.files[0].id;
  const f = await drive.files.create({
    requestBody: { name: month, mimeType: 'application/vnd.google-apps.folder', parents: [parentId] },
    fields: 'id', supportsAllDrives: true,
  });
  return f.data.id;
}

async function uploadToDrive(drive, folderId, fileName, base64, mimeType) {
  const { Readable } = require('stream');
  const buf = Buffer.from(base64, 'base64');
  const r = await drive.files.create({
    requestBody: { name: fileName, parents: [folderId] },
    media: { mimeType, body: Readable.from(buf) },
    fields: 'id,webViewLink', supportsAllDrives: true,
  });
  return r.data;
}

async function uploadViaAppsScript(base64, fileName, folderId, month) {
  const scriptUrl    = process.env.GDRIVE_APPS_SCRIPT_URL;
  const scriptSecret = process.env.GDRIVE_APPS_SCRIPT_SECRET || '';
  if (!scriptUrl) throw new Error('GDRIVE_APPS_SCRIPT_URL not set');

  const body = JSON.stringify({ secret: scriptSecret, pdf: base64, filename: fileName, folderId, month });
  const opts = { method: 'POST', headers: { 'Content-Type': 'application/json' }, body, redirect: 'manual' };
  let resp = await fetch(scriptUrl, opts);
  let hops = 0;
  while ([301,302,307,308].includes(resp.status) && hops++ < 5) {
    resp = await fetch(resp.headers.get('location'), opts);
  }
  const text = await resp.text();
  try {
    const result = JSON.parse(text);
    if (result.webViewLink) return result.webViewLink;
  } catch (_) {}
  throw new Error(`Apps Script upload failed for ${fileName}: ${text.slice(0, 200)}`);
}

app.post('/api/jhonny/save-pl-form', async (req, res) => {
  try {
    const { po, rowIndex, fields, files } = req.body || {};
    if (!po || !rowIndex) return res.status(400).json({ error: 'po and rowIndex required' });

    const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
    const sheets = google.sheets({ version: 'v4', auth });
    const TAB    = 'Warehouse Now Database';

    const hRes = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `${TAB}!1:1` });
    const H    = (hRes.data.values?.[0] || []).map(h => (h || '').trim());
    const idxH = (...keys) => H.findIndex(h => keys.some(k => h.toLowerCase().includes(k.toLowerCase())));
    const colL = i => { let col='',n=i; while(n>=0){col=String.fromCharCode(65+(n%26))+col;n=Math.floor(n/26)-1;} return col; };

    const fieldMap = {
      style:       idxH('style#', 'style #', 'style'),
      shipDate:    idxH('ship date tradestone', 'ship date', 'shipped date'),
      cancelDate:  idxH('cancel date', 'cancel by', 'cancel'),
      supplier:    idxH('supplier'),
      boxQty:      idxH('box qty', 'boxes', 'cartons'),
      exFactory:   idxH('ex factory', 'ex-factory', 'flight date'),
      arrivalDate: idxH('arrival date', 'eta'),
      supInvoice:  idxH('sup invoice', 'supplier invoice'),
      mawb:        idxH('mawb', 'hawb'),
      poQty:        idxH('po qty', 'po quantity'),
      awbFolder:    idxH('awb folder'),
      trackingNum:  idxH('tracking number', 'tracking'),
      shippingCost: idxH('shipping cost'),
    };
    fieldMap.subCategory = H.findIndex(h => h.toLowerCase().includes('sub') && h.toLowerCase().includes('categor'));
    fieldMap.category    = H.findIndex(h => h.toLowerCase().includes('categor') && !h.toLowerCase().includes('sub'));

    const data = [];
    for (const [key, val] of Object.entries(fields || {})) {
      const col = fieldMap[key];
      if (col >= 0 && val !== undefined) data.push({ range: `${TAB}!${colL(col)}${rowIndex}`, values: [[val]] });
    }

    // Upload files via Apps Script (handles storage quota)
    const month = MONTH_NAMES[new Date().getMonth()];

    const alLinkIdx  = idxH('al link', 'anthro label link');
    const plLinkIdx  = idxH('pl link', 'packing list link');
    const fxLinkIdx  = idxH('fedex label link', 'fedex link');
    const alCheckIdx = H.findIndex(h => h.toLowerCase().includes('anthro label') && !h.toLowerCase().includes('link') && !h.toLowerCase().includes('printed'));
    const plCheckIdx = H.findIndex(h => h.toLowerCase().includes('packing list') && !h.toLowerCase().includes('link') && !h.toLowerCase().includes('printed'));
    const fxCheckIdx = H.findIndex(h => h.toLowerCase().includes('fedex label') && !h.toLowerCase().includes('link') && !h.toLowerCase().includes('printed'));

    if (files?.al?.base64) {
      const ext  = (files.al.name.split('.').pop() || 'pdf').toLowerCase();
      const link = await uploadViaAppsScript(files.al.base64, `AL ${po}.${ext}`, AL_PL_FOLDER_ID, month);
      if (link) {
        if (alLinkIdx  >= 0) data.push({ range: `${TAB}!${colL(alLinkIdx)}${rowIndex}`,  values: [[link]] });
        if (alCheckIdx >= 0) data.push({ range: `${TAB}!${colL(alCheckIdx)}${rowIndex}`, values: [['✅']] });
      }
    }
    if (files?.pl?.base64) {
      const ext  = (files.pl.name.split('.').pop() || 'pdf').toLowerCase();
      const link = await uploadViaAppsScript(files.pl.base64, `PL ${po}.${ext}`, AL_PL_FOLDER_ID, month);
      if (link) {
        if (plLinkIdx  >= 0) data.push({ range: `${TAB}!${colL(plLinkIdx)}${rowIndex}`,  values: [[link]] });
        if (plCheckIdx >= 0) data.push({ range: `${TAB}!${colL(plCheckIdx)}${rowIndex}`, values: [['✅']] });
      }
    }
    if (files?.fedex?.base64) {
      const ext  = (files.fedex.name.split('.').pop() || 'pdf').toLowerCase();
      const link = await uploadViaAppsScript(files.fedex.base64, `FEDEX ${po}.${ext}`, FEDEX_FOLDER_ID, month);
      if (link) {
        if (fxLinkIdx  >= 0) data.push({ range: `${TAB}!${colL(fxLinkIdx)}${rowIndex}`,  values: [[link]] });
        if (fxCheckIdx >= 0) data.push({ range: `${TAB}!${colL(fxCheckIdx)}${rowIndex}`, values: [['✅']] });
      }
    }

    if (data.length) await sheets.spreadsheets.values.batchUpdate({ spreadsheetId: SHEET_ID, requestBody: { valueInputOption: 'RAW', data } });
    res.json({ ok: true, updated: data.length });
  } catch (e) {
    console.error('[save-pl-form]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── Jhonny: PO DETAIL — size breakdown for a single PO ─────────────────────
app.get('/api/jhonny/po-detail/:po', async (req, res) => {
  try {
    const po  = (req.params.po || '').trim();
    if (!po) return res.json([]);

    const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'] });
    const sheets = google.sheets({ version: 'v4', auth });

    const result = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'PO DETAIL' });
    const rows   = result.data.values || [];
    if (rows.length < 2) return res.json([]);

    const H    = rows[0].map(h => (h || '').trim());
    const idxH = (...keys) => H.findIndex(h => keys.some(k => h.toLowerCase().includes(k.toLowerCase())));
    const get  = (row, i) => i >= 0 ? (row[i] || '').trim() : '';

    const poCol       = idxH('purchase order', 'po#', 'po number', 'purch');
    const sizeCol     = idxH('size desc desc');
    const shipPackCol = idxH('ship pack', 'shippack');
    const totalQtyCol = idxH('total qty', 'total');

    if (poCol < 0) return res.json([]);

    // Group by display label and sum qty.
    // PREPACK rows use their Ship Pack value (e.g. "PPK") as the label.
    const sizeMap = new Map(); // label → total qty
    const sizeOrder = [];     // preserve first-seen order
    for (let i = 1; i < rows.length; i++) {
      if (get(rows[i], poCol) !== po) continue;
      const rawSize  = get(rows[i], sizeCol);
      const shipPack = get(rows[i], shipPackCol);
      const qty      = parseInt(get(rows[i], totalQtyCol)) || 0;
      if (!rawSize) continue;
      const label = rawSize.toLowerCase().includes('prepack') ? (shipPack || rawSize) : rawSize;
      if (!sizeMap.has(label)) { sizeMap.set(label, 0); sizeOrder.push(label); }
      sizeMap.set(label, sizeMap.get(label) + qty);
    }

    const items = sizeOrder.map(label => ({ sizeDesc: label, totalQty: String(sizeMap.get(label)) }));
    res.json(items);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ─── Jhonny: CAD image cache ──────────────────────────────────────────────────
const _cadSheetCache  = { rows: null, ts: 0 };          // sheet rows, 10-min TTL
const _cadImageCache  = new Map();                       // fileId → { imageData, mimeType }

async function getCadSheetRows() {
  if (_cadSheetCache.rows && Date.now() - _cadSheetCache.ts < 10 * 60 * 1000) return _cadSheetCache.rows;
  const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly', 'https://www.googleapis.com/auth/drive.readonly'] });
  const sheets = google.sheets({ version: 'v4', auth });
  const result = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: "'Production & PO DataBase'" });
  const rows   = result.data.values || [];
  _cadSheetCache.rows = rows;
  _cadSheetCache.ts   = Date.now();
  _cadSheetCache.auth = auth;
  return rows;
}

async function getCadImage(style) {
  const rows = await getCadSheetRows();
  if (rows.length < 2) return { found: false };

  const H    = rows[0].map(h => (h || '').trim());
  const idxH = (...keys) => H.findIndex(h => keys.some(k => h.toLowerCase().includes(k.toLowerCase())));
  const get  = (row, i) => i >= 0 ? (row[i] || '').trim() : '';

  const styleCol    = idxH('original style#', 'original style');
  const cadImageCol = idxH('cad image');
  const cadUrlCol   = idxH('cad url');
  if (styleCol < 0) return { found: false };

  let driveUrl = '';
  for (let i = 1; i < rows.length; i++) {
    if (get(rows[i], styleCol).toLowerCase() === style.toLowerCase()) {
      driveUrl = get(rows[i], cadImageCol) || get(rows[i], cadUrlCol);
      break;
    }
  }
  if (!driveUrl) return { found: false };

  const m = driveUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (!m) return { found: false };
  const fileId = m[1];

  if (_cadImageCache.has(fileId)) return { found: true, ..._cadImageCache.get(fileId) };

  const sa    = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const auth  = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/drive.readonly'] });
  const drive = google.drive({ version: 'v3', auth });
  const fileRes = await drive.files.get({ fileId, alt: 'media', supportsAllDrives: true }, { responseType: 'arraybuffer' });
  const mimeType  = fileRes.headers['content-type'] || 'image/jpeg';
  const imageData = Buffer.from(fileRes.data).toString('base64');
  _cadImageCache.set(fileId, { imageData, mimeType });

  return { found: true, imageData, mimeType };
}

// Single style
app.get('/api/jhonny/cad-image/:style', async (req, res) => {
  try {
    const style = (req.params.style || '').trim();
    if (!style) return res.json({ found: false });
    res.json(await getCadImage(style));
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Batch — POST { styles: ['A','B',...] } → { results: { style: { found, imageData, mimeType } } }
app.post('/api/jhonny/cad-images-batch', async (req, res) => {
  try {
    const styles = [...new Set((req.body?.styles || []).map(s => s.trim()).filter(Boolean))];
    if (!styles.length) return res.json({ results: {} });
    const entries = await Promise.all(styles.map(async s => [s, await getCadImage(s)]));
    res.json({ results: Object.fromEntries(entries) });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ─── Jhonny: Adjustments Email — buyer lookup helper ─────────────────────────
function findBuyersForCategory(brand, category) {
  const isNuuly = (brand || '').toLowerCase().includes('nuuly');
  const book    = isNuuly ? buyers.nuuly : buyers.anthropologie;
  const cat     = (category || '').toLowerCase();

  // Direct word-match pass
  for (const [key, contacts] of Object.entries(book)) {
    if (cat.split(/\s+/).some(w => w.length > 2 && key.toLowerCase().includes(w)))
      return { key, contacts };
  }
  // Semantic fallback
  if (cat.includes('dress') || cat.includes('romper'))
    return isNuuly ? { key:'DRESS, ROMPER & SWIMWEAR', contacts: book['DRESS, ROMPER & SWIMWEAR']||[] }
                   : { key:'DRESS & ROMPER',           contacts: book['DRESS & ROMPER']||[] };
  if (cat.includes('shirt') || cat.includes('blouse') || cat.includes('top'))
    return { key:'BLOUSES & SHIRTS', contacts: book['BLOUSES & SHIRTS']||[] };
  if (cat.includes('pant') || cat.includes('jump'))
    return isNuuly ? { key:'PANTS, JUMPSUIT, SKIRTS & SHORTS', contacts: book['PANTS, JUMPSUIT, SKIRTS & SHORTS']||[] }
                   : { key:'PANTS & JUMPSUIT', contacts: book['PANTS & JUMPSUIT']||[] };
  if (cat.includes('skirt') || cat.includes('short'))
    return isNuuly ? { key:'PANTS, JUMPSUIT, SKIRTS & SHORTS', contacts: book['PANTS, JUMPSUIT, SKIRTS & SHORTS']||[] }
                   : { key:'SKIRTS & SHORTS', contacts: book['SKIRTS & SHORTS']||[] };
  if (cat.includes('lounge'))  return { key:'LOUNGE',   contacts: book['LOUNGE']||[] };
  if (cat.includes('swim'))
    return isNuuly ? { key:'DRESS, ROMPER & SWIMWEAR', contacts: book['DRESS, ROMPER & SWIMWEAR']||[] }
                   : { key:'SWIMWEAR', contacts: book['SWIMWEAR']||[] };

  return { key:'UNKNOWN', contacts: [] };
}


// ─── Jhonny: send adjustment emails (drafts) ─────────────────────────────────
app.post('/api/jhonny/send-adjustment-emails', async (req, res) => {
  try {
    const { adjustments } = req.body || {};
    if (!Array.isArray(adjustments) || !adjustments.length) return res.status(400).json({ error: 'adjustments required' });

    // Same auth pattern as send-invoices
    const sessionEmail = req.session?.user?.email?.toLowerCase();
    const teamUser = TEAM_USERS.find(u => u.email.toLowerCase() === sessionEmail);
    if (!teamUser) return res.json({ gmail_not_connected: true });
    const token = userTokens[teamUser.id];
    if (!token?.refreshToken) return res.json({ gmail_not_connected: true });

    const oAuth2 = makeOAuth2Client();
    oAuth2.setCredentials({ refresh_token: token.refreshToken });
    const gmail = google.gmail({ version: 'v1', auth: oAuth2 });

    // Group by buyer contacts (category)
    const groups = new Map();
    for (const adj of adjustments) {
      const { contacts } = findBuyersForCategory(adj.brand, adj.category);
      const key = contacts.map(c => c.email).join(',') || 'unknown';
      if (!groups.has(key)) groups.set(key, { contacts, items: [] });
      groups.get(key).items.push(adj);
    }

    let sent = 0;
    const errors = [];

    for (const { contacts, items } of groups.values()) {
      if (!contacts.length) { const msg = `No buyers found for category "${items[0].category}" brand "${items[0].brand}"`; console.error('[adj-email]', msg); errors.push(msg); continue; }

      const styles  = [...new Set(items.map(a => a.style))];
      const poNums  = items.map(a => `PO# ${a.po}`).join(' ');
      const catName = items[0].category || 'Team';
      const buyerNames = contacts.map(c => c.name.split(' ')[0]).join(' and ');

      const poBlocks = items.map(adj => {
        const { po, origSizes: orig = {}, suppSizes: supp = {} } = adj;
        const sizes = [...new Set([...Object.keys(orig), ...Object.keys(supp)])];
        const lines = sizes
          .map(s => {
            const o = parseInt(orig[s]) || 0;
            const v = parseInt(supp[s]) || 0;
            if (o === 0 || Math.abs(v - o) / o <= 0.10) return null;
            const dir = v > o ? 'Increase' : 'Decrease';
            return `${dir} ${s} size from ${o} to ${v} units.`;
          })
          .filter(Boolean);
        return `• ${po} -\n${lines.map(l => `${l}`).join('\n')}`;
      }).join('\n\n');

      const subj = `Adjustment Request — Style ${styles.join(', ')} — ${poNums}`;
      const body =
`Hi ${buyerNames} and ${catName} Team,

I hope this message finds you well.

We kindly request an adjustment for Style ${styles.join(', ')} under ${items.map(a=>`PO# ${a.po}`).join(', ')} to align with supplier production.
We apologize for the discrepancies and appreciate your understanding.

${poBlocks}

Your prompt attention to this request is greatly appreciated.

Best regards,
${teamUser.name}
Logistics Team
305 CONSULTING AND PRODUCTION
1800 NW 15TH Avenue, Suite 110
Pompano Beach, Florida 33069`;

      const to  = contacts.map(c => c.email).join(', ');
      const cc  = 'support@creativetwotwelve.com, logistics@creativetwotwelve.com, inspection@creativetwotwelve.com';
      try {
        const rawBuf = await buildRawMime({ from: `"${teamUser.name}" <${teamUser.email}>`, to, cc, subject: subj, text: body });
        const encoded = rawBuf.toString('base64').replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');
        await gmail.users.drafts.create({ userId: 'me', requestBody: { message: { raw: encoded } } });
        sent++;
      } catch (e) { errors.push(e.message); }
    }

    res.json({ ok: true, sent, errors });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ─── Jhonny: create invoice email draft in sender's Gmail ────────────────────
app.post('/api/jhonny/send-invoices', async (req, res) => {
  try {
    const { pos } = req.body || {};
    if (!Array.isArray(pos) || !pos.length) return res.status(400).json({ error: 'pos array required' });

    // Find the logged-in user's token
    const sessionEmail = req.session?.user?.email?.toLowerCase();
    const teamUser = TEAM_USERS.find(u => u.email.toLowerCase() === sessionEmail);
    if (!teamUser) return res.status(401).json({ error: 'User not found in team list' });
    const token = userTokens[teamUser.id];
    if (!token?.refreshToken) return res.status(401).json({ error: 'gmail_not_connected', name: teamUser.name });

    // Read sheet for selected POs
    const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly', 'https://www.googleapis.com/auth/drive.readonly'] });
    const sheets = google.sheets({ version: 'v4', auth });
    const drive  = google.drive({ version: 'v3', auth });

    const sheetRes = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Warehouse Now Database' });
    const rows = sheetRes.data.values || [];
    const H    = rows[0].map(h => (h || '').trim().toLowerCase());
    const idxH = (...keys) => H.findIndex(h => keys.some(k => h.includes(k.toLowerCase())));
    const poCol  = idxH('po#', 'po number');
    const trkCol = idxH('tracking number', 'tracking');
    const lnkCol = idxH('invoice link');
    const get    = (row, i) => i >= 0 ? (row[i] || '').trim() : '';

    const posSet = new Set(pos);
    const poData = [];
    for (let i = 1; i < rows.length; i++) {
      const po = get(rows[i], poCol);
      if (posSet.has(po)) poData.push({ po, tracking: get(rows[i], trkCol), link: get(rows[i], lnkCol) });
    }
    poData.sort((a, b) => pos.indexOf(a.po) - pos.indexOf(b.po));

    // Download PDFs from Drive
    const attachments = [];
    const missingPdfs = [];
    for (const { po, link } of poData) {
      if (!link) { missingPdfs.push(po); continue; }
      const m = link.match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (!m) { missingPdfs.push(po); continue; }
      try {
        const r = await drive.files.get({ fileId: m[1], alt: 'media', supportsAllDrives: true }, { responseType: 'arraybuffer' });
        attachments.push({ filename: `INV ${po}.pdf`, content: Buffer.from(r.data), contentType: 'application/pdf' });
      } catch (_) { missingPdfs.push(po); }
    }
    if (missingPdfs.length) return res.status(400).json({ error: 'missing_pdfs', missing: missingPdfs });

    // Build email
    const subject = `Invoices POs ${pos.map(p => `#${p}`).join(' ')}`;
    const tracking = poData.map(({ po, tracking }) => `PO ${po}: Tracking Number ${tracking || '(not found)'}`).join('\n');
    const body = `Invoices are attached to this email for your records. Below are the tracking details for the respective PO numbers:\n\n${tracking}\n\nBest regards,\n${teamUser.name}\nLogistics Team\n305 CONSULTING AND PRODUCTION\n1800 NW 15TH Avenue, Suite 110\nPompano Beach, Florida 33069`;

    const TO = 'invoices@urbanout.com';
    const CC = ['support@creativetwotwelve.com','logistics@creativetwotwelve.com','inspection@creativetwotwelve.com','paula@creativetwotwelve.com','rafaela.neves@farmrio.com','ozan.guruscu@creativetwotwelve.com'].join(', ');

    const rawMime = await buildRawMime({ from: `"${teamUser.name}" <${teamUser.email}>`, to: TO, cc: CC, subject, text: body, attachments });
    const encoded = rawMime.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

    const authClient = makeOAuth2Client();
    authClient.setCredentials({ refresh_token: token.refreshToken });
    const gmail = google.gmail({ version: 'v1', auth: authClient });
    await gmail.users.messages.send({ userId: 'me', requestBody: { raw: encoded } });

    // Mark INVOICE EMAIL SENT in the sheet
    try {
      const saWrite  = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
      const authW    = new google.auth.GoogleAuth({ credentials: saWrite, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
      const sheetsW  = google.sheets({ version: 'v4', auth: authW });
      const hRes     = await sheetsW.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Warehouse Now Database!1:1' });
      const headers  = (hRes.data.values?.[0] || []).map(h => (h || '').trim().toLowerCase());
      const sentIdx  = headers.findIndex(h => h.includes('invoice email sent'));
      const poIdx2   = headers.findIndex(h => h.includes('po#') || h.includes('po number'));
      if (sentIdx >= 0 && poIdx2 >= 0) {
        const colLetter = i => { let col='',n=i; while(n>=0){col=String.fromCharCode(65+(n%26))+col;n=Math.floor(n/26)-1;} return col; };
        const poRes  = await sheetsW.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `Warehouse Now Database!${colLetter(poIdx2)}:${colLetter(poIdx2)}` });
        const poCells = (poRes.data.values || []).map(r => (r[0] || '').trim());
        const today  = new Date().toLocaleDateString('en-US', { month:'2-digit', day:'2-digit', year:'numeric' });
        const data   = [];
        for (const po of pos) {
          const row = poCells.findIndex((v, i) => i > 0 && v === po);
          if (row >= 0) data.push({ range: `Warehouse Now Database!${colLetter(sentIdx)}${row + 1}`, values: [[today]] });
        }
        if (data.length) await sheetsW.spreadsheets.values.batchUpdate({ spreadsheetId: SHEET_ID, requestBody: { valueInputOption: 'RAW', data } });
      }
    } catch (e) { console.warn('[send-invoices] sheet mark failed:', e.message); }

    res.json({ ok: true, message: 'Draft created in your Gmail Drafts folder' });
  } catch (e) {
    console.error('[send-invoices]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── Jhonny: mark all Ready-to-Ship POs as SHIPPED ───────────────────────────
app.post('/api/jhonny/ship-all', async (req, res) => {
  const { pos } = req.body || {};
  if (!Array.isArray(pos) || !pos.length) return res.status(400).json({ error: 'pos array required' });
  try {
    const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
    const sheets = google.sheets({ version: 'v4', auth });
    const TAB    = 'Warehouse Now Database';
    const r      = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: TAB });
    const rows   = r.data.values || [];
    if (rows.length < 2) return res.status(404).json({ error: 'No data' });

    const H         = rows[0].map(h => (h || '').trim());
    const colLetter = i => { let col = '', n = i; while (n >= 0) { col = String.fromCharCode(65 + (n % 26)) + col; n = Math.floor(n / 26) - 1; } return col; };
    const idx       = (...keys) => H.findIndex(h => keys.some(k => h.toLowerCase().includes(k.toLowerCase())));

    const statusIdx   = idx('warehouse status', 'wh status', 'status');
    const rtsIdx      = idx('ready to ship');
    const shipDateIdx = idx('shipped date', 'ship date', 'ex-factory', 'exfactory');
    const poIdx       = idx('po#', 'po number');
    if (statusIdx < 0) return res.status(400).json({ error: 'STATUS column not found' });

    const posSet = new Set(pos.map(p => p.trim()));
    const today  = new Date().toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: 'numeric' });
    const updates = [];
    for (let i = 1; i < rows.length; i++) {
      const po = (rows[i][poIdx] || '').trim();
      if (!posSet.has(po)) continue;
      const rowNum = i + 1;
      updates.push({ range: `'${TAB}'!${colLetter(statusIdx)}${rowNum}`, values: [['SHIPPED']] });
      if (rtsIdx >= 0)      updates.push({ range: `'${TAB}'!${colLetter(rtsIdx)}${rowNum}`,      values: [['']] });
      if (shipDateIdx >= 0) updates.push({ range: `'${TAB}'!${colLetter(shipDateIdx)}${rowNum}`, values: [[today]] });
    }
    if (updates.length) {
      await sheets.spreadsheets.values.batchUpdate({ spreadsheetId: SHEET_ID, requestBody: { valueInputOption: 'RAW', data: updates } });
    }
    res.json({ ok: true, updated: posSet.size });
  } catch (e) {
    console.error('[ship-all]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── Jhonny: mark PO as SHIPPED ──────────────────────────────────────────────
app.post('/api/jhonny/mark-shipped', async (req, res) => {
  const { po } = req.body || {};
  if (!po) return res.status(400).json({ error: 'po required' });
  try {
    const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
    const sheets = google.sheets({ version: 'v4', auth });
    const TAB    = 'Warehouse Now Database';
    const r      = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: TAB });
    const rows   = r.data.values || [];
    if (rows.length < 2) return res.status(404).json({ error: 'No data' });

    const H         = rows[0].map(h => (h || '').trim());
    const colLetter = i => { let col = '', n = i; while (n >= 0) { col = String.fromCharCode(65 + (n % 26)) + col; n = Math.floor(n / 26) - 1; } return col; };
    const idx       = (...keys) => H.findIndex(h => keys.some(k => h.toLowerCase().includes(k.toLowerCase())));

    const statusIdx  = idx('warehouse status', 'wh status', 'status');
    const rtsIdx     = idx('ready to ship');
    const shipDateIdx = idx('shipped date', 'ship date', 'ex-factory', 'exfactory');
    const poIdx      = idx('po#', 'po number');

    if (statusIdx < 0) return res.status(400).json({ error: 'STATUS column not found' });

    const rowIdx = rows.slice(1).findIndex(row => (row[poIdx] || '').trim() === po.trim());
    if (rowIdx < 0) return res.status(404).json({ error: 'PO not found' });

    const rowNum  = rowIdx + 2;
    const today   = new Date().toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: 'numeric' });
    const updates = [
      { range: `'${TAB}'!${colLetter(statusIdx)}${rowNum}`, values: [['SHIPPED']] },
    ];
    if (rtsIdx >= 0)      updates.push({ range: `'${TAB}'!${colLetter(rtsIdx)}${rowNum}`,      values: [['']] });
    if (shipDateIdx >= 0) updates.push({ range: `'${TAB}'!${colLetter(shipDateIdx)}${rowNum}`, values: [[today]] });

    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: { valueInputOption: 'RAW', data: updates },
    });
    res.json({ ok: true });
  } catch (e) {
    console.error('[mark-shipped]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── Jhonny: set / clear Ready to Ship flag ──────────────────────────────────
app.post('/api/jhonny/ready-to-ship', async (req, res) => {
  const { po, ready } = req.body || {};
  if (!po) return res.status(400).json({ error: 'po required' });
  try {
    const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
    const sheets = google.sheets({ version: 'v4', auth });
    const TAB    = 'Warehouse Now Database';
    const r      = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: TAB });
    const rows   = r.data.values || [];
    if (rows.length < 2) return res.status(404).json({ error: 'No data' });

    const H   = rows[0].map(h => (h || '').trim());
    const colLetter = i => { let col = '', n = i; while (n >= 0) { col = String.fromCharCode(65 + (n % 26)) + col; n = Math.floor(n / 26) - 1; } return col; };

    let rtsIdx = H.findIndex(h => h.toLowerCase().includes('ready to ship'));
    if (rtsIdx < 0) {
      rtsIdx = H.length;
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `'${TAB}'!${colLetter(rtsIdx)}1`,
        valueInputOption: 'RAW',
        requestBody: { values: [['READY TO SHIP']] },
      });
    }

    const poIdx  = H.findIndex(h => h.toLowerCase().includes('po#') || h.toLowerCase().includes('po number'));
    const rowIdx = rows.slice(1).findIndex(row => (row[poIdx] || '').trim() === po.trim());
    if (rowIdx < 0) return res.status(404).json({ error: 'PO not found' });

    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `'${TAB}'!${colLetter(rtsIdx)}${rowIdx + 2}`,
      valueInputOption: 'RAW',
      requestBody: { values: [[ready ? 'YES' : '']] },
    });
    res.json({ ok: true });
  } catch (e) {
    console.error('[ready-to-ship]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── Print queue (local agent polls this) ────────────────────────────────────
const printQueue = []; // in-memory; jobs are short-lived

app.post('/api/print-queue', (req, res) => {
  const { link, type, po } = req.body || {};
  if (!link || !type) return res.status(400).json({ error: 'link and type required' });
  const job = { id: Date.now() + '_' + Math.random().toString(36).slice(2), link, type, po: po || '', createdAt: Date.now() };
  printQueue.push(job);
  res.json({ ok: true, jobId: job.id });
});

app.get('/api/print-queue/pending', (req, res) => {
  // Expire jobs older than 5 minutes so queue doesn't grow forever
  const now = Date.now();
  while (printQueue.length && now - printQueue[0].createdAt > 300000) printQueue.shift();
  res.json(printQueue);
});

app.delete('/api/print-queue/:id', (req, res) => {
  const i = printQueue.findIndex(j => j.id === req.params.id);
  if (i >= 0) printQueue.splice(i, 1);
  res.json({ ok: true });
});

// ─── PDF proxy + auto-print ───────────────────────────────────────────────────
// Fetches a Google Drive PDF via service account and streams it same-origin
app.get('/api/pdf-proxy', async (req, res) => {
  const link = req.query.link || '';
  const match = link.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (!match) return res.status(400).send('Invalid Drive link');
  const fileId = match[1];
  try {
    const sa    = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth  = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/drive.readonly'] });
    const drive = google.drive({ version: 'v3', auth });
    const file  = await drive.files.get({ fileId, alt: 'media' }, { responseType: 'stream' });
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'inline');
    file.data.pipe(res);
  } catch (e) { res.status(500).send('Could not fetch PDF: ' + e.message); }
});

// Returns an HTML page that embeds the PDF and auto-triggers window.print()
app.get('/api/print-doc', (req, res) => {
  const link = req.query.link || '';
  if (!link) return res.status(400).send('link required');
  const pdfSrc = `/api/pdf-proxy?link=${encodeURIComponent(link)}`;
  res.setHeader('Content-Type', 'text/html');
  res.send(`<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<style>*{margin:0;padding:0;}html,body{width:100%;height:100vh;overflow:hidden;}
embed{width:100%;height:100%;display:block;}</style>
</head><body>
<embed id="pdf" src="${pdfSrc}" type="application/pdf" />
<script>
  // Give the PDF viewer ~1.5s to render, then trigger print dialog
  setTimeout(function(){ window.print(); }, 1500);
</script>
</body></html>`);
});

// ─── Gabriel: MAP DATA SYNC ───────────────────────────────────────────────────
app.post('/api/gabriel/map-sync', async (req, res) => {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    return res.status(500).json({ error: 'Google credentials not configured' });
  }
  try {
    const sa     = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth   = new google.auth.GoogleAuth({ credentials: sa, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
    const sheets = google.sheets({ version: 'v4', auth });

    // Read all 5 sheets in parallel — MAP is a separate spreadsheet
    // Target sheet (MAP) is read with FORMULA mode so we can detect and preserve formula cells
    const readSheet = async (sheetId, tab, formula = false) => {
      const r = await sheets.spreadsheets.values.get({
        spreadsheetId: sheetId,
        range: `'${tab}'`,
        valueRenderOption: formula ? 'FORMULA' : 'UNFORMATTED_VALUE',
        dateTimeRenderOption: 'FORMATTED_STRING',
      });
      return r.data.values || [];
    };
    const isFormula = v => typeof v === 'string' && v.startsWith('=');
    const [targetData, tsData, pdData, ptData, sourceData] = await Promise.all([
      readSheet(MAP_SHEET_ID, 'ANTHRO MAP 2026', true), // FORMULA mode — preserves formulas on write-back
      readSheet(SHEET_ID,     'TRADESTONE DATABASE'),
      readSheet(SHEET_ID,     'PO DETAIL'),
      readSheet(SHEET_ID,     'PO TRADE'),
      readSheet(SHEET_ID,     'Production & PO DataBase'),
    ]);

    const srcH  = sourceData[0] || [];
    const tgtH  = targetData[0] || [];
    const tsH   = tsData[0]    || [];
    const pdH   = pdData[0]    || [];

    const idx  = (H, name) => H.indexOf(name);
    const norm = s => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
    const smartIdx = (H, ...names) => {
      const nH = H.map(h => norm(h));
      for (const name of names) {
        const t = norm(name);
        let i = nH.indexOf(t); if (i !== -1) return i;
        i = nH.findIndex(h => h.includes(t) || t.includes(h)); if (i !== -1) return i;
      }
      return -1;
    };

    // ── PASS 1: refreshManu ─────────────────────────────────────────────────
    const statusAllowed  = ["PO'd", "Waiting PO", "PO'd + production ok"];
    const alwaysWrite    = ["COST", "EX FACTORY / FLIGHT DATE"];
    const onlyIfBlank    = ["SUPPLIER", "HTS CODE", "DUTY", "FREIGHT"];
    const allCols        = [...alwaysWrite, ...onlyIfBlank];

    const srcStyleCol  = smartIdx(srcH, 'ORIGINAL STYLE#', 'STYLE #', 'STYLE#', 'Style Number', 'Style');
    const srcStatusCol = smartIdx(srcH, 'STATUS');
    const tgtStyleCol  = smartIdx(tgtH, 'STYLE #', 'STYLE#', 'Style Number');
    const srcSubCat    = smartIdx(srcH, 'SUB-CATEGORY', 'SubCategory', 'Sub Category');
    const tgtSubCat    = smartIdx(tgtH, 'SUB-CATEGORY', 'SubCategory', 'Sub Category');

    const srcCI = {}, tgtCI = {};
    for (const c of allCols) {
      srcCI[c] = smartIdx(srcH, c);
      tgtCI[c] = smartIdx(tgtH, c);
    }

    const styleMap = {};
    for (let i = 1; i < sourceData.length; i++) {
      const row    = sourceData[i];
      const status = srcStatusCol >= 0 ? String(row[srcStatusCol] || '') : '';
      const style  = srcStyleCol  >= 0 ? String(row[srcStyleCol]  || '').trim() : '';
      if (!statusAllowed.includes(status) || !style) continue;
      const entry  = {};
      for (const c of allCols) { const v = srcCI[c] >= 0 ? row[srcCI[c]] : undefined; if (v != null && v !== '') entry[c] = v; }
      if (srcSubCat >= 0 && row[srcSubCat] != null && row[srcSubCat] !== '') entry['SUB-CATEGORY'] = row[srcSubCat];
      styleMap[style] = entry;
    }

    // ── All column indices pre-computed once ───────────────────────────────
    const poCol = smartIdx(tgtH, 'Purchase Order', 'PO#', 'PO');
    const tsPO  = smartIdx(tsH,  'Purchase Order', 'PO#', 'PO');
    const pdPO  = smartIdx(pdH,  'Purchase Order', 'PO#', 'PO');
    const tsNorm = tsH.map(h => h.toString().trim().toLowerCase());

    // IP CLASS → CATEGORY mapping
    const IP_TO_CATEGORY = {
      1501:'Blouses',1502:'SWTRS & SWTSHRTS',1504:'Pants',1505:'Skirts',
      1506:'SHORTS',1508:'Dresses',1509:'Dresses',1515:'Dresses',
      4110:'Blouses',4114:'Sweaters',4120:'Skirts',4123:'Pants',
      4124:'Jumpsuit',4125:'Shorts',4130:'Dresses',4134:'Dresses',
      4141:'LOUNGE',4142:'Swimwear',4148:'Blouses',4152:'Hats',
      4153:'Accessories',4157:'Wraps',4428:'Dresses',4915:'Waters Edge',
      4920:'Blouses',4925:'Skirts',4927:'Pants',4928:'Dresses',4942:'Swimwear',
    };
    const ipToCategory = ip => { const v = parseInt(ip); return IP_TO_CATEGORY[v] || 'OTHER'; };

    // Convert 0-based column index to spreadsheet letter (A, B, ... Z, AA, AB ...)
    const colLetter = n => {
      let s = '', m = n + 1;
      while (m > 0) { s = String.fromCharCode(64 + ((m - 1) % 26 + 1)) + s; m = Math.floor((m - 1) / 26); }
      return s;
    };

    const tgt = {
      period:        idx(tgtH, 'PERIOD'),
      invoiceDate:   idx(tgtH, 'URBN INVOICE DATE'),
      invoiceTotal:  idx(tgtH, 'URBN INVOICE TOTAL'),
      totalQty:      idx(tgtH, 'Total Qty'),
      wholesale:     idx(tgtH, 'PO WHOLESALE'),
      shipDate:      idx(tgtH, 'Ship Date'),
      cancelDate:    idx(tgtH, 'Cancel Date'),
      originCountry: idx(tgtH, 'Origin Country'),
      styleDesc:     idx(tgtH, 'Style Description'),
      vendorColor:   idx(tgtH, 'Vendor Color'),
      ipClass:       idx(tgtH, 'IP CLASS'),
      category:      idx(tgtH, 'CATEGORY'),
      customsDesc:   idx(tgtH, 'Customs Description'),
      brand:         idx(tgtH, 'BRAND'),
      deliverTo:     idx(tgtH, 'Deliver To'),
      fobPrice:      idx(tgtH, 'FOB Price'),
      channel:       idx(tgtH, 'CHANNEL'),
    };

    // ── STEP 1: new PO discovery ────────────────────────────────────────────
    const ts_style  = tsNorm.indexOf('vendor style #');
    const ts_po     = tsNorm.indexOf('po#');
    const ts_qty    = tsNorm.indexOf('total qty');
    const ts_wsp    = tsNorm.indexOf('po wholesale');
    const ts_fob    = tsNorm.indexOf('total fob');
    const ts_ship   = tsNorm.indexOf('ship date');
    const ts_cancel = tsNorm.indexOf('cancel date');
    const ts_status = tsNorm.indexOf('urbn status');

    const existingPOs = new Set(
      targetData.slice(1).map(r => String(r[poCol] || '').trim()).filter(Boolean)
    );

    const emptyRow = Array(tgtH.length).fill('');
    const newRows  = [];

    for (let i = 1; i < tsData.length; i++) {
      const r      = tsData[i];
      const status = String(r[ts_status] || '').toUpperCase();
      if (status.includes('CANCEL')) continue;

      const po = String(r[ts_po] || '').trim();
      if (!po || existingPOs.has(po)) continue;

      const cancelStr  = typeof r[ts_cancel] === 'string' ? r[ts_cancel] : '';
      const cancelYear = cancelStr ? new Date(cancelStr).getFullYear() : 0;
      if (cancelYear !== 2026) continue;

      const row = [...emptyRow];
      if (tgtStyleCol    >= 0) row[tgtStyleCol]    = r[ts_style]  || '';
      if (poCol          >= 0) row[poCol]           = po;
      if (tgt.totalQty   >= 0) row[tgt.totalQty]   = r[ts_qty]   || '';
      if (tgt.wholesale  >= 0) row[tgt.wholesale]   = r[ts_wsp]  || r[ts_fob] || '';
      if (tgt.shipDate   >= 0) row[tgt.shipDate]    = r[ts_ship]  || '';
      if (tgt.cancelDate >= 0) row[tgt.cancelDate]  = r[ts_cancel]|| '';
      newRows.push(row);
      existingPOs.add(po);
    }

    newRows.sort((a, b) =>
      String(a[tgtStyleCol]).localeCompare(String(b[tgtStyleCol])) ||
      String(a[poCol]).localeCompare(String(b[poCol]))
    );

    // ── Build mapRows (existing rows — passes applied below) ───────────────
    let pass1 = 0;
    const mapRows = targetData.slice(1).map(rawRow => [...rawRow]);

    const ts_ = {
      period:       smartIdx(tsH, 'PERIOD'),
      invoiceDate:  smartIdx(tsH, 'URBN INVOICE DATE', 'INVOICE DATE', 'Invoice Dt', 'Invoice Date'),
      invoiceTotal: smartIdx(tsH, 'URBN INVOICE TOTAL', 'INVOICE TOTAL', 'Invoice Amount', 'Invoice Total'),
      totalQty:     smartIdx(tsH, 'TOTAL QTY', 'Total Qty', 'Total Units', 'QTY'),
      wholesale:    smartIdx(tsH, 'TOTAL FOB', 'Total FOB', 'PO WHOLESALE', 'WHOLESALE', 'Unit Cost', 'Wholesale'),
      shipDate:     smartIdx(tsH, 'SHIP DATE', 'Ship Date', 'Shipment Date'),
      cancelDate:   smartIdx(tsH, 'CANCEL DATE', 'Cancel Date'),
    };

    const pd_ = {
      originCountry: smartIdx(pdH, 'Origin Country', 'ORIGIN COUNTRY', 'Country of Origin'),
      styleDesc:     smartIdx(pdH, 'Style Description', 'STYLE DESCRIPTION', 'Description'),
      vendorColor:   smartIdx(pdH, 'Vendor Color', 'VENDOR COLOR', 'Color'),
      ipClass:       smartIdx(pdH, 'IP CLASS', 'IP Class'),
      customsDesc:   smartIdx(pdH, 'Customs Description', 'CUSTOMS DESCRIPTION', 'Customs Desc'),
      brand:         smartIdx(pdH, 'Brand', 'BRAND'),
      deliverTo:     smartIdx(pdH, 'Deliver To', 'DELIVER TO', 'Destination'),
      fobPrice:      smartIdx(pdH, 'FOB Price', 'FOB PRICE', 'Unit FOB'),
    };

    const tsMap = new Map();
    if (tsPO >= 0) for (let i = 1; i < tsData.length; i++) { const po = norm(tsData[i][tsPO]); if (po) tsMap.set(po, tsData[i]); }
    const pdMap = new Map();
    if (pdPO >= 0) for (let i = 1; i < pdData.length; i++) { const po = norm(pdData[i][pdPO]); if (po) pdMap.set(po, pdData[i]); }

    // ptMap: find PO and CHANNEL columns by header, not hardcoded index
    const ptH       = ptData[0] || [];
    const ptPO      = smartIdx(ptH, 'Purchase Order', 'PO#', 'PO');
    const ptChannel = smartIdx(ptH, 'CHANNEL', 'Channel');
    const ptMap = new Map();
    if (ptPO >= 0 && ptChannel >= 0) {
      for (let i = 1; i < ptData.length; i++) {
        const po = norm(ptData[i][ptPO]);
        if (po) ptMap.set(po, ptData[i][ptChannel]);
      }
    }

    // helper: only write if src column found, value is non-empty, and target cell is NOT a formula
    const setSrc = (row, tgtIdx, srcRow, srcIdx) => {
      if (tgtIdx < 0 || srcIdx < 0) return;
      if (isFormula(row[tgtIdx])) return;
      const v = srcRow[srcIdx];
      if (v != null && v !== '') row[tgtIdx] = v;
    };

    const MONTH_NAMES = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    const monthName = v => {
      if (!v && v !== 0) return '';
      const d = new Date(v);
      return isNaN(d) ? '' : MONTH_NAMES[d.getMonth()];
    };

    // Enrich a single row with Tradestone, PO DETAIL, PO TRADE, PERIOD, CATEGORY
    const enrichRow = row => {
      const po = norm(row[poCol]);
      if (!po) return;
      const ts = tsMap.get(po);
      const pd = pdMap.get(po);
      if (ts) {
        setSrc(row, tgt.invoiceDate,  ts, ts_.invoiceDate);
        setSrc(row, tgt.invoiceTotal, ts, ts_.invoiceTotal);
        setSrc(row, tgt.totalQty,     ts, ts_.totalQty);
        setSrc(row, tgt.wholesale,    ts, ts_.wholesale);
        setSrc(row, tgt.shipDate,     ts, ts_.shipDate);
        setSrc(row, tgt.cancelDate,   ts, ts_.cancelDate);
      }
      // PERIOD rule: month name of URBN INVOICE DATE if populated, else Cancel Date
      if (tgt.period >= 0 && !isFormula(row[tgt.period])) {
        const invDateVal    = tgt.invoiceDate >= 0 ? row[tgt.invoiceDate] : '';
        const cancelDateVal = tgt.cancelDate  >= 0 ? row[tgt.cancelDate]  : '';
        const src = (invDateVal !== '' && invDateVal != null) ? invDateVal : cancelDateVal;
        const mn = monthName(src);
        if (mn) row[tgt.period] = mn;
      }
      if (pd) {
        setSrc(row, tgt.originCountry, pd, pd_.originCountry);
        setSrc(row, tgt.styleDesc,     pd, pd_.styleDesc);
        setSrc(row, tgt.vendorColor,   pd, pd_.vendorColor);
        setSrc(row, tgt.ipClass,       pd, pd_.ipClass);
        setSrc(row, tgt.customsDesc,   pd, pd_.customsDesc);
        setSrc(row, tgt.brand,         pd, pd_.brand);
        setSrc(row, tgt.deliverTo,     pd, pd_.deliverTo);
        setSrc(row, tgt.fobPrice,      pd, pd_.fobPrice);
        if (tgt.category >= 0 && pd_.ipClass >= 0 && !isFormula(row[tgt.category])) {
          const ip = pd[pd_.ipClass];
          if (ip != null && ip !== '') row[tgt.category] = ipToCategory(ip);
        }
      }
      if (ptMap.has(po) && tgt.channel >= 0) row[tgt.channel] = ptMap.get(po);
    };

    // Apply styleMap (Pass 1 / Production DB) to a row by STYLE#
    const applyStyleMap = row => {
      const style = String(row[tgtStyleCol] || '').trim();
      if (!style) return;
      const match = styleMap[style];
      if (!match) return;
      for (const c of allCols) {
        const ti = tgtCI[c];
        if (ti < 0) continue;
        if (isFormula(row[ti])) continue;
        const nv = match[c];
        if (nv === undefined) continue;
        const cur = row[ti];
        const blank = cur === '' || cur === null || cur === undefined;
        if (alwaysWrite.includes(c) || (onlyIfBlank.includes(c) && blank)) row[ti] = nv;
      }
      if (!isFormula(row[tgtSubCat])) {
        const subBlank = row[tgtSubCat] === '' || row[tgtSubCat] === null || row[tgtSubCat] === undefined;
        if (subBlank && match['SUB-CATEGORY'] !== undefined) row[tgtSubCat] = match['SUB-CATEGORY'];
      }
    };

    // Apply to existing rows (Pass 1 + Pass 2/3)
    let pass2 = 0;
    for (const row of mapRows) {
      const style = String(row[tgtStyleCol] || '').trim();
      if (style && styleMap[style]) { applyStyleMap(row); pass1++; }
      const po = norm(row[poCol]);
      if (po) { enrichRow(row); pass2++; }
    }

    // Apply ALL passes to new rows in the same run — no second sync needed
    for (const row of newRows) {
      applyStyleMap(row);
      enrichRow(row);
    }

    // Write only the specific columns we touched (preserves all formulas in other columns)
    const writeCols = [
      ...Object.values(tgtCI),                           // Pass 1 cols (COST, EX FACTORY, SUPPLIER, HTS CODE, DUTY, FREIGHT)
      tgtSubCat,                                          // SUB-CATEGORY
      tgt.period, tgt.invoiceDate, tgt.invoiceTotal,     // Tradestone
      tgt.totalQty, tgt.wholesale, tgt.shipDate, tgt.cancelDate,
      tgt.originCountry, tgt.styleDesc, tgt.vendorColor, // PO DETAIL
      tgt.ipClass, tgt.category, tgt.customsDesc,
      tgt.brand, tgt.deliverTo, tgt.fobPrice,
      tgt.channel,                                        // PO TRADE
    ].filter(ci => ci != null && ci >= 0);

    const uniqueCols = [...new Set(writeCols)];

    const batchData = uniqueCols.map(ci => ({
      range: `'ANTHRO MAP 2026'!${colLetter(ci)}2`,
      values: mapRows.map(r => [r[ci] ?? '']),
    }));

    const writes = [
      sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: MAP_SHEET_ID,
        requestBody: { valueInputOption: 'USER_ENTERED', data: batchData },
      }),
    ];

    if (newRows.length) {
      writes.push(sheets.spreadsheets.values.append({
        spreadsheetId: MAP_SHEET_ID,
        range: `'ANTHRO MAP 2026'!A1`,
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        requestBody: { values: newRows },
      }));
    }
    await Promise.all(writes);

    res.json({
      ok: true,
      newPOs: newRows.length,
      pass1Updated: pass1,
      pass2Updated: pass2,
      totalRows: mapRows.length + newRows.length,
      debug: {
        srcStyleCol, srcStatusCol, styleMapSize: Object.keys(styleMap).length,
        tsPO, pdPO, ptPO, ptChannel,
        tsMapSize: tsMap.size, pdMapSize: pdMap.size, ptMapSize: ptMap.size,
        ts_, pd_,
      }
    });

  } catch (e) {
    console.error('[map-sync]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── PO Weekly SUP Report — writes sheet + creates Gmail drafts per supplier ──
app.post('/api/susan/weekly-sup-report', async (req, res) => {
  try {
    const { sendingAs } = req.body;
    const token = userTokens[sendingAs];
    if (!token?.refreshToken) return res.status(401).json({ error: 'Gmail not connected for this user' });
    const sender = TEAM_USERS.find(u => u.id === sendingAs);

    const sheetsAuth = new google.auth.GoogleAuth({ credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON), scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
    const sheets = google.sheets({ version: 'v4', auth: sheetsAuth });

    const COL = { style:2, status:3, supplier:6, category:7, subcat:8, freight:35, cost:36, proto:17, sms:25, ship:51, tp:14 };
    const EXCLUDED = ["Canceled","On Hold","Other Supplier","PO'd + production ok","PO'd","Waiting PO","Changed supplier after tariffs","Other supplier"];
    const REPORT_HEADER = ["Style#","Status","Supplier","Category","Subcategory","Freight","Cost","Proto sent","SMS sent","Ship Date","TP sent"];
    const SUPPLIER_CONTACTS = {
      "GAIA":       { email: "gburan@fama-sourcing.com", name: "Gozde" },
      "HS FASHION": { email: ["miya.lin@hsfashion.cn","aindy.wang@hsfashion.cn"], name: "Miya" },
      "H&F":        { email: ["daisy.zhu@hfourwing.com.cn","abby.hu@hfourwing.com.cn"], name: "Daisy" },
      "JJ":         { email: "vivek@cmsassociates.net", name: "VIVEK" },
      "ECICO":      { email: ["elin@ecicogroup.com","hyacinth@ecicogroup.com"], name: "Elin" },
      "CASCADE":    { email: "shilparawal@cascadenterprises.com", name: "Shilpa" },
      "KONCEPTION": { email: "neha.shashi@konceptiondesigns.com", name: "Neha" },
      "S&S":        { email: "saintsandseers@gmail.com", name: "Ravi" },
      "PQSWIM":     { email: ["paola@pqswim.com","pldesign@pqswim.com"], name: "Dir" }
    };
    const CC_LIST = ["paula@creativetwotwelve.com","ozan.guruscu@creativetwotwelve.com","rafaela@showroom212.com","kamilla@creativetwotwelve.com"];

    const sheetRes = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Production & PO DataBase!A:BZ' });
    const data = sheetRes.data.values || [];

    const reportMap = {};
    const emailMap  = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const styleRaw = row[COL.style];
      const status   = (row[COL.status]   || '').toString().trim();
      const supplier = (row[COL.supplier] || '').toString().trim();
      if (!styleRaw || !status || !supplier) continue;
      if (EXCLUDED.includes(status)) continue;

      const style = (() => { const s = styleRaw.toString().trim(); const m = s.match(/(\d.*)/); return m ? m[1] : s; })();
      const cost  = row[COL.cost];

      if (!row[COL.proto] || !row[COL.sms] || !row[COL.ship] || !row[COL.tp]) {
        if (!reportMap[supplier]) reportMap[supplier] = [];
        reportMap[supplier].push([style, status, supplier, row[COL.category]||'', row[COL.subcat]||'', row[COL.freight]||'', cost ? `$${Number(cost).toFixed(2)}` : '', row[COL.proto]||'', row[COL.sms]||'', row[COL.ship]||'', row[COL.tp]||'']);
      }

      if (!SUPPLIER_CONTACTS[supplier]) continue;
      const comments = [];
      if (!cost) comments.push("Waiting Price");
      if (!row[COL.proto] || !row[COL.sms]) comments.push("Waiting Proto/SMS");
      let tpFmt = '';
      const tpDate = row[COL.tp] ? new Date(row[COL.tp]) : null;
      if (tpDate && !isNaN(tpDate)) {
        tpFmt = `${String(tpDate.getMonth()+1).padStart(2,'0')}/${String(tpDate.getDate()).padStart(2,'0')}/${tpDate.getFullYear()}`;
        if (Math.floor((new Date() - tpDate) / 86400000) > 15) comments.push('<span style="color:red;font-size:8pt;">Urgent</span>');
      }
      if (!comments.length) continue;
      if (!emailMap[supplier]) emailMap[supplier] = [];
      emailMap[supplier].push([style, row[COL.category]||'', comments.join(', '), tpFmt, '']);
    }

    // Write sheet report
    const now = new Date();
    const sheetName = `PO Weekly SUP - ${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}-${now.getFullYear()}`;
    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
    const existing = spreadsheet.data.sheets.find(s => s.properties.title === sheetName);
    const setupReqs = [];
    if (existing) setupReqs.push({ deleteSheet: { sheetId: existing.properties.sheetId } });
    setupReqs.push({ addSheet: { properties: { title: sheetName } } });
    const batchRes = await sheets.spreadsheets.batchUpdate({ spreadsheetId: SHEET_ID, requestBody: { requests: setupReqs } });
    const newSheetId = batchRes.data.replies.find(r => r.addSheet).addSheet.properties.sheetId;

    const values = []; const headerRows = []; let ri = 0;
    for (const sup of Object.keys(reportMap).sort()) {
      values.push(REPORT_HEADER); headerRows.push(ri++);
      for (const r of reportMap[sup]) { values.push(r); ri++; }
      values.push([]); ri++;
    }
    await sheets.spreadsheets.values.update({ spreadsheetId: SHEET_ID, range: `${sheetName}!A1`, valueInputOption: 'USER_ENTERED', requestBody: { values } });
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: SHEET_ID, requestBody: { requests: headerRows.map(hr => ({ repeatCell: { range: { sheetId: newSheetId, startRowIndex: hr, endRowIndex: hr+1, startColumnIndex: 0, endColumnIndex: REPORT_HEADER.length }, cell: { userEnteredFormat: { backgroundColor: { red:0.812, green:0.886, blue:0.953 }, textFormat: { bold:true } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } })) } });

    // Create Gmail drafts (same pattern as PO Breakdown)
    const authClient = makeOAuth2Client();
    authClient.setCredentials({ refresh_token: token.refreshToken });
    const gmail = google.gmail({ version: 'v1', auth: authClient });
    const todaySlash = `${String(now.getMonth()+1).padStart(2,'0')}/${String(now.getDate()).padStart(2,'0')}`;
    const draftsCreated = [];

    for (const sup of Object.keys(emailMap).sort()) {
      const contact = SUPPLIER_CONTACTS[sup];
      const rows    = emailMap[sup];
      if (!rows.length) continue;
      const to = Array.isArray(contact.email) ? contact.email.join(', ') : contact.email;
      let html = `Hi ${contact.name}!<br><br>I hope all is well!<br><br>Please pay special attention to the styles below. The following styles updates are urgent!<br><br>We need them by Monday:<br><br><table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;"><tr style="background-color:#cfe2f3;font-weight:bold;"><th>Style #</th><th>Category</th><th>Comments</th><th>TP sent</th><th>Updates</th></tr>`;
      for (const r of rows) html += `<tr><td>${r[0]}</td><td>${r[1]}</td><td>${r[2]}</td><td>${r[3]}</td><td>${r[4]}</td></tr>`;
      html += `</table><br>Best regards,<br>${sender?.name || ''}`;
      const rawMime = await buildRawMime({ from: `"${sender?.name || ''}" <${sender?.email || ''}>`, to, cc: CC_LIST.join(','), subject: `URGENT FUP - Development Status - ${sup} - ${todaySlash}`, html });
      const encoded = rawMime.toString('base64').replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');
      await gmail.users.drafts.create({ userId: 'me', requestBody: { message: { raw: encoded } } });
      draftsCreated.push({ supplier: sup, styles: rows.length });
    }

    res.json({ ok: true, sheetName, draftsCreated });
  } catch (e) {
    console.error('[weekly-sup-report]', e.message);
    res.status(500).json({ error: e.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`\n  305 WORKSPACE TEAM`);
  console.log(`  ─────────────────────────────`);
  console.log(`  http://localhost:${PORT}\n`);
});
