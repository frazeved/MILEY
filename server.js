require('dotenv').config();
const express = require('express');
const path = require('path');
const nodemailer = require('nodemailer');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const GITHUB_TOKEN = process.env.GITHUB_TOKEN;
const REPO = 'frazeved/SAMANTHA';

const SHEET_BASE = 'https://docs.google.com/spreadsheets/d/1y0iL7PJldbVQmPIAnJi9wvA2hvjB8_aK2bU2kxvUf5Q/edit';

const WORKFLOWS = [
  { id: 'po-detail.yml',   name: 'PO DETAIL',   tab: 'PO DETAIL',  sheetUrl: `${SHEET_BASE}?gid=2017761959#gid=2017761959` },
  { id: 'all-pos.yml',     name: 'ALL POs',      tab: 'PO TRADE',   sheetUrl: `${SHEET_BASE}?gid=890202899#gid=890202899`   },
  { id: 'po-headers.yml',  name: 'PO HEADERS',   tab: 'PO NEW',     sheetUrl: `${SHEET_BASE}?gid=406184613#gid=406184613`   },
  { id: 'po-invoiced.yml', name: 'PO INVOICED',  tab: 'PO INVOICE', sheetUrl: `${SHEET_BASE}?gid=1379989532#gid=1379989532` },
];

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

app.get('/api/status', async (req, res) => {
  try {
    const r = await ghFetch(`https://api.github.com/repos/${REPO}/actions/workflows/all-extractions.yml/runs?per_page=1`);
    const data = await r.json();
    const run = data.workflow_runs?.[0] || null;
    const duration =
      run && run.status === 'completed' && run.run_started_at && run.updated_at
        ? Math.round((new Date(run.updated_at) - new Date(run.run_started_at)) / 1000)
        : null;

    const results = WORKFLOWS.map(wf => ({
      id: wf.id,
      name: wf.name,
      tab: wf.tab,
      sheetUrl: wf.sheetUrl,
      status: run?.status || 'unknown',
      conclusion: run?.conclusion || null,
      started_at: run?.run_started_at || null,
      updated_at: run?.updated_at || null,
      duration,
    }));

    res.json(results);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/run-all', async (req, res) => {
  try {
    const r = await ghFetch(`https://api.github.com/repos/${REPO}/actions/workflows/all-extractions.yml/dispatches`, {
      method: 'POST',
      body: JSON.stringify({ ref: 'main' }),
    });
    if (r.status !== 204) {
      const body = await r.text();
      console.error(`GitHub dispatch failed: ${r.status} ${body}`);
      return res.status(500).json({ error: `GitHub returned ${r.status}: ${body}` });
    }
    res.json({ ok: true });
  } catch (e) {
    console.error('run-all error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

async function triggerJhonnyUpdate(req, res) {
  try {
    const r = await ghFetch(`https://api.github.com/repos/frazeved/JHONNY/actions/workflows/update-fedex-status.yml/dispatches`, {
      method: 'POST',
      body: JSON.stringify({ ref: 'main' }),
    });
    if (r.status !== 204) {
      const body = await r.text();
      return res.status(500).json({ error: `GitHub returned ${r.status}: ${body}` });
    }
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
}

app.post('/api/run-jhonny', triggerJhonnyUpdate);
app.post('/api/fedex/update', triggerJhonnyUpdate);

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

app.post('/api/po/search', async (req, res) => {
  try {
    const { style } = req.body;
    if (!style || !style.trim()) {
      return res.status(400).json({ error: 'Style # is required' });
    }

    const styleToSearch = style.toString().trim().toLowerCase();
    const csvUrl = 'https://docs.google.com/spreadsheets/d/1y0iL7PJldbVQmPIAnJi9wvA2hvjB8_aK2bU2kxvUf5Q/export?format=csv&gid=0';

    const response = await fetch(csvUrl);
    if (!response.ok) return res.status(500).json({ error: 'Failed to fetch sheet data' });

    const rows = parseCSV(await response.text());
    if (rows.length < 2) return res.status(500).json({ error: 'No data in sheet' });

    const headers = rows[0].map(h => h.trim().toLowerCase());

    const col = (...kws) => {
      for (const kw of kws) {
        const i = headers.findIndex(h => h === kw.toLowerCase());
        if (i >= 0) return i;
      }
      for (const kw of kws) {
        const i = headers.findIndex(h => h.includes(kw.toLowerCase()));
        if (i >= 0) return i;
      }
      return -1;
    };

    const C = {
      style:           col('style #', 'style#', 'style'),
      status:          col('status'),
      supplier:        col('supplier'),
      category:        col('category'),
      subcategory:     col('sub-category', 'subcategory'),
      piReceived:      col('buyer presentation', 'pi received', 'pi date'),
      topSent:         col('sms sent to anthro', 'top sent to anthro', 'sms sent'),
      topSampleStatus: col('top approval from anthro', 'top sample status', 'top approval'),
      basePo:          col('po info anthro', 'base po', 'base po import farm'),
      finalNDC:        col('final ndc', 'ndc'),
      topDeadline:     col('sms deadline', 'top deadline'),
      exFactory:       col('ex factory / flight date', 'ex factory', 'flight date'),
      cost:            col('cost'),
      freight:         col('freight'),
      duty:            col('duty'),
      hts:             col('hts code', 'hts'),
    };

    const get = (r, i) => i >= 0 ? (r[i] || '').trim() : '';

    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      const rowStyle  = get(r, C.style).toLowerCase();
      const rowStatus = get(r, C.status).toLowerCase();
      if (rowStyle !== styleToSearch) continue;
      if (!rowStatus.includes("po'd")) continue;

      const finalNDC   = get(r, C.finalNDC);
      const topDeadline = get(r, C.topDeadline);
      let topDeadlineDays = '';
      if (finalNDC && topDeadline) {
        const a = new Date(finalNDC), b = new Date(topDeadline);
        if (!isNaN(a) && !isNaN(b)) topDeadlineDays = Math.round((a - b) / 86400000).toString();
      }

      return res.json({
        style:           get(r, C.style),
        supplier:        get(r, C.supplier),
        category:        get(r, C.category),
        subcategory:     get(r, C.subcategory),
        piReceived:      get(r, C.piReceived),
        topSent:         get(r, C.topSent),
        topSampleStatus: get(r, C.topSampleStatus),
        basePo:          get(r, C.basePo),
        finalNDC,
        topDeadline,
        topDeadlineDays,
        exFactory:       get(r, C.exFactory),
        cost:            get(r, C.cost),
        freight:         get(r, C.freight),
        duty:            get(r, C.duty),
        hts:             get(r, C.hts),
      });
    }

    res.status(404).json({ error: 'Style not found or not in PO\'d status' });
  } catch (e) {
    console.error('PO search error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/contact', async (req, res) => {
  const { name, email, message } = req.body;
  if (!name || !message) return res.status(400).json({ error: 'Name and message are required.' });

  try {
    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS,
      },
    });

    await transporter.sendMail({
      from: `"305 Workspace" <${process.env.EMAIL_USER}>`,
      to: 'support@creativetwotwelve.com',
      replyTo: email || process.env.EMAIL_USER,
      subject: `Message from ${name} — 305 Workspace`,
      text: `From: ${name}\nEmail: ${email || 'not provided'}\n\n${message}`,
      html: `<p><strong>From:</strong> ${name}</p><p><strong>Email:</strong> ${email || 'not provided'}</p><br/><p>${message.replace(/\n/g, '<br/>')}</p>`,
    });

    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});


const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`\n  305 WORKSPACE TEAM`);
  console.log(`  ─────────────────────────────`);
  console.log(`  http://localhost:${PORT}\n`);
});
