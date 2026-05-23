# 305 Workspace — CLAUDE.md

Full project reference for AI assistants. Read this before touching any code.

---

## Overview

Multi-agent production management platform for Farm Rio / 305 Consulting. Built with **Node.js + Express** (`server.js`) and plain HTML/JS front-ends in `public/`. Deployed on **Render** — pushing to `main` on GitHub auto-deploys.

**GitHub repo:** `frazeved/SEBASTIAN`
**Deployed URL:** check Render dashboard (frazeved account)
**Start command:** `node --max-old-space-size=460 server.js` (460MB heap — do not lower)

---

## Architecture

```
server.js          ← all Express routes for all agents
public/
  production.html  ← Susan UI (email automations + Excel downloads)
  index.html       ← main dashboard / landing
  setup.html       ← setup page
contacts/
  suppliers.js     ← supplier emails + main contacts
  team305.js       ← internal 305 team CC list
  users.js         ← TEAM_USERS (id, name, email) for sender selection
  buyers.js        ← Anthropologie + Nuuly buyer contacts
  authUsers.js     ← login credentials config
.env               ← secrets (never commit)
```

---

## Environment Variables (required)

| Variable | Purpose |
|---|---|
| `GOOGLE_SERVICE_ACCOUNT_JSON` | Google Service Account JSON (stringified) for Sheets access |
| `GOOGLE_CLIENT_ID` | OAuth client for Gmail |
| `GOOGLE_CLIENT_SECRET` | OAuth client secret |
| `SESSION_SECRET` | Express session secret |
| `PORT` | Server port (Render sets automatically) |

---

## Google Sheet

**Sheet ID:** `1y0iL7PJldbVQmPIAnJi9wvA2hvjB8_aK2bU2kxvUf5Q`

| Tab name | gid | Purpose |
|---|---|---|
| Main PO | 0 | Primary production data — all Susan routes read from here |
| PO DETAIL | 2017761959 | PO line details |
| PO TRADE | 890202899 | Trade data |
| RLM | 1284509953 | RLM data |
| GMAIL TOKENS | — | OAuth refresh tokens persisted here |
| WORKSPACE AUTH | — | Login passwords (email → password) |

Sheet is fetched as CSV: `https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={gid}`

---

## Auth

### Google Sheets (read/write)
Service Account JSON from env var. Instantiate with:
```js
const sa = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
const auth = new google.auth.GoogleAuth({ credentials: sa, scopes: [...] });
```

### Gmail OAuth (drafts)
Users authorize via popup: `window.open('/auth/google?userId=...&return=popup', 'gmail-oauth', 'width=520,height=640')`
- **Always use `return=popup`** — never a page path like `/production`
- Tokens stored in `userTokens[userId]` in memory + persisted to GMAIL TOKENS sheet tab
- Default sender: Kamilla

**Susan NEVER sends emails — always creates Gmail DRAFTS via `gmail.users.drafts.create()`**

---

## Agents

### SUSAN — Production Manager (`/api/susan/*`)
UI: `public/production.html`
Automates outbound production emails and Excel reports.

### JHONNY — FedEx / Shipping (`/api/jhonny/*`, `/api/fedex/*`)
FedEx label creation, tracking updates, packing lists, invoices.

### SAMANTHA — URBN Invoice (`/api/samantha/*`)
Tradestone automation — generates invoices for Anthropologie/Nuuly POs.

### GABRIEL — MAP Data (`/api/gabriel/*`)
MAP (Minimum Advertised Price) data sync and reporting.

### PAUL — Draft Emails (`/api/paul/*`)
General email draft helper.

---

## Critical Coding Rules

### Column Detection — ALWAYS use `findCol`, NEVER fixed indexes
```js
const H = rows[0].map(h => (h || '').trim().toLowerCase());
const findCol = (...kws) => {
  for (const kw of kws) {
    const i = H.findIndex(h => h.includes(kw.toLowerCase()));
    if (i >= 0) return i;
  }
  return -1;
};
```
Pass multiple fallback keywords in priority order. Never use `row[27]` — headers can shift.

### CAD Images — two modes depending on whether email has a file attachment

| Situation | Method | Reason |
|---|---|---|
| Email with file attachment (PI Status, Farm x Anthro, Weekly SUP) | **data URI** `data:${mimeType};base64,${imageData}` | CID breaks when MIME has attachment part |
| Email without file attachment (TOP STATUS, URGENT FUP) | **CID inline** `cid:${uniqueId}` | Smaller payload, no conflict |

`getCadImage(style)` → `{ found: bool, imageData: string (base64), mimeType: string }`
- Cache capped at 30 entries (LRU eviction)
- Always try normalized style (strip prefix) if first lookup fails

### ExcelJS — NEVER use `header` in `ws.columns`
```js
// WRONG — causes duplicate header row:
ws.columns = [{ header: 'Style#', key: 'style', width: 15 }];

// CORRECT — width only, write header row manually:
ws.columns = [{ width: 15 }, { width: 20 }];
ws.addRow(['Style#', 'Category', ...]); // manual header
```

### Excel Download Browser Pattern
```js
const blob = new Blob([data], { type: '...' });
const url = URL.createObjectURL(blob);
const a = document.createElement('a');
a.href = url; a.download = 'filename.xlsx';
document.body.appendChild(a);
a.click();
document.body.removeChild(a);
setTimeout(() => URL.revokeObjectURL(url), 200);
```

---

## Susan — Email Automations

All emails are **drafts only**. Never `gmail.users.messages.send()`.

### 1. PO BREAKDOWN — `POST /api/po/breakdown-email`
Sends PO breakdown per buyer (Anthropologie/Nuuly). CAD image is data URI (has Excel attachment). CC: `team305.breakdownCC`.

### 2. PO OFFICIAL — `POST /api/po/official-email`
Official PO confirmation emails.

### 3. TOP STATUS — `POST /api/po/top-status-email`
Status overview email. CAD via **CID** (no file attachment).

### 4. PO WEEKLY SUP REPORT — `POST /api/susan/weekly-sup-report`

**Filter:** Supplier must be in SUPPLIER_CONTACTS. Exclude statuses: Canceled, On Hold, Other Supplier, PO'd + production ok, PO'd, Waiting PO, Changed supplier after tariffs, Other supplier. Skip if no comments.

**Comment logic:**
- "Waiting Price" → cost column empty
- "Waiting Proto/SMS" → proto or sms column empty
- "Urgent" (red) → TP sent > 15 days ago

**Column keywords:**
- Proto → `"proto sent by supplier"`
- SMS → `"sms sent from supplier"`
- TP → `"tp sent to supplier"`
- Ship → `"ex factory / flight date"`
- Cost → `"cost"` (strip `$`/commas before parsing)
- Freight → `"freight"`

**Email:** One draft per supplier. Table: CAD (data URI, 70px), Style#, Category, Comments, TP Sent, Updates.
Subject: `URGENT FUP - Development Status - {SUPPLIER} - MM/DD`
CC: paula, ozan, rafaela, kamilla.

### 5. PI STATUS REPORT — `POST /api/susan/pi-status-email`

**Filter:**
- Status = `"PO'd + production ok"`
- PI Received column empty
- YEAR >= current year; if same year, NDC month >= current month

**Column keywords:** style `'style #'`, status `'status'`, supplier `'supplier'`, year `'year'`, ndc `'ndc month/year'`, piReceived `'pi received'`

**NDC month filter:**
```js
if (yr === today.getFullYear()) {
  const MONTH_MAP = { jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12 };
  const m = MONTH_MAP[ndcVal.toLowerCase().match(/[a-z]+/)?.[0].slice(0,3)];
  if (m && m < today.getMonth() + 1) continue; // skip past months
}
```

**Email:** One draft per supplier. Table: CAD (data URI 70px), Style#, Category, NDC Month. CSV template attached `${sup}_PI_Status.csv`.
Subject: `PI Needed - MM/DD`. CC: paula, kamilla, ozan, rafaela.

### 6. [Farm x Anthro] Weekly Status — `POST /api/susan/farm-anthro-weekly-email`

**Filter:**
- Status in `["PO'd", "Waiting SMS approval", "Waiting PO"]`
- Skip if `poIssued && !fit` (PO issued but no fit comment)
- Category must match a class group keyword

**Column keywords (exact sheet headers):**
- style: `'style #'`
- status: `'status'`
- class/category: `'category'` ← sheet header is CATEGORY
- fabric: `'fabric'`
- sms: `'sms images approval'`, `'sms sent to anthro'`
- fit: `'1st fit comments'`, `'fit comment'`, `'fit'`
- poInfo: `'po info anthro'`, `'po info'`
- ndc: `'final ndc'`, `'ndc'`
- poIssued: `'po issued by anthro'`, `'po issued'`

**Class groups → Anthropologie contacts:**
| Group | Keywords | Greeting | To |
|---|---|---|---|
| Blouse | BLOUSE, SHIRT | Aly, Kate and Sophia | akauffman, kcreswell, sgarino @anthropologie.com |
| Dress | DRESS, ROMPER | Lizzy | ebarrett, jstampone, nmicchia @anthropologie.com |
| Bottoms | PANT, JUMPSUIT, BOTTOM | Caroline | smurphy, criley4 @anthropologie.com |
| Lounge | LOUNGE | Simone | sbutler, mkroh, oschwann, rrum @anthropologie.com |
| Skirt | SKIRT, SHORT | Abby and Bella | apodolsky, bkese @anthropologie.com |
| Swimwear | SWIM | Ivy | iheilman, cterral, rlongenderfer, mweston @anthropologie.com |

CC: paula, kamilla, ozan, rafaela.neves@farmrio.com
Sheet link in body: `https://docs.google.com/spreadsheets/d/1Z0ekHVSDZkainyLifwJI4fpuna76bmRp/edit?gid=1580130442`
Subject: `[Farm x Anthro] Weekly Status - {ClassName} - MM/dd`
Table: CAD (data URI 60px), Style#, Fabric, NDC, Status message
NDC period: range of dates in intro e.g. `May - June26`

### 7. URGENT FUP — `POST /api/susan/urgent-fup-email`

**Filter:**
- Status in `["Waiting SMS", "Waiting price from supplier", "Waiting revised SMS"]`
- Supplier must be in `suppliers.emails`

**Column keywords:** style `'original style#', 'original style'`, status `'status'`, supplier `'supplier'`, category `'category'`, tp `'tp sent to supplier'`

**Email:** One draft per supplier. CAD via **CID** (no file attachment — same pattern as TOP STATUS).
Subject: `URGENT FUP - Development Status - {SUPPLIER} - MM/dd/yy`
CC: business@creativetwotwelve.com, paula, ozan, rafaela, kamilla.

### 8. EMAIL 8 — Pending definition

---

## Susan — Excel Reports

All use ExcelJS. Dark blue header (`#1F3864`, white bold) unless noted.

| # | Name | Endpoint | Header color | Notes |
|---|---|---|---|---|
| 1 | PO WEEKLY SUP REPORT | `GET /api/susan/weekly-sup-excel` | Dark blue + light blue rows `#DEEAF1` | Grouped by supplier, spacer between groups |
| 2 | PI STATUS REPORT | `GET /api/susan/pi-status-excel` | Dark blue + light blue rows | Cols: Style#, Category, Supplier, PI Status, NDC Month |
| 3 | Farm x Anthro Weekly | `GET /api/susan/farm-anthro-weekly-excel` | Light blue `#D9EAF7` bold + `#EEF5FB` rows | Sorted by category. Filename: `Farm_Anthro_Weekly_MM-DD-YYYY.xlsx` |

---

## Contacts

### Suppliers (`contacts/suppliers.js`)
Keys used in `suppliers.emails` and `suppliers.mainContact`:

| Key | Main Contact | Emails |
|---|---|---|
| ECICO | Elin | Elin, Hyacinth, Helen, Daphne @ecicogroup.com |
| H&F | Daisy | daisy.zhu, hayley.wu, Chloe.Yang, jean.zhang, abby.hu @hfourwing.com.cn |
| HS FASHION | Miya | miya.lin, tony.lu, sendy.sheng, sharon.xu, aindy.wang @hsfashion.cn |
| S&S | Ravi | saintsandseers@gmail.com, Info@saintsandseers.com |
| KON | Neha | neha.shashi, kaveri.das, pradeep.mishra1 @konceptiondesigns.com |
| GAIA | Gozde | gozdeb, CerenT, IremE, BesteK @gaia-sourcing.com |
| JJ | Vivek | vivek@cmsassociates.net, anjanisinghania@hotmail.com, pd@jjexpoimpo.com, taran, sanjana @cmsassociates.net |
| PQSWIM | Paola | paola, headofdesign, pldesign, anne, internationaltrade, planning @pqswim.com |
| CASCADE | Shilpa | shilparawal, nanditachauhan, simranbhateja, Dolphy @cascadenterprises.com |

### Internal Team (`contacts/users.js` + `contacts/team305.js`)

| id | Name | Email |
|---|---|---|
| kamilla | Kamilla Aguiar | kamilla@creativetwotwelve.com |
| paula | Paula Erthal | paula@creativetwotwelve.com |
| ozan | Ozan Guruscu | ozan.guruscu@creativetwotwelve.com |
| flavio | Flavio Azevedo | support@creativetwotwelve.com |
| eduardo | Eduardo Moraes | logistics@creativetwotwelve.com |
| julian | Julian Fajardo | inspection@creativetwotwelve.com |
| manuela | Manuela Carvalho | business@creativetwotwelve.com |
| igo | Igo Gardel | samples@creativetwotwelve.com |
| rafaela | Rafaela Neves | rafaela.neves@farmrio.com |

### Anthropologie Buyers (`contacts/buyers.js`)

| Category | Buyers |
|---|---|
| BLOUSES & SHIRTS | Aly Kauffman, Kate Creswell, Danielle Coccerino, Madison Morley |
| DRESS & ROMPER | Lizzy Barrett, Ellary Billings, Julie Stampone, Dami Amato, Molly MacRae, Jami Brady |
| PANTS & JUMPSUIT | Samantha Murphy, Caroline Riley, Gabriella Scotto D'Antuono |
| LOUNGE | Simone Butler, Mackenzie Kroh, Olivia Schwann, Rozina Rum |
| SKIRTS & SHORTS | Abby Podolsky, Justin Leonard, Bella Kese |
| SWIMWEAR | Ivy Turner, Vivian Nguyen, Allisyn Blazier, Riley Longenderfer |

### Nuuly Buyers (`contacts/buyers.js`)
| Category | Buyers |
|---|---|
| BLOUSES & SHIRTS | Emily Gallant, nuulytops@urbn.com |
| DRESS, ROMPER & SWIMWEAR | Julia Dame, Aaron Cooperman, Shelby Jensen, nuulyonepieces@urbn.com |
| PANTS, JUMPSUIT, SKIRTS & SHORTS | Lydia Lepping, Katelyn Buwalda, Hanna Saxon, MC Miskuly, NuulyBottoms@urbn.com |

---

## Known Issues

- `saveTokenToSheets error: Internal error encountered.` — pre-existing Google Sheets API intermittent error on every deploy. Non-fatal, unrelated to features.
- Sheet tab creation is always wrapped in non-fatal try/catch so it never blocks email drafts.
- Render marks deploy complete ~30–60s before the new instance actually serves traffic.

---

## Adding New Susan Features

1. Add Express route to `server.js`
2. Add card + modal + JS function to `public/production.html`
3. Use `findCol()` for all column lookups — never fixed indexes
4. Use data URI for CAD if email has file attachment; CID if no attachment
5. Always create Gmail DRAFT — never send
6. OAuth popup: always `return=popup`
