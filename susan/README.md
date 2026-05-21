# Susan — AI Agent (Production Manager / Factory Coordination)

Susan handles all outbound production emails to suppliers and internal team, plus Excel report generation. All automated via Gmail OAuth.

**UI lives in:** `public/production.html` (agent card + all modals)  
**Backend routes:** `server.js`  
**Agent image:** `public/SUSAN.png`  
**No dedicated `susan.html` yet — all Susan functionality currently inside `production.html`**

---

## Data Sources

| Sheet Tab | GID | Used For |
|-----------|-----|----------|
| Main PO sheet | 0 | General production data — search by style |
| PO DETAIL | 2017761959 | Line items: size / qty / SKU per style |
| PO TRADE | 890202899 | Channel mapping: PO# → channel name |
| RLM | 1284509953 | Color code lookup (col F = style, col P = color code) |
| GMAIL TOKENS | auto-created | Persists OAuth refresh tokens across Render redeploys |

**Google Sheet ID:** `1y0iL7PJldbVQmPIAnJi9wvA2hvjB8_aK2bU2kxvUf5Q`

---

## Gmail OAuth — Token Persistence

Refresh tokens are saved to the `GMAIL TOKENS` sheet tab on every new connection.  
On server startup, tokens are loaded from that tab — so tokens **survive Render redeploys**.  
Users only need to connect their Gmail account once.

If a user tries to send without being connected, the app:
1. Saves the full form state to `sessionStorage`
2. Redirects directly to Google OAuth (skips `/setup`)
3. On return (`?gmailConnected=1`), re-opens the modal with all data restored

**Default sender:** Kamilla (always, no localStorage memory)

---

## Email Automations (8 total)

| # | Name | Status | Endpoint |
|---|------|--------|----------|
| 1 | PO BREAKDOWN | ✅ Built | `POST /api/po/breakdown-email` |
| 2 | PO OFFICIAL | ✅ Built | `POST /api/po/official-email` |
| 3 | TOP STATUS | TBD | — |
| 4 | PO WEEKLY SUP REPORT | TBD | — |
| 5 | PI STATUS REPORT | TBD | — |
| 6 | [Farm x Anthro] Weekly Status | TBD | — |
| 7 | URGENT FUP | TBD | — |
| 8 | EMAIL 8 | Pending definition | — |

---

## Excel Reports (6 total)

| # | Name | Status |
|---|------|--------|
| 1 | PO OFFICIAL REPORT | TBD |
| 2 | PO WEEKLY SUP REPORT | TBD |
| 3 | PI STATUS REPORT | TBD |
| 4 | BREAKDOWN REPORT | TBD |
| 5 | TOP STATUS REPORT | TBD |
| 6 | [Farm x Anthro] Weekly Status | TBD |

---

## PO Breakdown Email ✅

**Frontend:** `public/production.html` — card opens `#po-breakdown-modal`  
**Backend:** `server.js` — `POST /api/po/breakdown-email`  
**CC list:** `contacts/team305.js` → `breakdownCC`

### Form Inputs
- Style # + Search (auto-fills all fields from main PO sheet)
- Supplier, Category, Sub-Category, PI Received, Top Sent, Top Sample Status
- Base PO, Final NDC, Top Deadline, Ex-Factory date, Cost, Freight, Duty, HTS
- Optional custom message
- Sending As (defaults to Kamilla)

### Excel Attachment
- Filename: `BREAKDOWN STYLE# {style}.xlsx`, sheet: `Breakdown`
- Columns: PO#, Style#, Type, Pack Type, Item Number, Vendor Color, Size Code, Size, Short SKU, Qty, RATIO, Total Units
- Per-PO TOTAL rows + GRAND TOTAL row
- PPK ratio = units / GCD of all PPK units

### Email Structure
- **To:** `suppliers.emails[supplierKey]` (fallback: `logistics@creativetwotwelve.com`)
- **CC:** `breakdownCC` (10 recipients)
- **Subject:** `BREAKDOWN STYLE# {style} - PO #xxx - {supplier}`
- **Body:** HTML — greeting, style#, HTS/freight/cost/dates, PO list with channels, fabric composition request, shipping docs reminder, optional message, sign-off
- **Delivery:** Gmail draft via `gmail.users.drafts.create`

---

## PO Official Email ✅

**Frontend:** `public/production.html` — card opens `#po-official-modal`  
**Backend:** `server.js` — `POST /api/po/official-email`

### Form Inputs
- Style # + Search (auto-fills all fields from main PO sheet)
- Same full field set as PO Breakdown
- Sending As (defaults to Kamilla)

### Data Sources
- PO DETAIL (gid=2017761959): aggregate qty by size across all POs
- RLM (gid=1284509953): color code lookup — alerts popup if style not found in either sheet

### Excel Attachment (formatted with exceljs)
- Filename: `PO_{style}.xlsx`, sheet: `PO Import`
- **26 columns:** COMPANY, DIVISION, USER PO# (style+A26), Season, Year, PO EXF Date, PO Vendor, Warehouse, Style#, Fabric Code (SKU200), Length Code (SKU300), Color Code (RLM), Size, Cost (FOB/DDP), Quantity, PO PIW/ETA Date, PO Notes, PO Product Notes, PO Cancel Date, Ship Mode, Selling Prd, Selling Prd Year, PO Type, PO Terms (FOB ou DDP), Special Instructions, Comments/Special 01
- One row per size; sizes sorted standard order (XXS → XXL); ETA = exfDate + 7 days
- **Formatting:** blue header (#2E75B6) + white bold text, alternating light blue (#DEEAF1)/white rows, thin borders, frozen header, Cost `$#,##0.00`, Qty `0`

### Email Structure
- **To:** inbound@farmrio.com, danielle.gouvea@farmrio.com, anacarolina.azevedo@farmrio.com
- **CC:** paula, rafaela@showroom212, ozan, business, kamilla
- **Subject:** `[ANTHRO X FARM] official PO request - style # {digits-only}`
- **Body:** plain text — Hi Ana, style#, supplier, Anthro NDC, RLM confirmation, sender name
- **Delivery:** Gmail draft via `gmail.users.drafts.create`

---

## CC Contacts (`contacts/team305.js → breakdownCC`)
```
ozan.guruscu@creativetwotwelve.com
rafaela.neves@farmrio.com
inbound@farmrio.com
kamilla@creativetwotwelve.com
danielle.gouvea@farmrio.com
paula@creativetwotwelve.com
inspection@creativetwotwelve.com
logistics@creativetwotwelve.com
support@creativetwotwelve.com
business@creativetwotwelve.com
```

---

## Top-Level Actions (Coming Soon)
- **Run Susan** — trigger full automation pipeline
- **Database** — link to backing sheet
