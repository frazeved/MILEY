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
| Main PO sheet | 0 | General production data |
| PO DETAIL | 2017761959 | Line items: size / qty / SKU per style |
| PO TRADE | 890202899 | Channel mapping: PO# → channel name |

**Google Sheet ID:** `1y0iL7PJldbVQmPIAnJi9wvA2hvjB8_aK2bU2kxvUf5Q`

---

## Email Automations (8 total)

| # | Name | Status | Endpoint |
|---|------|--------|----------|
| 1 | PO BREAKDOWN | Built | `POST /api/po/breakdown-email` |
| 2 | PO OFFICIAL | TBD | — |
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

## PO Breakdown Email

**Frontend:** `public/production.html:317` — card opens `#po-breakdown-modal`  
**Backend:** `server.js:416` — `POST /api/po/breakdown-email`  
**CC list:** `contacts/team305.js` → `breakdownCC`

### Form Inputs
- Style # (searches PO DETAIL sheet)
- Supplier (auto-filled from search)
- Cost, HTS#, Freight type
- Ex-Factory date (handover date + invoice deadline = exFactory − 2 days)
- Optional custom message
- Sending As (Gmail OAuth user selector)

### Excel Attachment
- Filename: `BREAKDOWN STYLE# {style}.xlsx`
- Sheet: `Breakdown`
- Columns: PO#, Style#, Type, Pack Type, Item Number, Vendor Color, Size Code, Size, Short SKU, Qty, RATIO, Total Units
- Per-PO TOTAL rows + GRAND TOTAL row at bottom
- PPK ratio = units / GCD of all PPK units

### Email Structure
- **To:** `suppliers.emails[supplierKey]` (fallback: `logistics@creativetwotwelve.com`)
- **CC:** `breakdownCC` — 10 internal recipients
- **Subject:** `BREAKDOWN STYLE# {style} - PO #xxx - {supplier}`
- **Body:** greeting, style#, HTS / freight / cost / dates, PO list with channels, fabric composition request, shipping docs reminder, optional custom message, sign-off
- **Delivery:** creates a Gmail **draft** — does NOT send directly

### Gmail Draft Flow
- Each sender needs OAuth token in `userTokens[sendingAs]`
- OAuth connect flow at `/setup`
- Draft created via `gmail.users.drafts.create`

---

## Top-Level Actions (Coming Soon)
- **Run Susan** — trigger full automation pipeline
- **Database** — link to backing sheet

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
