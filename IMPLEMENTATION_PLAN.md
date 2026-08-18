# Logistics Control Tower Website — Step-by-Step Implementation Plan

> **Scope of this document:** Planning only. No code changes are made here.  
> **Power BI reference (read-only):** `Project\Air Freight\Dashboard\OceanFreightControlTower - CS.pbix`  
> **Existing Python pipeline:** `Project\ocean\v10_pkg\main.py`, `config.py`, `website.py`, `dashboard.html`

---

## 1. Goal

Build an IIS-hosted website that mirrors the Power BI **Ocean Freight Control Tower** dashboard:

| Website page | Power BI page | Status at launch |
|---|---|---|
| Landing / Menu | `Menu` | Air and Truck disabled; Ocean active |
| Ocean Analysis Dashboard | `Ocean Freight Analysis` | Full parity |
| Detailed Report | `Ocean Detailed Report` | Full parity |
| Performance Dashboard | `Detailed Analysis` (+ related performance visuals) | Full parity |
| Admin Portal | *(new — not in Power BI)* | Excel upload → validate → refresh data |

**Key requirement:** After initial IIS deployment, administrators upload new Excel files through the admin page. The dashboard refreshes **without redeploying or republishing the site**.

---

## 2. High-Level Architecture

```text
┌─────────────────────────────────────────────────────────────────┐
│                         IIS (HTTPS)                              │
│  Static files: HTML / CSS / JS / images                          │
│  Reverse proxy → Python web app (FastAPI or Flask)               │
└────────────────────────────┬────────────────────────────────────┘
                             │
         ┌───────────────────┼───────────────────┐
         ▼                   ▼                   ▼
   GET /api/data      POST /admin/upload    GET /api/status
         │                   │                   │
         └───────────┬───────┴───────────────────┘
                     ▼
         ┌───────────────────────────┐
         │  Python backend (v10_pkg) │
         │  • main.py  (ETL/clean)   │
         │  • export JSON for UI     │
         │  • atomic publish         │
         └─────────────┬─────────────┘
                       ▼
         data/current/dashboard-data.json   ← browser reads this
         data/current/Consolidated_Ocean_DSR.xlsx   ← audit/export
         data/archive/{timestamp}/...       ← uploaded sources + logs
```

**Design principle:** The browser never opens Excel. It loads one JSON payload and applies all filters/charts in JavaScript — same pattern as the current `dashboard.html`, which already implements most Power BI DAX logic client-side.

---

## 3. What Already Exists (Reuse, Do Not Rebuild)

| Asset | Location | Reuse for |
|---|---|---|
| ETL + business rules | `v10_pkg/main.py` | Data cleaning, merges, Hotlist, Open Order, GoComet, normalization |
| Paths & output names | `v10_pkg/config.py` | `INPUT_DIR`, `OUTPUT_DIR`, `Consolidated_Ocean_DSR.xlsx` |
| Record loader | `v10_pkg/website.py` → `load_records()` | Excel → JSON records (dates as ISO strings) |
| Full multi-page UI + DAX-in-JS | `v10_pkg/dashboard.html` | Landing, Analysis, Detail, Performance pages |
| Static export pattern | `v10_pkg/export_github_pages.py` | Template for JSON export + field whitelist |
| IIS admin concept | `v10_pkg/IIS_ADMIN_PORTAL_IMPLEMENTATION_PLAN.md` | Upload workflow, security, folder layout |

**Important:** The `.pbix` file is the **reference for visuals, filters, and DAX**. Do not modify it. Extract logic from it once, then implement in JavaScript/Python using the rules already captured in `dashboard.html`.

---

## 4. Power BI → Website Page Mapping

Extracted from `OceanFreightControlTower - CS.pbix` (read-only inspection):

| PBIX page name | Website route / section | Primary visuals to replicate |
|---|---|---|
| **Menu** | `/` (landing) | 3 mode tiles: AIR (disabled), OCEAN (link), TRUCK (disabled); VTC branding; hero background |
| **Ocean Freight Analysis** | `/ocean` or `#analysis` | 10 slicers, 6 KPI cards, ETA quick-filter chips, donut (ETA Performance), bar charts (Plants, Vendor), column chart (Total Delivery Variance), Reset All Filters |
| **Ocean Detailed Report** | `#detail` | Same slicers + ETA chips + full data table (all consolidated columns) |
| **Detailed Analysis** | `#performance` | 1-year sliding window, KPIs, Container Type, Arrival/Departure Performance, LCL vs FCL, Carrier/Forwarder stacked performance, Avg Port-to-Port Days |
| Page 1 (map) | *Phase 2 optional* | Azure/map visuals — skip for v1 unless explicitly required |

---

## 5. Filters & Slicers (Match Power BI Pic 2)

### Top filter bar (all dashboard pages except landing)

Replicate these slicers from the Analysis page:

| UI label | Source column (from `main.py` output) |
|---|---|
| Plant | `Delivery Location` |
| Carrier | `Steamship Line` |
| Port of Loading | `Port of Loading` |
| Port of Discharge | `Port Of Discharge` |
| Arrived/Delivered | Derived from `Vessel Arrived` / `Container Delivered` |
| Container No. | `Container Number` |
| PO # | `PO#` |
| Hotlist | `Hotlist` |
| Class | `Class` |
| Forwarder | Derived from `Source File` (see `website.py` → `_forwarder()`) |

Also include **Reset All Filters** button.

### Left / chip quick filters (ETA performance buckets)

These map to the Power BI advanced slicer buttons. Logic is already in `dashboard.html` → `etaPerformance()`, `matchesEtaFilter()`, `buildEtaButtons()`:

- Arriving but Delayed by 1 / 2 / 2+ weeks  
- Arriving but Early by 1 / 2 / 2+ weeks  
- Arriving On Time  
- Arriving Today  
- Delivered but Delayed by 1 / 2 / 2+ weeks  
- Delivered but Early by 1 / 2 / 2+ weeks  
- Delivered On Time  

---

## 6. KPI Cards (Analysis Page)

| KPI | Logic (matches Power BI / `dashboard.html`) |
|---|---|
| Container # | Distinct `Container Number` in filtered set |
| Arriving in next 7 Days # | ETA To Port of Discharge between today and today+7 |
| Arriving Next 14 days # | ETA between today and today+14 |
| Delayed Containers # | `etaPerformance()` starts with `"Delayed"` |
| Arriving Today # | ETA === today |
| Hotlist # | Distinct containers where `Hotlist = Yes` |

**Sliding window (Analysis & Detail):** Past **4 weeks** through all future records — uses `anchor()` date (ETA, delivered, ATD, or ETD). Already in `dashboard.html` → `base(28)`.

**Sliding window (Performance):** Past **1 year** — `base(365)`.

---

## 7. Charts (Analysis Page)

| Chart | Type | Field / measure |
|---|---|---|
| ETA Performance | Donut | Early / Delayed / On Time / Missing — from `etaPerformance()` |
| Total Shipment Count by Plants | Horizontal bar | `Delivery Location` |
| Total Delivery Variance | Column | Bucket counts from `etaPerformance()` labels |
| Total Vendor by Count | Horizontal bar | `Shipper Name` (top N) |

**Right-side info panel:** Show `Last Refreshed` timestamp and note: *"DSR data from arrived within 2 weeks past and all future records."*

---

## 8. Performance & Detailed Report Pages

### Performance (`#performance`)

| Visual | Logic source in `dashboard.html` |
|---|---|
| Container #, Delivered, In transit, Delayed, On-time | `state()`, `perf(r,'arr')` |
| Average port-to-port | `avgPort()` = days between `Vessel Arrived` and `Vessel Departed` |
| Container type | `CTR SIZE / LCL` |
| Arrival performance | Column + % line overlay |
| Departure performance | Compare `Vessel Departed` vs `Estimated Time of Departure` |
| LCL vs FCL | Year + mode split |
| Carrier / Forwarder performance | 100% stacked bars; toggle Forwarder vs Carrier |
| Avg Port-to-Port Days | Bar chart; toggle Carrier vs Forwarder |

### Detailed Report (`#detail`)

- Same filters and ETA chips as Analysis  
- Table with **all output columns** from consolidated Excel (scroll horizontally)  
- Client-side search; cap displayed rows (e.g. 600) for browser performance  

---

## 9. Core Business Logic (DAX → JavaScript)

Port these rules from Power BI / `dashboard.html`. **Do not guess** — validate each against the `.pbix` before sign-off.

### 9.1 ETA Performance (primary DAX equivalent)

```javascript
// From dashboard.html — matches Power BI ETA Performance measure
function etaPerformance(r) {
  let booked = r['Booked ETA Port'];
  let arrived = r['Vessel Arrived'];
  let etaPort = r['ETA To Port of Discharge'];
  let dlv = r['Dlv Date'];
  if (!etaPort) return '';
  let target = arrived ? booked : (dlv || booked);
  let compare = etaPort;
  if (!target) return '';
  let days = diff(compare, target);  // compare - target in days
  if (days === 0) return 'On Time';
  let weeks = Math.ceil(Math.abs(days) / 7);
  let bucket = weeks === 1 ? '1 week' : weeks === 2 ? '2 weeks' : '2+ weeks';
  return days < 0 ? 'Early by ' + bucket : 'Delayed by ' + bucket;
}
```

### 9.2 Container collapse (distinct container grain)

Power BI counts at container level. Use `collapse()` in `dashboard.html` to merge multiple PO rows per container before KPIs/charts.

### 9.3 Forwarder derivation

From `website.py`:

- ScanGlobal, CargoMar, Sinpex, Liyana from `Source File` name  
- GoComet / Detailed Tracking → `"Detailed Tracking"`  

### 9.4 ETL rules (Python — do not duplicate in JS)

All cleaning stays in `main.py`:

- Required inputs: **Liyana** + **GoComet/Detailed Tracking**  
- Optional: ScanGlobal, CargoMar, Sinpex, Open Order, Corporate Hot List  
- Normalization: plants, ports, shippers, carriers, dates, PO tokens, containers  
- Output: `Consolidated_Ocean_DSR.xlsx` sheet `Consolidated DSR`  

**Step during implementation:** Export DAX from Power BI Desktop (Tabular Editor or DAX query view) and diff against `dashboard.html` functions. Fix any mismatch in JS, not in `.pbix`.

---

## 10. Data Pipeline (Upload → Dashboard Refresh)

### 10.1 Normal daily flow (no admin action)

```text
Excel files in INPUT_DIR
    → python main.py
    → Consolidated_Ocean_DSR.xlsx
    → export dashboard-data.json
    → atomic copy to data/current/
    → dashboard shows new timestamp
```

### 10.2 Admin upload flow

```text
Admin uploads Excel file(s) via /admin
    → save to data/staging/{upload_id}/
    → validate: extension, size, required sheets/columns
    → copy/replace files in INPUT_DIR (or point INPUT_DIR to staging for this run)
    → run main.py programmatically (import _main or subprocess)
    → if main.py exits 0: export JSON via load_records()
    → validate row count > 0, JSON parseable
    → atomic publish: os.replace() into data/current/
    → archive staging → data/archive/{upload_id}/
    → write audit log (user, time, filename, rows, warnings)
    → on ANY failure: leave data/current unchanged, return error to admin
```

### 10.3 JSON export fields

Use the whitelist from `export_github_pages.py` → `PUBLIC_FIELDS` (extend only if a new visual needs a column).

### 10.4 Files to publish atomically

```text
data/current/
  dashboard-data.json      # UI payload: { records, modified, workbook, version }
  status.json              # { lastRefresh, rowCount, version, warnings[] }
  Consolidated_Ocean_DSR.xlsx   # optional download for auditors
```

---

## 11. Admin Page Requirements

| Feature | Description |
|---|---|
| Upload form | Multi-file upload for daily source workbooks (Liyana, GoComet, forwarders, Open Order, Hot List) **or** single-file mode — decide in Step 12 |
| Validation preview | Show required file types detected, missing files, sheet/column checks |
| Run consolidation | Trigger `main.py` server-side |
| Results panel | Row count, warnings from Summary/Validation_Report logic, link to open dashboard |
| History | List past uploads with timestamp, user, success/fail, row count |
| Rollback | Republish a previous archived `dashboard-data.json` + workbook |
| Auth | Windows/AD authentication; restrict to `OceanDashboardAdmins` group |

**No IIS redeploy** after upload — only `data/current` files change.

---

## 12. IIS Deployment Structure

### 12.1 Server folders (outside web root)

```text
D:\OceanDashboard\
  app\                      # Python app + static HTML/CSS/JS
  data\
    current\                # live JSON + workbook
    staging\                # in-progress uploads
    archive\                # timestamped versions
    input\                  # mirror of INPUT_DIR (or use INPUT_DIR directly)
  logs\
```

### 12.2 IIS setup steps

1. Install Python 3.x + `pip install -r requirements.txt` on server  
2. Deploy `v10_pkg` code to `D:\OceanDashboard\app`  
3. Set environment variables:  
   - `OCEAN_DSR_INPUT_DIR` → `D:\OceanDashboard\data\input`  
   - `OCEAN_DSR_OUTPUT_DIR` → `D:\OceanDashboard\data\staging\output` (during processing)  
4. Run Python as Windows Service (Waitress/Uvicorn) on localhost port (e.g. 8080)  
5. Create IIS site:  
   - Static content from `app/static`  
   - URL Rewrite: proxy `/api/*` and `/admin/*` to Python service  
6. Enable HTTPS + Windows Authentication for `/admin`  
7. Set max upload size in IIS Request Filtering + Python  
8. Configure app pool identity with write access to `data\` only  

---

## 13. Step-by-Step Implementation Phases

### Phase 0 — Discovery & sign-off (1–2 days)

- [ ] Open `.pbix` in Power BI Desktop **read-only**; export list of measures, calculated columns, and slicer fields  
- [ ] Compare each measure to `dashboard.html` JS functions; document gaps  
- [ ] Confirm upload model: **all daily files** vs **single consolidated upload**  
- [ ] Confirm AD groups for viewers vs admins  
- [ ] Save reference screenshots (Pic 1 Menu, Pic 2 Analysis) in `logistics_website/docs/`  

### Phase 1 — Project scaffold (1 day)

- [ ] Create `logistics_website/` app structure:

```text
logistics_website/
  static/
    index.html          # copy/adapt dashboard.html
    css/
    js/
      dashboard.js      # extract inline JS from dashboard.html
      logic.js          # etaPerformance, collapse, perf, etc.
    assets/             # hero image, logos, mode icons
  app/
    config.py           # import or symlink v10_pkg config
    etl.py              # wrapper around main.main()
    export_data.py      # JSON export from load_records()
    server.py           # FastAPI/Flask routes
    admin.py            # upload + validation + publish
  requirements.txt
```

- [ ] Pin `v10_pkg` as dependency (copy module or install as local package) — **do not fork `main.py` logic**  

### Phase 2 — Landing page (1 day)

- [ ] Build Menu page matching Pic 1:  
  - VTC logo + “LOGISTICS CONTROL TOWER”  
  - Subtitle: “Real-Time Visibility. Smarter Decisions. Better Performance.”  
  - Three tiles: **AIR** (disabled/greyed), **OCEAN** (clickable → dashboard), **TRUCK** (disabled)  
  - Background: logistics hero image (`ocean-control-tower-hero.png`)  
- [ ] Show last refresh stamp on landing footer (`/api/status`)  

### Phase 3 — Ocean Analysis dashboard (2–3 days)

- [ ] Port `dashboard.html` Analysis section to static HTML/CSS  
- [ ] Wire 10 slicers + Reset All Filters  
- [ ] Implement 6 KPI cards  
- [ ] Implement ETA chip filters (left sidebar style or horizontal chips)  
- [ ] Build 4 charts (donut, 2 bar, 1 column variance)  
- [ ] Add Last Refreshed panel  
- [ ] Verify numbers against Power BI with same source Excel on same date  

### Phase 4 — Detailed Report page (1 day)

- [ ] Same filters/chips as Analysis  
- [ ] Full-column scrollable table  
- [ ] Tab/button navigation: Analysis ↔ Detailed Report ↔ Performance  

### Phase 5 — Performance dashboard (2 days)

- [ ] Port Performance section from `dashboard.html`  
- [ ] 1-year window, all performance charts  
- [ ] Forwarder/Carrier toggle on stacked charts  

### Phase 6 — Python API layer (1–2 days)

- [ ] Replace `website.py` dev server with production app:  

| Route | Method | Purpose |
|---|---|---|
| `/api/data` | GET | Current `dashboard-data.json` |
| `/api/status` | GET | Refresh metadata |
| `/api/refresh` | POST | *(optional, admin)* manual re-run ETL |
| `/admin` | GET | Admin UI |
| `/admin/upload` | POST | Accept files, run pipeline |
| `/admin/history` | GET | Past uploads |
| `/admin/rollback/{id}` | POST | Restore archived version |

- [ ] Add `Cache-Control: no-store` and version query param on JSON  
- [ ] Implement atomic publish (same pattern as `write_workbook()` in `main.py`)  

### Phase 7 — Admin portal (2–3 days)

- [ ] Build admin HTML page (simple form + results table)  
- [ ] File validation before ETL  
- [ ] Surface `main.py` warnings (missing Hot List, date issues, etc.)  
- [ ] Serialize concurrent uploads (lock file or queue)  
- [ ] Rollback to archived version  

### Phase 8 — IIS deployment & hardening (2 days)

- [ ] Dev/staging IIS site test  
- [ ] HTTPS, Windows Auth, AD group authorization  
- [ ] Service auto-restart, log rotation  
- [ ] Backup job for `data/archive`  

### Phase 9 — UAT & go-live (3–5 days)

- [ ] Run parallel with Power BI for 1–2 weeks  
- [ ] Compare KPI totals daily  
- [ ] Train admins on upload workflow  
- [ ] Cut over; keep Power BI as fallback until stable  

---

## 14. Landing Page — AIR / TRUCK Disabled Behavior

| Tile | UI | Behavior |
|---|---|---|
| AIR | Purple theme, airplane icon | `opacity: 0.45`, `cursor: not-allowed`, tooltip “Coming soon” |
| OCEAN | Blue theme, ship icon | Navigates to Ocean Analysis dashboard |
| TRUCK | Green theme, truck icon | Same disabled treatment as AIR |

No backend routes for Air/Truck at v1.

---

## 15. Testing Checklist

### Data parity

- [ ] Container count matches Power BI with identical filters  
- [ ] ETA Performance donut percentages match within rounding  
- [ ] Each slicer correctly filters all visuals  
- [ ] Reset All Filters clears slicers and chips  
- [ ] 4-week vs 1-year windows behave correctly  

### Admin / refresh

- [ ] Valid upload refreshes JSON and timestamp  
- [ ] Missing Liyana or GoComet fails safely; old data remains live  
- [ ] Corrupt/wrong file type rejected  
- [ ] Concurrent upload handled (second waits or gets clear error)  
- [ ] Rollback restores previous dashboard  

### Security

- [ ] Anonymous user cannot POST to `/admin/upload`  
- [ ] Uploaded files not directly browsable via URL  
- [ ] HTTPS enforced  

### IIS / ops

- [ ] Server reboot → service and site recover automatically  
- [ ] Largest expected workbook completes within acceptable time  

---

## 16. Decisions Needed Before Coding

| # | Question | Options |
|---|---|---|
| 1 | What does admin upload? | A) All separate daily files (matches current `main.py`) · B) One pre-merged file |
| 2 | Authentication | A) Windows/AD internal only · B) Anonymous dashboard + secured admin |
| 3 | Who can view dashboard? | All company · Specific AD group |
| 4 | Admin AD group name | e.g. `OceanDashboardAdmins` |
| 5 | Max upload size | e.g. 50 MB per file |
| 6 | Archive retention | e.g. 90 days |
| 7 | Map page (PBIX Page 1) | Include in v1 or defer |
| 8 | Hosting path | `D:\OceanDashboard` vs existing IIS site folder |

---

## 17. Recommended Tech Stack

| Layer | Choice | Reason |
|---|---|---|
| ETL | Existing `main.py` | Already production-tested; same rules as Power BI source |
| Backend | FastAPI + Waitress | IIS-friendly, easy file upload, async optional |
| Frontend | HTML + CSS + vanilla JS | Matches current `dashboard.html`; no build step on IIS |
| Charts | CSS conic-gradient + div bars (current) or Chart.js later | Current approach already matches PBIX layout |
| Data to browser | JSON file via `/api/data` | Fast filtering; no Excel per request |

---

## 18. What You Will NOT Need to Redeploy on IIS

After initial setup, these change **without** touching IIS config or republishing the site:

- Daily Excel uploads (admin portal)  
- Consolidated workbook output  
- `dashboard-data.json` / `status.json`  
- Archived upload history  

Only **code changes** (new features, bug fixes) require redeploying the `app\` folder.

---

## 19. Quick Start (Development Machine)

For local testing before IIS:

```text
1. Put daily Excel files in INPUT_DIR (see config.py)
2. python main.py
3. python export_github_pages.py   # or future export_data.py
4. python server.py                # serves dashboard + /api/data
5. Open http://127.0.0.1:8050
```

Compare output to Power BI refresh on the same `Consolidated_Ocean_DSR.xlsx`.

---

## 20. Deliverables Summary

1. Multi-page static dashboard (Menu, Analysis, Detail, Performance)  
2. Python backend wrapping `main.py` + JSON export  
3. Admin upload portal with validation, history, rollback  
4. IIS deployment guide for IT  
5. Admin operations guide (upload steps, error meanings)  
6. UAT sign-off checklist vs Power BI  

---

## 21. Reference File Paths

| Item | Path |
|---|---|
| Power BI (read-only) | `Project\Air Freight\Dashboard\OceanFreightControlTower - CS.pbix` |
| ETL | `Project\ocean\v10_pkg\main.py` |
| Config | `Project\ocean\v10_pkg\config.py` |
| Dev website server | `Project\ocean\v10_pkg\website.py` |
| Dashboard UI prototype | `Project\ocean\v10_pkg\dashboard.html` |
| JSON export example | `Project\ocean\v10_pkg\export_github_pages.py` |
| This plan | `Project\ocean\logistics_website\IMPLEMENTATION_PLAN.md` |

---

*Document created for planning purposes. Implementation should begin only after Phase 0 sign-off and decisions in Section 16 are confirmed.*
