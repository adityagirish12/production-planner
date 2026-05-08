# FLOORPLAN — Production Planning Console

A focused, browser-based two-screen production system: a **planner dashboard**
for supervisors and a **floor console** for shop-floor operators. Built around
a single Excel workbook that the planner uploads — the floor screen reads it
to show operators their queue, captures their confirmations as they work, and
syncs that data back to the planner's dashboard in real time. The planner sees
what's *actually* happening, not just what's planned.

```
[FLOORPLAN]    PRODUCTION PLANNING CONSOLE · v2.0
└── PLANNER ──────── reads/writes the workbook, shows KPIs, schedule, backlog
└── FLOOR ────────── tablet-friendly operator data capture (start/stop/qty/scrap)
```

The two screens talk through the browser's local storage and the
`BroadcastChannel` API — opening them in two tabs (or two devices on the same
machine) gives you a working two-station demo.

- **No backend.** Both pages are static HTML/CSS/JS.
- **No database.** State is held in browser storage; export to Excel preserves
  everything (including the floor confirmations log).
- **No login, no telemetry, no third-party calls** other than loading the
  SheetJS Excel library and Google Fonts from their public CDNs.

---

## The two screens

### `index.html` — Planner / Supervisor dashboard
- **Daily KPI strip** — open orders, past-due, at-risk, shortages, low stock,
  WC overload, schedule adherence
- **Live Floor** panel — per work-center status (running / paused / idle) with
  the order numbers currently in progress, updated in real time as operators
  report on the floor screen
- **Backlog Aging** — past-due orders bucketed by 1–3 days, 4–7 days, 8+ days
  late, with a drill-down list
- **Today's Schedule** — open work orders sorted by past-due first, then
  priority, then earliest finish; reflects *actual* confirmed quantities from
  the floor, not just the static field from the upload
- **Material Shortages** — BOM-exploded against open orders
- **Capacity Load** — demand vs effective hours per WC
- **Inventory Health** — clickable cells for what-if editing
- **Export** — produces a `.xlsx` with all original sheets, computed sheets,
  KPI snapshot, **and the Confirmations log from the floor**

### `floor.html` — Operator console (tablet-optimized)
- **Pick a work center** from the dropdown (one tablet per station)
- **See the queue** — all open orders for that station, sorted by status
  (running > paused > idle) then priority then due date
- **Tap an order** to open an action drawer:
  - **▶ START** — begin work, timestamps the start
  - **❚❚ PAUSE / ▶ RESUME** — pause without losing context
  - **＋ REPORT QTY** — opens a number pad, logs units completed
  - **⚠ REPORT SCRAP** — logs scrap separately
  - **⏱ DOWNTIME** — logs minutes of downtime
  - **✓ COMPLETE ORDER** — sets the order to TECO status
- **Operator ID** field — every event is timestamped with who reported it
- **Recent Activity** — running log of the last 10 events on this WC
- All events flow into a `Confirmations` log that the planner sees instantly

The floor screen is designed for 5-second interactions: large tap targets
(76px+ buttons, WCAG 2.5.5 compliant), high contrast, and a number pad
instead of typing for any quantity entry.

---

## How the two screens connect

```
                    ┌──────────────────┐
                    │    Excel file    │
                    │  (uploaded by    │
                    │     planner)     │
                    └────────┬─────────┘
                             ▼
            ┌────────────────────────────────────┐
            │     localStorage + BroadcastChannel │
            │     (browser-local shared state)   │
            └──────┬───────────────────────┬─────┘
                   │                       │
                   ▼                       ▼
         ┌───────────────────┐   ┌───────────────────┐
         │   index.html      │   │   floor.html      │
         │   (Planner)       │   │   (Operator)      │
         │                   │   │                   │
         │  Reads: raw data  │   │  Reads: queue for │
         │  Writes: edits    │   │    selected WC    │
         │  Renders: KPIs,   │   │  Writes: confirm- │
         │    backlog, sched │   │    ations log     │
         └───────────────────┘   └───────────────────┘
                   ▲                       │
                   └───────────────────────┘
                  (operator confirmations
                  flow back to planner)
```

Open the two pages in two browser tabs (or two devices on the same machine,
on the same login) to see them sync. Real cross-device sync would need a
backend — see the *Limitations* section below.

---

## What this addresses

The original problem statement that motivated the floor screen:
- *"Information on the floor doesn't come back to the systems."*
- *"No proper tracking of what needs to be done."*
- *"Lots of backlog already."*

The two-screen design tackles each:
- **Floor → systems flow:** every operator action becomes a timestamped row
  in the Confirmations log, instantly visible to the planner.
- **What needs to be done:** the floor screen sorts the queue automatically;
  operators don't decide what's next, the system does.
- **Backlog visibility:** the Backlog Aging section makes past-due orders
  unmissable, bucketed so supervisors can triage by urgency.

---

## Required Excel structure

Same four sheets as before — `WorkOrders`, `Materials`, `BOM`, `WorkCenters`.
Column names are case- and space-insensitive. See the on-screen schema
reference at the top of the planner page.

### Optional fifth sheet on import: `Confirmations`
If your uploaded file has a `Confirmations` sheet (e.g. you exported it
yesterday and re-uploaded today), those rows are preserved.

Columns: `Timestamp`, `OrderNumber`, `WorkCenter`, `Operator`, `Action`,
`Quantity`, `Minutes`.

---

## Project structure

```
production-planner/
├── index.html      # Planner dashboard
├── floor.html      # Operator floor console
├── styles.css      # Shared planner-side styles
├── floor.css       # Floor-specific tablet-friendly styles
├── shared.js       # Shared data layer (localStorage + BroadcastChannel)
├── app.js          # Planner logic
├── floor.js        # Floor logic
├── README.md       # This file
├── LICENSE         # MIT
└── .gitignore
```

Eight code files. No build step. No package manager.

---

## Running locally

Same as before — open `index.html` directly, or run a tiny local server:

```bash
python3 -m http.server 8000
# or
npx http-server -p 8000
```

Then open <http://localhost:8000> for the planner, or
<http://localhost:8000/floor.html> for the floor screen.

---

## Demo workflow

1. Open `index.html`, click **▶ Try with sample data**.
2. Open `floor.html` in a second tab.
3. On the floor screen, pick a work center (e.g. `WC-ASSY-01`), enter an
   operator ID, then tap an order card.
4. Tap **START**, then **REPORT QTY**, enter a number (e.g. 50), confirm.
5. Switch back to the planner tab — the **Live Floor** section now shows
   `1 RUN` for that work center, the order's confirmed quantity has gone up,
   and the Today's Schedule reflects the new completion percentage.
6. Click **↓ Export workbook** — the resulting file includes a
   `Confirmations` sheet with every floor event timestamped.

---

## Limitations (be honest about what this is and isn't)

- **Single-device only by default.** localStorage and BroadcastChannel only
  share data within one browser on one machine. Real multi-device floor
  capture requires a backend (Node.js server, websockets, or Firebase). The
  current architecture is a working prototype, not a deployable MES.
- **No authentication.** Anyone with the URL can use the floor screen and
  enter any operator ID. Real MES requires badge-based login.
- **No machine integration.** Real factories pull cycle counts and downtime
  events from the machines themselves via OPC UA / MTConnect, not from
  manual operator taps. This demo is the *manual reporting* slice only.
- **Confirmations are "additive" to ConfirmedQty.** If your uploaded file
  already has `ConfirmedQty=100` and an operator reports 50 more, the planner
  shows 150. If your ERP is the source of truth, re-upload to reset.

---

## Deploying to GitHub Pages

(Same as before — settings → Pages → main / root.)

Once deployed, both pages are accessible:
- Planner: `https://<username>.github.io/production-planner/`
- Floor:   `https://<username>.github.io/production-planner/floor.html`

---

## License

MIT — see [LICENSE](LICENSE).


---

## Live demo

Open `index.html` in any modern browser (or deploy to GitHub Pages — see
below). Click **Try with sample data** on the landing screen to see every
feature without uploading anything.

---

## What it does

### 1. Intake — Excel upload with explicit schema
The landing screen lists every required column on every required sheet.
You can:
- **Download the starter template** (an empty `.xlsx` with the right headers and a README sheet).
- **Drag-and-drop** your file onto the dropzone, or click to browse.
- **Try the sample data** to see the dashboard with realistic numbers.

Column names are case- and space-insensitive (`Order Number`, `OrderNumber`,
and `order_number` all match). Missing columns are listed in a "Data Notices"
panel rather than blocking the upload.

### 2. Daily KPI strip
Seven KPIs that a production planner reviews every morning:
| KPI | Definition |
|---|---|
| **Open Orders** | Work orders not yet TECO/CLSD with open quantity remaining |
| **Past Due** | Open orders whose finish date is before today |
| **At Risk · 5d** | Open orders due within the next 5 days |
| **Shortages** | Unique component materials where total demand exceeds available stock |
| **Low Stock** | Materials at or below reorder point, or below 1.2× safety stock |
| **WC Overload** | Work centers with demand > 100% of effective weekly capacity |
| **Schedule Adherence** | % of closed orders that finished on or before their due date |

### 3. Today's schedule
Open work orders sorted by:
1. Past-due first
2. Priority (1 = highest)
3. Earliest finish date

Each row shows estimated hours (`OpenQty × HoursPerUnit`), days to finish,
and a flag pill (PAST DUE / AT RISK / ON TRACK).

### 4. Material shortages
Explodes the BOM against open work orders to find component shortages,
flagging any component where **total demand across all open orders exceeds
the on-hand-minus-reserved quantity**.

### 5. Work-center load
Effective capacity = `HoursPerDay × DaysPerWeek × Efficiency`.
Demand = sum of `OpenQty × HoursPerUnit` for all open orders at that WC.
Bars and flags fire at 85% (high) and 100% (overload).

### 6. Inventory health
Stockout / Below Reorder / Low Stock / OK status per material, with safety
stock and lead-time reference.

### 7. Export enriched workbook
Click **Export workbook** to download `floorplan_<date>.xlsx` containing:
- All four original input sheets (cleaned and canonicalized)
- A computed **Schedule** sheet
- A computed **Shortages** sheet
- A computed **CapacityLoad** sheet
- A computed **InventoryHealth** sheet
- A **KPI_Snapshot** sheet (single column of today's KPI values)

Use this exported file as next session's input to maintain continuity.

---

## Required Excel structure

The workbook must contain four sheets named `WorkOrders`, `Materials`,
`BOM`, and `WorkCenters`. Required columns are listed below; optional
columns enrich the analysis but won't break anything if absent.

### Sheet: `WorkOrders` (one row per work order)
| Column | Required | Notes |
|---|---|---|
| `OrderNumber` | yes | Unique ID |
| `Material` | yes | Finished good SKU |
| `Description` | no | Free text |
| `OrderQty` | yes | Number |
| `ConfirmedQty` | yes | Number, ≤ OrderQty |
| `WorkCenter` | yes | Must match a row in `WorkCenters` |
| `StartDate` | yes | Date |
| `FinishDate` | yes | Date |
| `Priority` | no | 1 (highest) – 5 (lowest); defaults to 3 |
| `Status` | yes | `CRTD` · `REL` · `TECO` · `CLSD` |

### Sheet: `Materials` (one row per SKU)
| Column | Required | Notes |
|---|---|---|
| `Material` | yes | SKU |
| `Description` | no | Free text |
| `OnHand` | yes | Total physical stock |
| `Reserved` | yes | Allocated to other orders |
| `SafetyStock` | yes | Number |
| `ReorderPoint` | yes | Number |
| `LeadTimeDays` | no | Replenishment lead time |
| `UoM` | no | EA, KG, etc. |

### Sheet: `BOM` (one row per parent-component link)
| Column | Required | Notes |
|---|---|---|
| `ParentMaterial` | yes | Finished good (must match `Materials`) |
| `Component` | yes | Component (must match `Materials`) |
| `QtyPerParent` | yes | Number |
| `UoM` | no | Component UoM |

### Sheet: `WorkCenters` (one row per work center)
| Column | Required | Notes |
|---|---|---|
| `WorkCenter` | yes | Code (must match `WorkOrders.WorkCenter`) |
| `Description` | no | Free text |
| `HoursPerDay` | yes | Available shift hours per day |
| `DaysPerWeek` | yes | Working days per week |
| `Efficiency` | no | Decimal 0–1; defaults to 0.90 |
| `HoursPerUnit` | no | Standard time per unit; defaults to 1.0 |

---

## Project structure

```
production-planner/
├── index.html      # Single-page HTML (intake + dashboard)
├── styles.css      # Industrial dark theme (Bebas Neue / IBM Plex / JetBrains Mono)
├── app.js          # Excel I/O, validation, KPI computation, scheduling, export
├── README.md       # This file
├── LICENSE         # MIT
└── .gitignore
```

Three files. No build step. No package manager.

---

## Running locally

Because the app loads SheetJS from a CDN, you can simply double-click
`index.html` in most browsers. If your browser blocks CDN scripts on
`file://` URLs, run a tiny local server:

```bash
# Any of these works:
python3 -m http.server 8000
# or
npx http-server -p 8000
# or, if you have it,
php -S localhost:8000
```

Then open <http://localhost:8000>.

---

## Deploying to GitHub Pages

1. Create a new repo on GitHub (e.g. `production-planner`).
2. Push these files to the `main` branch:
   ```bash
   git init
   git add .
   git commit -m "Initial commit: FLOORPLAN production planning console"
   git branch -M main
   git remote add origin https://github.com/<your-username>/production-planner.git
   git push -u origin main
   ```
3. In GitHub: **Settings → Pages → Build and deployment**
   - Source: **Deploy from a branch**
   - Branch: **`main`** / folder **`/ (root)`**
   - Click **Save**.
4. Wait ~30 seconds. Your site is live at
   `https://<your-username>.github.io/production-planner/`.

---

## Deploying anywhere else

Because this is a static site (three files), it works on:
- **Netlify** — drag the folder onto the Netlify dashboard
- **Vercel** — `vercel deploy` from the folder
- **Cloudflare Pages** — connect the GitHub repo
- **Any S3 bucket / nginx / Apache** — just serve the directory

---

## Customizing thresholds

Open `app.js` and edit the `SETTINGS` object near the top:

```js
const SETTINGS = {
  riskWindowDays: 5,         // "AT RISK" = due within this many days
  pastDueGraceDays: 0,       // ignore orders this many days past due
  lowStockMultiplier: 1.20,  // LOW STOCK = below 1.2 × SafetyStock
  capacityWarn: 0.85,        // HIGH LOAD threshold
  capacityCrit: 1.00,        // OVERLOAD threshold
  defaultHoursPerUnit: 1.0,  // fallback when WorkCenter has no HoursPerUnit
  defaultEfficiency: 0.90    // fallback when WorkCenter has no Efficiency
};
```

---

## Roadmap (open contributions welcome)

- [ ] Optional `ActualFinishDate` column for accurate schedule adherence
- [ ] Multi-day scheduling that allocates capacity across the week
- [ ] Power Query template for direct SAP OData / HANA connection
- [ ] CSV-only fallback for users without Excel
- [ ] Side-by-side workbook compare (yesterday vs today)
- [ ] Print-optimized layout for daily standup handouts

---

## Privacy

The application is a single static page. The only network requests it
makes are:
- Loading **SheetJS** from `cdn.jsdelivr.net` (the Excel parser)
- Loading **Google Fonts** stylesheets and font files

No file you upload is sent anywhere. Computation happens entirely in your
browser tab. To verify this, open your browser's Network panel before
uploading.

If you need a fully air-gapped build, vendor the SheetJS file locally
(`xlsx.full.min.js` from the SheetJS release) and replace the Google Fonts
`<link>` with self-hosted fonts.

---

## License

MIT — see [LICENSE](LICENSE).
