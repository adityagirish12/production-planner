# FLOORPLAN — Production Planning Console

A focused, single-page web app for production planners. Upload an Excel
workbook with your work orders, materials, BOM, and work centers — get back
a daily KPI dashboard, prioritized schedule, material-shortage flags, and
work-center load analysis. Export an enriched workbook at the end of the
session and re-upload it tomorrow to continue.

```
[FLOORPLAN]   PRODUCTION PLANNING CONSOLE · v1.0
```

- **No backend.** All parsing and computation happens in your browser.
  Your file never leaves your machine.
- **No database.** "Previous data" is the workbook you exported last time —
  upload it again to pick up where you left off.
- **No login, no telemetry, no third-party calls** other than loading the
  SheetJS Excel library and Google Fonts from their public CDNs.

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
