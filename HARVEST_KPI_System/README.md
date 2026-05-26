# HARVEST Food Hub — KPI Tracking System

A complete KPI tracking system for HARVEST Food Hub. Each staff member
keeps their own Excel log workbook. A master workbook reads from all
five logs via Power Query and renders dashboards for the JSC, funders,
and the Director's annual review. Every metric is tagged to the OFSA
Six Dimensions of Food Security.

## How it works

```
  Staff Log Workbooks (one per role, in SharePoint/OneDrive)
  ├── Kitchen_Operations_Log.xlsx       (Logistics Manager, weekly)
  ├── Culinary_Training_Log.xlsx        (Culinary Manager, per event)
  ├── Community_Engagement_Log.xlsx     (Community Outreach, per event)
  ├── Director_Partnership_Log.xlsx     (Director, monthly + events)
  └── UAC_Farmer_Engagement_Log.xlsx    (Director for UAC, monthly)
              │
              ▼  Power Query
  HARVEST_KPI_Master.xlsx  ◄── Food Corridor CSV
              │
              ▼
  Dashboards: Executive Summary, By OFSA Dimension,
              By Audience, 2026 Performance Goals
```

No SharePoint Lists, no Forms, no complex automation. Staff open
their workbook, add rows, save. The Director refreshes the master
workbook to see updated dashboards.

UAC farmer data is entered by HARVEST staff on behalf of UAC partners
(from a monthly call/email), so no external access is needed.

## Prerequisites

- **Microsoft 365** for the HARVEST team (any tier with Excel and
  SharePoint/OneDrive).
- **Food Corridor** export access for the Logistics Manager.
- **Monthly Power Automate flow or recurring calendar event** for
  reminders (optional but recommended — see `schemas/monthly_reminders.md`).

## File map

| File | Purpose |
|---|---|
| `HARVEST_KPI_Master.xlsx` | Master dashboard workbook. The deliverable. |
| `logs/Kitchen_Operations_Log.xlsx` | Logistics Manager's log |
| `logs/Culinary_Training_Log.xlsx` | Culinary Manager's log |
| `logs/Community_Engagement_Log.xlsx` | Community Outreach Specialist's log (includes Survey Responses tab) |
| `logs/Director_Partnership_Log.xlsx` | Director's log |
| `logs/UAC_Farmer_Engagement_Log.xlsx` | UAC farmer data (entered by HARVEST staff) |
| `build_kpi_workbook.py` | Regenerates all workbooks (Python, openpyxl) |
| `schemas/monthly_reminders.md` | Power Automate reminder setup |
| `docs/kpi_catalog.md` | Full enriched KPI catalog — source of truth |
| `docs/ofsa_dimensions_reference.md` | OFSA Six Dimensions definitions + tagging |
| `docs/role_assignments.md` | Who owns what; cadence |
| `docs/data_dictionary.md` | Field definitions + Power Query M code |
| `samples/food_corridor_sample.csv` | Sample Food Corridor export |
| `samples/seed_data_sample.csv` | Sample rows for testing |

## Setup checklist (one-time)

Estimated total: **45 minutes**.

### 1. Save files to SharePoint/OneDrive (5 min)

Upload the entire `HARVEST_KPI_System/` folder to the HARVEST
SharePoint site. Keep the folder structure intact — the master
workbook's Power Query connections reference the `logs/` subfolder.

### 2. Wire up Power Query in the master workbook (20 min)

1. Open `HARVEST_KPI_Master.xlsx` in Excel desktop.
2. Set the `LogsFolder` parameter to the SharePoint path of the
   `logs/` folder (instructions in `docs/data_dictionary.md`).
3. Add one Power Query connection per Data tab (7 total) using the
   M code in `docs/data_dictionary.md`.
4. Click **Data > Refresh All**. Dashboards populate from the sample
   data in the log workbooks.

### 3. Distribute log workbooks (5 min)

Send each role owner the SharePoint link to their log workbook:
- Logistics Manager → `Kitchen_Operations_Log.xlsx`
- Culinary Manager → `Culinary_Training_Log.xlsx`
- Community Outreach Specialist → `Community_Engagement_Log.xlsx`
- Director keeps `Director_Partnership_Log.xlsx` + `UAC_Farmer_Engagement_Log.xlsx`

### 4. Set up monthly reminders (10 min)

Follow `schemas/monthly_reminders.md` — either a Power Automate
scheduled flow or a recurring Outlook calendar event on the 1st of
each month.

### 5. Test (5 min)

Each role owner opens their log, adds a test row, saves. Director
opens the master workbook, clicks Refresh All, confirms the test row
appears on the corresponding Data tab and the dashboard tiles update.

Delete sample/test rows once real data is flowing.

---

## Operating rhythm

| Cadence | Who | What |
|---|---|---|
| Weekly | Logistics Manager | Add rows to Kitchen Operations Log |
| Per event | Culinary / Outreach | Add rows within 48 hours |
| Monthly (1st) | Director | Update Director + UAC logs; get Food Corridor CSV |
| Monthly (2nd week) | Director | Refresh master workbook, review dashboards |
| Before each JSC | Director | Refresh, screenshot 2026 Goals + Exec Summary |
| Quarterly | Director | Update earned revenue %, member sales growth |

---

## Troubleshooting

**The master workbook won't refresh.**
- Check Power Query connections: Data > Queries & Connections >
  right-click a query > Properties. Confirm the `LogsFolder` parameter
  points to the right SharePoint path.
- If you see credential errors: Data > Data Source Settings > select
  the SharePoint URL > Edit Permissions > sign in.

**A dashboard tile shows 0 when there should be data.**
- Open the corresponding Data tab. Are there rows?
- If empty, the Power Query connection isn't working — check the path.
- If rows exist, check that the table name is correct (e.g.,
  `tbl_Director`, not `tbl_Director_1`).

**A formula shows `#REF!` or `#NAME?`.**
- A table or column was likely renamed. Open the formula, check that
  structured references match actual table/column names.

**Staff can't open their log workbook.**
- Confirm SharePoint/OneDrive sharing permissions. Each role owner
  needs Edit access to their own log file.

---

## Maintenance

### Adding a new KPI

1. Add a row to `docs/kpi_catalog.md`.
2. Add columns to the relevant log workbook (open it, click the table
   header row, add a column to the right).
3. Update the Power Query M code in `docs/data_dictionary.md` if the
   new column needs type transformation.
4. Add a dashboard tile or row in the master workbook.

### Adding a new survey question

No structural change needed. In `Community_Engagement_Log.xlsx`,
open the **Survey Responses** tab and add rows with the new
`Question_Text`. The master workbook picks them up on next refresh.

### Retiring a KPI

Change its Status to "Retired" in the catalog. Leave the column in
the log workbook for historical data. Remove the KPI from dashboard
tabs.

---

## Out of scope

- **Power BI dashboards** — Excel is v1; Power BI is a later phase.
- **OFSA logic model** — KPI tagging is the bridge; the model is next.
- **Patient-level health data** — no PHI. Readmission KPI is a
  placeholder.
- **SharePoint Lists and Forms** — removed in favor of direct Excel
  entry per team feedback.
- **Structural documentation (Goal 2)** — tracked in the shared drive.
