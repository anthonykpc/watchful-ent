# HARVEST Food Hub — KPI Tracking System

A complete KPI tracking system for HARVEST Food Hub. Captures data
from HARVEST staff and external UAC partners, stores it in SharePoint,
tags every metric to the OFSA Six Dimensions of Food Security, and
rolls up to a master Excel dashboard for the JSC, funders, and the
Director's annual review.

## What this is

Five Microsoft Forms feed six SharePoint Lists. Power Automate routes
submissions to the right List. An Excel master workbook reads from
those Lists via Power Query and renders dashboards (Executive Summary,
By OFSA Dimension, By Audience, 2026 Performance Goals). Food Corridor
CSV exports drop into a dedicated tab to bring in kitchen rental hours
without manual re-entry.

The system is built on the M365 stack you already have. No new
licenses. No third-party SaaS. UAC partners submit via an anonymous
public Form link — no M365 license needed on their side.

## Architecture

```
                    +------------------------------+
                    |   Microsoft Forms (x5)       |
                    |   - Kitchen Operations       |
                    |   - Culinary & Training      |
                    |   - Community Engagement     |
                    |   - Director & Partnership   |
                    |   - UAC Farmer Engagement    |   <- external/public link
                    +--------------+---------------+
                                   |
                                   v
                    +------------------------------+
                    |   Power Automate flows       |
                    |   (route + notify)           |
                    +--------------+---------------+
                                   |
                                   v
                    +------------------------------+
                    |   SharePoint Lists (x6)      |
                    |   - Food Entrepreneurs       |
                    |   - Institution              |
                    |   - Community Orgs           |
                    |   - Investors-Funders        |
                    |   - Farmers                  |
                    |   - Environmental            |
                    +--------------+---------------+
                                   |
        +--------------------------+--------------------------+
        |                                                     |
        v                                                     v
+-----------------+                              +------------------------+
| Food Corridor   |  -- CSV drop -->             | Excel master workbook  |
| (external SaaS) |                              | (Power Query connects) |
+-----------------+                              +------------------------+
                                                              |
                                                              v
                                                  +-----------------------+
                                                  | Dashboards            |
                                                  | - Executive Summary   |
                                                  | - By OFSA Dimension   |
                                                  | - By Audience         |
                                                  | - 2026 Performance    |
                                                  +-----------------------+
```

## Prerequisites

- **Microsoft 365** for the HARVEST team (E3 or higher). UAC partners
  do not need a license.
- **SharePoint Online** site for HARVEST (existing site reused, or a
  new site created — either is fine).
- **Power Automate** standard connectors (no premium tier needed).
- **Food Corridor** export access for the Logistics Manager.
- **For the provisioning script (optional)**: PowerShell 7+ and the
  PnP.PowerShell module (`Install-Module -Name PnP.PowerShell -Scope CurrentUser`).
  If your tenant blocks the module, use the manual setup path —
  everything works without it.

## File map

| File | Purpose |
|---|---|
| `README.md` | This file. Start here. |
| `HARVEST_KPI_Master.xlsx` | The user-facing deliverable. Dashboards live here. |
| `build_kpi_workbook.py` | Regenerates the workbook (Python, openpyxl). Don't run unless you need a clean rebuild — you'll lose any Power Query connections you added. |
| `schemas/sharepoint_lists.md` | Column definitions for each List + manual setup walkthrough. |
| `schemas/microsoft_forms.md` | Question-by-question spec for each Form + setup walkthrough. |
| `schemas/power_automate_flows.md` | Flow logic, triggers, field mappings, error handling. |
| `scripts/provision_lists.ps1` | PnP PowerShell script that creates all six Lists. |
| `docs/kpi_catalog.md` | The full enriched KPI catalog. Source of truth. |
| `docs/ofsa_dimensions_reference.md` | OFSA Six Dimensions definitions + tagging guidance. |
| `docs/role_assignments.md` | Who owns what; weekly/monthly cadence. |
| `docs/data_dictionary.md` | Field definitions + Power Query M code (paste-ready). |
| `samples/food_corridor_sample.csv` | Sample Food Corridor export structure (v1 minimal schema). |
| `samples/seed_data_sample.csv` | Sample rows for each List, for dashboard testing. |

## Setup checklist (one-time)

Estimated total: **3–4 hours**, in order.

### 1. Prepare the SharePoint site (5 min)

- Pick an existing HARVEST SharePoint site, or create a new one
  (`https://[tenant].sharepoint.com/sites/HARVEST`). Note the full URL.

### 2. Create the SharePoint Lists (15 min via script, ~2 hours manual)

**Option A — script (recommended):**
```powershell
Install-Module -Name PnP.PowerShell -Scope CurrentUser
cd scripts
.\provision_lists.ps1 -SiteUrl https://[tenant].sharepoint.com/sites/HARVEST
```
The script is idempotent — safe to re-run if you need to add new columns later.

**Option B — manual UI:** Follow `schemas/sharepoint_lists.md`, section
"Manual setup walkthrough — Path B".

### 3. Enable external sharing on the Farmers List (5 min)

Required for UAC partners (external submitters) to see/edit it
directly. See `schemas/sharepoint_lists.md`, section "External sharing
setup (Farmers List only)". This may require a SharePoint admin
depending on your tenant's policy.

### 4. Create the five Microsoft Forms (60 min total)

For each Form, follow the per-Form setup in `schemas/microsoft_forms.md`:
- Kitchen Operations Log (Logistics Manager owns)
- Culinary & Training Log (Culinary Manager owns)
- Community Engagement Log (Community Outreach Specialist owns)
- Director & Partnership Log (Director owns)
- UAC Farmer Engagement Log (Director owns + Director adds co-owners)
  — **critical:** set sharing to "Anyone can respond" so external UAC
  partners can submit without an M365 license.

### 5. Build the Power Automate flows (60 min total)

Follow `schemas/power_automate_flows.md`:
- 5 Form → List flows (one per Form)
- 1 New entrepreneur onboarding alert
- 1 Overdue cadence reminder (weekly schedule)
- 1 Grant entered alert (Teams + email)

Test each Form → List flow with a dummy submission and confirm the row
appears in the destination List.

### 6. Wire up the Excel workbook (30 min)

1. Save `HARVEST_KPI_Master.xlsx` to the HARVEST SharePoint site (not
   your Downloads folder).
2. Open it in Excel desktop or Excel Online.
3. Set the `HarvestSite` parameter to your SharePoint site URL —
   instructions in `docs/data_dictionary.md`.
4. Add one Power Query connection per Data tab, using the M code in
   `docs/data_dictionary.md` (7 queries total: 6 Lists + Food Corridor).
5. Click **Data → Refresh All**. Dashboards populate.

### 7. Test with sample data (15 min)

Submit at least one response per Form (use a dummy "Test Test" name).
Confirm:
- The row appears in the right SharePoint List within 60 seconds.
- The Executive Summary dashboard tile increments after Refresh All.
- The 2026 Goals tab status indicator updates if relevant.

### 8. Share with the team (10 min)

- Send each role owner the link to their Form.
- Send the Director's calendar a recurring monthly reminder to refresh
  the workbook before the JSC meeting.
- Add a link to the workbook in the HARVEST Teams channel.

---

## Daily / weekly / monthly operating rhythm

See `docs/role_assignments.md` for the full breakdown. Summary:

| Cadence | Owner | What |
|---|---|---|
| Weekly (Mon) | Logistics Manager | Submit Kitchen Operations Log |
| Per event | Culinary / Outreach Mgrs | Submit Form within 48 hours |
| Monthly (1st) | UAC Partner | Submit Farmer Engagement Log |
| Monthly (1st) | Director | Submit Partnership Log; drop Food Corridor CSV |
| Monthly (2nd week) | Director | Refresh + review |
| Per JSC | Director | Refresh + screenshot for meeting |
| Quarterly | Director | Update SROI, earned revenue, sales growth |
| Annually | Director | Annual review using 2026 Goals tab |

---

## Troubleshooting

**A Form submission isn't appearing in the List.**
1. Check the Power Automate flow run history (Power Automate → My flows
   → click the flow → Run history). Look for a red X.
2. Click the failed run to see the exact error. Most common: a Form
   question was renamed but the flow still references the old name.
3. Fix the field mapping in the flow's "Create item" action.

**The Excel workbook won't refresh.**
1. Check the SharePoint connection in Data → Queries & Connections →
   right-click query → Properties. Confirm the site URL matches.
2. If you see "credentials" errors, click Data → Get Data → Data Source
   Settings → select the SharePoint URL → Edit Permissions → Edit →
   sign in with your RWJBH account.

**A dashboard cell shows `#REF!` or `#NAME?`.**
1. The most likely cause is a table or column was renamed. Open the
   formula and confirm `tbl_Funders[Activity_Type]` (etc.) match the
   actual table names.
2. Check Formulas → Name Manager for stale references.

**The UAC Form rejects an external submitter.**
1. In Microsoft Forms, open the Form → Settings (…) → confirm "Anyone
   can respond" is selected. If it's not, the form is tenant-restricted
   and external users get a sign-in prompt.

**The provisioning script fails on connect.**
1. If you see "Cannot find module PnP.PowerShell," install it:
   `Install-Module -Name PnP.PowerShell -Scope CurrentUser`.
2. If you see a tenant authorization error, your tenant may require an
   admin to consent to the PnP app. Ask SharePoint admin to register
   the PnP Management Shell app, OR use the manual UI path in
   `schemas/sharepoint_lists.md`.

**A dashboard formula returns 0 when there should be data.**
1. Confirm the Data tab actually has rows (open the tab; scroll down).
2. Confirm the Excel Table on that tab is named correctly
   (`tbl_Funders` not `tbl_Funders_1`). Tables get renumbered when you
   delete and recreate them.

---

## Maintenance

### Adding a new KPI

1. Add a row to `docs/kpi_catalog.md` with all tagging fields populated.
2. Decide which Form / List it lands in. Update
   `schemas/microsoft_forms.md` (add the question) and
   `schemas/sharepoint_lists.md` (add the column).
3. Update the Form in the Microsoft Forms UI to add the question.
4. Add the column to the SharePoint List (rerun the script with the new
   column added, or add via the UI).
5. Update the destination Power Automate flow to map the new question.
6. If the KPI should show on the dashboards, edit
   `build_kpi_workbook.py` (add to `KPI_CATALOG`) and either rerun the
   script for a full rebuild OR manually add a row to the KPI Catalog
   tab + a tile / row in the relevant dashboard tab.

### Retiring a KPI

1. Change its `Status` in the catalog to `Retired`.
2. Leave the historical data in the List — don't delete columns.
3. Remove the KPI from the dashboard tabs (it'll continue to be tagged
   in the catalog for historical reference).

### Changing a cadence

1. Update `docs/kpi_catalog.md`.
2. Update the overdue cadence reminder flow if the new cadence falls
   into a different alert window.

---

## Out of scope (intentional)

- **Power BI dashboards.** Excel is v1. Power BI is a planned later
  phase. The OFSA dimension tagging on every KPI is what makes that
  migration straightforward.
- **OFSA logic model.** The KPI tagging is the bridge to it. The model
  itself is built in the next phase using the OFSA template.
- **Patient-level health outcome data from RWJBH EHR systems.** No PHI
  is ingested. The "Reduction in readmission costs" KPI is a
  placeholder until RWJBH provides an aggregate, non-PHI methodology.
- **Mobile app or Power Apps.** Forms + Excel on mobile is sufficient
  for v1.
- **Structural documentation (Goal 2).** SOPs and guiding documents
  are tracked in the shared drive folder, not in this system. The 2026
  Performance Goals tab notes this explicitly.
