# Data Dictionary and Power Query M Code

Three things live in this file:
1. **Field definitions** — every column on every List, in plain English.
2. **Power Query M code** — paste-ready code to connect each Data tab in
   the master workbook to its SharePoint List.
3. **Judgment-call log** — small decisions made during the build that
   weren't worth pinging Anthony for.

---

## 1. Field definitions

### Common columns (every List)

| Column | Plain-English meaning |
|---|---|
| `Title` | A short, human-readable label for the row. For events, the event name. For roll-ups, e.g. "Q1 operational snapshot". |
| `Submission_ID` | Auto-assigned by SharePoint. Used for cross-referencing. |
| `Submitted_By` | Who submitted this row. Text (not Person picker) so external UAC submitters work without an M365 license. |
| `Submitted_Date` | Date and time the row was created. Set by Power Automate. |
| `Reporting_Period_Start` / `Reporting_Period_End` | The window of activity this row describes. For a per-event entry these are the event date. For a monthly roll-up they're the first/last day of the month. |
| `OFSA_Dimension` | Multi-select. Which of the OFSA Six Dimensions this row contributes to. See `ofsa_dimensions_reference.md` for definitions. |
| `Food_System_Sector` | Production / Aggregation / Processing / Distribution / Retail-Consumption / Recovery-Waste. The point in the food chain this row sits in. |
| `Partner_Sector` | Healthcare / Agriculture / Education / Government / Philanthropy / Community-Based Org / Private Sector / Food Enterprise. The sector of the partner/audience the row is about. |
| `Lead_Role` | HARVEST Lead / UAC Lead / Joint. Who's primarily driving the work this row describes. |
| `Notes` | Free-text. Include date prefixes for incremental notes (e.g., `2026-04-22: …`). |
| `Attachments` | Files attached to the row (photos, contracts, screenshots). |

### Per-List columns

See `schemas/sharepoint_lists.md` for the per-List column tables. Each column there is named and typed; this section adds the plain-English meaning where it isn't obvious from the name.

**Food Entrepreneurs List**
- `Stage` — Where this business sits in the HARVEST funnel: `Engaged` (had a conversation), `Onboarding` (signing paperwork), `Active Member` (renting kitchen), `Inactive` (was a member, not currently), `Graduated` (moved to their own facility), `Operational Snapshot` (the row is a weekly operations log, not a business record).
- `Sales_YoY_Change_%` — Self-reported by the business. Use the most recent 12 months vs. the prior 12 months. Allow negative values.
- `Demographic_Tags` — Self-identified by the business owner during onboarding. Don't infer.

**Community Orgs List**
- `Sector_Count` — How many distinct partner sectors were represented at this event. E.g., a teaching kitchen co-hosted with RWJBH (healthcare) + La Casa (CBO) + Newark Public Library (government) = 3.
- `Multilingual_Cultural` — `Yes` if the event was delivered in a language other than English OR was built around a specific cultural tradition. Mark `Yes` even if only part of the event qualified.

**Farmers List**
- `Origin_Region` — Geographic specificity at the county-or-finer level. "Newark, NJ" is fine; "Northeast" is too coarse.
- `Distribution_Destinations` — Where the produce ended up. Recipient orgs OR neighborhoods if not org-attributable.

**Environmental List**
- `Activity_Type = Food Miles Snapshot` — Used only for quarterly avg distance entries. lbs_Value is null for these rows; use `Food_Miles_Avg` instead.

### KPI Catalog tagging schema

Every KPI in the catalog carries the tags below. The catalog is the
source of truth — change a KPI here first, then propagate.

| Tag | Allowed values |
|---|---|
| Audience | Food Entrepreneurs / Institution / Community Organizations / Investors-Funders / Farmers / Environmental Sustainability |
| OFSA Dimension(s) | Availability / Access / Utilization / Stability / Agency / Sustainability (multi-select) |
| Food System Sector | Production / Aggregation / Processing / Distribution / Retail-Consumption / Recovery-Waste / n/a |
| Partner Sector | Healthcare / Agriculture / Education / Government / Philanthropy / Community-Based Org / Private Sector / Food Enterprise / n/a |
| Cadence | Real-time / Weekly / Monthly / Quarterly / Annual / Per event / Per session / Per training |
| Lead Role | HARVEST Lead / UAC Lead / Joint |
| Data Source | Form submission / Food Corridor CSV / Manual entry / System-calculated |
| Owner Role | Director / Culinary Manager / Logistics Manager / Community Outreach Specialist / UAC Partner |
| Status | Active / Planned — methodology TBD / Retired |

---

## 2. Power Query M code

Paste these into Excel: **Data → Get Data → From Other Sources → Blank
Query → Advanced Editor**. Name the query exactly as the table name
shown so the existing dashboard formulas keep working.

Before pasting, set this parameter once:

- Excel ribbon → **Data → Get Data → Data Source Settings → Manage
  Parameters → New** → Name = `HarvestSite`, Type = Text, Current Value
  = the URL of your HARVEST SharePoint site (e.g.,
  `https://rwjbh.sharepoint.com/sites/HARVEST`).

All queries below reference `HarvestSite`. If you don't want to use the
parameter, find and replace `HarvestSite` with the URL string directly.

### tbl_Entrepreneurs — Data — Food Entrepreneurs tab

```m
let
    Source = SharePoint.Tables(HarvestSite, [Implementation="2.0", ApiVersion=15]),
    EntrepreneursList = Source{[Title="Food Entrepreneurs List"]}[Items],
    KeepCols = Table.SelectColumns(EntrepreneursList, {
        "ID", "Title", "Submitted_By", "Submitted_Date",
        "Reporting_Period_Start", "Reporting_Period_End",
        "OFSA_Dimension", "Food_System_Sector", "Partner_Sector", "Lead_Role",
        "Business_Name", "Stage", "Date_Engaged", "Date_Converted",
        "Jobs_Created_FTE", "Jobs_Created_PTE", "Sales_Period_USD",
        "Sales_YoY_Change_Pct", "Sources_Local_Produce",
        "Wellness_Coaching_Sessions", "Demographic_Tags",
        "New_Members", "ServSafe_Passed", "ServSafe_Attempted", "Notes"
    }),
    Renamed = Table.RenameColumns(KeepCols, {
        {"ID", "Submission_ID"},
        {"Sales_Period_USD", "Sales_Period_$"},
        {"Sales_YoY_Change_Pct", "Sales_YoY_Change_%"}
    }),
    TypedDates = Table.TransformColumnTypes(Renamed, {
        {"Submitted_Date", type datetime},
        {"Reporting_Period_Start", type date},
        {"Reporting_Period_End", type date},
        {"Date_Engaged", type date},
        {"Date_Converted", type date}
    }),
    TypedNumbers = Table.TransformColumnTypes(TypedDates, {
        {"Jobs_Created_FTE", Int64.Type},
        {"Jobs_Created_PTE", Int64.Type},
        {"Sales_Period_$", Currency.Type},
        {"Sales_YoY_Change_%", type number},
        {"Wellness_Coaching_Sessions", Int64.Type},
        {"New_Members", Int64.Type},
        {"ServSafe_Passed", Int64.Type},
        {"ServSafe_Attempted", Int64.Type}
    })
in
    TypedNumbers
```

### tbl_Institution — Data — Institution tab

```m
let
    Source = SharePoint.Tables(HarvestSite, [Implementation="2.0", ApiVersion=15]),
    InstitutionList = Source{[Title="Institution List"]}[Items],
    KeepCols = Table.SelectColumns(InstitutionList, {
        "ID", "Title", "Submitted_By", "Submitted_Date",
        "Reporting_Period_Start", "Reporting_Period_End",
        "OFSA_Dimension", "Food_System_Sector", "Partner_Sector", "Lead_Role",
        "Institution_Name", "Institution_Type", "Activity_Type",
        "PO_Value_USD", "Items_On_Menu", "Connection_Notes",
        "Health_Outcome_Metric", "Health_Outcome_Value", "Notes"
    }),
    Renamed = Table.RenameColumns(KeepCols, {
        {"ID", "Submission_ID"},
        {"PO_Value_USD", "PO_Value_$"}
    }),
    Typed = Table.TransformColumnTypes(Renamed, {
        {"Submitted_Date", type datetime},
        {"Reporting_Period_Start", type date},
        {"Reporting_Period_End", type date},
        {"PO_Value_$", Currency.Type},
        {"Items_On_Menu", Int64.Type},
        {"Health_Outcome_Value", Currency.Type}
    })
in
    Typed
```

### tbl_Community — Data — Community Orgs tab

```m
let
    Source = SharePoint.Tables(HarvestSite, [Implementation="2.0", ApiVersion=15]),
    CommunityList = Source{[Title="Community Orgs List"]}[Items],
    KeepCols = Table.SelectColumns(CommunityList, {
        "ID", "Title", "Submitted_By", "Submitted_Date",
        "Reporting_Period_Start", "Reporting_Period_End",
        "OFSA_Dimension", "Food_System_Sector", "Partner_Sector", "Lead_Role",
        "Event_Name", "Event_Type", "Event_Date", "Attendance_Count",
        "Volunteer_Hours", "Orgs_Engaged", "Sector_Count",
        "Multilingual_Cultural", "Language_Tradition",
        "Survey_Satisfaction_Avg", "Survey_Learning_Avg",
        "Food_Recovery_lbs", "Food_Recovery_Recipient",
        "Direct_Service_Provider_Onsite", "Notes"
    }),
    Renamed = Table.RenameColumns(KeepCols, {{"ID", "Submission_ID"}}),
    Typed = Table.TransformColumnTypes(Renamed, {
        {"Submitted_Date", type datetime},
        {"Reporting_Period_Start", type date},
        {"Reporting_Period_End", type date},
        {"Event_Date", type date},
        {"Attendance_Count", Int64.Type},
        {"Volunteer_Hours", type number},
        {"Sector_Count", Int64.Type},
        {"Survey_Satisfaction_Avg", type number},
        {"Survey_Learning_Avg", type number},
        {"Food_Recovery_lbs", type number}
    })
in
    Typed
```

### tbl_Funders — Data — Investors-Funders tab

```m
let
    Source = SharePoint.Tables(HarvestSite, [Implementation="2.0", ApiVersion=15]),
    FundersList = Source{[Title="Investors-Funders List"]}[Items],
    KeepCols = Table.SelectColumns(FundersList, {
        "ID", "Title", "Submitted_By", "Submitted_Date",
        "Reporting_Period_Start", "Reporting_Period_End",
        "OFSA_Dimension", "Food_System_Sector", "Partner_Sector", "Lead_Role",
        "Funder_Name", "Activity_Type", "Amount_USD", "Period_Start",
        "Period_End", "Earned_Revenue_Pct", "Leveraged_Capital_Source",
        "SROI_Ratio", "Status", "Notes"
    }),
    Renamed = Table.RenameColumns(KeepCols, {
        {"ID", "Submission_ID"},
        {"Amount_USD", "Amount_$"},
        {"Earned_Revenue_Pct", "Earned_Revenue_%"}
    }),
    Typed = Table.TransformColumnTypes(Renamed, {
        {"Submitted_Date", type datetime},
        {"Period_Start", type date},
        {"Period_End", type date},
        {"Amount_$", Currency.Type},
        {"Earned_Revenue_%", type number},
        {"SROI_Ratio", type number}
    })
in
    Typed
```

### tbl_Farmers — Data — Farmers tab

```m
let
    Source = SharePoint.Tables(HarvestSite, [Implementation="2.0", ApiVersion=15]),
    FarmersList = Source{[Title="Farmers List"]}[Items],
    KeepCols = Table.SelectColumns(FarmersList, {
        "ID", "Title", "Submitted_By", "Submitted_Date",
        "Reporting_Period_Start", "Reporting_Period_End",
        "OFSA_Dimension", "Food_System_Sector", "Partner_Sector", "Lead_Role",
        "Farm_Name", "Farm_Location", "Farm_Contact",
        "Produce_Types", "Produce_lbs", "Produce_Value_USD",
        "Origin_Region", "Distribution_Destinations",
        "Value_Added_Product", "Value_Added_Description",
        "Demographic_Race_Ethnicity", "Demographic_Gender",
        "Demographic_Age_Band", "Notes"
    }),
    Renamed = Table.RenameColumns(KeepCols, {
        {"ID", "Submission_ID"},
        {"Produce_Value_USD", "Produce_Value_$"}
    }),
    Typed = Table.TransformColumnTypes(Renamed, {
        {"Submitted_Date", type datetime},
        {"Reporting_Period_Start", type date},
        {"Reporting_Period_End", type date},
        {"Produce_lbs", type number},
        {"Produce_Value_$", Currency.Type},
        {"Value_Added_Product", type logical}
    })
in
    Typed
```

### tbl_Environmental — Data — Environmental tab

```m
let
    Source = SharePoint.Tables(HarvestSite, [Implementation="2.0", ApiVersion=15]),
    EnvList = Source{[Title="Environmental List"]}[Items],
    KeepCols = Table.SelectColumns(EnvList, {
        "ID", "Title", "Submitted_By", "Submitted_Date",
        "Reporting_Period_Start", "Reporting_Period_End",
        "OFSA_Dimension", "Food_System_Sector", "Partner_Sector", "Lead_Role",
        "Activity_Type", "Activity_Date", "lbs_Value", "Dollar_Value",
        "Disposition", "Food_Miles_Avg", "Source_or_Recipient", "Notes"
    }),
    Renamed = Table.RenameColumns(KeepCols, {{"ID", "Submission_ID"}}),
    Typed = Table.TransformColumnTypes(Renamed, {
        {"Submitted_Date", type datetime},
        {"Reporting_Period_Start", type date},
        {"Reporting_Period_End", type date},
        {"Activity_Date", type date},
        {"lbs_Value", type number},
        {"Dollar_Value", Currency.Type},
        {"Food_Miles_Avg", type number}
    })
in
    Typed
```

### tbl_FoodCorridor — Food Corridor Import tab

This query loads the CSV from the same folder as the workbook. v1
schema is minimal (date, member, hours). Extend the column list when
you have a real export.

```m
let
    WorkbookFolder = Excel.CurrentWorkbook(){[Name="ThisWorkbookPath"]}[Content]{0}[Column1],
    CsvPath = WorkbookFolder & "food_corridor.csv",
    Source = Csv.Document(File.Contents(CsvPath), [Delimiter=",", Columns=3, Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Renamed = Table.RenameColumns(Promoted, {
        {"Booking Date", "Booking_Date"},
        {"Member Business", "Member_Business"},
        {"Hours", "Hours"}
    }),
    Typed = Table.TransformColumnTypes(Renamed, {
        {"Booking_Date", type date},
        {"Member_Business", type text},
        {"Hours", type number}
    })
in
    Typed
```

**Helper for `ThisWorkbookPath`.** Add a named range called
`ThisWorkbookPath` to the workbook that returns the folder path. The
formula:

```
=LEFT(CELL("filename",A1), FIND("[", CELL("filename",A1)) - 1)
```

Place it in any cell, then **Formulas → Define Name** → Name:
`ThisWorkbookPath`, Refers to: `=Sheet!cell`. If the workbook hasn't
been saved yet, `CELL("filename")` returns an empty string and the
query will fail — save to SharePoint/OneDrive first.

**Alternative (simpler):** hard-code the SharePoint folder path into
the M code. Less portable but easier to debug.

---

## 3. Judgment-call log

Small decisions made during the build that didn't need a clarifying
question:

- **Dropdown wording.** I used Title Case for all dropdown options
  (e.g., "Active Member" not "active member"). Microsoft Forms renders
  these as-is to submitters; Title Case reads cleaner in the UI.
- **Sector list.** I kept the Partner Sector list to 8 values
  (Healthcare, Agriculture, Education, Government, Philanthropy,
  Community-Based Org, Private Sector, Food Enterprise). Anthony's KPI
  list mentioned "different sectors engaged" without enumerating them —
  these 8 cover all the example partners in the brief.
- **Stage funnel.** Engaged → Onboarding → Active Member → Inactive /
  Graduated. Added "Operational Snapshot" as a non-business-record
  stage so weekly operations logs can sit in the same List without
  polluting business counts. The dashboards filter Operational Snapshot
  rows out of business counts.
- **CAB sectors.** Reused the Partner Sector list. The brief says
  "7–12 members across 4+ sectors" — I let the user pick any 4 from
  Partner Sector when adding a CAB member rather than defining a
  separate "CAB sector" enum.
- **Governance in Investors-Funders List.** JSC and CAB activity is
  written to the Investors-Funders List with governance-specific
  `Activity_Type` values, instead of creating a 7th SharePoint List
  for governance. Trade-off: keeps List count tight, costs a bit of
  semantic purity. Documented in `microsoft_forms.md`.
- **Kitchen Operations Form destination.** Operational-snapshot rows
  write to the Food Entrepreneurs List with `Stage = "Operational Snapshot"`
  rather than a separate 7th List. Same rationale as above. If your
  tenant allows, splitting this into a dedicated `Kitchen Operations Log`
  List is cleaner.
- **Sample seed data.** Every Data tab ships with 2–7 sample rows so
  the dashboards render numbers immediately without Power Query being
  wired up. Delete these once Power Query is connected to live Lists.
- **SharePoint column name encoding.** SharePoint Lists don't allow
  `$` or `%` in column internal names. Real columns are
  `PO_Value_USD`, `Earned_Revenue_Pct`, etc.; the workbook renames
  them to `PO_Value_$` and `Earned_Revenue_%` for display via the M
  code's `Table.RenameColumns` step.
- **No password on sheet protection.** Per the brief — keeps it
  recoverable if Anthony or any HARVEST staffer needs to edit a locked
  formula.
- **Food Corridor schema.** v1 is the minimum needed to compute
  "Kitchen hours rented" (date + member + hours). Will be expanded when
  a real export is available — flagged in the README and the import
  tab.
