# Data Dictionary and Power Query M Code

Three things live in this file:
1. **Field definitions** — every column in every log workbook.
2. **Power Query M code** — paste-ready code to connect the master
   workbook's Data tabs to the five log workbooks.
3. **Judgment-call log** — small decisions made during the build.

---

## 1. Field definitions

### Kitchen Operations Log

| Column | Meaning |
|---|---|
| `Date` | Date this row was entered. |
| `Reporting_Week_Start` / `_End` | Monday–Sunday of the week being reported. |
| `Entry_Type` | Weekly Log / Incident / New Member Onboarding. |
| `Active_Members` | Total active member businesses this week. |
| `New_Members` | Members who signed a rental agreement this week. |
| `New_Member_Business` | Business name (if new member). |
| `New_Member_Demographics` | Self-identified tags (multi-select dropdown). |
| `Equipment_Status` | All operational / Minor issue / Major issue. |
| `Equipment_Notes` | Details if equipment not fully operational. |
| `Incident_Type` | Safety / Sanitation / Member conflict / Other. |
| `Incident_Description` | What happened. |
| `Food_Waste_Diverted_lbs` | lbs of food waste diverted from landfill this week. |
| `Composted_lbs` | lbs composted this week. |
| `Notes` | Free text. |

### Culinary & Training Log

| Column | Meaning |
|---|---|
| `Date` | Date of the workshop/training/session. |
| `Event_Type` | Dropdown: ServSafe, Health/Nutrition Demo, Seminar types, 1:1 TA, Wellness Coaching, Other. |
| `Event_Title` | Short name for the event. |
| `Curriculum` | **NEW.** Name of the curriculum for health demos (e.g., "Diabetes-Friendly Cooking"). Dropdown with common options; type a custom one if needed. Feeds the "Health demo curriculums delivered" KPI. |
| `Duration_Hours` | Length of the event in hours. |
| `Attendance_Count` | Number of attendees. |
| `Businesses_Present` | Which member businesses attended (comma-separated). |
| `ServSafe_Passed` | # who passed ServSafe (if applicable). |
| `ServSafe_Attempted` | # who attempted ServSafe (denominator for the % KPI). |
| `TA_Topic` | Topic of the 1:1 TA session. |
| `TA_Recipient_Business` | Which business received the TA. |
| `Wellness_Coaching_Recipient` | Which business received wellness coaching. |
| `RD_InKind_Hours` | **NEW.** Hours the Registered Dietitian provided demos at no charge to the community. This is the in-kind hour tracking KPI. |
| `Food_Captured_Demo_lbs` | lbs of food used in demos instead of being wasted. |
| `Food_Captured_Demo_$` | Estimated $ value of that food. |
| `Notes` | Free text. |

### Community Engagement Log — Log tab

| Column | Meaning |
|---|---|
| `Date` | Date of the event. |
| `Event_Name` | Name of the event. |
| `Event_Type` | Dropdown: Teaching Kitchen, RD Workshop, Tour, Partner-Hosted, HARVEST-Hosted, Pop-Up, Other. |
| `Attendance_Count` | Total attendees. |
| `Volunteer_Hours` | Hours contributed by community volunteers. |
| `Orgs_Engaged` | Organizations present (comma-separated). |
| `Sector_Count` | Distinct partner sectors represented. |
| `Multilingual_Cultural` | Yes/No. |
| `Language_Tradition` | Which language or tradition (if Yes). |
| `Food_Recovery_lbs` | lbs diverted to Bridges or other providers. |
| `Food_Recovery_Recipient` | Who received the recovered food. |
| `DSP_Onsite` | Yes/No — did a direct service provider have a presence? |
| `Notes` | Free text. |

### Community Engagement Log — Survey Responses tab

| Column | Meaning |
|---|---|
| `Date` | Date of the event the survey is about. |
| `Event_Name` | Which event. |
| `Survey_Instrument` | Name/version of the survey (e.g., "Post-Event v1", "Teaching Kitchen Series v1"). Lets you track which survey instrument was used as they evolve. |
| `Question_Text` | The full text of the survey question. |
| `Response_Text` | The text response (for open-ended questions). |
| `Numeric_Score` | The numeric score (for Likert/scale questions, typically 1–5). |
| `Notes` | Free text. |

**Why row-per-question?** New survey questions get new rows, not new
columns. You never need to restructure the table. When a new survey
instrument is created, just start entering rows with the new
`Survey_Instrument` name and new `Question_Text` values.

### Director & Partnership Log

| Column | Meaning |
|---|---|
| `Date` | Date of the activity. |
| `Entry_Type` | Dropdown: Grant, Partnership, Sponsorship, Leveraged Capital, Institutional Connection, Purchase Order, Menu Placement, TA Partnership, JSC Meeting, CAB Meeting, CAB Member Added, Professional Development, Earned Revenue Snapshot. |
| `Name` | Name of the funder, partner, institution, or activity. |
| `Amount_$` | Dollar value (for grants, POs, leveraged capital). |
| `Period_Start` / `Period_End` | Grant or reporting period dates. |
| `Institution_Type` | Healthcare / Education / Government / Corporate / Other. |
| `Items_On_Menu` | Count of HARVEST items on institutional menus. |
| `CAB_Member_Name` | Name of new CAB member (if CAB Member Added). |
| `CAB_Member_Sector` | Which sector the CAB member represents. |
| `PD_Activity` | Type of professional development (dropdown). |
| `Earned_Revenue_%` | Earned revenue as % of total (quarterly snapshot). |
| `Avg_Sales_Increase_%` | Avg YoY sales increase for members (quarterly). |
| `Status` | Active / Pending / Closed. |
| `Notes` | Free text. |

### UAC Farmer Engagement Log

| Column | Meaning |
|---|---|
| `Date` | Date this row was entered. |
| `Reporting_Month` | Which month is being reported (e.g., "March 2026"). |
| `Entered_By` | Who entered this row (HARVEST staff on behalf of UAC). |
| `Farm_Name` | Name of the farm. One row per farm per month. |
| `Farm_Location` | City, State. |
| `Farm_Contact` | Name / email / phone. |
| `Produce_Types` | Comma-separated list of produce moved this month. |
| `Produce_lbs` | Total lbs moved. |
| `Produce_Value_$` | Total $ value moved. |
| `Origin_Region` | County or region of origin. |
| `Distribution_Destinations` | Where produce went (comma-separated). |
| `Value_Added_Product` | Yes/No — did this farmer make a value-added product? |
| `Value_Added_Description` | What was made (e.g., "Pepper hot sauce"). |
| `Demographic_Race_Ethnicity` | Self-identified. Dropdown. |
| `Demographic_Gender` | Self-identified. Dropdown. |
| `Demographic_Age_Band` | Self-identified. Dropdown. |
| `Notes` | Free text. |

---

## 2. Power Query M code

Paste these into the master workbook: **Data > Get Data > From Other
Sources > Blank Query > Advanced Editor**. Name each query exactly as
shown so the dashboard formulas keep working.

Before pasting, set this parameter:
- **Data > Get Data > Data Source Settings > Manage Parameters > New**
- Name: `LogsFolder`
- Type: Text
- Value: the full path to the `logs/` folder on SharePoint, e.g.
  `https://rwjbh.sharepoint.com/sites/HARVEST/Shared Documents/HARVEST_KPI_System/logs/`

If working locally, use a local path like `C:\Users\acapece\...\logs\`.

### tbl_KitchenOps

```m
let
    Source = Excel.Workbook(
        File.Contents(LogsFolder & "Kitchen_Operations_Log.xlsx"),
        null, true
    ),
    LogTable = Source{[Item="tbl_KitchenOps", Kind="Table"]}[Data],
    Typed = Table.TransformColumnTypes(LogTable, {
        {"Date", type date},
        {"Reporting_Week_Start", type date},
        {"Reporting_Week_End", type date},
        {"Active_Members", Int64.Type},
        {"New_Members", Int64.Type},
        {"Food_Waste_Diverted_lbs", type number},
        {"Composted_lbs", type number}
    })
in
    Typed
```

### tbl_Culinary

```m
let
    Source = Excel.Workbook(
        File.Contents(LogsFolder & "Culinary_Training_Log.xlsx"),
        null, true
    ),
    LogTable = Source{[Item="tbl_Culinary", Kind="Table"]}[Data],
    Typed = Table.TransformColumnTypes(LogTable, {
        {"Date", type date},
        {"Duration_Hours", type number},
        {"Attendance_Count", Int64.Type},
        {"ServSafe_Passed", Int64.Type},
        {"ServSafe_Attempted", Int64.Type},
        {"RD_InKind_Hours", type number},
        {"Food_Captured_Demo_lbs", type number},
        {"Food_Captured_Demo_$", Currency.Type}
    })
in
    Typed
```

### tbl_CommunityEvents

```m
let
    Source = Excel.Workbook(
        File.Contents(LogsFolder & "Community_Engagement_Log.xlsx"),
        null, true
    ),
    LogTable = Source{[Item="tbl_CommunityEvents", Kind="Table"]}[Data],
    Typed = Table.TransformColumnTypes(LogTable, {
        {"Date", type date},
        {"Attendance_Count", Int64.Type},
        {"Volunteer_Hours", type number},
        {"Sector_Count", Int64.Type},
        {"Food_Recovery_lbs", type number}
    })
in
    Typed
```

### tbl_SurveyResponses

```m
let
    Source = Excel.Workbook(
        File.Contents(LogsFolder & "Community_Engagement_Log.xlsx"),
        null, true
    ),
    SurveyTable = Source{[Item="tbl_SurveyResponses", Kind="Table"]}[Data],
    Typed = Table.TransformColumnTypes(SurveyTable, {
        {"Date", type date},
        {"Numeric_Score", type number}
    })
in
    Typed
```

### tbl_Director

```m
let
    Source = Excel.Workbook(
        File.Contents(LogsFolder & "Director_Partnership_Log.xlsx"),
        null, true
    ),
    LogTable = Source{[Item="tbl_Director", Kind="Table"]}[Data],
    Typed = Table.TransformColumnTypes(LogTable, {
        {"Date", type date},
        {"Amount_$", Currency.Type},
        {"Period_Start", type date},
        {"Period_End", type date},
        {"Items_On_Menu", Int64.Type},
        {"Earned_Revenue_%", type number},
        {"Avg_Sales_Increase_%", type number}
    })
in
    Typed
```

### tbl_Farmers

```m
let
    Source = Excel.Workbook(
        File.Contents(LogsFolder & "UAC_Farmer_Engagement_Log.xlsx"),
        null, true
    ),
    LogTable = Source{[Item="tbl_Farmers", Kind="Table"]}[Data],
    Typed = Table.TransformColumnTypes(LogTable, {
        {"Date", type date},
        {"Produce_lbs", type number},
        {"Produce_Value_$", Currency.Type}
    })
in
    Typed
```

### tbl_FoodCorridor

```m
let
    Source = Csv.Document(
        File.Contents(LogsFolder & "../food_corridor.csv"),
        [Delimiter=",", Columns=3, Encoding=65001,
         QuoteStyle=QuoteStyle.Csv]
    ),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Renamed = Table.RenameColumns(Promoted, {
        {"Booking Date", "Booking_Date"},
        {"Member Business", "Member_Business"}
    }),
    Typed = Table.TransformColumnTypes(Renamed, {
        {"Booking_Date", type date},
        {"Member_Business", type text},
        {"Hours", type number}
    })
in
    Typed
```

---

## 3. Judgment-call log

- **Dropdown wording.** Title Case for all options.
- **Curriculum dropdown.** Seeded with 8 common health demo curriculums.
  The dropdown allows "Other (specify in Notes)" and staff can type
  custom values — dropdowns in Excel show a warning but still accept
  the entry.
- **RD in-kind hours.** Added to the Culinary & Training Log since RD
  demo hours are tied to training events. The dashboard sums this
  across all events.
- **Survey structure.** Row-per-question in the Community Engagement Log.
  `Survey_Instrument` column lets you track which version of a survey
  was used. New survey questions = new rows, not new columns.
- **UAC data entry.** HARVEST staff enters UAC data on behalf of UAC
  partners, based on a monthly call/email. No external Excel sharing
  needed.
- **Governance data.** JSC and CAB activity lives in the Director's
  log with governance-specific `Entry_Type` values rather than a
  separate workbook.
- **No sheet protection on log workbooks.** Staff need full edit access
  to enter data. Dashboard tabs in the master workbook are protected
  (no password).
- **Sample data.** Every log ships with 2–4 sample rows so the master
  dashboard renders numbers immediately. Delete sample rows once real
  data is flowing.
