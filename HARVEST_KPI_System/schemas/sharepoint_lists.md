# SharePoint Lists — Schemas and Setup

Six Lists, one per audience category. Each List is the canonical
storage for that audience's KPI submissions. Forms write here; Excel
reads from here via Power Query; Power Automate routes between them.

## Setup paths

You have two ways to create these Lists. Pick one:

- **Path A — Run the PowerShell script** (`scripts/provision_lists.ps1`).
  Creates all six Lists with columns, default views, and permissions in
  about 60 seconds. Requires PnP.PowerShell module (no app registration
  needed; uses interactive login).
- **Path B — Create by hand in the SharePoint UI**. Follow the
  step-by-step in each List's section below. Takes about 20 minutes per
  List. Use this if PowerShell is blocked on your machine.

Both paths produce the same end state.

---

## Common columns — present on every List

These ten columns are identical across all six Lists. The script creates
them automatically. If you're building by hand, add them first before
the per-List columns.

| Column name | Type | Required | Default | Notes |
|---|---|---|---|---|
| `Title` | Single line of text | Yes | — | SharePoint default. Use as the short event/submission name |
| `Submission_ID` | Calculated | Auto | `=ID` | Renders the row ID; surfaced for cross-list reference |
| `Submitted_By` | Single line of text | Yes | — | Text (not Person picker) so external UAC submitters work |
| `Submitted_Date` | Date and Time | Yes | `=Today` | Set automatically by Power Automate on submission |
| `Reporting_Period_Start` | Date | Yes | — | First day of the period this row describes |
| `Reporting_Period_End` | Date | Yes | — | Last day of the period this row describes |
| `OFSA_Dimension` | Choice (multi-select) | Yes | — | Choices: Availability; Access; Utilization; Stability; Agency; Sustainability |
| `Food_System_Sector` | Choice | No | — | Choices: Production; Aggregation; Processing; Distribution; Retail-Consumption; Recovery-Waste |
| `Partner_Sector` | Choice | No | — | Choices: Healthcare; Agriculture; Education; Government; Philanthropy; Community-Based Org; Private Sector; Food Enterprise |
| `Lead_Role` | Choice | Yes | HARVEST Lead | Choices: HARVEST Lead; UAC Lead; Joint |
| `Notes` | Multiple lines of text (plain) | No | — | Free text for context |
| `Attachments` | Attachments | No | — | SharePoint built-in attachments — enable in List settings |

**Why text for `Submitted_By`, not Person picker?** Person pickers
require the submitter to be a recognized user in the RWJBH tenant. UAC
partners are external. Text accepts any submitter and Power Automate
populates it from the Form response.

---

## List 1 — Food Entrepreneurs List

**Purpose.** Member businesses, conversion funnel, jobs, growth,
business support activities.

**Permissions.** Tenant-restricted. Edit access: Director, Culinary
Manager, Logistics Manager, Community Outreach Specialist. Read access:
JSC members.

### Columns (in addition to common columns)

| Column name | Type | Required | Default | Choices / validation |
|---|---|---|---|---|
| `Business_Name` | Single line of text | Yes | — | |
| `Stage` | Choice | Yes | Engaged | Engaged; Onboarding; Active Member; Inactive; Graduated |
| `Date_Engaged` | Date | No | — | |
| `Date_Converted` | Date | No | — | First date business rented kitchen |
| `Jobs_Created_FTE` | Number | No | 0 | Integer ≥ 0 |
| `Jobs_Created_PTE` | Number | No | 0 | Integer ≥ 0 |
| `Sales_Period_$` | Currency | No | — | Self-reported sales for the period |
| `Sales_YoY_Change_%` | Number | No | — | Percentage; allow negative |
| `Sources_Local_Produce` | Yes/No | No | No | True if business buys via UAC |
| `Wellness_Coaching_Sessions` | Number | No | 0 | Sessions received in the period |
| `ServSafe_Passed` | Number | No | 0 | # of ServSafe attendees who passed at this training (event-based row) |
| `ServSafe_Attempted` | Number | No | 0 | # of ServSafe attendees who attempted at this training (denominator for the % KPI) |
| `Demographic_Tags` | Choice (multi-select) | No | — | Woman-owned; BIPOC-owned; Immigrant-owned; Veteran-owned; LGBTQ+-owned; First-time entrepreneur |

### Default views

- **All Items** — default
- **Active Members** — filter `Stage = Active Member`
- **Current Quarter** — filter `Submitted_Date ≥ [start of current quarter]`
- **By Owner** — group by `Submitted_By`

---

## List 2 — Institution List

**Purpose.** Institutional purchasing, menu placements, cross-institution
connections, TA partnerships, healthcare outcome placeholders.

**Permissions.** Tenant-restricted. Edit: Director. Read: JSC.

### Columns (in addition to common columns)

| Column name | Type | Required | Default | Choices / validation |
|---|---|---|---|---|
| `Institution_Name` | Single line of text | Yes | — | |
| `Institution_Type` | Choice | Yes | — | Healthcare; Education; Government; Corporate; Other |
| `Activity_Type` | Choice | Yes | — | Purchase Order; Menu Placement; TA Partnership; Cross-Institution Connection; Health Outcome Measure |
| `PO_Value_$` | Currency | No | — | If Activity_Type = Purchase Order |
| `Items_On_Menu` | Number | No | — | If Activity_Type = Menu Placement |
| `Connection_Notes` | Multiple lines of text | No | — | Details of the connection or partnership |
| `Health_Outcome_Metric` | Single line of text | No | — | Placeholder for readmission / outcome metric |
| `Health_Outcome_Value` | Currency | No | — | Placeholder — leave blank until methodology exists |

### Default views

- **All Items**
- **Current Quarter** — filter on `Submitted_Date`
- **By Activity Type** — group by `Activity_Type`

---

## List 3 — Community Orgs List

**Purpose.** Community engagement events, tours, attendance, surveys,
food recovery, volunteer hours, multilingual programming.

**Permissions.** Tenant-restricted. Edit: Community Outreach Specialist,
Director. Read: JSC.

### Columns (in addition to common columns)

| Column name | Type | Required | Default | Choices / validation |
|---|---|---|---|---|
| `Event_Name` | Single line of text | Yes | — | |
| `Event_Type` | Choice | Yes | — | Teaching Kitchen; RD Workshop; Tour; Partner-Hosted Event; HARVEST-Hosted Event; Pop-Up; Other |
| `Event_Date` | Date | Yes | — | |
| `Attendance_Count` | Number | No | 0 | Integer ≥ 0 |
| `Volunteer_Hours` | Number | No | 0 | Decimal ≥ 0 |
| `Orgs_Engaged` | Multiple lines of text | No | — | Comma-separated list of org names |
| `Sector_Count` | Number | No | 0 | Distinct sectors represented at this event |
| `Multilingual_Cultural` | Yes/No | No | No | True if event was in a non-English language or built around a cultural tradition |
| `Language_Tradition` | Single line of text | No | — | E.g., "Spanish", "Haitian Creole", "Lunar New Year" |
| `Survey_Satisfaction_Avg` | Number | No | — | 1.0–5.0 |
| `Survey_Learning_Avg` | Number | No | — | 1.0–5.0 |
| `Food_Recovery_lbs` | Number | No | 0 | lbs diverted at this event |
| `Food_Recovery_Recipient` | Single line of text | No | — | E.g., "Bridges", "St. James Pantry" |
| `Direct_Service_Provider_Onsite` | Yes/No | No | No | True if a DSP had a presence at this event |

### Default views

- **All Items**
- **Current Quarter** — filter on `Event_Date`
- **By Event Type** — group by `Event_Type`
- **Multilingual / Cultural** — filter `Multilingual_Cultural = Yes`

---

## List 4 — Investors-Funders List

**Purpose.** Grants, partnerships, funder relationships, leveraged
capital, earned revenue %, SROI placeholder.

**Permissions.** Tenant-restricted. Edit: Director. Read: JSC.

### Columns (in addition to common columns)

| Column name | Type | Required | Default | Choices / validation |
|---|---|---|---|---|
| `Funder_Name` | Single line of text | Yes | — | |
| `Activity_Type` | Choice | Yes | — | Grant; Partnership; Sponsorship; Leveraged Capital; SROI Calculation; Earned Revenue Snapshot |
| `Amount_$` | Currency | No | — | $ value where applicable |
| `Period_Start` | Date | No | — | Grant period start |
| `Period_End` | Date | No | — | Grant period end |
| `Earned_Revenue_%` | Number | No | — | Quarterly snapshot, 0–100 |
| `Leveraged_Capital_Source` | Single line of text | No | — | Description of matching grant/investment |
| `SROI_Ratio` | Number | No | — | Placeholder — leave blank until methodology exists |
| `Status` | Choice | No | Active | Active; Pending; Closed |

### Default views

- **All Items**
- **Active Grants** — filter `Activity_Type = Grant` and `Status = Active`
- **By Activity Type** — group by `Activity_Type`

---

## List 5 — Farmers List

**Purpose.** Farm relationships, produce moved, distribution, origins,
value-added participation, farmer demographics. **This is the UAC List.**

**Permissions.** **External sharing on.** Edit: UAC Partner (external),
Director. Read: JSC, Logistics Manager. See "External sharing setup"
below.

### Columns (in addition to common columns)

| Column name | Type | Required | Default | Choices / validation |
|---|---|---|---|---|
| `Farm_Name` | Single line of text | Yes | — | |
| `Farm_Location` | Single line of text | No | — | City, State |
| `Farm_Contact` | Single line of text | No | — | Name / email / phone |
| `Produce_Types` | Multiple lines of text | No | — | Comma-separated list of produce types this period |
| `Produce_lbs` | Number | No | 0 | Total lbs moved this period |
| `Produce_Value_$` | Currency | No | — | Total $ value moved this period |
| `Origin_Region` | Single line of text | No | — | Region of origin |
| `Distribution_Destinations` | Multiple lines of text | No | — | Comma-separated list of recipient orgs / locations |
| `Value_Added_Product` | Yes/No | No | No | True if farmer used HARVEST for value-added |
| `Value_Added_Description` | Single line of text | No | — | E.g., "Hot sauce from peppers" |
| `Demographic_Race_Ethnicity` | Choice (multi-select) | No | — | Black/African American; Hispanic/Latino; Asian; White; Native American/Indigenous; Pacific Islander; Multiracial; Prefer not to say |
| `Demographic_Gender` | Choice | No | — | Woman; Man; Non-binary; Prefer not to say |
| `Demographic_Age_Band` | Choice | No | — | Under 25; 25–34; 35–44; 45–54; 55–64; 65+; Prefer not to say |

### Default views

- **All Items**
- **Current Month** — filter on `Reporting_Period_Start`
- **By Farm** — group by `Farm_Name`
- **Value-Added Participants** — filter `Value_Added_Product = Yes`

### External sharing setup (Farmers List only)

This is the one piece of configuration that requires SharePoint admin
involvement.

1. Open the Farmers List → **Settings (gear) → List settings**.
2. Under **Permissions and Management**, click **Permissions for this list**.
3. Click **Stop Inheriting Permissions** → confirm.
4. Click **Share** → add the UAC Partner's external email address →
   set permission level to **Contribute** → uncheck "Send email" if you
   don't want a notification, otherwise leave it on.
5. If the share fails with a tenant-policy error, ask SharePoint admin
   to enable external sharing on the HARVEST site at the
   "Existing guests" or "New and existing guests" level. The Microsoft
   docs page on this is titled "Manage sharing settings for SharePoint
   and OneDrive in Microsoft 365".

The UAC Form (separate from this List) is configured for anonymous
public-link submission, which is the primary entry path. Direct List
edit access is a secondary path for corrections.

---

## List 6 — Environmental List

**Purpose.** Food waste diversion, composting, food miles, food captured
for nutritional demos.

**Permissions.** Tenant-restricted. Edit: Logistics Manager, Culinary
Manager. Read: JSC.

### Columns (in addition to common columns)

| Column name | Type | Required | Default | Choices / validation |
|---|---|---|---|---|
| `Activity_Type` | Choice | Yes | — | Food Waste Diverted; Composted; Food Captured for Demo; Food Miles Snapshot |
| `Activity_Date` | Date | Yes | — | |
| `lbs_Value` | Number | No | 0 | lbs (used for all activity types except Food Miles) |
| `Dollar_Value` | Currency | No | — | $ value (typically for Food Captured for Demo) |
| `Disposition` | Choice | No | — | Composted On-Site; Composted Off-Site; Donated; Used in Demo; Other |
| `Food_Miles_Avg` | Number | No | — | Avg miles (only for Food Miles Snapshot) |
| `Source_or_Recipient` | Single line of text | No | — | E.g., "Local City Compost", "On-site demo kitchen" |

### Default views

- **All Items**
- **Current Quarter** — filter on `Activity_Date`
- **By Activity Type** — group by `Activity_Type`

---

## Manual setup walkthrough — Path B (per List)

If you're not running the PowerShell script, do this once per List:

1. Go to the HARVEST SharePoint site → **+ New → List → Blank list**.
2. Name it exactly as the List name above (e.g., `Food Entrepreneurs List`).
3. Click **Create**.
4. **Add common columns** in the order shown in the Common Columns table
   above. For each:
   - Click **+ Add column** → pick the type → enter the name → set
     required / default per the table.
   - For `OFSA_Dimension` (multi-choice): click "More options" and
     enable **"Allow multiple selections"**.
5. **Add per-List columns** from the relevant section.
6. **Enable attachments**: List settings → Advanced settings →
   Attachments → **Enabled** → OK.
7. **Create default views**: All Items → **+ Create new view** → pick
   Grid view → name it per the "Default views" section → set filters /
   grouping accordingly.
8. **Set permissions** per the section above. For the Farmers List,
   follow the External sharing setup steps.
9. Repeat for all six Lists.

Total time, hand-built: ~2 hours for all six Lists.
