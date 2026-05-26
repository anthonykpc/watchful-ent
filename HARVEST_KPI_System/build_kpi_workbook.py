"""
Generates the HARVEST Food Hub KPI tracking system workbooks.

Produces:
  - 5 individual role-based log workbooks (logs/ folder)
  - 1 master dashboard workbook (HARVEST_KPI_Master.xlsx)

The log workbooks are where staff enter data directly. The master
workbook uses Power Query to pull from all five logs plus the Food
Corridor CSV, and renders dashboards.

Run:  python3 build_kpi_workbook.py
"""

from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    PatternFill,
    Side,
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.protection import SheetProtection

BASE_DIR = Path(__file__).parent
LOGS_DIR = BASE_DIR / "logs"
MASTER_FILE = BASE_DIR / "HARVEST_KPI_Master.xlsx"

DEEP_GREEN = "2D6A4F"
LIGHT_GREEN = "E8F3EC"
WHITE = "FFFFFF"
TEXT_DARK = "1B3A2B"
BORDER_GREEN = "B7D9C2"
STATUS_GREEN = "C6EFCE"
STATUS_YELLOW = "FFEB9C"
STATUS_RED = "FFC7CE"


# ---------------------------------------------------------------------------
# Style helpers
# ---------------------------------------------------------------------------

def header_font():
    return Font(name="Calibri", size=11, bold=True, color=WHITE)


def header_fill():
    return PatternFill("solid", fgColor=DEEP_GREEN)


def body_font():
    return Font(name="Calibri", size=11, color=TEXT_DARK)


def thin_border():
    side = Side(style="thin", color=BORDER_GREEN)
    return Border(left=side, right=side, top=side, bottom=side)


def write_header_row(ws, row, headers, widths=None):
    for idx, label in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=idx, value=label)
        cell.font = header_font()
        cell.fill = header_fill()
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        cell.border = thin_border()
        if widths and idx - 1 < len(widths):
            ws.column_dimensions[get_column_letter(idx)].width = widths[idx - 1]
    ws.row_dimensions[row].height = 32


def add_table(ws, name, ref):
    tbl = Table(displayName=name, ref=ref)
    style = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    tbl.tableStyleInfo = style
    ws.add_table(tbl)


def add_dropdown(ws, col_letter, options, start_row=2, end_row=500):
    dv = DataValidation(
        type="list",
        formula1='"{}"'.format(",".join(options)),
        allow_blank=True,
        showDropDown=False,
    )
    dv.error = "Pick a value from the dropdown."
    dv.errorTitle = "Invalid entry"
    ws.add_data_validation(dv)
    dv.add(f"{col_letter}{start_row}:{col_letter}{end_row}")


def write_instructions(ws, title, lines):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 100
    t = ws.cell(row=2, column=2, value=title)
    t.font = Font(name="Calibri", size=16, bold=True, color=WHITE)
    t.fill = header_fill()
    t.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 34
    for i, line in enumerate(lines, start=4):
        c = ws.cell(row=i, column=2, value=line)
        c.font = body_font()
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[i].height = max(20, 18 * (line.count("\n") + 1))


# ---------------------------------------------------------------------------
# Log workbook specs
# ---------------------------------------------------------------------------

LOG_SPECS = {
    "Kitchen_Operations_Log": {
        "file": "Kitchen_Operations_Log.xlsx",
        "owner": "Logistics Manager",
        "cadence": "Weekly",
        "table_name": "tbl_KitchenOps",
        "instructions_title": "Kitchen Operations Log",
        "instructions": [
            "Owner: Logistics Manager",
            "Cadence: one row per week, plus one row per incident as they happen.",
            "Kitchen rental hours come from the Food Corridor CSV — do NOT enter them here.",
            "",
            "How to use:",
            "  1. Open this file from SharePoint/OneDrive (not your Downloads folder).",
            "  2. Scroll to the next empty row in the 'Log' tab.",
            "  3. Fill in the columns. Use the dropdowns where provided.",
            "  4. Save. The master workbook picks up your data on its next refresh.",
            "",
            "Food waste and composting entries here automatically feed the Environmental dashboard.",
        ],
        "headers": [
            "Date", "Reporting_Week_Start", "Reporting_Week_End",
            "Entry_Type", "Active_Members", "New_Members",
            "New_Member_Business", "New_Member_Demographics",
            "Equipment_Status", "Equipment_Notes",
            "Incident_Type", "Incident_Description",
            "Food_Waste_Diverted_lbs", "Composted_lbs",
            "Notes",
        ],
        "widths": [
            12, 14, 14,
            16, 14, 12,
            24, 26,
            18, 30,
            16, 30,
            18, 14,
            36,
        ],
        "dropdowns": {
            "D": ["Weekly Log", "Incident", "New Member Onboarding"],
            "H": ["Woman-owned", "BIPOC-owned", "Immigrant-owned", "Veteran-owned", "LGBTQ+-owned", "First-time entrepreneur"],
            "I": ["All operational", "Minor issue (logged)", "Major issue (offline)"],
            "K": ["Safety", "Sanitation", "Member conflict", "Other"],
        },
        "samples": [
            [date(2026, 3, 21), date(2026, 3, 15), date(2026, 3, 21),
             "Weekly Log", 2, 1, "Newark Eats LLC", "Woman-owned; BIPOC-owned",
             "All operational", None, None, None, 85, 60,
             "First full week with founding cohort."],
            [date(2026, 3, 28), date(2026, 3, 22), date(2026, 3, 28),
             "Weekly Log", 2, 1, "Casa Sabor", "Immigrant-owned; First-time entrepreneur",
             "All operational", None, None, None, 72, 50,
             "Second member onboarded."],
            [date(2026, 4, 4), date(2026, 3, 29), date(2026, 4, 4),
             "Weekly Log", 2, 0, None, None,
             "Minor issue (logged)", "Walk-in cooler thermostat replaced",
             None, None, 90, 65, None],
        ],
    },
    "Culinary_Training_Log": {
        "file": "Culinary_Training_Log.xlsx",
        "owner": "Culinary Manager",
        "cadence": "Per workshop / training / TA session",
        "table_name": "tbl_Culinary",
        "instructions_title": "Culinary & Training Log",
        "instructions": [
            "Owner: Culinary Manager",
            "Cadence: one row per workshop, training, or 1:1 TA session. Submit within 48 hours.",
            "",
            "New fields in this version:",
            "  - Curriculum: name of the curriculum used for health demos (e.g., 'Diabetes-Friendly Cooking').",
            "  - RD_InKind_Hours: hours the Registered Dietitian provided demos at no charge.",
            "    This feeds the in-kind tracking KPI on the master dashboard.",
            "",
            "How to use:",
            "  1. Open this file from SharePoint/OneDrive.",
            "  2. Scroll to the next empty row in the 'Log' tab.",
            "  3. Fill in the columns. Most have dropdowns.",
            "  4. Save.",
        ],
        "headers": [
            "Date", "Event_Type", "Event_Title", "Curriculum",
            "Duration_Hours", "Attendance_Count",
            "Businesses_Present",
            "ServSafe_Passed", "ServSafe_Attempted",
            "TA_Topic", "TA_Recipient_Business",
            "Wellness_Coaching_Recipient",
            "RD_InKind_Hours",
            "Food_Captured_Demo_lbs", "Food_Captured_Demo_$",
            "Notes",
        ],
        "widths": [
            12, 24, 28, 28,
            14, 14,
            28,
            14, 14,
            22, 24,
            24,
            14,
            18, 18,
            36,
        ],
        "dropdowns": {
            "B": [
                "ServSafe Training", "Health/Nutrition Demo",
                "Business Building Seminar", "Marketing Seminar",
                "Capital Access Seminar", "1:1 Technical Assistance",
                "Wellness Coaching", "Other Workshop",
            ],
            "D": [
                "Diabetes-Friendly Cooking", "Heart Health Meals",
                "Cooking with Local Produce", "Meal Prep on a Budget",
                "Infant & Toddler Nutrition", "Cultural Heritage Cooking",
                "Food Safety Basics", "Other (specify in Notes)",
            ],
            "J": [
                "Business building", "Capital access",
                "Marketing", "Operations", "Other",
            ],
        },
        "samples": [
            [date(2026, 3, 18), "ServSafe Training", "ServSafe Food Handler — Cohort 1", None,
             4, 2, "Newark Eats LLC; Casa Sabor",
             2, 2, None, None, None, 0, 0, 0,
             "Both founding members passed."],
            [date(2026, 4, 5), "Health/Nutrition Demo", "Teaching Kitchen — Diabetes Friendly", "Diabetes-Friendly Cooking",
             2.5, 18, None,
             0, 0, None, None, None, 2.5,
             12, 45,
             "RD-led; surplus produce from UAC. 12-week series start."],
            [date(2026, 4, 10), "1:1 Technical Assistance", "Marketing TA — Newark Eats", None,
             1, 1, "Newark Eats LLC",
             0, 0, "Marketing", "Newark Eats LLC", None, 0, 0, 0,
             "Social media strategy session."],
            [date(2026, 4, 14), "Health/Nutrition Demo", "Heart Health Cooking Class", "Heart Health Meals",
             2, 22, None,
             0, 0, None, None, None, 3,
             8, 30,
             "Partnership with RWJBH cardiology outreach."],
        ],
    },
    "Community_Engagement_Log": {
        "file": "Community_Engagement_Log.xlsx",
        "owner": "Community Outreach Specialist",
        "cadence": "Per event",
        "table_name": "tbl_CommunityEvents",
        "instructions_title": "Community Engagement Log",
        "instructions": [
            "Owner: Community Outreach Specialist",
            "Cadence: one row per event. Submit within 48 hours.",
            "",
            "This workbook has two tabs:",
            "  - Log: one row per event (attendance, orgs, food recovery, etc.)",
            "  - Survey Responses: one row per survey question per event.",
            "    Add new survey questions as new rows — the structure never needs to change.",
            "",
            "How to use:",
            "  1. Open from SharePoint/OneDrive.",
            "  2. Add a row to 'Log' for the event.",
            "  3. If you collected survey data, add rows to 'Survey Responses' — one row per question.",
            "  4. Save.",
        ],
        "headers": [
            "Date", "Event_Name", "Event_Type",
            "Attendance_Count", "Volunteer_Hours",
            "Orgs_Engaged", "Sector_Count",
            "Multilingual_Cultural", "Language_Tradition",
            "Food_Recovery_lbs", "Food_Recovery_Recipient",
            "DSP_Onsite",
            "Notes",
        ],
        "widths": [
            12, 28, 22,
            14, 14,
            36, 10,
            18, 22,
            14, 22,
            12,
            36,
        ],
        "dropdowns": {
            "C": [
                "Teaching Kitchen", "RD Workshop", "Tour",
                "Partner-Hosted Event", "HARVEST-Hosted Event",
                "Pop-Up", "Other",
            ],
            "H": ["Yes", "No"],
            "L": ["Yes", "No"],
        },
        "samples": [
            [date(2026, 3, 22), "Open House — March", "HARVEST-Hosted Event",
             62, 8.5, "Bridges; Newark Public Library; La Casa", 3,
             "No", None, 24, "Bridges", "Yes",
             "First public open house."],
            [date(2026, 4, 5), "Teaching Kitchen — Diabetes Friendly", "Teaching Kitchen",
             18, 4, "RWJBH Community Health", 1,
             "No", None, 0, None, "No",
             "RD-led; 12-week series start."],
            [date(2026, 4, 19), "Latine Food Heritage Night", "Pop-Up",
             95, 12, "La Casa; Lincoln Park Coast Cultural District", 2,
             "Yes", "Spanish", 35, "Bridges", "No",
             "Member businesses showcased."],
        ],
        "extra_tabs": ["survey"],
    },
    "Director_Partnership_Log": {
        "file": "Director_Partnership_Log.xlsx",
        "owner": "Director",
        "cadence": "Monthly + event-based for grants and partnerships",
        "table_name": "tbl_Director",
        "instructions_title": "Director & Partnership Log",
        "instructions": [
            "Owner: Director",
            "Cadence: monthly governance roll-up, plus one row per grant/partnership/connection as they happen.",
            "",
            "Entry types — use the dropdown to classify each row:",
            "  - Grant: a new grant secured. Include $ amount and funder.",
            "  - Partnership: a formal partnership executed.",
            "  - Sponsorship: a sponsorship secured.",
            "  - Leveraged Capital: matching grants, investments, in-kind.",
            "  - Institutional Connection: new relationship with Rutgers, NPS, Audible, etc.",
            "  - Purchase Order: institutional PO for HARVEST/UAC products.",
            "  - Menu Placement: HARVEST items on an institutional menu.",
            "  - TA Partnership: an institution providing TA to members.",
            "  - JSC Meeting: Joint Steering Committee meeting.",
            "  - CAB Meeting: Community Advisory Board meeting.",
            "  - CAB Member Added: new CAB member recruited.",
            "  - Professional Development: conference, cert, presentation.",
            "  - Earned Revenue Snapshot: quarterly earned-revenue %.",
            "",
            "How to use: open from SharePoint/OneDrive, add rows, save.",
        ],
        "headers": [
            "Date", "Entry_Type", "Name",
            "Amount_$", "Period_Start", "Period_End",
            "Institution_Type", "Items_On_Menu",
            "CAB_Member_Name", "CAB_Member_Sector",
            "PD_Activity",
            "Earned_Revenue_%", "Avg_Sales_Increase_%",
            "Status", "Notes",
        ],
        "widths": [
            12, 24, 30,
            14, 12, 12,
            18, 14,
            22, 22,
            28,
            14, 14,
            12, 36,
        ],
        "dropdowns": {
            "B": [
                "Grant", "Partnership", "Sponsorship", "Leveraged Capital",
                "Institutional Connection", "Purchase Order",
                "Menu Placement", "TA Partnership",
                "JSC Meeting", "CAB Meeting", "CAB Member Added",
                "Professional Development", "Earned Revenue Snapshot",
            ],
            "G": ["Healthcare", "Education", "Government", "Corporate", "Other"],
            "J": [
                "Healthcare", "Agriculture", "Education", "Government",
                "Philanthropy", "Community-Based Org", "Private Sector",
                "Food Enterprise",
            ],
            "K": [
                "Shared Kitchen Summit", "ServSafe Manager Cert",
                "Community Presentation", "Conference", "Other",
            ],
            "N": ["Active", "Pending", "Closed"],
        },
        "samples": [
            [date(2026, 2, 12), "Grant", "Whole Cities Foundation — 2026",
             75000, date(2026, 2, 1), date(2026, 12, 31),
             None, None, None, None, None, None, None,
             "Active", "Anchor 2026 grant."],
            [date(2026, 1, 18), "JSC Meeting", "JSC — January 2026",
             None, None, None, None, None, None, None, None, None, None,
             "Closed", "Reviewed Q4 2025 ops; approved 2026 KPI framework."],
            [date(2026, 3, 14), "JSC Meeting", "JSC — March 2026",
             None, None, None, None, None, None, None, None, None, None,
             "Closed", "Founding cohort selection."],
            [date(2026, 3, 28), "CAB Member Added", "D. Okafor",
             None, None, None, None, None,
             "D. Okafor", "Healthcare", None, None, None,
             "Active", "Healthcare sector."],
            [date(2026, 4, 4), "CAB Member Added", "M. Vega",
             None, None, None, None, None,
             "M. Vega", "Food Enterprise", None, None, None,
             "Active", "Food Enterprise sector."],
            [date(2026, 3, 31), "Earned Revenue Snapshot", "Q1 2026",
             None, date(2026, 1, 1), date(2026, 3, 31),
             None, None, None, None, None, 18, None,
             "Closed", "Mostly rentals; expect to rise as cohort scales."],
            [date(2026, 3, 31), "Institutional Connection", "Rutgers New Brunswick",
             None, None, None, "Education", None, None, None, None, None, None,
             "Active", "Joint research interest on food-as-medicine outcomes. Lead: Dr. Patel."],
            [date(2026, 3, 31), "Purchase Order", "RWJBH Newark — Q1 PO",
             8500, date(2026, 1, 1), date(2026, 3, 31),
             "Healthcare", None, None, None, None, None, None,
             "Closed", "First institutional PO."],
            [date(2026, 4, 22), "Professional Development", "ServSafe Manager Cert — Director",
             None, None, None, None, None, None, None,
             "ServSafe Manager Cert", None, None,
             "Closed", "Director ServSafe Manager Cert — passed."],
        ],
    },
    "UAC_Farmer_Engagement_Log": {
        "file": "UAC_Farmer_Engagement_Log.xlsx",
        "owner": "Director / designated HARVEST staff (on behalf of UAC)",
        "cadence": "Monthly",
        "table_name": "tbl_Farmers",
        "instructions_title": "UAC Farmer Engagement Log",
        "instructions": [
            "Owner: Director or designated HARVEST staff member",
            "Cadence: monthly, by the 5th of each month for the prior month.",
            "",
            "HARVEST staff enters this data on behalf of UAC partners.",
            "Source: monthly call/email with UAC contact, who provides:",
            "  - Which farms were active, what produce moved, volumes, and $ values.",
            "  - Distribution destinations.",
            "  - Any new value-added products.",
            "  - Farmer demographic updates (ask once, update as needed).",
            "",
            "One row per farm per month. If 4 farms were active in March, that's 4 rows.",
            "",
            "How to use: open from SharePoint/OneDrive, add rows, save.",
        ],
        "headers": [
            "Date", "Reporting_Month", "Entered_By",
            "Farm_Name", "Farm_Location", "Farm_Contact",
            "Produce_Types", "Produce_lbs", "Produce_Value_$",
            "Origin_Region", "Distribution_Destinations",
            "Value_Added_Product", "Value_Added_Description",
            "Demographic_Race_Ethnicity", "Demographic_Gender",
            "Demographic_Age_Band", "Notes",
        ],
        "widths": [
            12, 14, 18,
            22, 22, 24,
            28, 12, 14,
            22, 32,
            14, 26,
            26, 16, 14, 28,
        ],
        "dropdowns": {
            "L": ["Yes", "No"],
            "N": [
                "Black/African American", "Hispanic/Latino", "Asian",
                "White", "Native American/Indigenous", "Pacific Islander",
                "Multiracial", "Prefer not to say",
            ],
            "O": ["Woman", "Man", "Non-binary", "Prefer not to say"],
            "P": [
                "Under 25", "25-34", "35-44",
                "45-54", "55-64", "65+", "Prefer not to say",
            ],
        },
        "samples": [
            [date(2026, 4, 3), "March 2026", "A. Capece (Director)",
             "Greenfields Farm", "Sussex County, NJ", "rowens@greenfields.example",
             "kale; collards; turnips", 480, 1200,
             "Sussex County, NJ", "RWJBH Newark cafeteria; Bridges",
             "No", None,
             "Black/African American", "Man", "45-54", None],
            [date(2026, 4, 3), "March 2026", "A. Capece (Director)",
             "Sunrise Acres", "Hunterdon County, NJ", "info@sunriseacres.example",
             "tomatoes; peppers", 320, 980,
             "Hunterdon County, NJ", "Newark Public Schools pilot",
             "Yes", "Pepper hot sauce — Newark Eats LLC",
             "Hispanic/Latino", "Woman", "35-44",
             "Value-added through Newark Eats LLC."],
            [date(2026, 5, 2), "April 2026", "A. Capece (Director)",
             "Liberty Roots", "Essex County, NJ", "liberty@roots.example",
             "salad greens; radishes", 210, 640,
             "Essex County, NJ", "RWJBH Newark cafeteria",
             "No", None,
             "Black/African American", "Woman", "25-34",
             "New farmer this period."],
            [date(2026, 5, 2), "April 2026", "A. Capece (Director)",
             "Hudson Hill Farm", "Warren County, NJ", "hudson@hill.example",
             "carrots; beets; squash", 540, 1450,
             "Warren County, NJ", "Bridges; Newark Public Schools pilot",
             "No", None,
             "White", "Man", "55-64", None],
        ],
    },
}

SURVEY_HEADERS = [
    "Date", "Event_Name", "Survey_Instrument",
    "Question_Text", "Response_Text", "Numeric_Score",
    "Notes",
]
SURVEY_WIDTHS = [12, 28, 22, 40, 40, 12, 28]
SURVEY_SAMPLES = [
    [date(2026, 3, 22), "Open House — March", "Post-Event v1",
     "Overall satisfaction (1-5)", None, 4.5, None],
    [date(2026, 3, 22), "Open House — March", "Post-Event v1",
     "I learned something new today (1-5)", None, 4.3, None],
    [date(2026, 3, 22), "Open House — March", "Post-Event v1",
     "What did you enjoy most?", "The cooking demo and trying the food", None, None],
    [date(2026, 4, 5), "Teaching Kitchen — Diabetes Friendly", "Teaching Kitchen Series v1",
     "Overall satisfaction (1-5)", None, 4.7, None],
    [date(2026, 4, 5), "Teaching Kitchen — Diabetes Friendly", "Teaching Kitchen Series v1",
     "I feel confident I can make this recipe at home (1-5)", None, 4.6, None],
    [date(2026, 4, 5), "Teaching Kitchen — Diabetes Friendly", "Teaching Kitchen Series v1",
     "What would you like to learn next?", "More low-sugar desserts", None, None],
    [date(2026, 4, 19), "Latine Food Heritage Night", "Post-Event v1",
     "Overall satisfaction (1-5)", None, 4.8, None],
    [date(2026, 4, 19), "Latine Food Heritage Night", "Post-Event v1",
     "I learned something new today (1-5)", None, 4.5, None],
]


# ---------------------------------------------------------------------------
# Build individual log workbooks
# ---------------------------------------------------------------------------

def build_log_workbook(spec):
    wb = Workbook()

    # Instructions tab
    ws_instr = wb.active
    ws_instr.title = "How to Use"
    write_instructions(
        ws_instr,
        f"HARVEST — {spec['instructions_title']}",
        spec["instructions"],
    )

    # Log tab
    ws = wb.create_sheet("Log")
    header_row = 1
    write_header_row(ws, header_row, spec["headers"], spec["widths"])

    # Sample data
    for ridx, row in enumerate(spec["samples"], start=header_row + 1):
        for cidx, val in enumerate(row, start=1):
            c = ws.cell(row=ridx, column=cidx, value=val)
            c.font = body_font()
            c.alignment = Alignment(vertical="top", wrap_text=True)
            if isinstance(val, date):
                c.number_format = "yyyy-mm-dd"
            if spec["headers"][cidx - 1].endswith("_$"):
                if isinstance(val, (int, float)) and val:
                    c.number_format = "$#,##0"

    last_row = header_row + len(spec["samples"])
    last_col = get_column_letter(len(spec["headers"]))
    add_table(ws, spec["table_name"], f"A{header_row}:{last_col}{last_row}")

    # Dropdowns
    for col_letter, options in spec.get("dropdowns", {}).items():
        add_dropdown(ws, col_letter, options)

    ws.freeze_panes = "A2"
    ws.sheet_view.tabSelected = False
    wb.active = 1  # Open on Log tab

    # Survey Responses tab (Community Engagement Log only)
    if "survey" in spec.get("extra_tabs", []):
        ws_survey = wb.create_sheet("Survey Responses")
        write_header_row(ws_survey, 1, SURVEY_HEADERS, SURVEY_WIDTHS)
        for ridx, row in enumerate(SURVEY_SAMPLES, start=2):
            for cidx, val in enumerate(row, start=1):
                c = ws_survey.cell(row=ridx, column=cidx, value=val)
                c.font = body_font()
                c.alignment = Alignment(vertical="top", wrap_text=True)
                if isinstance(val, date):
                    c.number_format = "yyyy-mm-dd"
        last_row_s = 1 + len(SURVEY_SAMPLES)
        last_col_s = get_column_letter(len(SURVEY_HEADERS))
        add_table(ws_survey, "tbl_SurveyResponses",
                  f"A1:{last_col_s}{last_row_s}")
        ws_survey.freeze_panes = "A2"

    out_path = LOGS_DIR / spec["file"]
    wb.save(out_path)
    print(f"  {out_path.name}")


# ---------------------------------------------------------------------------
# KPI catalog (drives master workbook)
# ---------------------------------------------------------------------------

KPI_CATALOG_HEADERS = [
    "KPI", "Definition", "Unit", "Audience", "OFSA Dimension(s)",
    "Food System Sector", "Partner Sector", "Cadence", "Lead Role",
    "Data Source", "2026 Target", "Owner Role", "Source Log",
    "Status",
]
KPI_CATALOG_WIDTHS = [38, 50, 8, 22, 26, 22, 22, 12, 14, 22, 12, 26, 28, 22]

# (KPI, Definition, Unit, Audience, OFSA_Dims, Sector, Partner, Cadence,
#  Lead, DataSource, Target2026, OwnerRole, SourceLog, Status)
KPI_CATALOG = [
    # Food Entrepreneurs
    ("Entrepreneurs engaged", "Unique individuals reached via outreach, intake, or info session", "count", "Food Entrepreneurs", "Agency", "Processing", "Food Enterprise", "Monthly", "HARVEST Lead", "Community Engagement Log", 40, "Community Outreach Specialist", "Community_Engagement_Log.xlsx", "Active"),
    ("Entrepreneurs converted to renting members", "Engaged entrepreneurs who signed a kitchen rental agreement", "count", "Food Entrepreneurs", "Agency; Availability", "Processing", "Food Enterprise", "Monthly", "HARVEST Lead", "Kitchen Operations Log", 8, "Logistics Manager", "Kitchen_Operations_Log.xlsx", "Active"),
    ("% members achieving ServSafe", "Members who pass ServSafe via HARVEST", "%", "Food Entrepreneurs", "Utilization", "Processing", "Food Enterprise", "Per training", "HARVEST Lead", "Culinary & Training Log", 100, "Culinary Manager", "Culinary_Training_Log.xlsx", "Active"),
    ("Avg % sales increase for members", "YoY revenue change for members in kitchen >=6 months", "%", "Food Entrepreneurs", "Stability; Agency", "Processing", "Food Enterprise", "Quarterly", "HARVEST Lead", "Director & Partnership Log", None, "Director", "Director_Partnership_Log.xlsx", "Active"),
    ("Businesses supported", "Distinct member businesses receiving any HARVEST service", "count", "Food Entrepreneurs", "Agency", "Processing", "Food Enterprise", "Monthly", "HARVEST Lead", "System-calculated", 8, "Director", "Multiple logs", "Active"),
    ("Workforce dev seminars held", "Seminars on business building, food safety, marketing, capital", "count", "Food Entrepreneurs", "Utilization; Agency", "Processing", "Education", "Per event", "HARVEST Lead", "Culinary & Training Log", 6, "Culinary Manager", "Culinary_Training_Log.xlsx", "Active"),
    ("Kitchen hours rented", "Total billable kitchen hours in the period", "hours", "Food Entrepreneurs", "Availability; Stability", "Processing", "Food Enterprise", "Real-time", "HARVEST Lead", "Food Corridor CSV", 1500, "Logistics Manager", "Food Corridor CSV", "Active"),
    ("Jobs created by member businesses", "FTEs and PTEs hired (self-reported)", "count", "Food Entrepreneurs", "Stability; Agency", "Processing", "Food Enterprise", "Quarterly", "HARVEST Lead", "Director & Partnership Log", None, "Director", "Director_Partnership_Log.xlsx", "Active"),
    ("Businesses buying local produce via UAC", "Members sourcing produce from UAC channels", "count", "Food Entrepreneurs", "Access; Sustainability", "Aggregation", "Agriculture", "Monthly", "Joint", "UAC Farmer Engagement Log", 3, "Director", "UAC_Farmer_Engagement_Log.xlsx", "Active"),
    ("Events showcasing local entrepreneurs", "Public events featuring member businesses", "count", "Food Entrepreneurs", "Agency", "Retail-Consumption", "Food Enterprise", "Per event", "HARVEST Lead", "Community Engagement Log", 4, "Community Outreach Specialist", "Community_Engagement_Log.xlsx", "Active"),
    ("Technical assistance sessions", "1:1 sessions for business building, capital, marketing", "count", "Food Entrepreneurs", "Utilization; Agency", "Processing", "Food Enterprise", "Per session", "HARVEST Lead", "Culinary & Training Log", 30, "Culinary Manager", "Culinary_Training_Log.xlsx", "Active"),
    ("Businesses receiving wellness coaching", "Member businesses in wellness/nutrition coaching", "count", "Food Entrepreneurs", "Utilization", "Processing", "Healthcare", "Per session", "HARVEST Lead", "Culinary & Training Log", None, "Culinary Manager", "Culinary_Training_Log.xlsx", "Active"),
    # Institution
    ("$ purchasing orders from HARVEST/UAC vendors", "Total $ value of POs by institutions", "$", "Institution", "Access; Availability", "Distribution", "Healthcare", "Monthly", "Joint", "Director & Partnership Log", None, "Director", "Director_Partnership_Log.xlsx", "Active"),
    ("HARVEST-produced items on institutional menus", "Distinct items on institutional menus", "count", "Institution", "Access; Availability", "Retail-Consumption", "Healthcare", "Quarterly", "HARVEST Lead", "Director & Partnership Log", None, "Director", "Director_Partnership_Log.xlsx", "Active"),
    ("Cross-institution connections", "Active relationships with Rutgers, NPS, etc.", "count", "Institution", "Stability; Access", "n/a", "Education", "Quarterly", "HARVEST Lead", "Director & Partnership Log", 4, "Director", "Director_Partnership_Log.xlsx", "Active"),
    ("TA partnerships", "External institutions providing TA to members", "count", "Institution", "Utilization; Stability", "n/a", "Education", "Quarterly", "Joint", "Director & Partnership Log", 3, "Director", "Director_Partnership_Log.xlsx", "Active"),
    ("Reduction in readmission costs", "RWJBH readmission $ reduction from HARVEST", "$", "Institution", "Utilization; Access", "Retail-Consumption", "Healthcare", "Annual", "Joint", "Manual entry", None, "Director", "Director_Partnership_Log.xlsx", "Planned — methodology TBD"),
    # Community Organizations
    ("Community orgs engaged", "Distinct orgs participating in any HARVEST activity", "count", "Community Organizations", "Agency; Access", "n/a", "Community-Based Org", "Monthly", "HARVEST Lead", "Community Engagement Log", 15, "Community Outreach Specialist", "Community_Engagement_Log.xlsx", "Active"),
    ("Event participants", "Total participants across events", "count", "Community Organizations", "Utilization; Agency", "n/a", "Community-Based Org", "Per event", "HARVEST Lead", "Community Engagement Log", 300, "Community Outreach Specialist", "Community_Engagement_Log.xlsx", "Active"),
    ("Community volunteer hours", "Hours from community volunteers at HARVEST", "hours", "Community Organizations", "Agency; Stability", "n/a", "Community-Based Org", "Per event", "HARVEST Lead", "Community Engagement Log", None, "Community Outreach Specialist", "Community_Engagement_Log.xlsx", "Active"),
    ("Community events run by staff", "Events run by HARVEST staff", "count", "Community Organizations", "Agency", "n/a", "Community-Based Org", "Per event", "HARVEST Lead", "Community Engagement Log", 24, "Community Outreach Specialist", "Community_Engagement_Log.xlsx", "Active"),
    ("Survey feedback", "Row-per-question survey data (satisfaction, learning, open-ended)", "mixed", "Community Organizations", "Utilization", "n/a", "Community-Based Org", "Per event", "HARVEST Lead", "Community Engagement Log", None, "Community Outreach Specialist", "Community_Engagement_Log.xlsx", "Active"),
    ("Tours given", "Facility tours for external visitors", "count", "Community Organizations", "Agency", "n/a", "Community-Based Org", "Per event", "HARVEST Lead", "Community Engagement Log", 24, "Community Outreach Specialist", "Community_Engagement_Log.xlsx", "Active"),
    ("Orgs hosting events at HARVEST", "External orgs using HARVEST as venue", "count", "Community Organizations", "Agency; Access", "n/a", "Community-Based Org", "Per event", "HARVEST Lead", "Community Engagement Log", 6, "Community Outreach Specialist", "Community_Engagement_Log.xlsx", "Active"),
    ("Sectors engaged at HARVEST", "Distinct partner sectors across engagements", "count", "Community Organizations", "Stability", "n/a", "Community-Based Org", "Quarterly", "HARVEST Lead", "System-calculated", 5, "Director", "Community_Engagement_Log.xlsx", "Active"),
    ("Multilingual / cultural events", "Events in non-English language or cultural tradition", "count", "Community Organizations", "Agency; Access", "n/a", "Community-Based Org", "Per event", "HARVEST Lead", "Community Engagement Log", 6, "Community Outreach Specialist", "Community_Engagement_Log.xlsx", "Active"),
    ("Food recovery to service providers (lbs)", "lbs diverted to Bridges and others", "lbs", "Community Organizations", "Access; Sustainability", "Recovery-Waste", "Community-Based Org", "Per event", "HARVEST Lead", "Community Engagement Log", 500, "Community Outreach Specialist", "Community_Engagement_Log.xlsx", "Active"),
    ("DSPs with presence at HARVEST", "Direct service providers with recurring presence", "count", "Community Organizations", "Access", "n/a", "Community-Based Org", "Quarterly", "HARVEST Lead", "Community Engagement Log", 3, "Community Outreach Specialist", "Community_Engagement_Log.xlsx", "Active"),
    # Investors / Funders
    ("Grants and partnerships secured", "New grants or partnerships executed", "count", "Investors-Funders", "Stability", "n/a", "Philanthropy", "Per event", "HARVEST Lead", "Director & Partnership Log", 4, "Director", "Director_Partnership_Log.xlsx", "Active"),
    ("% earned revenue", "Earned revenue as % of total revenue", "%", "Investors-Funders", "Stability; Sustainability", "Processing", "Philanthropy", "Quarterly", "HARVEST Lead", "Director & Partnership Log", None, "Director", "Director_Partnership_Log.xlsx", "Active"),
    ("Funders and sponsors", "Distinct active funders/sponsors", "count", "Investors-Funders", "Stability", "n/a", "Philanthropy", "Quarterly", "HARVEST Lead", "Director & Partnership Log", None, "Director", "Director_Partnership_Log.xlsx", "Active"),
    ("Leveraged capital ($)", "Matching grants, investments, in-kind", "$", "Investors-Funders", "Stability; Sustainability", "n/a", "Philanthropy", "Quarterly", "HARVEST Lead", "Director & Partnership Log", None, "Director", "Director_Partnership_Log.xlsx", "Active"),
    ("Social return on investment (SROI)", "Ratio of social value to dollars invested", "ratio", "Investors-Funders", "Sustainability; Stability", "n/a", "Philanthropy", "Annual", "HARVEST Lead", "Manual entry", None, "Director", "Director_Partnership_Log.xlsx", "Planned — methodology TBD"),
    ("JSC meetings held", "Joint Steering Committee meetings convened", "count", "Investors-Funders", "Stability", "n/a", "Philanthropy", "Per event", "Joint", "Director & Partnership Log", 10, "Director", "Director_Partnership_Log.xlsx", "Active"),
    ("CAB members onboarded", "Community Advisory Board members across sectors", "count", "Investors-Funders", "Agency; Stability", "n/a", "Community-Based Org", "Per event", "HARVEST Lead", "Director & Partnership Log", 10, "Director", "Director_Partnership_Log.xlsx", "Active"),
    ("CAB meetings held", "CAB meetings convened", "count", "Investors-Funders", "Agency; Stability", "n/a", "Community-Based Org", "Per event", "HARVEST Lead", "Director & Partnership Log", 2, "Director", "Director_Partnership_Log.xlsx", "Active"),
    ("Director community presentations", "Community-facing presentations by Director", "count", "Investors-Funders", "Agency", "n/a", "Philanthropy", "Per event", "HARVEST Lead", "Director & Partnership Log", 4, "Director", "Director_Partnership_Log.xlsx", "Active"),
    # Farmers
    ("Farms engaged", "Distinct farms in active relationship", "count", "Farmers", "Agency; Stability", "Production", "Agriculture", "Monthly", "UAC Lead", "UAC Farmer Engagement Log", 4, "Director", "UAC_Farmer_Engagement_Log.xlsx", "Active"),
    ("$ farmer business via HARVEST", "Total $ farmer sales via HARVEST/UAC", "$", "Farmers", "Stability; Access", "Aggregation", "Agriculture", "Monthly", "UAC Lead", "UAC Farmer Engagement Log", None, "Director", "UAC_Farmer_Engagement_Log.xlsx", "Active"),
    ("Types of produce moved", "Distinct produce types moved", "count", "Farmers", "Availability", "Aggregation", "Agriculture", "Monthly", "UAC Lead", "UAC Farmer Engagement Log", None, "Director", "UAC_Farmer_Engagement_Log.xlsx", "Active"),
    ("lbs of produce moved", "Total lbs aggregated/distributed", "lbs", "Farmers", "Availability; Access", "Aggregation", "Agriculture", "Monthly", "UAC Lead", "UAC Farmer Engagement Log", None, "Director", "UAC_Farmer_Engagement_Log.xlsx", "Active"),
    ("Farmers making value-added products", "Farmers using HARVEST for value-added", "count", "Farmers", "Agency; Stability", "Processing", "Agriculture", "Monthly", "UAC Lead", "UAC Farmer Engagement Log", None, "Director", "UAC_Farmer_Engagement_Log.xlsx", "Active"),
    # Environmental
    ("Food waste diversion (lbs)", "lbs diverted from landfill (any pathway)", "lbs", "Environmental Sustainability", "Sustainability", "Recovery-Waste", "Community-Based Org", "Monthly", "HARVEST Lead", "Kitchen Operations Log", 1000, "Logistics Manager", "Kitchen_Operations_Log.xlsx", "Active"),
    ("Composted (lbs)", "lbs of food waste composted", "lbs", "Environmental Sustainability", "Sustainability", "Recovery-Waste", "Community-Based Org", "Monthly", "HARVEST Lead", "Kitchen Operations Log", None, "Logistics Manager", "Kitchen_Operations_Log.xlsx", "Active"),
    ("Local food miles (avg)", "Avg distance from source to HARVEST", "miles", "Environmental Sustainability", "Sustainability", "Distribution", "Agriculture", "Quarterly", "UAC Lead", "System-calculated", None, "Director", "UAC_Farmer_Engagement_Log.xlsx", "Active"),
    ("Food captured for demos (lbs)", "lbs used in demos instead of wasted", "lbs", "Environmental Sustainability", "Sustainability; Utilization", "Recovery-Waste", "Healthcare", "Per event", "HARVEST Lead", "Culinary & Training Log", None, "Culinary Manager", "Culinary_Training_Log.xlsx", "Active"),
    # NEW KPIs
    ("Health demo curriculums delivered", "Distinct curriculums delivered in health/nutrition demos", "count", "Food Entrepreneurs", "Utilization; Agency", "Processing", "Healthcare", "Per event", "HARVEST Lead", "Culinary & Training Log", None, "Culinary Manager", "Culinary_Training_Log.xlsx", "Active"),
    ("RD in-kind hours", "Hours Registered Dietitians provided free demos to community", "hours", "Community Organizations", "Utilization; Access", "n/a", "Healthcare", "Per event", "HARVEST Lead", "Culinary & Training Log", None, "Culinary Manager", "Culinary_Training_Log.xlsx", "Active"),
]


# ---------------------------------------------------------------------------
# Master workbook data tabs (seeded from log samples)
# ---------------------------------------------------------------------------

MASTER_DATA_TABS = [
    {
        "sheet": "Data — Kitchen Ops",
        "table": "tbl_KitchenOps",
        "source_log": "Kitchen_Operations_Log",
    },
    {
        "sheet": "Data — Culinary",
        "table": "tbl_Culinary",
        "source_log": "Culinary_Training_Log",
    },
    {
        "sheet": "Data — Community Events",
        "table": "tbl_CommunityEvents",
        "source_log": "Community_Engagement_Log",
    },
    {
        "sheet": "Data — Survey Responses",
        "table": "tbl_SurveyResponses",
        "source_log": None,  # built from SURVEY_SAMPLES
    },
    {
        "sheet": "Data — Director",
        "table": "tbl_Director",
        "source_log": "Director_Partnership_Log",
    },
    {
        "sheet": "Data — Farmers",
        "table": "tbl_Farmers",
        "source_log": "UAC_Farmer_Engagement_Log",
    },
]


# ---------------------------------------------------------------------------
# 2026 Performance Goals
# ---------------------------------------------------------------------------

GOALS_2026 = [
    ("1. Kitchen launch / operations", "Founding cohort — members converted",
     "Entrepreneurs converted to renting members", 8,
     '=SUMIFS(tbl_KitchenOps[New_Members],tbl_KitchenOps[Reporting_Week_End],">="&DATE(2026,1,1),tbl_KitchenOps[Reporting_Week_End],"<="&DATE(2026,12,31))',
     "Kitchen Operations Log"),
    ("1. Kitchen launch / operations", "ServSafe pass rate",
     "% members achieving ServSafe", 100,
     '=IFERROR(SUM(tbl_Culinary[ServSafe_Passed])/SUM(tbl_Culinary[ServSafe_Attempted])*100,0)',
     "Culinary & Training Log"),
    ("1. Kitchen launch / operations", "Kitchen rental hours",
     "Kitchen hours rented", 1500,
     '=SUM(tbl_FoodCorridor[Hours])',
     "Food Corridor CSV"),
    ("2. Structural documentation", "SOPs, guiding docs complete",
     "—", None, '="See shared drive folder"',
     "Tracked outside KPI system"),
    ("3. UAC partnership maintenance", "JSC meetings held", "JSC meetings held",
     10, '=COUNTIFS(tbl_Director[Entry_Type],"JSC Meeting")',
     "Director & Partnership Log"),
    ("3. UAC partnership maintenance", "Shared data framework operational",
     "Binary status", 1, '=1',
     "This workbook is live = 1"),
    ("3. UAC partnership maintenance", "UAC-affiliated farmers engaged",
     "Farms engaged", 4,
     '=IFERROR(ROWS(UNIQUE(FILTER(tbl_Farmers[Farm_Name],tbl_Farmers[Farm_Name]<>""))),0)',
     "UAC Farmer Engagement Log"),
    ("4. CAB development", "CAB members across sectors", "CAB members onboarded",
     10, '=COUNTIFS(tbl_Director[Entry_Type],"CAB Member Added")',
     "Director & Partnership Log"),
    ("4. CAB development", "CAB meetings held by year-end", "CAB meetings held",
     2, '=COUNTIFS(tbl_Director[Entry_Type],"CAB Meeting")',
     "Director & Partnership Log"),
    ("5. Professional development", "Shared Kitchen Summit attendance",
     "PD activity log", 1,
     '=COUNTIFS(tbl_Director[PD_Activity],"Shared Kitchen Summit")',
     "Director & Partnership Log"),
    ("5. Professional development", "ServSafe Manager Cert (Director)",
     "PD activity log", 1,
     '=COUNTIFS(tbl_Director[PD_Activity],"ServSafe Manager Cert")',
     "Director & Partnership Log"),
    ("5. Professional development", "Community presentations",
     "Director community presentations", 4,
     '=COUNTIFS(tbl_Director[PD_Activity],"Community Presentation")',
     "Director & Partnership Log"),
]

FOOD_CORRIDOR_HEADERS = ["Booking_Date", "Member_Business", "Hours"]
FOOD_CORRIDOR_WIDTHS = [16, 32, 12]
FOOD_CORRIDOR_SAMPLES = [
    [date(2026, 3, 12), "Newark Eats LLC", 6],
    [date(2026, 3, 14), "Casa Sabor", 4],
    [date(2026, 3, 19), "Newark Eats LLC", 8],
    [date(2026, 3, 26), "Casa Sabor", 5],
    [date(2026, 4, 2), "Newark Eats LLC", 6],
    [date(2026, 4, 9), "Casa Sabor", 4],
    [date(2026, 4, 16), "Newark Eats LLC", 6],
]


# ---------------------------------------------------------------------------
# Master workbook builders
# ---------------------------------------------------------------------------

def build_readme(wb):
    ws = wb.active
    ws.title = "README"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 110

    t = ws.cell(row=2, column=2, value="HARVEST Food Hub — KPI Master Workbook")
    t.font = Font(name="Calibri", size=20, bold=True, color=WHITE)
    t.fill = header_fill()
    t.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 38

    sub = ws.cell(row=3, column=2, value=(
        "Real-time KPI tracking across all audiences, tagged to the OFSA "
        "Six Dimensions. Pulls from individual staff log workbooks via Power Query."
    ))
    sub.font = Font(name="Calibri", size=11, italic=True, color=TEXT_DARK)
    sub.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[3].height = 22

    sections = [
        ("Open the right tab first",
         "- JSC meeting? Open '2026 Performance Goals' + 'Dashboard — Executive Summary'.\n"
         "- OFSA logic model? Open 'Dashboard — By OFSA Dimension'.\n"
         "- Funder report? Open 'Dashboard — By Audience' and filter.\n"
         "- What does a metric mean? Open 'KPI Catalog'."),
        ("Refresh the data",
         "1. Click the 'Data' tab in the Excel ribbon.\n"
         "2. Click 'Refresh All'.\n"
         "3. Wait for spinning to stop (~10-30 seconds).\n"
         "Power Query pulls from the five log workbooks in the logs/ folder."),
        ("Loading Food Corridor data",
         "1. Export the kitchen rentals CSV from Food Corridor.\n"
         "2. Open the 'Food Corridor Import' tab.\n"
         "3. Paste or refresh the CSV data.\n"
         "4. Hit 'Refresh All'. Kitchen hours KPI updates."),
        ("How it works",
         "Each staff member has their own Excel log workbook:\n"
         "  - Kitchen_Operations_Log.xlsx (Logistics Manager, weekly)\n"
         "  - Culinary_Training_Log.xlsx (Culinary Manager, per event)\n"
         "  - Community_Engagement_Log.xlsx (Community Outreach, per event)\n"
         "  - Director_Partnership_Log.xlsx (Director, monthly + event)\n"
         "  - UAC_Farmer_Engagement_Log.xlsx (Director enters for UAC, monthly)\n\n"
         "Staff open their log, add rows, save. This workbook reads from all five."),
        ("What's NOT here",
         "- SOPs / structural documentation (Goal 2) — tracked in shared drive.\n"
         "- Patient-level health data — out of scope.\n"
         "- Power BI — Excel is v1; Power BI is a later phase."),
    ]

    row = 5
    for heading, body in sections:
        h = ws.cell(row=row, column=2, value=heading)
        h.font = Font(name="Calibri", size=13, bold=True, color=DEEP_GREEN)
        h.alignment = Alignment(vertical="center")
        ws.row_dimensions[row].height = 22
        row += 1
        b = ws.cell(row=row, column=2, value=body)
        b.font = body_font()
        b.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = max(18 * (body.count("\n") + 1), 28)
        row += 2

    ws.protection = SheetProtection(sheet=True, password=None,
                                    selectLockedCells=False,
                                    selectUnlockedCells=False)


def build_kpi_catalog(wb):
    ws = wb.create_sheet("KPI Catalog")
    ws.sheet_view.showGridLines = False
    write_header_row(ws, 1, KPI_CATALOG_HEADERS, KPI_CATALOG_WIDTHS)

    for ridx, row in enumerate(KPI_CATALOG, start=2):
        for cidx, val in enumerate(row, start=1):
            c = ws.cell(row=ridx, column=cidx, value=val)
            c.font = body_font()
            c.alignment = Alignment(vertical="top", wrap_text=True)

    last_col = get_column_letter(len(KPI_CATALOG_HEADERS))
    last_row = 1 + len(KPI_CATALOG)
    add_table(ws, "tbl_KPI_Catalog", f"A1:{last_col}{last_row}")

    planned_fill = PatternFill("solid", fgColor=STATUS_YELLOW)
    ws.conditional_formatting.add(
        f"A2:{last_col}{last_row}",
        FormulaRule(
            formula=[f'$N2="Planned — methodology TBD"'],
            fill=planned_fill,
        ),
    )
    ws.freeze_panes = "A2"
    ws.protection = SheetProtection(sheet=True, password=None,
                                    selectLockedCells=False,
                                    selectUnlockedCells=False)


def build_goals_2026(wb):
    ws = wb.create_sheet("2026 Performance Goals")
    ws.sheet_view.showGridLines = False

    t = ws.cell(row=1, column=1, value="HARVEST — 2026 Performance Goals")
    t.font = Font(name="Calibri", size=18, bold=True, color=WHITE)
    t.fill = header_fill()
    t.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells("A1:G1")
    ws.row_dimensions[1].height = 36

    sub = ws.cell(row=2, column=1, value=(
        "Live status against the Director's 2026 annual review goals. "
        "Refresh (Data > Refresh All) before screenshotting."
    ))
    sub.font = Font(name="Calibri", size=11, italic=True, color=TEXT_DARK)
    sub.alignment = Alignment(vertical="center", indent=1)
    ws.merge_cells("A2:G2")
    ws.row_dimensions[2].height = 22

    headers = ["Goal", "Sub-goal", "Linked metric", "Target", "Current", "% Progress", "Status"]
    widths = [32, 38, 32, 12, 14, 14, 14]
    write_header_row(ws, 4, headers, widths)

    for ridx, (goal, sub, metric, target, formula, notes) in enumerate(GOALS_2026, start=5):
        ws.cell(row=ridx, column=1, value=goal).font = body_font()
        ws.cell(row=ridx, column=2, value=sub).font = body_font()
        ws.cell(row=ridx, column=3, value=metric).font = body_font()
        ws.cell(row=ridx, column=4, value=target).font = body_font()
        ws.cell(row=ridx, column=4).alignment = Alignment(horizontal="center")
        ws.cell(row=ridx, column=5, value=formula).font = body_font()
        ws.cell(row=ridx, column=5).alignment = Alignment(horizontal="center")

        if target:
            pct = ws.cell(row=ridx, column=6, value=f"=IFERROR(E{ridx}/D{ridx},0)")
            pct.number_format = "0%"
            pct.alignment = Alignment(horizontal="center")
            pct.font = body_font()
            st = ws.cell(row=ridx, column=7, value=(
                f'=IF(D{ridx}="","",IF(F{ridx}>=1,"On track",'
                f'IF(F{ridx}>=0.5,"In progress","At risk")))'
            ))
            st.alignment = Alignment(horizontal="center")
            st.font = body_font()
        else:
            ws.cell(row=ridx, column=7, value=notes).font = body_font()
            ws.cell(row=ridx, column=7).alignment = Alignment(horizontal="center", wrap_text=True)

        for c in range(1, 8):
            ws.cell(row=ridx, column=c).alignment = Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[ridx].height = 30

    last_row = 4 + len(GOALS_2026)
    for op, val, color in [
        ("equal", '"On track"', STATUS_GREEN),
        ("equal", '"In progress"', STATUS_YELLOW),
        ("equal", '"At risk"', STATUS_RED),
    ]:
        ws.conditional_formatting.add(
            f"G5:G{last_row}",
            CellIsRule(operator=op, formula=[val],
                       fill=PatternFill("solid", fgColor=color)),
        )
    ws.freeze_panes = "A5"
    ws.protection = SheetProtection(sheet=True, password=None,
                                    selectLockedCells=False,
                                    selectUnlockedCells=False)


def build_executive_summary(wb):
    ws = wb.create_sheet("Dashboard — Executive Summary")
    ws.sheet_view.showGridLines = False

    t = ws.cell(row=1, column=1, value="HARVEST — Executive Summary")
    t.font = Font(name="Calibri", size=18, bold=True, color=WHITE)
    t.fill = header_fill()
    t.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells("A1:F1")
    ws.row_dimensions[1].height = 36

    sub = ws.cell(row=2, column=1, value="Top-line metrics. Refresh: Data > Refresh All.")
    sub.font = Font(name="Calibri", size=11, italic=True, color=TEXT_DARK)
    sub.alignment = Alignment(vertical="center", indent=1)
    ws.merge_cells("A2:F2")

    tiles = [
        ("Active member businesses",
         '=IFERROR(ROWS(UNIQUE(FILTER(tbl_KitchenOps[New_Member_Business],tbl_KitchenOps[New_Member_Business]<>""))),0)',
         "0"),
        ("Kitchen hours rented (YTD)", "=SUM(tbl_FoodCorridor[Hours])", "0"),
        ("Workshops & TA sessions (YTD)",
         '=ROWS(FILTER(tbl_Culinary[Event_Type],tbl_Culinary[Event_Type]<>"",""))',
         "0"),
        ("Community event participants (YTD)",
         "=SUM(tbl_CommunityEvents[Attendance_Count])", "0"),
        ("Food recovery to providers (lbs)",
         "=SUM(tbl_CommunityEvents[Food_Recovery_lbs])",
         "0"),
        ("Food waste diverted (lbs)",
         '=SUM(tbl_KitchenOps[Food_Waste_Diverted_lbs])',
         "0"),
        ("Grants secured ($)",
         '=SUMIFS(tbl_Director[Amount_$],tbl_Director[Entry_Type],"Grant")',
         "$#,##0"),
        ("Farms engaged",
         '=IFERROR(ROWS(UNIQUE(FILTER(tbl_Farmers[Farm_Name],tbl_Farmers[Farm_Name]<>""))),0)',
         "0"),
        ("RD in-kind hours (YTD)",
         "=SUM(tbl_Culinary[RD_InKind_Hours])",
         "0"),
        ("Health demo curriculums delivered",
         '=IFERROR(ROWS(UNIQUE(FILTER(tbl_Culinary[Curriculum],tbl_Culinary[Curriculum]<>""))),0)',
         "0"),
        ("Produce moved (lbs)", "=SUM(tbl_Farmers[Produce_lbs])", "0"),
        ("Survey responses collected", "=ROWS(tbl_SurveyResponses)-1", "0"),
    ]

    for idx, (label, formula, fmt) in enumerate(tiles):
        col = (idx % 3) * 2 + 1
        row = 4 + (idx // 3) * 4
        lbl = ws.cell(row=row, column=col, value=label)
        lbl.font = Font(name="Calibri", size=10, bold=True, color=DEEP_GREEN)
        lbl.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        lbl.fill = PatternFill("solid", fgColor=LIGHT_GREEN)
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 1)
        ws.row_dimensions[row].height = 22
        val = ws.cell(row=row + 1, column=col, value=formula)
        val.font = Font(name="Calibri", size=22, bold=True, color=TEXT_DARK)
        val.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        val.fill = PatternFill("solid", fgColor=LIGHT_GREEN)
        val.number_format = fmt
        ws.merge_cells(start_row=row + 1, start_column=col, end_row=row + 2, end_column=col + 1)
        ws.row_dimensions[row + 1].height = 30
        ws.row_dimensions[row + 2].height = 8
        for c in range(col, col + 2):
            ws.column_dimensions[get_column_letter(c)].width = 22

    base = 4 + (((len(tiles) - 1) // 3) + 1) * 4 + 1
    ws.cell(row=base, column=1, value="By OFSA Dimension — count of KPIs tagged").font = \
        Font(name="Calibri", size=13, bold=True, color=DEEP_GREEN)
    ws.merge_cells(start_row=base, start_column=1, end_row=base, end_column=6)

    dims = ["Availability", "Access", "Utilization", "Stability", "Agency", "Sustainability"]
    for i, dim in enumerate(dims):
        ws.cell(row=base + 1, column=i + 1, value=dim).font = header_font()
        ws.cell(row=base + 1, column=i + 1).fill = header_fill()
        ws.cell(row=base + 1, column=i + 1).alignment = Alignment(horizontal="center")
        ws.cell(row=base + 2, column=i + 1,
                value=f'=COUNTIF(tbl_KPI_Catalog[OFSA Dimension(s)],"*{dim}*")')
        ws.cell(row=base + 2, column=i + 1).alignment = Alignment(horizontal="center")
        ws.cell(row=base + 2, column=i + 1).font = Font(name="Calibri", size=16, bold=True, color=TEXT_DARK)

    ws.protection = SheetProtection(sheet=True, password=None,
                                    selectLockedCells=False,
                                    selectUnlockedCells=False)


def build_ofsa_dashboard(wb):
    ws = wb.create_sheet("Dashboard — By OFSA Dimension")
    ws.sheet_view.showGridLines = False

    t = ws.cell(row=1, column=1, value="HARVEST — By OFSA Dimension")
    t.font = Font(name="Calibri", size=18, bold=True, color=WHITE)
    t.fill = header_fill()
    t.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells("A1:E1")
    ws.row_dimensions[1].height = 36

    sub = ws.cell(row=2, column=1, value="KPIs tagged to each OFSA dimension. Bridge to the logic model.")
    sub.font = Font(name="Calibri", size=11, italic=True, color=TEXT_DARK)
    sub.alignment = Alignment(vertical="center", indent=1)
    ws.merge_cells("A2:E2")

    dims = ["Availability", "Access", "Utilization", "Stability", "Agency", "Sustainability"]
    row = 4
    for dim in dims:
        h = ws.cell(row=row, column=1, value=dim)
        h.font = Font(name="Calibri", size=14, bold=True, color=WHITE)
        h.fill = header_fill()
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        ws.row_dimensions[row].height = 28
        row += 1
        for i, sh in enumerate(["KPI", "Audience", "Unit", "2026 Target", "Status"]):
            c = ws.cell(row=row, column=i + 1, value=sh)
            c.font = Font(name="Calibri", size=10, bold=True, color=TEXT_DARK)
            c.fill = PatternFill("solid", fgColor=LIGHT_GREEN)
        row += 1
        for kpi in [k for k in KPI_CATALOG if dim in (k[4] or "")]:
            ws.cell(row=row, column=1, value=kpi[0]).font = body_font()
            ws.cell(row=row, column=2, value=kpi[3]).font = body_font()
            ws.cell(row=row, column=3, value=kpi[2]).font = body_font()
            ws.cell(row=row, column=4, value=kpi[10]).font = body_font()
            ws.cell(row=row, column=5, value=kpi[13]).font = body_font()
            for c in range(1, 6):
                ws.cell(row=row, column=c).alignment = Alignment(vertical="top", wrap_text=True)
            row += 1
        row += 1

    for i, w in enumerate([42, 24, 8, 12, 22]):
        ws.column_dimensions[get_column_letter(i + 1)].width = w
    ws.freeze_panes = "A4"
    ws.protection = SheetProtection(sheet=True, password=None,
                                    selectLockedCells=False,
                                    selectUnlockedCells=False)


def build_audience_dashboard(wb):
    ws = wb.create_sheet("Dashboard — By Audience")
    ws.sheet_view.showGridLines = False

    t = ws.cell(row=1, column=1, value="HARVEST — By Audience")
    t.font = Font(name="Calibri", size=18, bold=True, color=WHITE)
    t.fill = header_fill()
    ws.merge_cells("A1:E1")
    ws.row_dimensions[1].height = 36

    sub = ws.cell(row=2, column=1, value="KPIs by audience. Filter for audience-specific reports.")
    sub.font = Font(name="Calibri", size=11, italic=True, color=TEXT_DARK)
    ws.merge_cells("A2:E2")

    audiences = [
        "Food Entrepreneurs", "Institution", "Community Organizations",
        "Investors-Funders", "Farmers", "Environmental Sustainability",
    ]
    row = 4
    for aud in audiences:
        h = ws.cell(row=row, column=1, value=aud)
        h.font = Font(name="Calibri", size=14, bold=True, color=WHITE)
        h.fill = header_fill()
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        ws.row_dimensions[row].height = 28
        row += 1
        for i, sh in enumerate(["KPI", "Unit", "Cadence", "2026 Target", "Owner"]):
            c = ws.cell(row=row, column=i + 1, value=sh)
            c.font = Font(name="Calibri", size=10, bold=True, color=TEXT_DARK)
            c.fill = PatternFill("solid", fgColor=LIGHT_GREEN)
        row += 1
        for kpi in [k for k in KPI_CATALOG if k[3] == aud]:
            ws.cell(row=row, column=1, value=kpi[0]).font = body_font()
            ws.cell(row=row, column=2, value=kpi[2]).font = body_font()
            ws.cell(row=row, column=3, value=kpi[7]).font = body_font()
            ws.cell(row=row, column=4, value=kpi[10]).font = body_font()
            ws.cell(row=row, column=5, value=kpi[11]).font = body_font()
            for c in range(1, 6):
                ws.cell(row=row, column=c).alignment = Alignment(vertical="top", wrap_text=True)
            row += 1
        row += 1

    for i, w in enumerate([42, 8, 14, 12, 28]):
        ws.column_dimensions[get_column_letter(i + 1)].width = w
    ws.freeze_panes = "A4"
    ws.protection = SheetProtection(sheet=True, password=None,
                                    selectLockedCells=False,
                                    selectUnlockedCells=False)


def build_food_corridor_import(wb):
    ws = wb.create_sheet("Food Corridor Import")
    ws.sheet_view.showGridLines = False

    t = ws.cell(row=1, column=1, value="Food Corridor — CSV Import")
    t.font = Font(name="Calibri", size=18, bold=True, color=WHITE)
    t.fill = header_fill()
    ws.merge_cells("A1:F1")
    ws.row_dimensions[1].height = 36

    instructions = [
        "How to refresh kitchen rental data:",
        "  1. In Food Corridor, go to Reports > Bookings > Export to CSV.",
        "  2. Save the file as 'food_corridor.csv' to the same folder as this workbook.",
        "  3. In Excel, click Data > Refresh All.",
        "",
        "v1 schema: Booking Date, Member Business, Hours only.",
        "Expand when Anthony shares a real Food Corridor export.",
        "",
        "If Power Query is not yet wired up, paste rows directly below the headers.",
    ]
    for i, line in enumerate(instructions):
        c = ws.cell(row=2 + i, column=1, value=line)
        c.font = body_font()
        c.alignment = Alignment(vertical="center", wrap_text=True)
        ws.merge_cells(start_row=2 + i, start_column=1, end_row=2 + i, end_column=6)

    header_row = 2 + len(instructions) + 2
    write_header_row(ws, header_row, FOOD_CORRIDOR_HEADERS, FOOD_CORRIDOR_WIDTHS)

    for ridx, row in enumerate(FOOD_CORRIDOR_SAMPLES, start=header_row + 1):
        for cidx, val in enumerate(row, start=1):
            c = ws.cell(row=ridx, column=cidx, value=val)
            c.font = body_font()
            if cidx == 1:
                c.number_format = "yyyy-mm-dd"

    last_row = header_row + len(FOOD_CORRIDOR_SAMPLES)
    last_col = get_column_letter(len(FOOD_CORRIDOR_HEADERS))
    add_table(ws, "tbl_FoodCorridor", f"A{header_row}:{last_col}{last_row}")
    ws.freeze_panes = f"A{header_row + 1}"


def build_data_tab(wb, tab_spec):
    ws = wb.create_sheet(tab_spec["sheet"])
    ws.sheet_view.showGridLines = False

    note = ws.cell(row=1, column=1, value=(
        f"Back-end data tab. Populated by Power Query from the corresponding "
        f"log workbook. Sample rows seeded so dashboards render. "
        f"Replace via Data > Get Data > From Workbook — M code in docs/data_dictionary.md."
    ))
    note.font = Font(name="Calibri", size=10, italic=True, color=TEXT_DARK)
    note.alignment = Alignment(vertical="center", wrap_text=True)

    if tab_spec["source_log"] and tab_spec["source_log"] in LOG_SPECS:
        src = LOG_SPECS[tab_spec["source_log"]]
        headers = src["headers"]
        widths = src["widths"]
        samples = src["samples"]
    elif tab_spec["table"] == "tbl_SurveyResponses":
        headers = SURVEY_HEADERS
        widths = SURVEY_WIDTHS
        samples = SURVEY_SAMPLES
    else:
        return

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))

    header_row = 3
    write_header_row(ws, header_row, headers, widths)

    for ridx, row in enumerate(samples, start=header_row + 1):
        for cidx, val in enumerate(row, start=1):
            c = ws.cell(row=ridx, column=cidx, value=val)
            c.font = body_font()
            c.alignment = Alignment(vertical="top", wrap_text=True)
            if isinstance(val, date):
                c.number_format = "yyyy-mm-dd"
            if cidx <= len(headers) and headers[cidx - 1].endswith("_$"):
                if isinstance(val, (int, float)) and val:
                    c.number_format = "$#,##0"

    last_row = header_row + len(samples)
    last_col = get_column_letter(len(headers))
    add_table(ws, tab_spec["table"], f"A{header_row}:{last_col}{last_row}")
    ws.freeze_panes = f"A{header_row + 1}"


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    LOGS_DIR.mkdir(exist_ok=True)

    print("Building individual log workbooks:")
    for key, spec in LOG_SPECS.items():
        build_log_workbook(spec)

    print(f"\nBuilding master workbook:")
    wb = Workbook()
    build_readme(wb)
    build_kpi_catalog(wb)
    build_goals_2026(wb)
    build_executive_summary(wb)
    build_ofsa_dashboard(wb)
    build_audience_dashboard(wb)
    build_food_corridor_import(wb)
    for tab_spec in MASTER_DATA_TABS:
        build_data_tab(wb, tab_spec)

    wb.active = 0
    wb.save(MASTER_FILE)
    print(f"  {MASTER_FILE.name}")
    print("\nDone.")


if __name__ == "__main__":
    main()
