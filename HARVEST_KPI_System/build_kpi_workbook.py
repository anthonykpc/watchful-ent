"""
Generates HARVEST_KPI_Master.xlsx — the user-facing deliverable of the
HARVEST Food Hub KPI tracking system.

Tabs produced (in this order):
  1. README                       — quick start, refresh instructions
  2. KPI Catalog                  — full enriched catalog (source of truth)
  3. 2026 Performance Goals       — goal tracker with R/Y/G status
  4. Dashboard — Executive Summary
  5. Dashboard — By OFSA Dimension
  6. Dashboard — By Audience
  7. Food Corridor Import         — drop zone for the CSV export
  8. Data — Food Entrepreneurs    \
  9. Data — Institution            \
 10. Data — Community Orgs          > Back-end data tabs, populated from
 11. Data — Investors-Funders      /  SharePoint Lists via Power Query.
 12. Data — Farmers                /  Seeded with sample rows so the
 13. Data — Environmental         /   dashboards render before Power Query
                                      is wired up.

Run:  python3 build_kpi_workbook.py
Output: HARVEST_KPI_Master.xlsx in the script's directory.
"""

from datetime import date, timedelta
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

OUTPUT_FILE = Path(__file__).parent / "HARVEST_KPI_Master.xlsx"

# Brand colors (matched to existing HARVEST_Contact_Tracker.xlsx)
DEEP_GREEN = "2D6A4F"
LIGHT_GREEN = "E8F3EC"
WHITE = "FFFFFF"
TEXT_DARK = "1B3A2B"
BORDER_GREEN = "B7D9C2"
STATUS_GREEN = "C6EFCE"
STATUS_YELLOW = "FFEB9C"
STATUS_RED = "FFC7CE"

# Row count for pre-formatted data ranges in each Data tab
DATA_MAX_ROW = 2000


# -----------------------------------------------------------------------------
# Shared style helpers
# -----------------------------------------------------------------------------

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
    """Add an Excel Table with structured references. Style matches brand."""
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


# -----------------------------------------------------------------------------
# KPI catalog data — drives multiple tabs
# -----------------------------------------------------------------------------

# Each row: (KPI, Definition, Unit, Audience, OFSA_Dimensions, Food_System_Sector,
#            Partner_Sector, Cadence, Lead_Role, Data_Source, Target_2026,
#            Owner_Role, Source_Form, Destination_List, Status)
KPI_CATALOG = [
    # Food Entrepreneurs
    ("Entrepreneurs engaged", "Unique individuals reached via outreach, intake, or info session", "count", "Food Entrepreneurs", "Agency", "Processing", "Food Enterprise", "Monthly", "HARVEST Lead", "Form submission", 40, "Community Outreach Specialist", "Community Engagement Log", "Food Entrepreneurs List", "Active"),
    ("Entrepreneurs converted to renting members", "Engaged entrepreneurs who signed a kitchen rental agreement", "count", "Food Entrepreneurs", "Agency; Availability", "Processing", "Food Enterprise", "Monthly", "HARVEST Lead", "Form submission", 8, "Logistics Manager", "Kitchen Operations Log", "Food Entrepreneurs List", "Active"),
    ("% members achieving ServSafe", "Members who pass ServSafe Food Handler or Manager via HARVEST", "%", "Food Entrepreneurs", "Utilization", "Processing", "Food Enterprise", "Per training", "HARVEST Lead", "Form submission", 100, "Culinary Manager", "Culinary & Training Log", "Food Entrepreneurs List", "Active"),
    ("Avg % sales increase for members", "YoY revenue change for members in kitchen ≥6 months", "%", "Food Entrepreneurs", "Stability; Agency", "Processing", "Food Enterprise", "Quarterly", "HARVEST Lead", "Manual entry", None, "Director", "Director & Partnership Log", "Food Entrepreneurs List", "Active"),
    ("Businesses supported", "Distinct member businesses receiving any HARVEST service", "count", "Food Entrepreneurs", "Agency", "Processing", "Food Enterprise", "Monthly", "HARVEST Lead", "System-calculated", 8, "Director", None, "Food Entrepreneurs List", "Active"),
    ("Workforce dev seminars held", "Seminars on business building, food safety, marketing, capital", "count", "Food Entrepreneurs", "Utilization; Agency", "Processing", "Education", "Per event", "HARVEST Lead", "Form submission", 6, "Culinary Manager", "Culinary & Training Log", "Food Entrepreneurs List", "Active"),
    ("Kitchen hours rented", "Total billable kitchen hours rented in the period", "hours", "Food Entrepreneurs", "Availability; Stability", "Processing", "Food Enterprise", "Real-time", "HARVEST Lead", "Food Corridor CSV", 1500, "Logistics Manager", None, "Food Entrepreneurs List", "Active"),
    ("Jobs created by member businesses", "FTEs and PTEs hired by member businesses (self-reported)", "count", "Food Entrepreneurs", "Stability; Agency", "Processing", "Food Enterprise", "Quarterly", "HARVEST Lead", "Form submission", None, "Director", "Director & Partnership Log", "Food Entrepreneurs List", "Active"),
    ("Growth of businesses by hours rented", "Avg hours/member over time", "hours", "Food Entrepreneurs", "Stability", "Processing", "Food Enterprise", "Quarterly", "HARVEST Lead", "System-calculated", None, "Director", None, "Food Entrepreneurs List", "Active"),
    ("Businesses buying local produce via UAC", "Member businesses sourcing produce from UAC channels", "count", "Food Entrepreneurs", "Access; Sustainability", "Aggregation", "Agriculture", "Monthly", "Joint", "Form submission", 3, "UAC Partner", "UAC Farmer Engagement Log", "Food Entrepreneurs List", "Active"),
    ("Events showcasing local entrepreneurs", "Public events featuring member businesses", "count", "Food Entrepreneurs", "Agency", "Retail-Consumption", "Food Enterprise", "Per event", "HARVEST Lead", "Form submission", 4, "Community Outreach Specialist", "Community Engagement Log", "Food Entrepreneurs List", "Active"),
    ("Technical assistance sessions", "1:1 sessions for business building, capital, marketing", "count", "Food Entrepreneurs", "Utilization; Agency", "Processing", "Food Enterprise", "Per session", "HARVEST Lead", "Form submission", 30, "Culinary Manager", "Culinary & Training Log", "Food Entrepreneurs List", "Active"),
    ("Businesses receiving wellness coaching", "Member businesses in HARVEST wellness/nutrition coaching", "count", "Food Entrepreneurs", "Utilization", "Processing", "Healthcare", "Per session", "HARVEST Lead", "Form submission", None, "Community Outreach Specialist", "Culinary & Training Log", "Food Entrepreneurs List", "Active"),
    # Institution
    ("$ purchasing orders from HARVEST/UAC vendors", "Total $ value of POs by institutions for HARVEST/UAC products", "$", "Institution", "Access; Availability", "Distribution", "Healthcare", "Monthly", "Joint", "Form submission", None, "Director", "Director & Partnership Log", "Institution List", "Active"),
    ("HARVEST-produced items on institutional menus", "Distinct HARVEST/UAC items on institutional menus", "count", "Institution", "Access; Availability", "Retail-Consumption", "Healthcare", "Quarterly", "HARVEST Lead", "Form submission", None, "Director", "Director & Partnership Log", "Institution List", "Active"),
    ("Cross-institution connections", "Active relationships with Rutgers, NPS, Audible, etc.", "count", "Institution", "Stability; Access", "n/a", "Education", "Quarterly", "HARVEST Lead", "Form submission", 4, "Director", "Director & Partnership Log", "Institution List", "Active"),
    ("TA partnerships", "External institutions providing TA to HARVEST members", "count", "Institution", "Utilization; Stability", "n/a", "Education", "Quarterly", "Joint", "Form submission", 3, "Director", "Director & Partnership Log", "Institution List", "Active"),
    ("Reduction in readmission costs through HARVEST", "RWJBH patient readmission $ reduction attributable to HARVEST", "$", "Institution", "Utilization; Access", "Retail-Consumption", "Healthcare", "Annual", "Joint", "Manual entry", None, "Director", None, "Institution List", "Planned — methodology TBD"),
    # Community Orgs
    ("Community orgs engaged", "Distinct orgs participating in any HARVEST activity", "count", "Community Organizations", "Agency; Access", "n/a", "Community-Based Org", "Monthly", "HARVEST Lead", "Form submission", 15, "Community Outreach Specialist", "Community Engagement Log", "Community Orgs List", "Active"),
    ("Event participants", "Total participants across teaching kitchen, RD workshops, etc.", "count", "Community Organizations", "Utilization; Agency", "n/a", "Community-Based Org", "Per event", "HARVEST Lead", "Form submission", 300, "Community Outreach Specialist", "Community Engagement Log", "Community Orgs List", "Active"),
    ("Community volunteer hours", "Hours contributed by community volunteers at HARVEST", "hours", "Community Organizations", "Agency; Stability", "n/a", "Community-Based Org", "Per event", "HARVEST Lead", "Form submission", None, "Community Outreach Specialist", "Community Engagement Log", "Community Orgs List", "Active"),
    ("Community events run by staff", "Events run by HARVEST staff (on-site and external)", "count", "Community Organizations", "Agency", "n/a", "Community-Based Org", "Per event", "HARVEST Lead", "Form submission", 24, "Community Outreach Specialist", "Community Engagement Log", "Community Orgs List", "Active"),
    ("Event survey — satisfaction (avg)", "Avg satisfaction from post-event survey (1–5)", "rating", "Community Organizations", "Utilization", "n/a", "Community-Based Org", "Per event", "HARVEST Lead", "Form submission", 4.0, "Community Outreach Specialist", "Community Engagement Log", "Community Orgs List", "Active"),
    ("Event survey — learning (avg)", "Avg self-reported learning from post-event survey (1–5)", "rating", "Community Organizations", "Utilization", "n/a", "Community-Based Org", "Per event", "HARVEST Lead", "Form submission", 4.0, "Community Outreach Specialist", "Community Engagement Log", "Community Orgs List", "Active"),
    ("Tours given", "Facility tours given to external visitors", "count", "Community Organizations", "Agency", "n/a", "Community-Based Org", "Per event", "HARVEST Lead", "Form submission", 24, "Community Outreach Specialist", "Community Engagement Log", "Community Orgs List", "Active"),
    ("Orgs hosting events at HARVEST", "External orgs using HARVEST as their venue", "count", "Community Organizations", "Agency; Access", "n/a", "Community-Based Org", "Per event", "HARVEST Lead", "Form submission", 6, "Community Outreach Specialist", "Community Engagement Log", "Community Orgs List", "Active"),
    ("Sectors engaged at HARVEST", "Distinct partner sectors represented across engagements", "count", "Community Organizations", "Stability", "n/a", "Community-Based Org", "Quarterly", "HARVEST Lead", "System-calculated", 5, "Director", None, "Community Orgs List", "Active"),
    ("Multilingual / cultural events", "Events in a non-English language or built around a tradition", "count", "Community Organizations", "Agency; Access", "n/a", "Community-Based Org", "Per event", "HARVEST Lead", "Form submission", 6, "Community Outreach Specialist", "Community Engagement Log", "Community Orgs List", "Active"),
    ("Food recovery to service providers (lbs)", "lbs diverted to Bridges and other direct service providers", "lbs", "Community Organizations", "Access; Sustainability", "Recovery-Waste", "Community-Based Org", "Per event", "HARVEST Lead", "Form submission", 500, "Community Outreach Specialist", "Community Engagement Log", "Community Orgs List", "Active"),
    ("DSPs with presence at HARVEST", "Direct service providers with recurring on-site presence", "count", "Community Organizations", "Access", "n/a", "Community-Based Org", "Quarterly", "HARVEST Lead", "Form submission", 3, "Community Outreach Specialist", "Community Engagement Log", "Community Orgs List", "Active"),
    # Investors / Funders
    ("Grants and partnerships secured", "New grants or formal partnerships executed", "count", "Investors-Funders", "Stability", "n/a", "Philanthropy", "Per event", "HARVEST Lead", "Form submission", 4, "Director", "Director & Partnership Log", "Investors-Funders List", "Active"),
    ("% earned revenue", "Earned revenue as % of total revenue", "%", "Investors-Funders", "Stability; Sustainability", "Processing", "Philanthropy", "Quarterly", "HARVEST Lead", "Manual entry", None, "Director", "Director & Partnership Log", "Investors-Funders List", "Active"),
    ("Funders and sponsors", "Distinct active funders/sponsors in the period", "count", "Investors-Funders", "Stability", "n/a", "Philanthropy", "Quarterly", "HARVEST Lead", "Form submission", None, "Director", "Director & Partnership Log", "Investors-Funders List", "Active"),
    ("Leveraged capital ($)", "Matching grants, investments, in-kind contributions leveraged", "$", "Investors-Funders", "Stability; Sustainability", "n/a", "Philanthropy", "Quarterly", "HARVEST Lead", "Form submission", None, "Director", "Director & Partnership Log", "Investors-Funders List", "Active"),
    ("Social return on investment (SROI)", "Ratio of social value created to dollars invested", "ratio", "Investors-Funders", "Sustainability; Stability", "n/a", "Philanthropy", "Annual", "HARVEST Lead", "Manual entry", None, "Director", None, "Investors-Funders List", "Planned — methodology TBD"),
    ("JSC meetings held", "Joint Steering Committee meetings convened", "count", "Investors-Funders", "Stability", "n/a", "Philanthropy", "Per event", "Joint", "Form submission", 10, "Director", "Director & Partnership Log", "Investors-Funders List", "Active"),
    ("CAB members onboarded", "Community Advisory Board members across sectors", "count", "Investors-Funders", "Agency; Stability", "n/a", "Community-Based Org", "Per event", "HARVEST Lead", "Form submission", 10, "Director", "Director & Partnership Log", "Investors-Funders List", "Active"),
    ("CAB meetings held", "CAB meetings convened in the period", "count", "Investors-Funders", "Agency; Stability", "n/a", "Community-Based Org", "Per event", "HARVEST Lead", "Form submission", 2, "Director", "Director & Partnership Log", "Investors-Funders List", "Active"),
    ("Director community presentations", "Community-facing presentations by the Director", "count", "Investors-Funders", "Agency", "n/a", "Philanthropy", "Per event", "HARVEST Lead", "Form submission", 4, "Director", "Director & Partnership Log", "Investors-Funders List", "Active"),
    # Farmers
    ("Farms engaged", "Distinct farms in active relationship with HARVEST/UAC", "count", "Farmers", "Agency; Stability", "Production", "Agriculture", "Monthly", "UAC Lead", "Form submission", 4, "UAC Partner", "UAC Farmer Engagement Log", "Farmers List", "Active"),
    ("$ farmer business via HARVEST", "Total $ farmer sales attributable to HARVEST/UAC channels", "$", "Farmers", "Stability; Access", "Aggregation", "Agriculture", "Monthly", "UAC Lead", "Form submission", None, "UAC Partner", "UAC Farmer Engagement Log", "Farmers List", "Active"),
    ("Types of produce moved", "Distinct produce types moved through HARVEST/UAC", "count", "Farmers", "Availability", "Aggregation", "Agriculture", "Monthly", "UAC Lead", "Form submission", None, "UAC Partner", "UAC Farmer Engagement Log", "Farmers List", "Active"),
    ("lbs of produce moved", "Total lbs of produce aggregated/distributed via HARVEST/UAC", "lbs", "Farmers", "Availability; Access", "Aggregation", "Agriculture", "Monthly", "UAC Lead", "Form submission", None, "UAC Partner", "UAC Farmer Engagement Log", "Farmers List", "Active"),
    ("Farmers making value-added products", "Farmers using HARVEST for value-added production", "count", "Farmers", "Agency; Stability", "Processing", "Agriculture", "Monthly", "UAC Lead", "Form submission", None, "UAC Partner", "UAC Farmer Engagement Log", "Farmers List", "Active"),
    # Environmental
    ("Food waste diversion (lbs)", "lbs of food waste diverted from landfill (any pathway)", "lbs", "Environmental Sustainability", "Sustainability", "Recovery-Waste", "Community-Based Org", "Monthly", "HARVEST Lead", "Form submission", 1000, "Logistics Manager", "Kitchen Operations Log", "Environmental List", "Active"),
    ("Composted (lbs)", "lbs of food waste composted", "lbs", "Environmental Sustainability", "Sustainability", "Recovery-Waste", "Community-Based Org", "Monthly", "HARVEST Lead", "Form submission", None, "Logistics Manager", "Kitchen Operations Log", "Environmental List", "Active"),
    ("Local food miles (avg)", "Avg distance (miles) from source to HARVEST", "miles", "Environmental Sustainability", "Sustainability", "Distribution", "Agriculture", "Quarterly", "UAC Lead", "System-calculated", None, "UAC Partner", None, "Environmental List", "Active"),
    ("Food captured for demos (lbs)", "lbs of food used in demos instead of being wasted", "lbs", "Environmental Sustainability", "Sustainability; Utilization", "Recovery-Waste", "Healthcare", "Per event", "HARVEST Lead", "Form submission", None, "Culinary Manager", "Culinary & Training Log", "Environmental List", "Active"),
]

KPI_CATALOG_HEADERS = [
    "KPI", "Definition", "Unit", "Audience", "OFSA Dimension(s)",
    "Food System Sector", "Partner Sector", "Cadence", "Lead Role",
    "Data Source", "2026 Target", "Owner Role", "Source Form",
    "Destination List", "Status",
]
KPI_CATALOG_WIDTHS = [38, 50, 8, 22, 26, 22, 22, 12, 14, 22, 12, 26, 28, 26, 22]


# -----------------------------------------------------------------------------
# 2026 Performance Goals
# -----------------------------------------------------------------------------

# (Goal, Sub-goal, Linked metric, Target, Current value formula, Tracking notes)
GOALS_2026 = [
    ("1. Kitchen launch / operations", "Founding cohort — members converted",
     "Entrepreneurs converted to renting members", 8,
     '=SUMIFS(tbl_Entrepreneurs[New_Members],tbl_Entrepreneurs[Reporting_Period_End],">="&DATE(2026,1,1),tbl_Entrepreneurs[Reporting_Period_End],"<="&DATE(2026,12,31))',
     "KPI system"),
    ("1. Kitchen launch / operations", "ServSafe pass rate",
     "% members achieving ServSafe", 100,
     '=IFERROR(SUMIFS(tbl_Entrepreneurs[ServSafe_Passed],tbl_Entrepreneurs[Reporting_Period_End],">="&DATE(2026,1,1))/SUMIFS(tbl_Entrepreneurs[ServSafe_Attempted],tbl_Entrepreneurs[Reporting_Period_End],">="&DATE(2026,1,1))*100,0)',
     "KPI system"),
    ("1. Kitchen launch / operations", "Kitchen rental hours",
     "Kitchen hours rented", 1500,
     '=SUM(tbl_FoodCorridor[Hours])',
     "Food Corridor CSV"),
    ("2. Structural documentation", "SOPs, guiding docs complete",
     "—", None, '="See shared drive folder"',
     "Tracked outside KPI system"),
    ("3. UAC partnership maintenance", "JSC meetings held", "JSC meetings held",
     10, '=COUNTIFS(tbl_Funders[Activity_Type],"JSC Meeting",tbl_Funders[Submitted_Date],">="&DATE(2026,1,1))',
     "KPI system"),
    ("3. UAC partnership maintenance", "Shared data framework operational",
     "Binary status", 1, '=1',
     "This workbook is live = 1"),
    ("3. UAC partnership maintenance", "UAC-affiliated farmers engaged",
     "Farms engaged", 4,
     '=IFERROR(ROWS(UNIQUE(FILTER(tbl_Farmers[Farm_Name],tbl_Farmers[Farm_Name]<>""))),0)',
     "Distinct farms in Farmers data tab"),
    ("4. CAB development", "CAB members across sectors", "CAB members onboarded",
     10, '=COUNTIFS(tbl_Funders[Activity_Type],"CAB Member Added",tbl_Funders[Submitted_Date],">="&DATE(2026,1,1))',
     "KPI system"),
    ("4. CAB development", "CAB meetings held by year-end", "CAB meetings held",
     2, '=COUNTIFS(tbl_Funders[Activity_Type],"CAB Meeting",tbl_Funders[Submitted_Date],">="&DATE(2026,1,1))',
     "KPI system"),
    ("5. Professional development", "Shared Kitchen Summit attendance",
     "PD activity log", 1,
     '=COUNTIFS(tbl_Funders[Activity_Type],"Professional Development",tbl_Funders[Notes],"*Shared Kitchen Summit*")',
     "KPI system"),
    ("5. Professional development", "ServSafe Manager Cert (Director)",
     "PD activity log", 1,
     '=COUNTIFS(tbl_Funders[Activity_Type],"Professional Development",tbl_Funders[Notes],"*ServSafe Manager*")',
     "KPI system"),
    ("5. Professional development", "Community presentations",
     "Director community presentations", 4,
     '=COUNTIFS(tbl_Funders[Activity_Type],"Professional Development",tbl_Funders[Notes],"*Community Presentation*")',
     "KPI system"),
]


# -----------------------------------------------------------------------------
# Data tab specs: (sheet name, table name, headers, sample rows)
# -----------------------------------------------------------------------------

DATA_TABS = [
    {
        "sheet": "Data — Food Entrepreneurs",
        "table": "tbl_Entrepreneurs",
        "headers": [
            "Submission_ID", "Title", "Submitted_By", "Submitted_Date",
            "Reporting_Period_Start", "Reporting_Period_End", "OFSA_Dimension",
            "Food_System_Sector", "Partner_Sector", "Lead_Role",
            "Business_Name", "Stage", "Date_Engaged", "Date_Converted",
            "Jobs_Created_FTE", "Jobs_Created_PTE", "Sales_Period_$",
            "Sales_YoY_Change_%", "Sources_Local_Produce",
            "Wellness_Coaching_Sessions", "Demographic_Tags",
            "New_Members", "ServSafe_Passed", "ServSafe_Attempted", "Notes",
        ],
        "widths": [12, 32, 22, 14, 16, 16, 24, 18, 18, 14, 24, 14, 12, 12, 10, 10, 14, 12, 10, 12, 24, 12, 12, 12, 36],
        "samples": [
            [1, "Newark Eats LLC — onboarding", "L. Reyes (Logistics)", date(2026, 3, 15), date(2026, 3, 1), date(2026, 3, 31),
             "Agency; Availability", "Processing", "Food Enterprise", "HARVEST Lead",
             "Newark Eats LLC", "Onboarding", date(2026, 3, 10), date(2026, 3, 14),
             1, 2, 4800, None, True, 0, "Woman-owned; BIPOC-owned", 1, 0, 0, "First member of founding cohort."],
            [2, "Casa Sabor — onboarding", "L. Reyes (Logistics)", date(2026, 3, 22), date(2026, 3, 1), date(2026, 3, 31),
             "Agency", "Processing", "Food Enterprise", "HARVEST Lead",
             "Casa Sabor", "Onboarding", date(2026, 3, 18), date(2026, 3, 21),
             0, 1, 2400, None, True, 0, "Immigrant-owned; First-time entrepreneur", 1, 0, 0, "Multilingual onboarding."],
            [3, "Q1 operational snapshot", "L. Reyes (Logistics)", date(2026, 3, 31), date(2026, 1, 1), date(2026, 3, 31),
             "Availability; Stability", "Processing", "Food Enterprise", "HARVEST Lead",
             None, "Operational Snapshot", None, None, None, None, None, None, None, None, None, 0, 0, 0, "First quarter open — 2 founding cohort members."],
            [4, "ServSafe — April cohort", "M. Chen (Culinary)", date(2026, 4, 12), date(2026, 4, 12), date(2026, 4, 12),
             "Utilization", "Processing", "Education", "HARVEST Lead",
             None, "Operational Snapshot", None, None, None, None, None, None, None, None, None, 0, 2, 2, "ServSafe Food Handler — 2 of 2 passed."],
            [5, "Spring TA cohort", "M. Chen (Culinary)", date(2026, 4, 10), date(2026, 4, 1), date(2026, 4, 30),
             "Utilization; Agency", "Processing", "Food Enterprise", "HARVEST Lead",
             "Newark Eats LLC", "Active Member", None, None, None, None, None, None, None, 1, None, 0, 0, 0, "1:1 marketing TA."],
        ],
    },
    {
        "sheet": "Data — Institution",
        "table": "tbl_Institution",
        "headers": [
            "Submission_ID", "Title", "Submitted_By", "Submitted_Date",
            "Reporting_Period_Start", "Reporting_Period_End", "OFSA_Dimension",
            "Food_System_Sector", "Partner_Sector", "Lead_Role",
            "Institution_Name", "Institution_Type", "Activity_Type",
            "PO_Value_$", "Items_On_Menu", "Connection_Notes",
            "Health_Outcome_Metric", "Health_Outcome_Value", "Notes",
        ],
        "widths": [12, 36, 22, 14, 16, 16, 24, 18, 18, 14, 26, 16, 22, 14, 12, 36, 26, 14, 36],
        "samples": [
            [1, "RWJBH Newark — Q1 PO", "A. Capece (Director)", date(2026, 3, 31), date(2026, 1, 1), date(2026, 3, 31),
             "Access; Availability", "Distribution", "Healthcare", "Joint",
             "RWJBH Newark", "Healthcare", "Purchase Order",
             8500, None, "First institutional PO — UAC produce + 2 HARVEST member items.",
             None, None, "Confirms institutional pathway is live."],
            [2, "Rutgers connection", "A. Capece (Director)", date(2026, 4, 5), date(2026, 4, 1), date(2026, 4, 30),
             "Stability; Access", "n/a", "Education", "HARVEST Lead",
             "Rutgers New Brunswick", "Education", "Cross-Institution Connection",
             None, None, "Joint research interest on food-as-medicine outcomes.",
             None, None, "Lead: Dr. Patel."],
        ],
    },
    {
        "sheet": "Data — Community Orgs",
        "table": "tbl_Community",
        "headers": [
            "Submission_ID", "Title", "Submitted_By", "Submitted_Date",
            "Reporting_Period_Start", "Reporting_Period_End", "OFSA_Dimension",
            "Food_System_Sector", "Partner_Sector", "Lead_Role",
            "Event_Name", "Event_Type", "Event_Date", "Attendance_Count",
            "Volunteer_Hours", "Orgs_Engaged", "Sector_Count",
            "Multilingual_Cultural", "Language_Tradition",
            "Survey_Satisfaction_Avg", "Survey_Learning_Avg",
            "Food_Recovery_lbs", "Food_Recovery_Recipient",
            "Direct_Service_Provider_Onsite", "Notes",
        ],
        "widths": [12, 32, 22, 14, 16, 16, 24, 18, 22, 14, 28, 22, 12, 12, 14, 36, 10, 18, 22, 14, 14, 12, 22, 22, 32],
        "samples": [
            [1, "Open House — March", "J. Tran (Outreach)", date(2026, 3, 22), date(2026, 3, 22), date(2026, 3, 22),
             "Agency; Access", "n/a", "Community-Based Org", "HARVEST Lead",
             "Open House — March", "HARVEST-Hosted Event", date(2026, 3, 22),
             62, 8.5, "Bridges; Newark Public Library; La Casa", 3, False, None,
             4.5, 4.3, 24, "Bridges", True, "First public open house."],
            [2, "Teaching Kitchen — Diabetes Friendly", "J. Tran (Outreach)", date(2026, 4, 5), date(2026, 4, 5), date(2026, 4, 5),
             "Utilization; Agency", "n/a", "Community-Based Org", "HARVEST Lead",
             "Teaching Kitchen — Diabetes Friendly", "Teaching Kitchen", date(2026, 4, 5),
             18, 4, "RWJBH Community Health", 1, False, None,
             4.7, 4.6, 0, None, False, "RD-led; 12-week series start."],
            [3, "Latine Food Heritage Night", "J. Tran (Outreach)", date(2026, 4, 19), date(2026, 4, 19), date(2026, 4, 19),
             "Agency; Access", "Retail-Consumption", "Community-Based Org", "HARVEST Lead",
             "Latine Food Heritage Night", "Pop-Up", date(2026, 4, 19),
             95, 12, "La Casa; Lincoln Park Coast Cultural District", 2, True, "Spanish",
             4.8, 4.5, 35, "Bridges", False, "Member businesses showcased."],
        ],
    },
    {
        "sheet": "Data — Investors-Funders",
        "table": "tbl_Funders",
        "headers": [
            "Submission_ID", "Title", "Submitted_By", "Submitted_Date",
            "Reporting_Period_Start", "Reporting_Period_End", "OFSA_Dimension",
            "Food_System_Sector", "Partner_Sector", "Lead_Role",
            "Funder_Name", "Activity_Type", "Amount_$", "Period_Start",
            "Period_End", "Earned_Revenue_%", "Leveraged_Capital_Source",
            "SROI_Ratio", "Status", "Notes",
        ],
        "widths": [12, 38, 22, 14, 16, 16, 24, 18, 18, 14, 26, 26, 14, 14, 14, 14, 28, 12, 14, 38],
        "samples": [
            [1, "Whole Cities Foundation — 2026 grant", "A. Capece (Director)", date(2026, 2, 12), date(2026, 1, 1), date(2026, 12, 31),
             "Stability", "n/a", "Philanthropy", "HARVEST Lead",
             "Whole Cities Foundation", "Grant", 75000, date(2026, 2, 1), date(2026, 12, 31),
             None, None, None, "Active", "Anchor 2026 grant."],
            [2, "JSC — January 2026", "A. Capece (Director)", date(2026, 1, 18), date(2026, 1, 18), date(2026, 1, 18),
             "Stability", "n/a", "Philanthropy", "Joint",
             "JSC", "JSC Meeting", None, None, None, None, None, None, "Closed", "Reviewed Q4 2025 ops; approved 2026 KPI framework."],
            [3, "JSC — March 2026", "A. Capece (Director)", date(2026, 3, 14), date(2026, 3, 14), date(2026, 3, 14),
             "Stability", "n/a", "Philanthropy", "Joint",
             "JSC", "JSC Meeting", None, None, None, None, None, None, "Closed", "Founding cohort selection."],
            [4, "CAB member — D. Okafor (Healthcare)", "A. Capece (Director)", date(2026, 3, 28), date(2026, 3, 28), date(2026, 3, 28),
             "Agency; Stability", "n/a", "Community-Based Org", "HARVEST Lead",
             "CAB", "CAB Member Added", None, None, None, None, None, None, "Active", "D. Okafor — Healthcare sector."],
            [5, "CAB member — M. Vega (Food Enterprise)", "A. Capece (Director)", date(2026, 4, 4), date(2026, 4, 4), date(2026, 4, 4),
             "Agency; Stability", "n/a", "Community-Based Org", "HARVEST Lead",
             "CAB", "CAB Member Added", None, None, None, None, None, None, "Active", "M. Vega — Food Enterprise sector."],
            [6, "Q1 earned revenue snapshot", "A. Capece (Director)", date(2026, 3, 31), date(2026, 1, 1), date(2026, 3, 31),
             "Stability; Sustainability", "Processing", "Philanthropy", "HARVEST Lead",
             "HARVEST", "Earned Revenue Snapshot", None, date(2026, 1, 1), date(2026, 3, 31),
             18, None, None, "Closed", "Mostly rentals; expect to rise as cohort scales."],
            [7, "Professional development — ServSafe Manager Cert", "A. Capece (Director)", date(2026, 4, 22), date(2026, 4, 22), date(2026, 4, 22),
             "Stability", "n/a", "Philanthropy", "HARVEST Lead",
             "HARVEST", "Professional Development", None, None, None, None, None, None, "Closed", "Director ServSafe Manager Cert — passed."],
        ],
    },
    {
        "sheet": "Data — Farmers",
        "table": "tbl_Farmers",
        "headers": [
            "Submission_ID", "Title", "Submitted_By", "Submitted_Date",
            "Reporting_Period_Start", "Reporting_Period_End", "OFSA_Dimension",
            "Food_System_Sector", "Partner_Sector", "Lead_Role",
            "Farm_Name", "Farm_Location", "Farm_Contact",
            "Produce_Types", "Produce_lbs", "Produce_Value_$",
            "Origin_Region", "Distribution_Destinations",
            "Value_Added_Product", "Value_Added_Description",
            "Demographic_Race_Ethnicity", "Demographic_Gender",
            "Demographic_Age_Band", "Notes",
        ],
        "widths": [12, 36, 18, 14, 16, 16, 26, 16, 14, 12, 22, 22, 24, 28, 12, 14, 22, 32, 14, 26, 26, 16, 14, 28],
        "samples": [
            [1, "Greenfields Farm — March 2026", "R. Owens (UAC)", date(2026, 4, 3), date(2026, 3, 1), date(2026, 3, 31),
             "Availability; Access; Agency", "Aggregation", "Agriculture", "UAC Lead",
             "Greenfields Farm", "Sussex County, NJ", "rowens@greenfields.example",
             "kale; collards; turnips", 480, 1200, "Sussex County, NJ",
             "RWJBH Newark cafeteria; Bridges", False, None,
             "Black/African American", "Man", "45–54", None],
            [2, "Sunrise Acres — March 2026", "R. Owens (UAC)", date(2026, 4, 3), date(2026, 3, 1), date(2026, 3, 31),
             "Availability; Access; Stability", "Aggregation", "Agriculture", "UAC Lead",
             "Sunrise Acres", "Hunterdon County, NJ", "info@sunriseacres.example",
             "tomatoes; peppers", 320, 980, "Hunterdon County, NJ",
             "Newark Public Schools pilot", True, "Pepper hot sauce — Newark Eats LLC",
             "Hispanic/Latino", "Woman", "35–44", "Value-added through Newark Eats LLC."],
            [3, "Liberty Roots — April 2026", "R. Owens (UAC)", date(2026, 5, 2), date(2026, 4, 1), date(2026, 4, 30),
             "Availability; Access; Agency", "Aggregation", "Agriculture", "UAC Lead",
             "Liberty Roots", "Essex County, NJ", "liberty@roots.example",
             "salad greens; radishes", 210, 640, "Essex County, NJ",
             "RWJBH Newark cafeteria", False, None,
             "Black/African American", "Woman", "25–34", "New farmer this period."],
            [4, "Hudson Hill Farm — April 2026", "R. Owens (UAC)", date(2026, 5, 2), date(2026, 4, 1), date(2026, 4, 30),
             "Availability; Sustainability", "Aggregation", "Agriculture", "UAC Lead",
             "Hudson Hill Farm", "Warren County, NJ", "hudson@hill.example",
             "carrots; beets; squash", 540, 1450, "Warren County, NJ",
             "Bridges; Newark Public Schools pilot", False, None,
             "White", "Man", "55–64", None],
        ],
    },
    {
        "sheet": "Data — Environmental",
        "table": "tbl_Environmental",
        "headers": [
            "Submission_ID", "Title", "Submitted_By", "Submitted_Date",
            "Reporting_Period_Start", "Reporting_Period_End", "OFSA_Dimension",
            "Food_System_Sector", "Partner_Sector", "Lead_Role",
            "Activity_Type", "Activity_Date", "lbs_Value", "Dollar_Value",
            "Disposition", "Food_Miles_Avg", "Source_or_Recipient", "Notes",
        ],
        "widths": [12, 36, 18, 14, 16, 16, 24, 18, 22, 14, 24, 14, 12, 14, 22, 12, 26, 28],
        "samples": [
            [1, "Week 12 — waste diverted", "L. Reyes (Logistics)", date(2026, 3, 21), date(2026, 3, 15), date(2026, 3, 21),
             "Sustainability", "Recovery-Waste", "Community-Based Org", "HARVEST Lead",
             "Food Waste Diverted", date(2026, 3, 21), 85, None,
             "Composted On-Site", None, "On-site compost bin", None],
            [2, "Open House — donation to Bridges", "J. Tran (Outreach)", date(2026, 3, 22), date(2026, 3, 22), date(2026, 3, 22),
             "Sustainability; Access", "Recovery-Waste", "Community-Based Org", "HARVEST Lead",
             "Food Waste Diverted", date(2026, 3, 22), 24, None,
             "Donated", None, "Bridges", None],
            [3, "Teaching Kitchen — food captured", "M. Chen (Culinary)", date(2026, 4, 5), date(2026, 4, 5), date(2026, 4, 5),
             "Sustainability; Utilization", "Recovery-Waste", "Healthcare", "HARVEST Lead",
             "Food Captured for Demo", date(2026, 4, 5), 12, 45,
             "Used in Demo", None, "On-site demo kitchen", "Surplus produce from UAC delivery."],
            [4, "Q1 food miles snapshot", "R. Owens (UAC)", date(2026, 3, 31), date(2026, 1, 1), date(2026, 3, 31),
             "Sustainability", "Distribution", "Agriculture", "UAC Lead",
             "Food Miles Snapshot", date(2026, 3, 31), None, None,
             None, 38.5, "Avg across all UAC farms", None],
        ],
    },
]

# Food Corridor import tab — minimal v1 schema (date, member, hours).
# Documented gap per the brief: full schema deferred until Anthony shares
# a real export.
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


# -----------------------------------------------------------------------------
# Tab builders
# -----------------------------------------------------------------------------

def build_readme(wb):
    ws = wb.active
    ws.title = "README"
    ws.sheet_view.showGridLines = False

    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 110

    ws.merge_cells("B2:B2")
    title = ws["B2"]
    title.value = "HARVEST Food Hub — KPI Master Workbook"
    title.font = Font(name="Calibri", size=20, bold=True, color=WHITE)
    title.fill = header_fill()
    title.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 38

    subtitle = ws["B3"]
    subtitle.value = (
        "Real-time KPI tracking across all six audience categories, tagged to "
        "the OFSA Six Dimensions of Food Security. Refreshes from SharePoint."
    )
    subtitle.font = Font(name="Calibri", size=11, italic=True, color=TEXT_DARK)
    subtitle.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[3].height = 22

    sections = [
        ("Open the right tab first",
         "• Going to a JSC meeting? Open '2026 Performance Goals' and "
         "'Dashboard — Executive Summary'.\n"
         "• Working on the OFSA logic model? Open 'Dashboard — By OFSA Dimension'.\n"
         "• Funder report? Open 'Dashboard — By Audience' and filter to the "
         "relevant audience.\n"
         "• Need to know what a metric means? Open 'KPI Catalog'."),
        ("Refresh the data",
         "1. Click the 'Data' tab in the Excel ribbon.\n"
         "2. Click 'Refresh All'.\n"
         "3. Wait for the spinning circles to stop (~10–30 seconds).\n"
         "All dashboards update automatically. Power Query connections to "
         "SharePoint Lists are set up once (see Setup below); after that, "
         "Refresh All is all you do."),
        ("Loading the Food Corridor data",
         "1. In Food Corridor, export the kitchen rentals CSV for the period.\n"
         "2. Open the 'Food Corridor Import' tab in this workbook.\n"
         "3. Follow the instructions at the top of that tab.\n"
         "4. Hit 'Refresh All' — the 'Kitchen hours rented' KPI updates."),
        ("One-time setup",
         "1. Open the 'KPI Catalog' tab — verify the catalog matches your team's intent.\n"
         "2. Provision the SharePoint Lists (run scripts/provision_lists.ps1 OR "
         "follow the manual steps in schemas/sharepoint_lists.md).\n"
         "3. Build the Forms per schemas/microsoft_forms.md.\n"
         "4. Build the Power Automate flows per schemas/power_automate_flows.md.\n"
         "5. In each 'Data — …' tab, replace the seeded sample rows with a "
         "Power Query connection to the corresponding SharePoint List. The M "
         "code is in docs/data_dictionary.md — copy/paste into Data > Get Data > "
         "Blank Query > Advanced Editor."),
        ("Sheet protection",
         "Dashboards are locked from accidental edits. Right-click any locked "
         "cell → 'Unprotect Sheet' (no password) if you need to modify a formula. "
         "Data tabs are unlocked because Power Query needs write access there."),
        ("What's NOT in this workbook",
         "• SOPs and structural documentation (Goal 2) — tracked in the shared "
         "drive folder.\n"
         "• Patient-level health data — out of scope; readmission KPI is a "
         "placeholder until RWJBH provides a methodology.\n"
         "• Power BI dashboards — Excel is v1; Power BI is a later phase."),
        ("Help",
         "Owner: HARVEST Director (anthony@…). For workbook bugs, see "
         "scripts/build_kpi_workbook.py — re-running it regenerates the file "
         "(you'll lose Power Query connections, so back them up first)."),
    ]

    row = 5
    for heading, body in sections:
        h = ws.cell(row=row, column=2, value=heading)
        h.font = Font(name="Calibri", size=13, bold=True, color=DEEP_GREEN)
        h.alignment = Alignment(vertical="center")
        ws.row_dimensions[row].height = 22
        row += 1
        b = ws.cell(row=row, column=2, value=body)
        b.font = Font(name="Calibri", size=11, color=TEXT_DARK)
        b.alignment = Alignment(wrap_text=True, vertical="top")
        line_count = body.count("\n") + 1
        ws.row_dimensions[row].height = max(18 * line_count, 28)
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

    ws.freeze_panes = "A2"

    # Light shade for "Planned" KPIs to flag them visually
    planned_fill = PatternFill("solid", fgColor=STATUS_YELLOW)
    ws.conditional_formatting.add(
        f"A2:{last_col}{last_row}",
        FormulaRule(
            formula=[f'$O2="Planned — methodology TBD"'],
            fill=planned_fill,
        ),
    )

    ws.protection = SheetProtection(sheet=True, password=None,
                                    selectLockedCells=False,
                                    selectUnlockedCells=False)


def build_goals_2026(wb):
    ws = wb.create_sheet("2026 Performance Goals")
    ws.sheet_view.showGridLines = False

    title = ws.cell(row=1, column=1, value="HARVEST — 2026 Performance Goals")
    title.font = Font(name="Calibri", size=18, bold=True, color=WHITE)
    title.fill = header_fill()
    title.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells("A1:G1")
    ws.row_dimensions[1].height = 36

    subtitle = ws.cell(row=2, column=1, value=(
        "Live status against the Director's 2026 annual review goals. "
        "Refresh the workbook (Data → Refresh All) before screenshotting."
    ))
    subtitle.font = Font(name="Calibri", size=11, italic=True, color=TEXT_DARK)
    subtitle.alignment = Alignment(vertical="center", indent=1)
    ws.merge_cells("A2:G2")
    ws.row_dimensions[2].height = 22

    headers = [
        "Goal", "Sub-goal", "Linked metric",
        "Target", "Current", "% Progress", "Status",
    ]
    widths = [32, 38, 32, 12, 14, 14, 14]
    write_header_row(ws, 4, headers, widths)

    for ridx, (goal, sub, metric, target, formula, notes) in enumerate(GOALS_2026, start=5):
        ws.cell(row=ridx, column=1, value=goal).font = body_font()
        ws.cell(row=ridx, column=2, value=sub).font = body_font()
        ws.cell(row=ridx, column=3, value=metric).font = body_font()
        tgt_cell = ws.cell(row=ridx, column=4, value=target)
        tgt_cell.font = body_font()
        tgt_cell.alignment = Alignment(horizontal="center")
        cur_cell = ws.cell(row=ridx, column=5, value=formula)
        cur_cell.font = body_font()
        cur_cell.alignment = Alignment(horizontal="center")

        # % progress
        if target:
            pct_formula = f"=IFERROR(E{ridx}/D{ridx},0)"
            pct_cell = ws.cell(row=ridx, column=6, value=pct_formula)
            pct_cell.number_format = "0%"
            pct_cell.alignment = Alignment(horizontal="center")
            pct_cell.font = body_font()
            # Status
            status_formula = (
                f'=IF(D{ridx}="","",'
                f'IF(F{ridx}>=1,"On track",'
                f'IF(F{ridx}>=0.5,"In progress","At risk")))'
            )
            status_cell = ws.cell(row=ridx, column=7, value=status_formula)
            status_cell.alignment = Alignment(horizontal="center")
            status_cell.font = body_font()
        else:
            ws.cell(row=ridx, column=6, value="").alignment = Alignment(horizontal="center")
            sc = ws.cell(row=ridx, column=7, value=notes)
            sc.font = body_font()
            sc.alignment = Alignment(horizontal="center", wrap_text=True)

        for cidx in range(1, 8):
            ws.cell(row=ridx, column=cidx).alignment = Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[ridx].height = 30

    last_row = 4 + len(GOALS_2026)

    # Conditional formatting on Status column (G)
    ws.conditional_formatting.add(
        f"G5:G{last_row}",
        CellIsRule(operator="equal", formula=['"On track"'],
                   fill=PatternFill("solid", fgColor=STATUS_GREEN))
    )
    ws.conditional_formatting.add(
        f"G5:G{last_row}",
        CellIsRule(operator="equal", formula=['"In progress"'],
                   fill=PatternFill("solid", fgColor=STATUS_YELLOW))
    )
    ws.conditional_formatting.add(
        f"G5:G{last_row}",
        CellIsRule(operator="equal", formula=['"At risk"'],
                   fill=PatternFill("solid", fgColor=STATUS_RED))
    )

    ws.freeze_panes = "A5"
    ws.protection = SheetProtection(sheet=True, password=None,
                                    selectLockedCells=False,
                                    selectUnlockedCells=False)


def build_executive_summary(wb):
    ws = wb.create_sheet("Dashboard — Executive Summary")
    ws.sheet_view.showGridLines = False

    title = ws.cell(row=1, column=1, value="HARVEST — Executive Summary")
    title.font = Font(name="Calibri", size=18, bold=True, color=WHITE)
    title.fill = header_fill()
    title.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells("A1:F1")
    ws.row_dimensions[1].height = 36

    subtitle = ws.cell(row=2, column=1, value=(
        "Top-line metrics across all six audiences. "
        "Refresh: Data → Refresh All."
    ))
    subtitle.font = Font(name="Calibri", size=11, italic=True, color=TEXT_DARK)
    subtitle.alignment = Alignment(vertical="center", indent=1)
    ws.merge_cells("A2:F2")
    ws.row_dimensions[2].height = 22

    # KPI tiles — (label, formula, format)
    tiles = [
        ("Active member businesses",
         '=IFERROR(ROWS(UNIQUE(FILTER(tbl_Entrepreneurs[Business_Name],((tbl_Entrepreneurs[Stage]="Active Member")+(tbl_Entrepreneurs[Stage]="Onboarding"))*(tbl_Entrepreneurs[Business_Name]<>"")))),0)',
         "0"),
        ("Kitchen hours rented (YTD)", "=SUM(tbl_FoodCorridor[Hours])", "0"),
        ("Workshops & TA sessions (YTD)",
         '=COUNTIFS(tbl_Entrepreneurs[Notes],"*workshop*")+COUNTIFS(tbl_Entrepreneurs[Notes],"*TA*")',
         "0"),
        ("Community event participants (YTD)",
         "=SUM(tbl_Community[Attendance_Count])", "0"),
        ("Food recovery to providers (lbs)",
         "=SUM(tbl_Community[Food_Recovery_lbs])+SUMIFS(tbl_Environmental[lbs_Value],tbl_Environmental[Disposition],\"Donated\")",
         "0"),
        ("Food waste diverted (lbs)",
         '=SUMIFS(tbl_Environmental[lbs_Value],tbl_Environmental[Activity_Type],"Food Waste Diverted")',
         "0"),
        ("Grants secured ($)",
         '=SUMIFS(tbl_Funders[Amount_$],tbl_Funders[Activity_Type],"Grant")',
         "$#,##0"),
        ("Farms engaged",
         '=IFERROR(ROWS(UNIQUE(FILTER(tbl_Farmers[Farm_Name],tbl_Farmers[Farm_Name]<>""))),0)',
         "0"),
        ("Produce moved (lbs)", "=SUM(tbl_Farmers[Produce_lbs])", "0"),
    ]

    # Layout: 3 columns of tiles, starting at row 4
    for idx, (label, formula, fmt) in enumerate(tiles):
        col = (idx % 3) * 2 + 1
        row = 4 + (idx // 3) * 4
        # Tile label
        lbl_cell = ws.cell(row=row, column=col, value=label)
        lbl_cell.font = Font(name="Calibri", size=10, bold=True, color=DEEP_GREEN)
        lbl_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        lbl_cell.fill = PatternFill("solid", fgColor=LIGHT_GREEN)
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 1)
        ws.row_dimensions[row].height = 22
        # Tile value
        val_cell = ws.cell(row=row + 1, column=col, value=formula)
        val_cell.font = Font(name="Calibri", size=22, bold=True, color=TEXT_DARK)
        val_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        val_cell.fill = PatternFill("solid", fgColor=LIGHT_GREEN)
        val_cell.number_format = fmt
        ws.merge_cells(start_row=row + 1, start_column=col, end_row=row + 2, end_column=col + 1)
        ws.row_dimensions[row + 1].height = 30
        ws.row_dimensions[row + 2].height = 8  # spacer
        for c in range(col, col + 2):
            ws.column_dimensions[get_column_letter(c)].width = 22

    # OFSA dimension summary (bottom section)
    base = 4 + ((len(tiles) + 2) // 3) * 4 + 1
    ws.cell(row=base, column=1, value="By OFSA Dimension — count of KPIs tagged").font = \
        Font(name="Calibri", size=13, bold=True, color=DEEP_GREEN)
    ws.merge_cells(start_row=base, start_column=1, end_row=base, end_column=6)
    ws.row_dimensions[base].height = 26

    dims = ["Availability", "Access", "Utilization", "Stability", "Agency", "Sustainability"]
    for i, dim in enumerate(dims):
        ws.cell(row=base + 1, column=i + 1, value=dim).font = header_font()
        ws.cell(row=base + 1, column=i + 1).fill = header_fill()
        ws.cell(row=base + 1, column=i + 1).alignment = Alignment(horizontal="center")
        ws.cell(row=base + 2, column=i + 1,
                value=f'=COUNTIF(tbl_KPI_Catalog[OFSA Dimension(s)],"*{dim}*")')
        ws.cell(row=base + 2, column=i + 1).alignment = Alignment(horizontal="center")
        ws.cell(row=base + 2, column=i + 1).font = Font(name="Calibri", size=16, bold=True, color=TEXT_DARK)
    ws.row_dimensions[base + 1].height = 22
    ws.row_dimensions[base + 2].height = 30

    ws.protection = SheetProtection(sheet=True, password=None,
                                    selectLockedCells=False,
                                    selectUnlockedCells=False)


def build_ofsa_dashboard(wb):
    ws = wb.create_sheet("Dashboard — By OFSA Dimension")
    ws.sheet_view.showGridLines = False

    title = ws.cell(row=1, column=1, value="HARVEST — By OFSA Dimension")
    title.font = Font(name="Calibri", size=18, bold=True, color=WHITE)
    title.fill = header_fill()
    title.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells("A1:E1")
    ws.row_dimensions[1].height = 36

    subtitle = ws.cell(row=2, column=1, value=(
        "Every KPI tagged to each of the OFSA Six Dimensions. "
        "Bridge to the logic model phase."
    ))
    subtitle.font = Font(name="Calibri", size=11, italic=True, color=TEXT_DARK)
    subtitle.alignment = Alignment(vertical="center", indent=1)
    ws.merge_cells("A2:E2")
    ws.row_dimensions[2].height = 22

    dims = ["Availability", "Access", "Utilization", "Stability", "Agency", "Sustainability"]
    row = 4
    for dim in dims:
        # Dimension header
        h = ws.cell(row=row, column=1, value=dim)
        h.font = Font(name="Calibri", size=14, bold=True, color=WHITE)
        h.fill = header_fill()
        h.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        ws.row_dimensions[row].height = 28
        row += 1

        # Sub-header
        sub_headers = ["KPI", "Audience", "Unit", "2026 Target", "Status"]
        for i, sh in enumerate(sub_headers):
            c = ws.cell(row=row, column=i + 1, value=sh)
            c.font = Font(name="Calibri", size=10, bold=True, color=TEXT_DARK)
            c.fill = PatternFill("solid", fgColor=LIGHT_GREEN)
            c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[row].height = 18
        row += 1

        # Filter catalog to KPIs tagged with this dimension
        matching = [k for k in KPI_CATALOG if dim in (k[4] or "")]
        for kpi in matching:
            ws.cell(row=row, column=1, value=kpi[0]).font = body_font()
            ws.cell(row=row, column=2, value=kpi[3]).font = body_font()
            ws.cell(row=row, column=3, value=kpi[2]).font = body_font()
            ws.cell(row=row, column=4, value=kpi[10]).font = body_font()
            ws.cell(row=row, column=5, value=kpi[14]).font = body_font()
            for c in range(1, 6):
                ws.cell(row=row, column=c).alignment = Alignment(vertical="top", wrap_text=True)
            row += 1
        row += 1  # spacer

    for i, w in enumerate([42, 24, 8, 12, 22]):
        ws.column_dimensions[get_column_letter(i + 1)].width = w

    ws.freeze_panes = "A4"
    ws.protection = SheetProtection(sheet=True, password=None,
                                    selectLockedCells=False,
                                    selectUnlockedCells=False)


def build_audience_dashboard(wb):
    ws = wb.create_sheet("Dashboard — By Audience")
    ws.sheet_view.showGridLines = False

    title = ws.cell(row=1, column=1, value="HARVEST — By Audience")
    title.font = Font(name="Calibri", size=18, bold=True, color=WHITE)
    title.fill = header_fill()
    title.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells("A1:E1")
    ws.row_dimensions[1].height = 36

    subtitle = ws.cell(row=2, column=1, value=(
        "Every KPI grouped by audience category. "
        "Filter / screenshot a single section for audience-specific reports."
    ))
    subtitle.font = Font(name="Calibri", size=11, italic=True, color=TEXT_DARK)
    subtitle.alignment = Alignment(vertical="center", indent=1)
    ws.merge_cells("A2:E2")
    ws.row_dimensions[2].height = 22

    audiences = [
        "Food Entrepreneurs", "Institution", "Community Organizations",
        "Investors-Funders", "Farmers", "Environmental Sustainability",
    ]
    row = 4
    for aud in audiences:
        h = ws.cell(row=row, column=1, value=aud)
        h.font = Font(name="Calibri", size=14, bold=True, color=WHITE)
        h.fill = header_fill()
        h.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        ws.row_dimensions[row].height = 28
        row += 1

        sub_headers = ["KPI", "Unit", "Cadence", "2026 Target", "Owner"]
        for i, sh in enumerate(sub_headers):
            c = ws.cell(row=row, column=i + 1, value=sh)
            c.font = Font(name="Calibri", size=10, bold=True, color=TEXT_DARK)
            c.fill = PatternFill("solid", fgColor=LIGHT_GREEN)
            c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[row].height = 18
        row += 1

        matching = [k for k in KPI_CATALOG if k[3] == aud]
        for kpi in matching:
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

    # Title
    title = ws.cell(row=1, column=1, value="Food Corridor — CSV Import")
    title.font = Font(name="Calibri", size=18, bold=True, color=WHITE)
    title.fill = header_fill()
    title.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells("A1:F1")
    ws.row_dimensions[1].height = 36

    # Instructions
    instructions = [
        "How to refresh kitchen rental data:",
        "  1. In Food Corridor, go to Reports → Bookings → Export to CSV.",
        "  2. Save the file as 'food_corridor.csv' to the same folder as this workbook.",
        "  3. In Excel, click Data → Refresh All.",
        "  4. The 'Kitchen hours rented' KPI on the Executive Summary updates automatically.",
        "",
        "v1 schema is minimal — only Booking Date, Member Business, and Hours.",
        "When Anthony shares a real Food Corridor export, expand this table to include the",
        "additional columns (rate, total, kitchen/zone, status, etc.) and update the Power Query",
        "in docs/data_dictionary.md.",
        "",
        "If Power Query is not yet wired up, you can paste rows directly below the headers.",
    ]
    for i, line in enumerate(instructions):
        c = ws.cell(row=2 + i, column=1, value=line)
        c.font = Font(name="Calibri", size=11, color=TEXT_DARK,
                      italic=(line.startswith("v1") or line.startswith("When")))
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


def build_data_tab(wb, spec):
    ws = wb.create_sheet(spec["sheet"])
    ws.sheet_view.showGridLines = False

    # Small note at top
    note = ws.cell(row=1, column=1, value=(
        f"Back-end data tab. Populated by Power Query from the corresponding "
        f"SharePoint List. Sample rows seeded so dashboards render before Power Query is wired up. "
        f"Replace via Data → Get Data → Blank Query → paste M code from docs/data_dictionary.md."
    ))
    note.font = Font(name="Calibri", size=10, italic=True, color=TEXT_DARK)
    note.alignment = Alignment(vertical="center", wrap_text=True)
    ws.merge_cells(start_row=1, start_column=1,
                   end_row=1, end_column=len(spec["headers"]))
    ws.row_dimensions[1].height = 32

    header_row = 3
    write_header_row(ws, header_row, spec["headers"], spec["widths"])

    # Sample data
    for ridx, row in enumerate(spec["samples"], start=header_row + 1):
        for cidx, val in enumerate(row, start=1):
            c = ws.cell(row=ridx, column=cidx, value=val)
            c.font = body_font()
            c.alignment = Alignment(vertical="top", wrap_text=True)
            # Date formatting for date-like columns
            header = spec["headers"][cidx - 1]
            if isinstance(val, date):
                c.number_format = "yyyy-mm-dd"
            elif header.endswith("_$") or "Amount" in header or "Value" in header:
                if isinstance(val, (int, float)):
                    c.number_format = "$#,##0"

    last_row = header_row + len(spec["samples"])
    last_col = get_column_letter(len(spec["headers"]))
    add_table(ws, spec["table"], f"A{header_row}:{last_col}{last_row}")

    ws.freeze_panes = f"A{header_row + 1}"


# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------

def main():
    wb = Workbook()
    build_readme(wb)
    build_kpi_catalog(wb)
    build_goals_2026(wb)
    build_executive_summary(wb)
    build_ofsa_dashboard(wb)
    build_audience_dashboard(wb)
    build_food_corridor_import(wb)
    for spec in DATA_TABS:
        build_data_tab(wb, spec)

    # README first; dashboards next; data tabs at the end
    wb.active = 0
    wb.save(OUTPUT_FILE)
    print(f"Wrote {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
