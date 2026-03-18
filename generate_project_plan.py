"""
generate_project_plan.py
Generates SharePoint-Migration-ProjectPlan.xlsx
Hawaiian Air → Alaska Air cross-tenant SharePoint Online migration project plan.

Usage:
    pip install openpyxl
    python generate_project_plan.py
"""

from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side
)
from openpyxl.formatting.rule import CellIsRule, Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Colour palette
# ---------------------------------------------------------------------------
DARK_BLUE    = "003366"
WHITE        = "FFFFFF"
LIGHT_BLUE   = "E8F0FE"
BORDER_COLOR = "B0B0B0"
GOLD         = "FFD700"
RED_FONT     = "FF0000"

# Conditional-format fills
CF_GREEN  = "C6EFCE"
CF_YELLOW = "FFEB9C"
CF_RED    = "FFC7CE"
CF_ORANGE = "FCE4D6"

# Alternating row fills
FILL_LIGHT = PatternFill("solid", fgColor=LIGHT_BLUE)
FILL_WHITE = PatternFill("solid", fgColor=WHITE)

# Header fill/font
HEADER_FILL = PatternFill("solid", fgColor=DARK_BLUE)
HEADER_FONT = Font(bold=True, color=WHITE, name="Calibri", size=11)

# Thin border
def thin_border():
    side = Side(style="thin", color=BORDER_COLOR)
    return Border(left=side, right=side, top=side, bottom=side)

def header_style(cell, text):
    cell.value = text
    cell.fill  = HEADER_FILL
    cell.font  = HEADER_FONT
    cell.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
    cell.border = thin_border()

def data_style(cell, value, row_idx, date_fmt=False):
    cell.value  = value
    row_fill    = FILL_LIGHT if row_idx % 2 == 0 else FILL_WHITE
    cell.fill   = row_fill
    cell.font   = Font(name="Calibri", size=10)
    cell.alignment = Alignment(vertical="top", wrap_text=True)
    cell.border = thin_border()
    if date_fmt and isinstance(value, datetime):
        cell.number_format = "MM/DD/YYYY"

def set_column_widths(ws, widths):
    for col_idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = w

def add_autofilter(ws, header_row=1):
    ws.auto_filter.ref = ws.dimensions

def freeze_header(ws, cell="A2"):
    ws.freeze_panes = ws[cell]

def add_status_cf(ws, col_letter, first_data_row, last_row):
    """Add conditional formatting to a status column."""
    rng = f"{col_letter}{first_data_row}:{col_letter}{last_row}"

    green_vals  = ["Complete", "Pass", "Closed", "On Track", "Mitigated",
                   "Yes", "Done", "Signed Off"]
    yellow_vals = ["In Progress", "At Risk", "Open", "Not Tested",
                   "Pending", "Not Started"]
    red_vals    = ["Blocked", "Fail", "Critical", "Missed"]
    orange_vals = ["Decision Required"]

    for val in green_vals:
        ws.conditional_formatting.add(
            rng,
            CellIsRule(operator="equal", formula=[f'"{val}"'],
                       fill=PatternFill("solid", fgColor=CF_GREEN)))
    for val in yellow_vals:
        ws.conditional_formatting.add(
            rng,
            CellIsRule(operator="equal", formula=[f'"{val}"'],
                       fill=PatternFill("solid", fgColor=CF_YELLOW)))
    for val in red_vals:
        ws.conditional_formatting.add(
            rng,
            CellIsRule(operator="equal", formula=[f'"{val}"'],
                       fill=PatternFill("solid", fgColor=CF_RED)))
    for val in orange_vals:
        ws.conditional_formatting.add(
            rng,
            CellIsRule(operator="equal", formula=[f'"{val}"'],
                       fill=PatternFill("solid", fgColor=CF_ORANGE)))

# ---------------------------------------------------------------------------
# Sheet 1 – Executive Summary
# ---------------------------------------------------------------------------
def build_executive_summary(wb):
    ws = wb.create_sheet("Executive Summary")

    # Title block
    title_fill = PatternFill("solid", fgColor=DARK_BLUE)
    title_font = Font(bold=True, color=WHITE, name="Calibri", size=14)

    ws.merge_cells("A1:G1")
    c = ws["A1"]
    c.value = "SharePoint Online Cross-Tenant Migration — Project Plan"
    c.fill = title_fill
    c.font = title_font
    c.alignment = Alignment(horizontal="center", vertical="center")

    meta = [
        ("Source Tenant",  "hawaiianair.sharepoint.com"),
        ("Target Tenant",  "alaskaair.sharepoint.com"),
        ("Deadline",       "May 30, 2026"),
        ("Prepared",       "March 17, 2026"),
        ("Generated",      datetime.now().strftime("%B %d, %Y %H:%M")),
    ]

    for i, (label, val) in enumerate(meta, start=2):
        ws.row_dimensions[i].height = 18
        lc = ws.cell(row=i, column=1, value=label)
        lc.font = Font(bold=True, name="Calibri", size=11)
        lc.fill = PatternFill("solid", fgColor="D9E1F2")
        lc.border = thin_border()
        lc.alignment = Alignment(vertical="center")

        vc = ws.cell(row=i, column=2, value=val)
        vc.font = Font(name="Calibri", size=11)
        vc.border = thin_border()
        vc.alignment = Alignment(vertical="center")

    # Comparison table header
    comp_row = 8
    ws.merge_cells(f"A{comp_row}:G{comp_row}")
    hc = ws.cell(row=comp_row, column=1, value="Site Comparison")
    hc.fill = HEADER_FILL
    hc.font = HEADER_FONT
    hc.alignment = Alignment(horizontal="center", vertical="center")

    comp_headers = ["Item", "In-Flight Services (IFS)", "Flight Operations"]
    for col, h in enumerate(comp_headers, start=1):
        header_style(ws.cell(row=comp_row + 1, column=col), h)

    comp_data = [
        ("Source URL",               "/sites/ifs",
                                     "/sites/flightoperations"),
        ("Site Template",            "Communication Site",
                                     "Communication Site"),
        ("Language",                 "English US (1033)",
                                     "English US (1033)"),
        ("Created (source)",         "August 26, 2018",
                                     "June 11, 2020"),
        ("Export Size",              "1,369.9 MB",
                                     "108.8 MB"),
        ("Lists / Libraries",        "122 (103 visible)",
                                     "32 (14 visible)"),
        ("Pages (SitePages)",        "356",
                                     "20 + 27 in library"),
        ("Documents to migrate",     "~1,272 files",
                                     "~81 files"),
        ("Content Types",            "65",
                                     "65"),
        ("Site Columns",             "166",
                                     "Standard only"),
    ]

    for r_idx, row in enumerate(comp_data, start=comp_row + 2):
        for c_idx, val in enumerate(row, start=1):
            data_style(ws.cell(row=r_idx, column=c_idx), val,
                       r_idx - (comp_row + 2))

    set_column_widths(ws, [30, 30, 30, 15, 15, 15, 15])
    ws.row_dimensions[1].height = 30
    freeze_header(ws, "A2")

# ---------------------------------------------------------------------------
# Sheet 2 – Project Plan (51 tasks)
# ---------------------------------------------------------------------------
def build_project_plan(wb):
    ws = wb.create_sheet("Project Plan")

    headers = ["ID", "Phase", "Task", "Site", "Owner",
               "Start Date", "End Date", "Duration (Days)",
               "Status", "% Complete", "Notes/Dependencies"]
    for col, h in enumerate(headers, start=1):
        header_style(ws.cell(row=1, column=col), h)

    # (id, phase, task, site, owner, start, end, dur, status, pct, notes)
    tasks = [
        # PREP
        (1,  "PREP", "Confirm Alaska Air tenant admin URL",
         "BOTH", "IT Admin",
         datetime(2026,3,17), datetime(2026,3,19), 2,
         "Not Started", "0%", ""),
        (2,  "PREP", "Confirm PnP PowerShell version (UseWebLogin vs Interactive)",
         "BOTH", "Engineer",
         datetime(2026,3,17), datetime(2026,3,19), 2,
         "Not Started", "0%", ""),
        (3,  "PREP", "Identify/create migration service account",
         "BOTH", "IT Admin",
         datetime(2026,3,18), datetime(2026,3,21), 3,
         "Not Started", "0%", ""),
        (4,  "PREP", "Export Term Store terms from hawaiianair tenant",
         "BOTH", "Engineer",
         datetime(2026,3,18), datetime(2026,3,21), 3,
         "Not Started", "0%", ""),
        (5,  "PREP",
         "DECISION: Re-populate high-volume list data? "
         "(FlightSchedules 12,950; Deadhead 599; PBS Archives ~2,200+; Events 483)",
         "IFS", "Stakeholder",
         datetime(2026,3,17), datetime(2026,3,24), 5,
         "Not Started", "0%", "Decision required by Mar 24"),
        (6,  "PREP",
         "DECISION: Re-populate FlightOps list data? "
         "(Daily Flight Schedules 4,471; Instructor/Evaluator/LCA 141)",
         "FlightOps", "Stakeholder",
         datetime(2026,3,17), datetime(2026,3,24), 5,
         "Not Started", "0%", "Decision required by Mar 24"),
        (7,  "PREP", "Identify all Power Automate flows on both sites",
         "BOTH", "BA/Owner",
         datetime(2026,3,18), datetime(2026,3,24), 5,
         "Not Started", "0%", ""),
        (8,  "PREP", "Identify AAD groups/users needed in target",
         "BOTH", "IT Admin",
         datetime(2026,3,18), datetime(2026,3,24), 5,
         "Not Started", "0%", ""),
        # WEEK 1-2
        (9,  "Week 1-2: Site Creation & Initial Build",
         "[IFS] Create Communication Site in target tenant",
         "IFS", "IT Admin",
         datetime(2026,3,24), datetime(2026,3,24), 1,
         "Not Started", "0%", ""),
        (10, "Week 1-2: Site Creation & Initial Build",
         "[IFS] Run Build-IFSSite.ps1 — applies template, creates 103 lists/libraries, sets navigation",
         "IFS", "Engineer",
         datetime(2026,3,25), datetime(2026,3,26), 2,
         "Not Started", "0%", "Depends on Task 9"),
        (11, "Week 1-2: Site Creation & Initial Build",
         "[IFS] Verify site is up, Home page loads, Quick Launch visible",
         "IFS", "Engineer",
         datetime(2026,3,26), datetime(2026,3,26), 1,
         "Not Started", "0%", "Depends on Task 10"),
        (12, "Week 1-2: Site Creation & Initial Build",
         "[FlightOps] Create Communication Site",
         "FlightOps", "IT Admin",
         datetime(2026,3,24), datetime(2026,3,24), 1,
         "Not Started", "0%", ""),
        (13, "Week 1-2: Site Creation & Initial Build",
         "[FlightOps] IMMEDIATELY verify Crew Planning section exists — "
         "DAY-ONE PRIORITY (/sites/flightoperations/CrewPlanning)",
         "FlightOps", "Engineer",
         datetime(2026,3,24), datetime(2026,3,24), 1,
         "Not Started", "0%", "⚠ DAY-ONE PRIORITY — must be verified same day as site creation"),
        (14, "Week 1-2: Site Creation & Initial Build",
         "[FlightOps] Run Build-FlightOpsSite.ps1 — applies template, creates 14 lists, sets nav",
         "FlightOps", "Engineer",
         datetime(2026,3,25), datetime(2026,3,26), 2,
         "Not Started", "0%", "Depends on Task 12"),
        (15, "Week 1-2: Site Creation & Initial Build",
         "[FlightOps] Verify all 7 Top Nav nodes resolve "
         "(Home, On the Line, Crew Planning, Publications, Training, Forms, Admin)",
         "FlightOps", "Engineer",
         datetime(2026,3,26), datetime(2026,3,26), 1,
         "Not Started", "0%", "Depends on Task 14"),
        (16, "Week 1-2: Site Creation & Initial Build",
         "[IFS] Upload documents (1,272 files / 1.37 GB) — run off-peak hours",
         "IFS", "Engineer",
         datetime(2026,3,27), datetime(2026,4,2), 5,
         "Not Started", "0%", "Off-peak upload; PnP throttle retry"),
        (17, "Week 1-2: Site Creation & Initial Build",
         "[FlightOps] Upload documents (81 files)",
         "FlightOps", "Engineer",
         datetime(2026,3,27), datetime(2026,3,28), 1,
         "Not Started", "0%", ""),
        (18, "Week 1-2: Site Creation & Initial Build",
         "[FlightOps] Verify Crew Planning page content loaded correctly from template",
         "FlightOps", "Content",
         datetime(2026,3,30), datetime(2026,3,31), 2,
         "Not Started", "0%", ""),
        (19, "Week 1-2: Site Creation & Initial Build",
         "[FlightOps] Verify EFB page renders",
         "FlightOps", "Content",
         datetime(2026,3,30), datetime(2026,3,30), 1,
         "Not Started", "0%", ""),
        # WEEK 3-4
        (20, "Week 3-4: Page Fixes, Navigation & Flows",
         "[IFS] Page review — all 356 pages, fix broken web parts (person, news, list rollup)",
         "IFS", "Content",
         datetime(2026,4,2), datetime(2026,4,14), 10,
         "Not Started", "0%", ""),
        (21, "Week 3-4: Page Fixes, Navigation & Flows",
         "[IFS] Verify and fix all 31 Quick Launch nodes",
         "IFS", "Engineer",
         datetime(2026,4,3), datetime(2026,4,5), 2,
         "Not Started", "0%", ""),
        (22, "Week 3-4: Page Fixes, Navigation & Flows",
         "[IFS] Verify Top Nav (Forms link)",
         "IFS", "Engineer",
         datetime(2026,4,3), datetime(2026,4,3), 1,
         "Not Started", "0%", ""),
        (23, "Week 3-4: Page Fixes, Navigation & Flows",
         "[FlightOps] Page review — all 47 pages",
         "FlightOps", "Content",
         datetime(2026,4,2), datetime(2026,4,10), 7,
         "Not Started", "0%", ""),
        (24, "Week 3-4: Page Fixes, Navigation & Flows",
         "[FlightOps] Crew Planning — deep content review, verify all linked pages under /CrewPlanning",
         "FlightOps", "Content",
         datetime(2026,4,2), datetime(2026,4,7), 4,
         "Not Started", "0%", ""),
        (25, "Week 3-4: Page Fixes, Navigation & Flows",
         "[FlightOps] Verify and fix all 8 Quick Launch nodes",
         "FlightOps", "Engineer",
         datetime(2026,4,3), datetime(2026,4,4), 2,
         "Not Started", "0%", ""),
        (26, "Week 3-4: Page Fixes, Navigation & Flows",
         "[FlightOps] Verify all 7 Top Nav nodes",
         "FlightOps", "Engineer",
         datetime(2026,4,3), datetime(2026,4,4), 2,
         "Not Started", "0%", ""),
        (27, "Week 3-4: Page Fixes, Navigation & Flows",
         "[FlightOps] On the Line (/otl) section review",
         "FlightOps", "Content",
         datetime(2026,4,5), datetime(2026,4,8), 3,
         "Not Started", "0%", ""),
        (28, "Week 3-4: Page Fixes, Navigation & Flows",
         "[FlightOps] Training section review",
         "FlightOps", "Content",
         datetime(2026,4,8), datetime(2026,4,10), 3,
         "Not Started", "0%", ""),
        (29, "Week 3-4: Page Fixes, Navigation & Flows",
         "Re-create Term Store terms (from PREP export)",
         "BOTH", "Engineer",
         datetime(2026,4,2), datetime(2026,4,9), 5,
         "Not Started", "0%", "Depends on Task 4"),
        (30, "Week 3-4: Page Fixes, Navigation & Flows",
         "Rebuild Power Automate flows (GAP-04)",
         "BOTH", "Engineer",
         datetime(2026,4,7), datetime(2026,4,18), 10,
         "Not Started", "0%", "Depends on Task 7"),
        # WEEK 5-6
        (31, "Week 5-6: Permissions, Data & URL Fixes",
         "[IFS] Assign permissions — owners, members, visitors groups from Alaska Air AAD",
         "IFS", "IT Admin",
         datetime(2026,4,14), datetime(2026,4,17), 3,
         "Not Started", "0%", ""),
        (32, "Week 5-6: Permissions, Data & URL Fixes",
         "[FlightOps] Assign permissions",
         "FlightOps", "IT Admin",
         datetime(2026,4,14), datetime(2026,4,17), 3,
         "Not Started", "0%", ""),
        (33, "Week 5-6: Permissions, Data & URL Fixes",
         "[IFS] Re-populate high-volume lists (if approved) — FlightSchedules, Deadhead Booking, PBS Archives",
         "IFS", "Data/BA",
         datetime(2026,4,14), datetime(2026,4,30), 12,
         "Not Started", "0%", "Depends on Task 5 decision"),
        (34, "Week 5-6: Permissions, Data & URL Fixes",
         "[FlightOps] Re-populate lists (if approved) — Daily Flight Schedules, Instructor/Evaluator",
         "FlightOps", "Data/BA",
         datetime(2026,4,14), datetime(2026,4,25), 8,
         "Not Started", "0%", "Depends on Task 6 decision"),
        (35, "Week 5-6: Permissions, Data & URL Fixes",
         "Update embedded URLs / CDN refs in all pages",
         "BOTH", "Content",
         datetime(2026,4,17), datetime(2026,4,24), 5,
         "Not Started", "0%", ""),
        (36, "Week 5-6: Permissions, Data & URL Fixes",
         "Fix person columns / user @domain references",
         "BOTH", "IT Admin",
         datetime(2026,4,17), datetime(2026,4,24), 5,
         "Not Started", "0%", "Depends on Task 8"),
        (37, "Week 5-6: Permissions, Data & URL Fixes",
         "Re-upload / re-link FlightOps video content (85 videos in Stream/SharePoint)",
         "FlightOps", "Content",
         datetime(2026,4,17), datetime(2026,4,25), 7,
         "Not Started", "0%", "GAP-06"),
        # WEEK 7-8
        (38, "Week 7-8: UAT",
         "[IFS] UAT round 1 — content owners test all pages, lists, navigation, search",
         "IFS", "Stakeholder",
         datetime(2026,4,28), datetime(2026,5,5), 6,
         "Not Started", "0%", ""),
        (39, "Week 7-8: UAT",
         "[FlightOps] UAT round 1 — pilots and crew mgmt test Crew Planning, OTL, Training, EFB, Daily Flight Schedules, Forms",
         "FlightOps", "Stakeholder",
         datetime(2026,4,28), datetime(2026,5,5), 6,
         "Not Started", "0%", ""),
        (40, "Week 7-8: UAT",
         "[IFS] UAT defect fixes",
         "IFS", "Engineer",
         datetime(2026,5,4), datetime(2026,5,12), 7,
         "Not Started", "0%", ""),
        (41, "Week 7-8: UAT",
         "[FlightOps] UAT defect fixes",
         "FlightOps", "Engineer",
         datetime(2026,5,4), datetime(2026,5,12), 7,
         "Not Started", "0%", ""),
        (42, "Week 7-8: UAT",
         "[IFS] UAT round 2 — sign-off test pass",
         "IFS", "Stakeholder",
         datetime(2026,5,9), datetime(2026,5,14), 4,
         "Not Started", "0%", ""),
        (43, "Week 7-8: UAT",
         "[FlightOps] UAT round 2 — sign-off test pass",
         "FlightOps", "Stakeholder",
         datetime(2026,5,9), datetime(2026,5,14), 4,
         "Not Started", "0%", ""),
        # WEEK 9-10
        (44, "Week 9-10: Final Checks & Go-Live",
         "Trigger search crawl on both sites",
         "BOTH", "IT Admin",
         datetime(2026,5,12), datetime(2026,5,13), 1,
         "Not Started", "0%", ""),
        (45, "Week 9-10: Final Checks & Go-Live",
         "External integrations re-pointed (GAP-05) — FlightSchedules feed, Deadhead system",
         "BOTH", "IT/Dev",
         datetime(2026,5,12), datetime(2026,5,20), 6,
         "Not Started", "0%", "GAP-05"),
        (46, "Week 9-10: Final Checks & Go-Live",
         "Final content review — spot-check key pages",
         "BOTH", "Stakeholder",
         datetime(2026,5,13), datetime(2026,5,20), 5,
         "Not Started", "0%", ""),
        (47, "Week 9-10: Final Checks & Go-Live",
         "Stakeholder sign-off received",
         "BOTH", "Stakeholder",
         datetime(2026,5,20), datetime(2026,5,20), 1,
         "Not Started", "0%", ""),
        (48, "Week 9-10: Final Checks & Go-Live",
         "Go-live user communications sent",
         "BOTH", "Comms Lead",
         datetime(2026,5,19), datetime(2026,5,22), 3,
         "Not Started", "0%", ""),
        (49, "Week 9-10: Final Checks & Go-Live",
         "[IFS] Site go-live — users redirected",
         "IFS", "IT Admin",
         datetime(2026,5,26), datetime(2026,5,26), 1,
         "Not Started", "0%", "Depends on Task 47"),
        (50, "Week 9-10: Final Checks & Go-Live",
         "[FlightOps] Site go-live — users redirected",
         "FlightOps", "IT Admin",
         datetime(2026,5,27), datetime(2026,5,27), 1,
         "Not Started", "0%", "Depends on Task 47"),
        (51, "Week 9-10: Final Checks & Go-Live",
         "Buffer / contingency for both sites",
         "BOTH", "All",
         datetime(2026,5,27), datetime(2026,5,30), 3,
         "Not Started", "0%", ""),
    ]

    for row_idx, task in enumerate(tasks, start=2):
        (tid, phase, name, site, owner, start, end,
         dur, status, pct, notes) = task

        values = [tid, phase, name, site, owner, start, end, dur,
                  status, pct, notes]

        is_task13 = (tid == 13)

        for col_idx, val in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            is_date = col_idx in (6, 7)
            data_style(cell, val, row_idx, date_fmt=is_date)
            if is_date:
                cell.number_format = "MM/DD/YYYY"

            if is_task13:
                cell.fill = PatternFill("solid", fgColor=GOLD)
                cell.font = Font(bold=True, color=RED_FONT,
                                 name="Calibri", size=10)

    # Column widths
    set_column_widths(ws, [5, 30, 60, 10, 12, 12, 12, 10, 14, 10, 45])

    freeze_header(ws)
    add_autofilter(ws)
    add_status_cf(ws, "I", 2, len(tasks) + 1)

    ws.row_dimensions[1].height = 30
    for r in range(2, len(tasks) + 2):
        ws.row_dimensions[r].height = 40

# ---------------------------------------------------------------------------
# Sheet 3 – Key Milestones
# ---------------------------------------------------------------------------
def build_key_milestones(wb):
    ws = wb.create_sheet("Key Milestones")

    headers = ["#", "Milestone Date", "Milestone Description", "Status"]
    for col, h in enumerate(headers, start=1):
        header_style(ws.cell(row=1, column=col), h)

    milestones = [
        (1,  datetime(2026,3,24), "Both sites created in target tenant",                              "Not Started"),
        (2,  datetime(2026,3,24), "Crew Planning verified accessible (FlightOps day-one check)",       "Not Started"),
        (3,  datetime(2026,3,26), "Both build scripts completed — structure live",                     "Not Started"),
        (4,  datetime(2026,4,2),  "All documents uploaded",                                            "Not Started"),
        (5,  datetime(2026,4,10), "All pages reviewed, navigation verified on both sites",             "Not Started"),
        (6,  datetime(2026,4,18), "Power Automate flows rebuilt",                                      "Not Started"),
        (7,  datetime(2026,4,30), "Data re-population complete (if approved)",                         "Not Started"),
        (8,  datetime(2026,5,14), "UAT complete, all defects resolved",                                "Not Started"),
        (9,  datetime(2026,5,20), "Final sign-off from stakeholders",                                  "Not Started"),
        (10, datetime(2026,5,26), "IFS go-live",                                                       "Not Started"),
        (11, datetime(2026,5,27), "Flight Operations go-live",                                         "Not Started"),
        (12, datetime(2026,5,30), "DEADLINE — both sites live in target tenant",                       "Not Started"),
    ]

    for row_idx, (num, date, desc, status) in enumerate(milestones, start=2):
        for col_idx, val in enumerate([num, date, desc, status], start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            data_style(cell, val, row_idx, date_fmt=(col_idx == 2))
            if col_idx == 2:
                cell.number_format = "MM/DD/YYYY"

    set_column_widths(ws, [5, 16, 60, 14])
    freeze_header(ws)
    add_autofilter(ws)
    add_status_cf(ws, "D", 2, len(milestones) + 1)

# ---------------------------------------------------------------------------
# Sheet 4 – IFS Libraries and Lists
# ---------------------------------------------------------------------------
def build_ifs_libraries(wb):
    ws = wb.create_sheet("IFS Libraries and Lists")

    headers = ["#", "Name", "Type", "Items",
               "Priority/Decision", "Migration Notes", "Status"]
    for col, h in enumerate(headers, start=1):
        header_style(ws.cell(row=1, column=col), h)

    rows = []

    # Document Libraries
    doc_libs = [
        ("Admin", 26),
        ("Briefing Sheets", 166),
        ("Cabin Seat Agreements", 20),
        ("Catering and Service", 51),
        ("Catering and Service Communication", 55),
        ("Corporate Safety and Security", 22),
        ("Credit Card", 31),
        ("Crew Scheduling", 81),
        ("Documents", 25),
        ("FAQ", 8),
        ("Graphics Library", 167),
        ("IFP Communication", 373),
        ("IFS Training - Practical Training", 14),
        ("IFS Training - Recurrent Awards", 304),
        ("Safety Communication", 20),
        ("Safety Reporting Tool", 2),
        ("Shared Documents", 4),
        ("Uniform and Grooming", 3),
        ("Vacation", 24),
    ]
    for name, items in doc_libs:
        rows.append((name, "Document Library", items, "Migrate", "", "Not Started"))

    # Custom/Specialty Lists
    custom_lists = [
        ("2023 P and V Awards", 2056),
        ("FA Base Month 2023 for IFS", 1788),
        ("FA Base Month 2024", 2188),
        ("FA Base Month 2024 v2", 2188),
        ("flying to/from", 18),
        ("RC3 Door Openings and Closings Playlist", 8),
        ("Revision Cycle 2 Playlist", 3),
        ("Town Hall Questions", 0),
        ("B2B Library", 11),
    ]
    for name, items in custom_lists:
        rows.append((name, "Custom List", items, "Migrate", "", "Not Started"))

    # Archive Lists
    archive_lists = [
        # 17.N. Contractual Leave Archives 2019-2026
        ("17.N. Contractual Leave Archive 2019", "Archive List", 0, "Low priority", "", "Not Started"),
        ("17.N. Contractual Leave Archive 2020", "Archive List", 0, "Low priority", "", "Not Started"),
        ("17.N. Contractual Leave Archive 2021", "Archive List", 0, "Low priority", "", "Not Started"),
        ("17.N. Contractual Leave Archive 2022", "Archive List", 0, "Low priority", "", "Not Started"),
        ("17.N. Contractual Leave Archive 2023", "Archive List", 0, "Low priority", "", "Not Started"),
        ("17.N. Contractual Leave Archive 2024", "Archive List", 0, "Low priority", "", "Not Started"),
        ("17.N. Contractual Leave Archive 2025", "Archive List", 0, "Low priority", "", "Not Started"),
        ("17.N. Contractual Leave Archive 2026", "Archive List", 0, "Low priority", "", "Not Started"),
        # Low Time Flying Archives
        ("Low Time Flying Archive 2019", "Archive List", 0, "Low priority", "", "Not Started"),
        ("Low Time Flying Archive 2020", "Archive List", 0, "Low priority", "", "Not Started"),
        ("Low Time Flying Archive 2021", "Archive List", 0, "Low priority", "", "Not Started"),
        ("Low Time Flying Archive 2022", "Archive List", 0, "Low priority", "", "Not Started"),
        ("Low Time Flying Archive 2023", "Archive List", 0, "Low priority", "", "Not Started"),
        ("Low Time Flying Archive 2024", "Archive List", 0, "Low priority", "", "Not Started"),
        ("Low Time Flying Archive 2025", "Archive List", 0, "Low priority", "", "Not Started"),
        ("Low Time Flying Archive 2026", "Archive List", 0, "Low priority", "", "Not Started"),
        # PBS Award Archives 2016-2019
        ("PBS Award Archive 2016", "Archive List", 0, "Low priority", "", "Not Started"),
        ("PBS Award Archive 2017", "Archive List", 0, "Low priority", "", "Not Started"),
        ("PBS Award Archive 2018", "Archive List", 0, "Low priority", "", "Not Started"),
        ("PBS Award Archive 2019", "Archive List", 0, "Low priority", "", "Not Started"),
        # PBS Award Archives 2020-2026
        ("PBS Award Archive 2020", "Archive List", 0, "Low priority", "", "Not Started"),
        ("PBS Award Archive 2021", "Archive List", 0, "Low priority", "", "Not Started"),
        ("PBS Award Archive 2022", "Archive List", 0, "Low priority", "", "Not Started"),
        ("PBS Award Archive 2023", "Archive List", 0, "Low priority", "", "Not Started"),
        ("PBS Award Archive 2024", "Archive List", 0, "Low priority", "", "Not Started"),
        ("PBS Award Archive 2025", "Archive List", 0, "Low priority", "", "Not Started"),
        ("PBS Award Archive 2026", "Archive List", 0, "Low priority", "", "Not Started"),
        # Vacancy Bid Archives 2021-2026
        ("Vacancy Bid Archive 2021", "Archive List", 0, "Low priority", "", "Not Started"),
        ("Vacancy Bid Archive 2022", "Archive List", 0, "Low priority", "", "Not Started"),
        ("Vacancy Bid Archive 2023", "Archive List", 0, "Low priority", "", "Not Started"),
        ("Vacancy Bid Archive 2024", "Archive List", 0, "Low priority", "", "Not Started"),
        ("Vacancy Bid Archive 2025", "Archive List", 0, "Low priority", "", "Not Started"),
        ("Vacancy Bid Archive 2026", "Archive List", 0, "Low priority", "", "Not Started"),
        # Vacation Annual Bidding 2021-2026
        ("Vacation Annual Bidding 2021", "Archive List", 0, "Low priority", "", "Not Started"),
        ("Vacation Annual Bidding 2022", "Archive List", 0, "Low priority", "", "Not Started"),
        ("Vacation Annual Bidding 2023", "Archive List", 0, "Low priority", "", "Not Started"),
        ("Vacation Annual Bidding 2024", "Archive List", 0, "Low priority", "", "Not Started"),
        ("Vacation Annual Bidding 2025", "Archive List", 0, "Low priority", "", "Not Started"),
        ("Vacation Annual Bidding 2026", "Archive List", 0, "Low priority", "", "Not Started"),
        # Others
        ("Vacation (Bonus Days)", "Archive List", 0, "Low priority", "", "Not Started"),
        ("LOA", "Archive List", 0, "Low priority", "", "Not Started"),
        ("COVID LOA", "Archive List", 0, "Low priority", "", "Not Started"),
    ]
    rows.extend(archive_lists)

    # High-Volume Data Lists
    high_vol = [
        ("FlightSchedules", "High-Volume List", 12950,
         "DECISION REQUIRED", "Re-populate from source or start fresh?", "Decision Required"),
        ("Deadhead Booking", "High-Volume List", 599,
         "DECISION REQUIRED", "Re-populate from source or start fresh?", "Decision Required"),
        ("IFP Communication list", "High-Volume List", 373,
         "DECISION REQUIRED", "Re-populate from source or start fresh?", "Decision Required"),
        ("Events/Calendar", "High-Volume List", 483,
         "Migrate", "", "Not Started"),
        ("Random Acts of Aloha", "High-Volume List", 492,
         "DECISION REQUIRED", "Re-populate from source or start fresh?", "Decision Required"),
        ("In-Flight Training Videos", "High-Volume List", 197,
         "Migrate", "", "Not Started"),
        ("Daily Briefings", "High-Volume List", 106,
         "Migrate", "", "Not Started"),
    ]
    rows.extend(high_vol)

    for r_idx, row in enumerate(rows, start=2):
        if len(row) == 6:
            name, typ, items, priority, notes, status = row
        else:
            name, typ, items, priority, notes, status = row

        vals = [r_idx - 1, name, typ, items, priority, notes, status]
        for c_idx, val in enumerate(vals, start=1):
            data_style(ws.cell(row=r_idx, column=c_idx), val, r_idx)

    set_column_widths(ws, [5, 40, 20, 8, 20, 45, 18])
    freeze_header(ws)
    add_autofilter(ws)
    add_status_cf(ws, "G", 2, len(rows) + 1)

# ---------------------------------------------------------------------------
# Sheet 5 – FlightOps Libraries and Lists
# ---------------------------------------------------------------------------
def build_flightops_libraries(wb):
    ws = wb.create_sheet("FlightOps Libraries and Lists")

    headers = ["#", "Name", "Type", "Items",
               "Priority/Decision", "Migration Notes", "Status"]
    for col, h in enumerate(headers, start=1):
        header_style(ws.cell(row=1, column=col), h)

    rows = [
        (1, "Documents",               "Document Library", 8,    "Migrate",           "", "Not Started"),
        (2, "Flight Operations Videos","Document Library", 85,   "Migrate (GAP-06)",  "Re-upload to Stream/SharePoint target", "Not Started"),
        (3, "KCM",                     "Document Library", 5,    "Migrate",           "", "Not Started"),
        (4, "Security",                "Document Library", 4,    "Migrate",           "", "Not Started"),
        (5, "Shared Documents",        "Document Library", 0,    "Migrate",           "", "Not Started"),
        (6, "Daily Flight Schedules",  "High-Volume List", 4471, "DECISION REQUIRED", "Re-populate from source or start fresh?", "Decision Required"),
        (7, "Demo-Daily Flight Schedules","Custom List",   560,  "Migrate",           "", "Not Started"),
        (8, "Instructor/Evaluator/LCA","High-Volume List", 141,  "DECISION REQUIRED", "Re-populate from source or start fresh?", "Decision Required"),
        (9, "Technical Pilot Application List","Custom List", 1, "Migrate",           "", "Not Started"),
    ]

    for row in rows:
        r_idx = row[0] + 1
        vals = list(row)
        for c_idx, val in enumerate(vals, start=1):
            data_style(ws.cell(row=r_idx, column=c_idx), val, r_idx)

    set_column_widths(ws, [5, 40, 20, 8, 20, 45, 18])
    freeze_header(ws)
    add_autofilter(ws)
    add_status_cf(ws, "G", 2, len(rows) + 1)

# ---------------------------------------------------------------------------
# Sheet 6 – Risk Register
# ---------------------------------------------------------------------------
def build_risk_register(wb):
    ws = wb.create_sheet("Risk Register")

    headers = ["Risk ID", "Risk Description", "Likelihood", "Impact",
               "Risk Score", "Mitigation", "Owner", "Status"]
    for col, h in enumerate(headers, start=1):
        header_style(ws.cell(row=1, column=col), h)

    risks = [
        ("RISK-01",
         "No tenant admin access to alaskaair.sharepoint.com",
         "HIGH", "HIGH", "Critical",
         "Escalate to Alaska Air IT Admin immediately. "
         "Build-IFSSite.ps1 works if admin creates sites manually and grants "
         "Site Collection Admin rights to engineer.",
         "IT Admin", "Open"),
        ("RISK-02",
         "PnP template apply fails on target (pages/branding errors)",
         "MEDIUM", "MEDIUM", "Medium",
         "Script has 4 fallback handler exclusion attempts. "
         "List/library build runs independently and is not blocked by template failure.",
         "Engineer", "Open"),
        ("RISK-03",
         "High-volume list data (FlightSchedules, PBS) not available",
         "HIGH", "HIGH", "Critical",
         "Decision needed by Mar 24. If re-population required, contract a data migration "
         "tool (ShareGate, Metalogix) or export CSV from source.",
         "Stakeholder", "Open"),
        ("RISK-04",
         "Managed metadata / Term Store not set up in target tenant",
         "MEDIUM", "MEDIUM", "Medium",
         "Export terms from source using Export-PnPTermGroupToXml and import to target "
         "before applying templates.",
         "Engineer", "Open"),
        ("RISK-05",
         "User accounts do not exist in target tenant AAD",
         "HIGH", "MEDIUM", "High",
         "Audit all person columns and web part references. "
         "Create/sync accounts before UAT starts.",
         "IT Admin", "Open"),
        ("RISK-06",
         "External video embeds / Stream links broken",
         "HIGH", "MEDIUM", "High",
         "Re-upload videos to SharePoint/Stream in target tenant and update embed URLs.",
         "Content", "Open"),
        ("RISK-07",
         "IFS export is 1.37 GB — upload may be slow / throttled",
         "MEDIUM", "LOW", "Low",
         "Run uploads during off-peak hours. Script handles throttling via PnP retry logic.",
         "Engineer", "Open"),
    ]

    for r_idx, row in enumerate(risks, start=2):
        for c_idx, val in enumerate(row, start=1):
            data_style(ws.cell(row=r_idx, column=c_idx), val, r_idx)

    set_column_widths(ws, [10, 50, 12, 12, 12, 60, 12, 14])
    freeze_header(ws)
    add_autofilter(ws)
    add_status_cf(ws, "E", 2, len(risks) + 1)   # Risk Score
    add_status_cf(ws, "H", 2, len(risks) + 1)   # Status

# ---------------------------------------------------------------------------
# Sheet 7 – Go-Live Checklist
# ---------------------------------------------------------------------------
def build_go_live_checklist(wb):
    ws = wb.create_sheet("Go-Live Checklist")

    headers = ["#", "Category", "Checklist Item",
               "Status", "Tested By", "Date Tested", "Notes"]
    for col, h in enumerate(headers, start=1):
        header_style(ws.cell(row=1, column=col), h)

    items = []

    ifs_items = [
        "Home page loads without errors",
        "All 15 Quick Launch links resolve",
        "Bidding (PBS) page functional",
        "Daily Briefing visible and current",
        "Training page loads correctly",
        "FlightSchedules list accessible with correct columns",
        "Search returns results from site content",
        '"Contact Us" form works (if Power Automate flow rebuilt)',
        "Mobile view renders correctly",
        "Permissions confirmed (members can view not edit unless intended)",
    ]
    for item in ifs_items:
        items.append(("IFS", item))

    fo_items = [
        "Home page loads without errors",
        "EFB page loads",
        "On the Line sub-site navigation works",
        "Crew Planning section functional",
        "Daily Flight Schedules list accessible",
        "Instructor / Evaluator list accessible with correct columns",
        "Top Navigation all 7 links resolve",
        "Forms section functional",
        "Search returns results",
    ]
    for item in fo_items:
        items.append(("Flight Operations", item))

    both_items = [
        "No @hawaiianair.com user references visible in web parts",
        "No broken images on any key page",
        "SSL / HTTPS valid",
        "Site logos / branding updated to Alaska Air",
        "Power Automate flows tested end-to-end",
        "External integrations verified",
        "Stakeholder sign-off received",
    ]
    for item in both_items:
        items.append(("Both Sites", item))

    for r_idx, (cat, desc) in enumerate(items, start=2):
        row_vals = [r_idx - 1, cat, desc, "Not Tested", "", "", ""]
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws.cell(row=r_idx, column=c_idx)
            data_style(cell, val, r_idx)

    set_column_widths(ws, [5, 18, 60, 14, 14, 14, 30])
    freeze_header(ws)
    add_autofilter(ws)
    add_status_cf(ws, "D", 2, len(items) + 1)

# ---------------------------------------------------------------------------
# Sheet 8 – Out of Scope and Gaps
# ---------------------------------------------------------------------------
def build_gaps(wb):
    ws = wb.create_sheet("Out of Scope and Gaps")

    headers = ["GAP ID", "Description", "Impact",
               "Decision/Action Required", "Owner", "Status"]
    for col, h in enumerate(headers, start=1):
        header_style(ws.cell(row=1, column=col), h)

    gaps = [
        ("GAP-01",
         "List item data for high-volume lists not exported",
         "HIGH",
         "DECISION REQUIRED: re-populate from source or start fresh?",
         "Stakeholder", "Decision Required"),
        ("GAP-02",
         "User profile data references old @hawaiianair.com accounts",
         "MEDIUM",
         "Re-map to @alaskaair.com accounts",
         "IT Admin", "Open"),
        ("GAP-03",
         "Permissions not migrated, all security groups must be rebuilt",
         "HIGH",
         "Create all AAD groups/users in target before UAT",
         "IT Admin", "Open"),
        ("GAP-04",
         "Power Automate / Logic Apps not captured in schema export",
         "MEDIUM",
         "Identify all flows; rebuild in target tenant",
         "Engineer", "Open"),
        ("GAP-05",
         "External integrations need re-pointing",
         "HIGH",
         "Identify all integration endpoints; re-point to new site URLs",
         "IT/Dev", "Open"),
        ("GAP-06",
         "Video content (85 items) likely reference broken Stream/embed URLs",
         "MEDIUM",
         "Re-upload videos and update embed URLs",
         "Content", "Open"),
        ("GAP-07",
         "Taxonomy / Term Store — TaxCatchAll columns suggest managed metadata in use",
         "MEDIUM",
         "Export terms; import to target before applying templates",
         "Engineer", "Open"),
    ]

    for r_idx, row in enumerate(gaps, start=2):
        for c_idx, val in enumerate(row, start=1):
            data_style(ws.cell(row=r_idx, column=c_idx), val, r_idx)

    set_column_widths(ws, [10, 50, 10, 50, 12, 18])
    freeze_header(ws)
    add_autofilter(ws)
    add_status_cf(ws, "F", 2, len(gaps) + 1)

# ---------------------------------------------------------------------------
# Sheet 9 – Site Columns (IFS)
# ---------------------------------------------------------------------------
def build_site_columns(wb):
    ws = wb.create_sheet("Site Columns (IFS)")

    headers = ["#", "InternalName", "Type", "Notes", "Created in Target?"]
    for col, h in enumerate(headers, start=1):
        header_style(ws.cell(row=1, column=col), h)

    columns = [
        (1,  "ComplianceAssetId",    "Text",          "Compliance asset identifier",           "Not Started"),
        (2,  "IsFeatured",           "Boolean",       "Feature flag for content",               "Not Started"),
        (3,  "TaskOutcome",          "OutcomeChoice", "Task outcome field",                     "Not Started"),
        (4,  "TaxKeywordTaxHTField", "Note",          "Taxonomy keyword hidden field",          "Not Started"),
        (5,  "TaxCatchAll",          "LookupMulti",   "Managed metadata catch-all (hidden)",    "Not Started"),
        (6,  "TaxCatchAllLabel",     "LookupMulti",   "Managed metadata label (hidden)",        "Not Started"),
        (7,  "WSEnabled",            "Boolean",       "Web service enabled flag",               "Not Started"),
        (8,  "WSPublishState",       "Integer",       "Web service publish state",              "Not Started"),
        (9,  "MyEditor",             "User",          "Custom editor user field",               "Not Started"),
        (10, "ParentFolderId",       "Integer",       "Parent folder identifier",               "Not Started"),
        (11, "ParentID",             "Lookup",        "Parent item ID lookup",                  "Not Started"),
        (12, "ParentItemEditor",     "User",          "Parent item editor user field",          "Not Started"),
        (13, "ParentItemID",         "Integer",       "Parent item integer ID",                 "Not Started"),
    ]

    for r_idx, row in enumerate(columns, start=2):
        for c_idx, val in enumerate(row, start=1):
            data_style(ws.cell(row=r_idx, column=c_idx), val, r_idx)

    set_column_widths(ws, [5, 25, 16, 40, 18])
    freeze_header(ws)
    add_autofilter(ws)
    add_status_cf(ws, "E", 2, len(columns) + 1)

# ---------------------------------------------------------------------------
# Sheet 10 – Navigation
# ---------------------------------------------------------------------------
def build_navigation(wb):
    ws = wb.create_sheet("Navigation")

    headers = ["#", "Site", "Nav Type", "Node Name", "URL", "Verified?"]
    for col, h in enumerate(headers, start=1):
        header_style(ws.cell(row=1, column=col), h)

    nav_rows = []

    # IFS Quick Launch
    ifs_ql = [
        ("Home",                        "/sites/ifs"),
        ("NEWS",                        "/sites/ifs/SitePages/News-Hub.aspx"),
        ("Daily Briefing",              "/sites/ifs/SitePages/Daily-Briefing.aspx"),
        ("Bidding (PBS)",               "/sites/ifs/SitePages/PBS-BIDDING.aspx"),
        ("Crew Scheduling",             "/sites/ifs/SitePages/Scheduling.aspx"),
        ("Catering and Service",        "/sites/ifs/SitePages/Catering-and-Service.aspx"),
        ("IFM Safety and Security",     "/sites/ifs/SitePages/IFM-Cabin-Safety.aspx"),
        ("Training",                    "/sites/ifs/SitePages/The-NEW-Train.aspx"),
        ("IMD",                         "/sites/ifs/SitePages/IMD.aspx"),
        ("Uniform and Grooming",        "/sites/ifs/SitePages/Uniform.aspx"),
        ("Forms",                       "/sites/ifs/SitePages/Forms.aspx"),
        ("Crew Accommodations",         "/sites/ifs/SitePages/Destinations.aspx"),
        ("Wings of Wellness",           "/sites/ifs/SitePages/Flight-Attendant-Mental-Health.aspx"),
        ("Inflight Admin",              "/sites/ifs/SitePages/Admin.aspx"),
        ("Contact Us!",                 "/sites/ifs/SitePages/Contact-Us!.aspx"),
        ("Other Links (header)",        ""),
        ("PBS Archive 2016",            "/sites/ifs/Lists/PBSArchive2016"),
        ("PBS Archive 2017",            "/sites/ifs/Lists/PBSArchive2017"),
        ("PBS Archive 2018",            "/sites/ifs/Lists/PBSArchive2018"),
        ("PBS Archive 2019",            "/sites/ifs/Lists/PBSArchive2019"),
        ("PBS Archive 2020",            "/sites/ifs/Lists/PBSArchive2020"),
        ("PBS Archive 2021",            "/sites/ifs/Lists/PBSArchive2021"),
        ("PBS Archive 2022",            "/sites/ifs/Lists/PBSArchive2022"),
        ("PBS Archive 2023",            "/sites/ifs/Lists/PBSArchive2023"),
        ("Vacancy Bid 2021",            "/sites/ifs/Lists/VacancyBid2021"),
        ("Vacancy Bid 2022",            "/sites/ifs/Lists/VacancyBid2022"),
        ("Vacancy Bid 2023",            "/sites/ifs/Lists/VacancyBid2023"),
        ("LTF Archive",                 "/sites/ifs/Lists/LTFArchive"),
        ("Vacation Archive",            "/sites/ifs/Lists/VacationArchive"),
        ("Leave Archive",               "/sites/ifs/Lists/LeaveArchive"),
        ("Recent (auto-node)",          ""),
    ]
    for name, url in ifs_ql:
        nav_rows.append(("IFS", "Quick Launch", name, url, "Not Started"))

    # IFS Top Navigation
    nav_rows.append(("IFS", "Top Navigation", "Forms", "/sites/ifs/Forms", "Not Started"))

    # FlightOps Quick Launch
    fo_ql = [
        ("Home",               "/sites/flightoperations"),
        ("On the Line",        "/sites/flightoperations/otl"),
        ("Crew Planning",      "/sites/flightoperations/CrewPlanning"),
        ("Training",           "/sites/flightoperations/training"),
        ("Crew Accommodations","/sites/flightoperations/crewaccommodations"),
        ("EFB",                "/sites/flightoperations/SitePages/EFB.aspx"),
        ("Forms",              "/sites/flightoperations/forms"),
        ("Admin",              "/sites/flightoperations/admin"),
    ]
    for name, url in fo_ql:
        nav_rows.append(("FlightOps", "Quick Launch", name, url, "Not Started"))

    # FlightOps Top Navigation
    fo_topnav = [
        ("Home",         "/sites/flightoperations"),
        ("On the Line",  "/sites/flightoperations/otl"),
        ("Crew Planning","/sites/flightoperations/CrewPlanning"),
        ("Publications", "/sites/flightoperations/publications"),
        ("Training",     "/sites/flightoperations/training"),
        ("Forms",        "/sites/flightoperations/forms"),
        ("Admin",        "/sites/flightoperations/admin"),
    ]
    for name, url in fo_topnav:
        nav_rows.append(("FlightOps", "Top Navigation", name, url, "Not Started"))

    for r_idx, (site, nav_type, node, url, verified) in \
            enumerate(nav_rows, start=2):
        vals = [r_idx - 1, site, nav_type, node, url, verified]
        for c_idx, val in enumerate(vals, start=1):
            data_style(ws.cell(row=r_idx, column=c_idx), val, r_idx)

    set_column_widths(ws, [5, 12, 16, 35, 60, 12])
    freeze_header(ws)
    add_autofilter(ws)
    add_status_cf(ws, "F", 2, len(nav_rows) + 1)

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    build_executive_summary(wb)
    build_project_plan(wb)
    build_key_milestones(wb)
    build_ifs_libraries(wb)
    build_flightops_libraries(wb)
    build_risk_register(wb)
    build_go_live_checklist(wb)
    build_gaps(wb)
    build_site_columns(wb)
    build_navigation(wb)

    output = "SharePoint-Migration-ProjectPlan.xlsx"
    wb.save(output)
    print(f"✅  Saved: {output}")
    print(f"   Sheets: {[s.title for s in wb.worksheets]}")

if __name__ == "__main__":
    main()
