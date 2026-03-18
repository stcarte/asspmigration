# Alaska Air / Hawaiian Air SharePoint Migration

SharePoint Online cross-tenant migration project plan and tooling.

- **Source Tenant:** hawaiianair.sharepoint.com
- **Target Tenant:** alaskaair.sharepoint.com
- **Deadline:** May 30, 2026

## Contents

- `generate_project_plan.py` — Python script to generate the Excel project plan
- `SharePoint-Migration-ProjectPlan.xlsx` — Full migration project workbook (10 sheets)

---

## How to get the Excel file

### Option 1 — Download directly from GitHub (no Python required)

1. Go to the repository on GitHub:
   <https://github.com/stcarte/asspmigration>
2. Click **`SharePoint-Migration-ProjectPlan.xlsx`** in the file list.
3. Click the **Download raw file** button (the down-arrow icon, top-right of the file view).
4. Open the downloaded file in Microsoft Excel, Google Sheets, or LibreOffice Calc.

### Option 2 — Clone the repo

```bash
git clone https://github.com/stcarte/asspmigration.git
cd asspmigration
# The .xlsx file is already present — open it directly:
open SharePoint-Migration-ProjectPlan.xlsx   # macOS
# or
start SharePoint-Migration-ProjectPlan.xlsx  # Windows
```

### Option 3 — Regenerate the file yourself (Python)

Use this if you want to customise the data and rebuild the workbook.

**Prerequisites:** Python 3.8+ with `pip`

```bash
git clone https://github.com/stcarte/asspmigration.git
cd asspmigration
pip install openpyxl
python generate_project_plan.py
```

The script writes `SharePoint-Migration-ProjectPlan.xlsx` to the current directory and
prints a confirmation with the list of sheets created.

---

## Workbook sheets

| Sheet | Description |
|---|---|
| Executive Summary | Tenant details and high-level IFS vs FlightOps comparison |
| Project Plan | 51 tasks across 10 weeks with owners, dates, and status |
| Key Milestones | 12 key milestones from kick-off to go-live |
| IFS Libraries and Lists | All document libraries and lists for the IFS site |
| FlightOps Libraries and Lists | All document libraries and lists for the FlightOps site |
| Risk Register | 7 identified risks with likelihood, impact, and mitigation |
| Go-Live Checklist | 26 pre-go-live validation items |
| Out of Scope and Gaps | 7 items requiring decisions or excluded from scope |
| Site Columns (IFS) | 13 custom site columns and their types |
| Navigation | 47 navigation nodes (Quick Launch + Top Nav) for both sites |