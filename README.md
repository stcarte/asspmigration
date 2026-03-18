# Alaska Air / Hawaiian Air SharePoint Migration

SharePoint Online cross-tenant migration project plan and tooling.

- **Source Tenant:** hawaiianair.sharepoint.com
- **Target Tenant:** alaskaair.sharepoint.com
- **Deadline:** May 30, 2026

## Contents

- `generate_project_plan.py` — Python script to generate the Excel project plan
- `SharePoint-Migration-ProjectPlan.xlsx` — Full migration project workbook (10 sheets)

## Usage

```bash
pip install openpyxl
python generate_project_plan.py
```