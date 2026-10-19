import openpyxl
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TARGETS = ["MM", "LTM", "CLEAN", "EDELWEISS"]

print("="*60)
print("COMPREHENSIVE SCAN FOR UNEXPECTED NEGATIVE VALUES")
print("="*60)

for t in TARGETS:
    p = ROOT / "companies" / t / "template_v6" / f"{t}-FullTemplate.xlsx"
    wb = openpyxl.load_workbook(p, data_only=True)
    print(f"\n--- Checking {t} ---")
    
    for sname in wb.sheetnames:
        ws = wb[sname]
        # Ignore sheets where negative values are expected and legitimate:
        # Cash Flow (outflows are negative), RM (stock return differences can be negative),
        # Beta (returns can be negative), Intrinsic Growth (if growth is negative or deductions)
        if sname in ["Cash Flow", "Beta", "RM"]:
            continue
            
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                v = ws.cell(r, c).value
                if isinstance(v, (int, float)) and v < -0.01:
                    coord = f"{sname}!{openpyxl.utils.get_column_letter(c)}{r}"
                    # Check what row it is
                    lbl = str(ws.cell(r, 2).value or ws.cell(r, 1).value or '')
                    # On Main Financials or Data Sheet, check if it's financing/investing or normal
                    print(f"  NEGATIVE: [{coord}] label='{lbl}': value={v}")
