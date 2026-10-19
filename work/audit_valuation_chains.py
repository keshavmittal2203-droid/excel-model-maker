import openpyxl
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TARGETS = ["MM", "LTM", "CLEAN", "EDELWEISS"]

print("="*60)
print("INDEPENDENT AUDIT: WACC, BETA, AND DCF VALUATION CHAINS")
print("="*60)

for t in TARGETS:
    p = ROOT / "companies" / t / "template_v6" / f"{t}-FullTemplate.xlsx"
    wb = openpyxl.load_workbook(p, data_only=False)
    wb_val = openpyxl.load_workbook(p, data_only=True)
    
    dcf = wb["DCF 1st way"]
    dcf_val = wb_val["DCF 1st way"]
    model = wb["Model ->"]
    model_val = wb_val["Model ->"]
    wacc = wb["WACC-DCF"]
    wacc_val = wb_val["WACC-DCF"]
    beta = wb["Beta"]
    beta_val = wb_val["Beta"]
    
    print(f"\n================ {t} ================")
    
    # 1. Cover sheet link
    c9_f = model["C9"].value
    c9_v = model_val["C9"].value
    e42_f = dcf["E42"].value
    e42_v = dcf_val["E42"].value
    print(f"  Cover Sheet 'Model ->'!C9: formula={c9_f} | value={c9_v}")
    print(f"  'DCF 1st way'!E42:         formula={e42_f} | value={e42_v}")
    assert str(c9_f).replace(" ", "") in ["='DCF1stway'!E42", "='DCF 1st way'!E42"], f"Invalid cover link: {c9_f}"
    print("  [PASS] Cover sheet links directly to 'DCF 1st way'!E42")

    # 2. DCF Bridge Cells (E36, E38, E39, E40, E41, E42)
    bridge_cells = ["E36", "E38", "E39", "E40", "E41", "E42"]
    print("  Valuation Bridge:")
    for bc in bridge_cells:
        f = dcf[bc].value
        v = dcf_val[bc].value
        print(f"    {bc}: formula={f} | value={v}")

    # Check for hardcoded numeric literals in bridge cells
    # E38 should be a link to Cash, E39 to Debt, E40 to =E36+E38-E39, E41 to Shares, E42 to =E40/E41
    assert str(dcf["E40"].value).replace(" ", "") == "=E36+E38-E39", f"Invalid E40 formula: {dcf['E40'].value}"
    assert str(dcf["E42"].value).replace(" ", "") == "=E40/E41", f"Invalid E42 formula: {dcf['E42'].value}"
    print("  [PASS] Valuation bridge formulas correctly structured")

    # 3. Terminal Value Formula & Gordon Growth Denominator
    # In template: E35 is PV of Terminal Value, E34 is Terminal Value (or E35 is direct formula)
    # Check E31 (WACC), E20 (Terminal growth), E29 (Terminal year FCFF)
    print("  DCF Key Valuation Drivers:")
    print(f"    WACC: cell E31 formula={dcf['E31'].value} | value={dcf_val['E31'].value}")
    print(f"    Terminal Growth: cell E20 formula={dcf['E20'].value} | value={dcf_val['E20'].value}")
    print(f"    Terminal Year FCFF: cell E29 formula={dcf['E29'].value} | value={dcf_val['E29'].value}")
    print(f"    Terminal Value: cell E34 formula={dcf['E34'].value} | value={dcf_val['E34'].value}")
    print(f"    PV of Terminal Value: cell E35 formula={dcf['E35'].value} | value={dcf_val['E35'].value}")
    print(f"    PV of 5-Yr Forecast: cell E30 formula={dcf['E30'].value} | value={dcf_val['E30'].value}")
    print(f"    Enterprise Value: cell E36 formula={dcf['E36'].value} | value={dcf_val['E36'].value}")

    # 4. WACC-DCF linkage
    print("  WACC-DCF Schedule:")
    print(f"    Cost of Equity (B15): formula={wacc['B15'].value} | value={wacc_val['B15'].value}")
    print(f"    Cost of Debt (B23):   formula={wacc['B23'].value} | value={wacc_val['B23'].value}")
    print(f"    Weight Equity (B26):  formula={wacc['B26'].value} | value={wacc_val['B26'].value}")
    print(f"    Weight Debt (B27):    formula={wacc['B27'].value} | value={wacc_val['B27'].value}")
    print(f"    WACC Output (B28):    formula={wacc['B28'].value} | value={wacc_val['B28'].value}")
    print(f"    DCF WACC Link (E31):  formula={dcf['E31'].value} | value={dcf_val['E31'].value}")

    # 5. Beta Linkage
    print("  Beta Linkage:")
    print(f"    Regression Beta (Beta!P17): formula={beta['P17'].value} | value={beta_val['P17'].value}")
    print(f"    Blume Adjusted (Beta!P19):  formula={beta['P19'].value} | value={beta_val['P19'].value}")
    print(f"    WACC Beta Input (WACC!B12): formula={wacc['B12'].value} | value={wacc_val['B12'].value}")
