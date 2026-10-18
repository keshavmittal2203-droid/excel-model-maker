import openpyxl
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TARGETS = ["MM", "LTM", "CLEAN", "EDELWEISS"]

print("="*60)
print("DEEP AUDIT: BALANCE SHEET INTEGRITY ACROSS ALL 4 WORKBOOKS")
print("="*60)

cols = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"]

for t in TARGETS:
    p = ROOT / "companies" / t / "template_v6" / f"{t}-FullTemplate.xlsx"
    wb = openpyxl.load_workbook(p, data_only=True)
    bs = wb["BS"]
    
    print(f"\n================ TARGET: {t} ================")
    
    # 1. Check for negative numbers
    negatives = []
    for r in range(4, 39):
        label = bs.cell(r, 2).value
        for c in cols:
            val = bs[f"{c}{r}"].value
            if isinstance(val, (int, float)) and val < -0.01:
                col_hdr = bs[f"{c}2"].value
                negatives.append((f"{c}{r}", col_hdr, label, val))
                
    if negatives:
        print(f"  [FAIL] Negative values found ({len(negatives)}):")
        for coord, hdr, lbl, val in negatives:
            print(f"    {coord} ({hdr} - {lbl}): {val}")
    else:
        print("  [PASS] Zero negative values in Balance Sheet (Rows 4-38, Cols C-N)")

    # 2. Check component sums and balance sheet ties
    bs_ties = []
    comp_ties_liab = []
    comp_ties_assets = []
    
    for c in cols:
        hdr = bs[f"{c}2"].value
        if not hdr:
            continue
            
        tot_liab = bs[f"{c}16"].value
        tot_asset = bs[f"{c}38"].value
        
        if tot_liab is not None and tot_asset is not None:
            diff = round(abs(float(tot_liab) - float(tot_asset)), 4)
            bs_ties.append((c, hdr, tot_liab, tot_asset, diff))
            
            # Check liabilities components:
            # Row 4 (Equity), Row 5 (Reserves), Borrowings (7..10), Other Liab (12..15)
            eq = bs[f"{c}4"].value or 0.0
            res = bs[f"{c}5"].value or 0.0
            borr_sum = sum((bs[f"{c}{r}"].value or 0.0) for r in range(7, 11))
            oliab_sum = sum((bs[f"{c}{r}"].value or 0.0) for r in range(12, 16))
            calc_liab = round(eq + res + borr_sum + oliab_sum, 2)
            liab_diff = round(abs(calc_liab - float(tot_liab)), 2)
            comp_ties_liab.append((c, hdr, calc_liab, tot_liab, liab_diff))
            
            # Check assets components:
            gb = bs[f"{c}28"].value or 0.0
            dep = bs[f"{c}29"].value or 0.0
            net_fa = round(gb - dep, 2)
            cwip = bs[f"{c}30"].value or 0.0
            inv = bs[f"{c}31"].value or 0.0
            oth_assets = sum((bs[f"{c}{r}"].value or 0.0) for r in range(33, 38))
            calc_asset = round(net_fa + cwip + inv + oth_assets, 2)
            asset_diff = round(abs(calc_asset - float(tot_asset)), 2)
            comp_ties_assets.append((c, hdr, calc_asset, tot_asset, asset_diff))

    # Summary of Historical BS ties (Cols C-M)
    hist_ties = [item for item in bs_ties if item[0] != "N"]
    fy26_tie = [item for item in bs_ties if item[0] == "N"][0]
    hist_mismatches = [item for item in hist_ties if item[4] > 2.0]
    if hist_mismatches:
        print(f"  [FAIL] Historical Balance sheet mismatches ({len(hist_mismatches)}):")
        for c, hdr, l, a, d in hist_mismatches:
            print(f"    Col {c} ({hdr}): Liab={l}, Assets={a}, Diff={d}")
    else:
        print(f"  [PASS] Historical Balance sheet reconciled across all 11 columns C-M (Max diff: {max(x[4] for x in hist_ties)} Cr)")
    
    print(f"  [FY26E Col N] Honest Un-plugged Variance: Liab={fy26_tie[2]}, Assets={fy26_tie[3]}, Diff={fy26_tie[4]} Cr")

    # Component tie checks
    liab_mismatches = [item for item in comp_ties_liab if item[4] > 0.10]
    if liab_mismatches:
        print(f"  [WARN] Liability component breakdown diffs ({len(liab_mismatches)}):")
        for c, hdr, calc, tot, d in liab_mismatches:
            print(f"    Col {c} ({hdr}): Calc={calc}, Stated={tot}, Diff={d}")
    else:
        print(f"  [PASS] Liabilities components sum exactly to Total Liabilities")

    asset_mismatches = [item for item in comp_ties_assets if item[4] > 0.10]
    if asset_mismatches:
        print(f"  [WARN] Asset component breakdown diffs ({len(asset_mismatches)}):")
        for c, hdr, calc, tot, d in asset_mismatches:
            print(f"    Col {c} ({hdr}): Calc={calc}, Stated={tot}, Diff={d}")
    else:
        print(f"  [PASS] Assets components sum exactly to Total Assets")

    # Specific company spot checks
    if t == "CLEAN":
        print("  Checking CLEAN pre-FY18 standalone actuals:")
        print(f"    FY15 Total Assets (Col C): {bs['C38'].value} (Equity: {bs['C4'].value}, Reserves: {bs['C5'].value})")
        print(f"    FY16 Total Assets (Col D): {bs['D38'].value} (Equity: {bs['D4'].value}, Reserves: {bs['D5'].value})")
        print(f"    FY17 Total Assets (Col E): {bs['E38'].value} (Equity: {bs['E4'].value}, Reserves: {bs['E5'].value})")
    elif t == "MM":
        print("  Checking MM FY25 & FY26E Other Assets & Net Block:")
        print(f"    FY25 (Col M): Net Block = {bs['M28'].value}, Loans = {bs['M36'].value}, Other Assets = {bs['M37'].value}, Total = {bs['M38'].value}")
        print(f"    FY26E (Col N): Net Block = {bs['N28'].value}, Loans = {bs['N36'].value}, Other Assets = {bs['N37'].value}, Total = {bs['N38'].value}")
    elif t == "LTM":
        print("  Checking LTM FY26E Net Block:")
        print(f"    FY26E (Col N): Gross Block = {bs['N28'].value}, Accum Dep = {bs['N29'].value}, Net Block = {(bs['N28'].value or 0)-(bs['N29'].value or 0)}")
    elif t == "EDELWEISS":
        print("  Checking EDELWEISS FY26E Net Block:")
        print(f"    FY26E (Col N): Gross Block = {bs['N28'].value}, Accum Dep = {bs['N29'].value}, Net Block = {(bs['N28'].value or 0)-(bs['N29'].value or 0)}")
