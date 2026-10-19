import openpyxl

for target in ['LTM', 'CLEAN', 'EDELWEISS']:
    path = f'C:/Users/kesha/Documents/Codex/2026-10-18/excel-model-maker-v9-remediation/companies/{target}/template_v15_excel_com/{target}-FullTemplate-v15-ExcelCOM.xlsx'
    wb_v = openpyxl.load_workbook(path, data_only=True)
    bs = wb_v['BS']
    cf = wb_v['Cash Flow']
    ds = wb_v['Data Sheet']

    print(f"\n==================== {target} DETAILED DELTAS ====================")
    d_reserves = bs['N5'].value - bs['M5'].value
    pat = ds['K30'].value
    div = ds['K31'].value
    print(f"Reserves Delta: {d_reserves} (PAT: {pat}, Div: {div}, Net: {pat - div})")
    
    # Liabilities deltas
    liab_deltas = {}
    for r in [4, 5, 7, 8, 9, 10, 12, 13, 14, 15]:
        b = bs[f'B{r}'].value
        if b and bs[f'M{r}'].value is not None:
            diff = (bs[f'N{r}'].value or 0) - (bs[f'M{r}'].value or 0)
            liab_deltas[r] = diff
            print(f"  Liab Row {r:2d} ({b}): delta = {diff:+.4f}")
    
    total_liab_delta = bs['N16'].value - bs['M16'].value
    print(f"Total Liab Delta (N16 - M16): {total_liab_delta:+.4f}")

    # Assets deltas
    asset_deltas = {}
    # Row 28 - Row 29
    gb_delta = bs['N28'].value - bs['M28'].value
    dep_delta = bs['N29'].value - bs['M29'].value
    nb_delta = gb_delta - dep_delta
    print(f"  Net Block Delta: {nb_delta:+.4f} (Gross Block delta: {gb_delta:+.4f}, Dep delta: {dep_delta:+.4f})")
    print(f"    Capex in CF!N15: {cf['N15'].value}, Dep in DS!K26: {ds['K26'].value}")
    
    for r in [30, 31, 33, 34, 35, 36, 37]:
        b = bs[f'B{r}'].value
        if b and bs[f'M{r}'].value is not None:
            diff = (bs[f'N{r}'].value or 0) - (bs[f'M{r}'].value or 0)
            asset_deltas[r] = diff
            print(f"  Asset Row {r:2d} ({b}): delta = {diff:+.4f}")

    total_asset_delta = bs['N38'].value - bs['M38'].value
    print(f"Total Asset Delta (N38 - M38): {total_asset_delta:+.4f}")
    print(f"Asset Delta - Liab Delta: {total_asset_delta - total_liab_delta:+.4f}")
    print(f"N40 Balance check: {bs['N40'].value}")

    print("\n--- Cash Flow Breakdown ---")
    cfo = cf['N5'].value
    cfi = cf['N14'].value
    cff = cf['N23'].value
    net_cf_sum = cfo + cfi + cff
    reported_net_cf = cf['N31'].value
    print(f"  CFO: {cfo}")
    print(f"  CFI: {cfi}")
    print(f"  CFF: {cff}")
    print(f"  Sum of CFO+CFI+CFF: {net_cf_sum}")
    print(f"  CF!N31 value: {reported_net_cf}")
    print(f"  Discrepancy in CF!N31: {reported_net_cf - net_cf_sum}")
    print(f"  BS Cash Delta: {bs['N35'].value - bs['M35'].value}")

    wc_cf = cf['N12'].value
    print(f"  CF!N12 (Working Capital changes): {wc_cf}")
    for r in range(7, 12):
        print(f"    CF Row {r} ({cf[f'B{r}'].value}): {cf[f'N{r}'].value}")
