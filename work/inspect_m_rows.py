import openpyxl

for target in ['LTM', 'CLEAN', 'EDELWEISS', 'MM']:
    path = f'C:/Users/kesha/Documents/Codex/2026-10-18/excel-model-maker-v9-remediation/companies/{target}/template_v15_excel_com/{target}-FullTemplate-v15-ExcelCOM.xlsx'
    wb = openpyxl.load_workbook(path, data_only=True)
    bs = wb['BS']
    print(f"\n=== {target} Column M (FY25) ===")
    print("Liabilities M:")
    s_liab = 0
    for r in [4, 5, 7, 8, 9, 10, 12, 13, 14, 15]:
        v = bs[f'M{r}'].value or 0
        s_liab += v
        print(f"  Row {r:2d} ({bs[f'B{r}'].value}): {v}")
    print(f"Sum of Liab rows: {s_liab} | Cell M16: {bs['M16'].value}")

    print("Assets M:")
    s_assets = 0
    for r in [28, 29, 30, 31, 33, 34, 35, 36, 37]:
        v = bs[f'M{r}'].value or 0
        if r == 29 and target != 'MM': # accumulated depreciation
            s_assets -= v
        else:
            s_assets += v
        print(f"  Row {r:2d} ({bs[f'B{r}'].value}): {v}")
    print(f"Sum of Asset rows: {s_assets} | Cell M38: {bs['M38'].value}")
    print(f"M40 value: {bs['M40'].value}")
