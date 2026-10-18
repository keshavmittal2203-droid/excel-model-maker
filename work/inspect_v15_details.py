import openpyxl

for target in ['MM', 'LTM', 'CLEAN', 'EDELWEISS']:
    path = f'C:/Users/kesha/Documents/Codex/2026-10-18/excel-model-maker-v9-remediation/companies/{target}/template_v15_excel_com/{target}-FullTemplate-v15-ExcelCOM.xlsx'
    wb_v = openpyxl.load_workbook(path, data_only=True)
    wb_f = openpyxl.load_workbook(path, data_only=False)
    bs_v = wb_v['BS']
    bs_f = wb_f['BS']
    print(f'********************* {target} *********************')
    print(f'N16 (Total Liab): {bs_v["N16"].value} | Formula: {bs_f["N16"].value}')
    print(f'N38 (Total Assets): {bs_v["N38"].value} | Formula: {bs_f["N38"].value}')
    print(f'N40 (Diff): {bs_v["N40"].value} | Formula: {bs_f["N40"].value}')
    print('--- Liabilities ---')
    for r in range(4, 16):
        desc = bs_v[f'B{r}'].value
        v = bs_v[f'N{r}'].value
        f = bs_f[f'N{r}'].value
        if v is not None or f is not None:
            print(f'  Row {r:2d} ({desc}): val={v} | f={f}')
    print('--- Assets ---')
    for r in range(18, 38):
        desc = bs_v[f'B{r}'].value
        v = bs_v[f'N{r}'].value
        f = bs_f[f'N{r}'].value
        if v is not None or f is not None:
            print(f'  Row {r:2d} ({desc}): val={v} | f={f}')
