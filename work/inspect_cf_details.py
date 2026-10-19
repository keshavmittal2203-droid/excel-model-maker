import openpyxl

for target in ['MM', 'LTM', 'CLEAN', 'EDELWEISS']:
    path = f'C:/Users/kesha/Documents/Codex/2026-10-18/excel-model-maker-v9-remediation/companies/{target}/template_v15_excel_com/{target}-FullTemplate-v15-ExcelCOM.xlsx'
    wb_v = openpyxl.load_workbook(path, data_only=True)
    wb_f = openpyxl.load_workbook(path, data_only=False)
    cf_v = wb_v['Cash Flow']
    cf_f = wb_f['Cash Flow']
    print(f'********************* {target} Cash Flow *********************')
    for r in range(4, 33):
        desc = cf_v[f'B{r}'].value
        v_m = cf_v[f'M{r}'].value
        v_n = cf_v[f'N{r}'].value
        f_n = cf_f[f'N{r}'].value
        if desc or v_n is not None or f_n is not None:
            print(f'Row {r:2d} | {str(desc):35s} | M(FY25)={v_m} | N(FY26E)={v_n} | formula={f_n}')
