import openpyxl

for target in ['LTM', 'CLEAN', 'EDELWEISS', 'MM']:
    path = f'C:/Users/kesha/Documents/Codex/2026-10-18/excel-model-maker-v9-remediation/companies/{target}/template_v15_excel_com/{target}-FullTemplate-v15-ExcelCOM.xlsx'
    wb_f = openpyxl.load_workbook(path, data_only=False)
    wb_v = openpyxl.load_workbook(path, data_only=True)
    bs_f = wb_f['BS']
    bs_v = wb_v['BS']
    print(f"\n==================== {target} BS Formulas across Cols C to N ====================")
    cols = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']
    for col in cols:
        f16 = bs_f[f'{col}16'].value
        v16 = bs_v[f'{col}16'].value
        f38 = bs_f[f'{col}38'].value
        v38 = bs_v[f'{col}38'].value
        f40 = bs_f[f'{col}40'].value
        v40 = bs_v[f'{col}40'].value
        print(f"Col {col} | Row 16: val={v16} (f={f16}) | Row 38: val={v38} (f={f38}) | Row 40: val={v40} (f={f40})")
