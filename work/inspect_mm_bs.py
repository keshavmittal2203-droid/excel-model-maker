import openpyxl

path = 'C:/Users/kesha/Documents/Codex/2026-10-18/excel-model-maker-v9-remediation/companies/MM/template_v15_excel_com/MM-FullTemplate-v15-ExcelCOM.xlsx'
wb_v = openpyxl.load_workbook(path, data_only=True)
wb_f = openpyxl.load_workbook(path, data_only=False)
bs_v = wb_v['BS']
bs_f = wb_f['BS']

print("=== ALL ROWS IN BS FOR MM ===")
for r in range(1, 42):
    b = bs_v[f'B{r}'].value
    m_v = bs_v[f'M{r}'].value
    m_f = bs_f[f'M{r}'].value
    n_v = bs_v[f'N{r}'].value
    n_f = bs_f[f'N{r}'].value
    if b or m_v is not None or n_v is not None:
        print(f"Row {r:2d} | {str(b):40s} | M(FY25): {m_v} ({m_f}) | N(FY26E): {n_v} ({n_f})")
