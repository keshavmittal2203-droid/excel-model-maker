import openpyxl
path = 'C:/Users/kesha/Documents/Codex/2026-10-18/excel-model-maker-v9-remediation/companies/LTM/template_v15_excel_com/LTM-FullTemplate-v15-ExcelCOM.xlsx'
wb = openpyxl.load_workbook(path, data_only=False)
ds = wb['Data Sheet']
for r in range(16, 36):
    print(f'Row {r:2d} | {str(ds[f"A{r}"].value):30s} | {str(ds[f"B{r}"].value):30s} | J(FY25)={ds[f"J{r}"].value} | K(FY26E)={ds[f"K{r}"].value}')
