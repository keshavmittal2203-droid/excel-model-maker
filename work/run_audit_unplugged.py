import os, openpyxl, win32com.client

ROOT_CAND = 'C:/Users/kesha/Documents/Codex/2026-10-18/excel-model-maker-v9-remediation/companies'

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

for target in ['MM', 'LTM', 'CLEAN', 'EDELWEISS']:
    p = f'{ROOT_CAND}/{target}/template_v15_excel_com/{target}-FullTemplate-v15-ExcelCOM.xlsx'
    wb_py = openpyxl.load_workbook(p)
    cf = wb_py['Cash Flow']
    cf['N31'].value = '=N5+N14+N23'
    for col in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']:
        cf[f'{col}31'].value = f'={col}5+{col}14+{col}23'
    
    bs = wb_py['BS']
    for col in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']:
        bs[f'{col}40'].value = f'={col}38-{col}16'
    
    bs['N35'].value = "=M35+'Cash Flow'!N31"
    # DO NOT TOUCH M37 or M15! Leave exactly as reported by Screener!
    
    test_path = os.path.abspath(f'work/test_audit_{target}.xlsx')
    wb_py.save(test_path)
    
    wb_xl = excel.Workbooks.Open(test_path)
    excel.CalculateFullRebuild()
    
    ws_bs = wb_xl.Sheets('BS')
    row40 = [ws_bs.Cells(40, c).Value for c in range(3, 15)]
    row16 = [ws_bs.Cells(16, c).Value for c in range(3, 15)]
    row38 = [ws_bs.Cells(38, c).Value for c in range(3, 15)]
    
    # Check formula errors across all sheets
    err_count = 0
    for s in wb_xl.Sheets:
        used = s.UsedRange
        vals = used.Value
        if vals:
            if not isinstance(vals, tuple):
                vals = ((vals,),)
            for row in vals:
                if row:
                    for cell in row:
                        if isinstance(cell, int) and cell in [-2146826281, -2146826246, -2146826259, -2146826273, -2146826265]:
                            err_count += 1
    
    ws_dcf = wb_xl.Sheets('DCF 1st way')
    val = ws_dcf.Range('D71').Value
    
    print(f'=== {target} (Unplugged Source Data) ===')
    print('Row 40 Checks (C to N):', [round(x, 4) if x is not None else None for x in row40])
    print(f'FY26E: Assets (Row 38) = {row38[-1]}, Liab (Row 16) = {row16[-1]}, Diff (Row 40) = {row40[-1]}')
    print(f'Formula Errors: {err_count}, Valuation: {val}')
    wb_xl.Close(False)

excel.Quit()
