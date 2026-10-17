import win32com.client
import os

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

targets = ['MM', 'LTM', 'CLEAN', 'EDELWEISS']
results = {}

for tgt in targets:
    path = os.path.abspath(f'companies/{tgt}/template_v2/{tgt}-FullTemplate.xlsx')
    print(f'=== Testing COM on {tgt} ({path}) ===')
    try:
        wb = excel.Workbooks.Open(path)
        excel.CalculateFullRebuild()
        print('CalculateFullRebuild succeeded!')
        
        info = {}
        for sname in ['Data Sheet', 'Main Financials', 'Ratio Analysis', 'WACC-DCF', 'Beta', 'Comps', 'DCF 1st way']:
            ws = wb.Sheets(sname)
            hdr = ws.Range("B1").Text or ws.Range("A1").Text
            info[sname] = {'header': hdr}
            if sname == 'WACC-DCF':
                info[sname]['wacc'] = ws.Range("J39").Text
                info[sname]['ke'] = ws.Range("J25").Text
                info[sname]['beta'] = ws.Range("J24").Text
                print(f'    WACC (J39): {ws.Range("J39").Text}, Ke (J25): {ws.Range("J25").Text}, Beta (J24): {ws.Range("J24").Text}')
            elif sname == 'Beta':
                info[sname]['slope'] = ws.Range("O6").Text
                info[sname]['raw_beta'] = ws.Range("J7").Text
                info[sname]['adj_beta'] = ws.Range("J13").Text
                print(f'    Beta Slope (O6): {ws.Range("O6").Text}, Raw (J7): {ws.Range("J7").Text}, Adj (J13): {ws.Range("J13").Text}')
            elif sname == 'DCF 1st way':
                info[sname]['val_per_share'] = ws.Range("E42").Text
                info[sname]['cmp'] = ws.Range("E44").Text
                print(f'    Value per share (E42): {ws.Range("E42").Text}, Mkt Price (E44): {ws.Range("E44").Text}')
        
        wb.Save()
        wb.Close(SaveChanges=True)
        results[tgt] = {'status': 'success', 'info': info}
        print(f'{tgt} saved successfully after recalculation.\n')
    except Exception as e:
        results[tgt] = {'status': 'error', 'error': str(e)}
        print(f'Error processing {tgt}: {e}\n')

excel.Quit()
print('All COM recalculations completed.')
