import win32com.client
import os
import re

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

targets = ['MM', 'LTM', 'CLEAN', 'EDELWEISS']
exported = {}

for tgt in targets:
    xlsx_path = os.path.abspath(f'companies/{tgt}/template_v3/{tgt}-FullTemplate.xlsx')
    preview_dir = os.path.abspath(f'companies/{tgt}/template_v3/previews')
    os.makedirs(preview_dir, exist_ok=True)
    
    print(f'=== Exporting Previews for {tgt} Using Original Native Print Settings ===')
    wb = excel.Workbooks.Open(xlsx_path)
    sheet_files = []
    
    for idx, ws in enumerate(wb.Sheets, start=1):
        sname = ws.Name
        safe_name = re.sub(r'[^a-zA-Z0-9_\-]', '_', sname)
        pdf_name = f'{idx:02d}_{safe_name}.pdf'
        pdf_path = os.path.join(preview_dir, pdf_name)
        
        try:
            # STRICT INSTRUCTION: Render and inspect the sheets using the original print settings;
            # do not compress or rescale them for a cleaner preview.
            ws.ExportAsFixedFormat(0, pdf_path)
            size = os.path.getsize(pdf_path)
            sheet_files.append((sname, pdf_name, size))
            print(f'  [{idx:02d}] {sname} -> {pdf_name} ({size} bytes, native print settings preserved)')
        except Exception as e:
            print(f'  [{idx:02d}] Error exporting {sname}: {e}')
            
    wb.Close(SaveChanges=False)
    exported[tgt] = sheet_files
    print(f'Exported {len(sheet_files)} sheets for {tgt}.\n')

excel.Quit()
print('All sheet preview exports using original native print settings completed.')
