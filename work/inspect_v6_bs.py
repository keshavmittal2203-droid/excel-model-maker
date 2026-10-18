import openpyxl

for t in ['MM', 'LTM', 'CLEAN', 'EDELWEISS']:
    fn = f'companies/{t}/template_v6/{t}-FullTemplate.xlsx'
    wb = openpyxl.load_workbook(fn, data_only=False)
    bs = wb['BS']
    print(f'=== {t} BS Sheet ({fn}) ===')
    cols = ['C','D','E','F','G','H','I','J','K','L','M','N']
    print('Row 16 (Total Liabilities):')
    for c in cols:
        print(f'{c}: {bs[f"{c}16"].value}', end=' | ')
    print('\nRow 37 (Other asset items):')
    for c in cols:
        print(f'{c}: {bs[f"{c}37"].value}', end=' | ')
    print('\nRow 38 (Total Assets):')
    for c in cols:
        print(f'{c}: {bs[f"{c}38"].value}', end=' | ')
    print('\n')
