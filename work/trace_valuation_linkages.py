import openpyxl

for t in ['MM', 'LTM', 'CLEAN', 'EDELWEISS']:
    wb = openpyxl.load_workbook(f'companies/{t}/template_v6/{t}-FullTemplate.xlsx', data_only=False)
    dcf = wb['DCF 1st way']
    wacc = wb['WACC-DCF']
    ds = wb['Data Sheet']
    print(f"{'='*60}\nCOMPANY: {t}\n{'='*60}")
    print(f"Cover Model -> C9: {wb['Model ->']['C9'].value}")
    print(f"DCF 1st way E42: {dcf['E42'].value}")
    
    if t in ['MM', 'EDELWEISS']:
        print("SOTP Schedule:")
        for r in range(45, 62):
            cells = [f"{c}{r}={dcf[f'{c}{r}'].value}" for c in ['B', 'C', 'D', 'E', 'F', 'G'] if dcf[f'{c}{r}'].value is not None]
            if cells:
                print(f"  Row {r:2d}: " + ' | '.join(cells))
    else:
        print("FCFF DCF Schedule:")
        for r in range(1, 20):
            cells = [f"{c}{r}={dcf[f'{c}{r}'].value}" for c in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K'] if dcf[f'{c}{r}'].value is not None]
            if cells:
                print(f"  Row {r:2d}: " + ' | '.join(cells))
        for r in range(20, 45):
            cells = [f"{c}{r}={dcf[f'{c}{r}'].value}" for c in ['D', 'E'] if dcf[f'{c}{r}'].value is not None]
            if cells:
                print(f"  Row {r:2d}: " + ' | '.join(cells))
                
    print("\nData Sheet FY25 (Col J) and FY26E (Col K):")
    for r in [17, 26, 27, 28, 29, 30, 31]:
        lbl = ds[f'A{r}'].value
        print(f"  Row {r:2d} ({lbl}): J={ds[f'J{r}'].value} | K={ds[f'K{r}'].value}")
        
    print("\nWACC-DCF Key Cells:")
    for cell_id in ['D18', 'E22', 'E23', 'E24', 'E30', 'E31', 'E32', 'J22', 'J23', 'J24', 'J25', 'J32', 'J39']:
        c = wacc[cell_id]
        cmt = f" [Comment: {c.comment.text[:30]}...]" if c.comment else ""
        print(f"  {cell_id}: {c.value}{cmt}")
    print()
