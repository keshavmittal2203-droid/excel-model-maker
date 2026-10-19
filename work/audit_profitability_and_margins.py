import openpyxl

targets = ['MM', 'LTM', 'CLEAN', 'EDELWEISS']

for t in targets:
    wb = openpyxl.load_workbook(f'companies/{t}/template_v6/{t}-FullTemplate.xlsx', data_only=True)
    ds = wb['Data Sheet']
    sales_j = ds['J17'].value
    sales_k = ds['K17'].value
    growth = (sales_k / sales_j) - 1.0
    
    print(f"{'='*60}")
    print(f"COMPANY: {t}")
    print(f"Sales: FY25 = {sales_j:,.2f} Cr | FY26E = {sales_k:,.2f} Cr (+{growth*100:.1f}%)")
    print(f"{'='*60}")
    
    exp_k_total = 0.0
    for r in range(18, 25):
        val_j = ds[f'J{r}'].value
        lbl = ds[f'A{r}'].value
        if val_j is not None:
            exp_k = round(sales_k * (val_j / sales_j), 2)
            exp_k_total += exp_k
            print(f"  Row {r:2d} ({lbl[:22]:22s}): FY25={val_j:10,.2f} ({val_j/sales_j*100:5.2f}%) -> FY26E={exp_k:10,.2f} ({exp_k/sales_k*100:5.2f}%)")
            
    oth_inc_k = ds['K25'].value or 0.0
    dep_k = ds['K26'].value or 0.0
    int_k = ds['K27'].value or 0.0
    
    pbt_k = round(sales_k - exp_k_total + oth_inc_k - dep_k - int_k, 2)
    tax_rate = 0.2517
    tax_k = round(pbt_k * tax_rate, 2)
    pat_k = round(pbt_k - tax_k, 2)
    div_k = round(pat_k * 0.20, 2)
    margin_k = (pat_k / sales_k) * 100
    margin_j = (ds['J30'].value / sales_j) * 100
    
    print(f"  Other Income (Row 25)    : {oth_inc_k:10,.2f}")
    print(f"  Depreciation (Row 26)    : {dep_k:10,.2f}")
    print(f"  Interest (Row 27)        : {int_k:10,.2f}")
    print(f"  Operating Profit         : {sales_k - exp_k_total:10,.2f} ({(sales_k - exp_k_total)/sales_k*100:5.2f}%)")
    print(f"  PBT (Row 28)             : {pbt_k:10,.2f} ({pbt_k/sales_k*100:5.2f}%)")
    print(f"  Tax (Row 29, 25.17%)     : {tax_k:10,.2f}")
    print(f"  PAT (Row 30)             : {pat_k:10,.2f}")
    print(f"  Dividends (Row 31, 20%)  : {div_k:10,.2f}")
    print(f"  --> FY26E Net Margin     : {margin_k:.2f}% (vs FY25 Actual: {margin_j:.2f}%)\n")
