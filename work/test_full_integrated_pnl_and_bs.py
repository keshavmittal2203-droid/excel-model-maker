import openpyxl

targets = ['MM', 'LTM', 'CLEAN', 'EDELWEISS']

for t in targets:
    wb = openpyxl.load_workbook(f'companies/{t}/template_v6/{t}-FullTemplate.xlsx', data_only=True)
    ds = wb['Data Sheet']
    bs = wb['BS']
    cf = wb['Cash Flow']
    
    sales_j = ds['J17'].value
    sales_k = ds['K17'].value
    
    exp_k_total = 0.0
    for r in range(18, 25):
        val_j = ds[f'J{r}'].value
        if val_j is not None:
            exp_k_total += round(sales_k * (val_j / sales_j), 2)
            
    oth_inc_k = ds['K25'].value or 0.0
    dep_k = ds['K26'].value or 0.0
    int_k = ds['K27'].value or 0.0
    
    pbt_k = round(sales_k - exp_k_total + oth_inc_k - dep_k - int_k, 2)
    tax_k = round(pbt_k * 0.2517, 2)
    pat_k = round(pbt_k - tax_k, 2)
    div_k = round(pat_k * 0.20, 2)
    
    # Opening BS (M)
    eq_open = bs['M4'].value or 0.0
    res_open = bs['M5'].value or 0.0
    st_borr_open = bs['M7'].value or 0.0
    lease_open = bs['M8'].value or 0.0
    pref_open = bs['M9'].value or 0.0
    oth_borr_open = bs['M10'].value or 0.0
    nci_open = bs['M12'].value or 0.0
    pay_open = bs['M13'].value or 0.0
    adv_open = bs['M14'].value or 0.0
    oliab_open = bs['M15'].value or 0.0
    
    gb_open = bs['M28'].value or 0.0
    dep_open = bs['M29'].value or 0.0
    cwip_open = bs['M30'].value or 0.0
    inv_open = bs['M31'].value or 0.0
    inven_open = bs['M33'].value or 0.0
    rec_open = bs['M34'].value or 0.0
    cash_open = bs['M35'].value or 0.0
    loans_open = bs['M36'].value or 0.0
    oth_ast_open = bs['M37'].value or 0.0
    
    # Forecast BS (N)
    eq_close = eq_open
    res_close = round(res_open + pat_k - div_k, 2)
    st_borr_close = st_borr_open
    lease_close = lease_open
    pref_close = pref_open
    oth_borr_close = oth_borr_open
    nci_close = nci_open
    adv_close = adv_open
    pay_close = round(pay_open * 1.05, 2) if pay_open > 0 else 0.0
    oliab_close = round(oliab_open * 1.05, 2) if oliab_open > 0 else 0.0
    
    capex_rate = 1.25 if t == 'MM' else 1.20
    capex = round(dep_k * capex_rate, 2)
    
    if t == 'MM':
        gb_close = round(gb_open + capex - dep_k, 2)
        dep_close = 0.0
        net_fa_close = gb_close
        net_fa_open = gb_open
    else:
        gb_close = round(gb_open + capex, 2)
        dep_close = round(dep_open + dep_k, 2)
        net_fa_close = round(gb_close - dep_close, 2)
        net_fa_open = round(gb_open - dep_open, 2)
        
    cwip_close = cwip_open
    inv_close = inv_open
    inven_close = round(inven_open * 1.05, 2) if inven_open > 0 else 0.0
    rec_close = round(rec_open * 1.05, 2) if rec_open > 0 else 0.0
    loans_close = loans_open
    oth_ast_close = round(oth_ast_open * 1.05, 2) if (oth_ast_open > 0 and t != 'MM') else 0.0
    
    # Cash Flow Integration
    delta_rec = rec_close - rec_open
    delta_inven = inven_close - inven_open
    delta_pay = pay_close - pay_open
    delta_loans = loans_close - loans_open
    delta_oth_ast = oth_ast_close - oth_ast_open
    delta_oliab = oliab_close - oliab_open
    delta_adv = adv_close - adv_open
    
    wc_changes = round(-delta_rec - delta_inven + delta_pay - delta_loans - delta_oth_ast + delta_oliab + delta_adv, 2)
    cfo = round(pat_k + dep_k + wc_changes, 2)
    cfi = round(-capex, 2)
    cff = round(-div_k, 2)
    net_cf = round(cfo + cfi + cff, 2)
    cash_close = round(cash_open + net_cf, 2)
    
    tot_liab_open = round(eq_open + res_open + st_borr_open + lease_open + pref_open + oth_borr_open + nci_open + pay_open + adv_open + oliab_open, 2)
    tot_liab_close = round(eq_close + res_close + st_borr_close + lease_close + pref_close + oth_borr_close + nci_close + pay_close + adv_close + oliab_close, 2)
    
    tot_ast_open = round(net_fa_open + cwip_open + inv_open + inven_open + rec_open + cash_open + loans_open + oth_ast_open, 2)
    tot_ast_close = round(net_fa_close + cwip_close + inv_close + inven_close + rec_close + cash_close + loans_close + oth_ast_close, 2)
    
    delta_liab = round(tot_liab_close - tot_liab_open, 2)
    delta_ast = round(tot_ast_close - tot_ast_open, 2)
    
    diff_open = round(tot_ast_open - tot_liab_open, 2)
    diff_close = round(tot_ast_close - tot_liab_close, 2)
    
    print(f"=== {t} ===")
    print(f"  Sales: {sales_k:,.2f} Cr, PAT: {pat_k:,.2f} Cr (Net Margin: {pat_k/sales_k*100:.2f}%)")
    print(f"  CFO: {cfo:,.2f} Cr, CFI: {cfi:,.2f} Cr, CFF: {cff:,.2f} Cr, Net CF: {net_cf:,.2f} Cr")
    print(f"  Delta Assets: {delta_ast:,.2f} Cr, Delta Liab: {delta_liab:,.2f} Cr => Net Integration Variance: {delta_ast - delta_liab:.4f} Cr")
    print(f"  Opening FY25 Check (Ast - Liab): {diff_open:+.2f} Cr")
    print(f"  Closing FY26E Check (Ast - Liab): {diff_close:+.2f} Cr\n")
