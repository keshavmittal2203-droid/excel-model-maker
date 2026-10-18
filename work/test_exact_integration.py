import sys
sys.path.insert(0, '.')
import json
import openpyxl

targets = ["MM", "LTM", "CLEAN", "EDELWEISS"]

for t in targets:
    wb_path = f"companies/{t}/template_v6/{t}-FullTemplate.xlsx"
    wb = openpyxl.load_workbook(wb_path, data_only=True)
    ds = wb["Data Sheet"]
    bs = wb["BS"]
    cf = wb["Cash Flow"]
    
    print(f"\n{'='*70}\nEXACT INTEGRATION AUDIT: {t}\n{'='*70}")
    
    # 1. P&L Forecast (Data Sheet Col K)
    sales = ds["K17"].value
    depr = ds["K26"].value
    interest = ds["K27"].value
    pbt = ds["K28"].value
    tax = ds["K29"].value
    pat = ds["K30"].value
    div = ds["K31"].value
    ret_earnings = round(pat - div, 2)
    
    # 2. Opening Balance Sheet (Col M - FY25)
    eq_open = bs["M4"].value
    res_open = bs["M5"].value
    st_borr_open = bs["M7"].value or 0.0
    lease_open = bs["M8"].value or 0.0
    pref_open = bs["M9"].value or 0.0
    oth_borr_open = bs["M10"].value or 0.0
    tot_borr_open = round(st_borr_open + lease_open + pref_open + oth_borr_open, 2)
    
    nci_open = bs["M12"].value or 0.0
    pay_open = bs["M13"].value or 0.0
    adv_open = bs["M14"].value or 0.0
    oliab_open = bs["M15"].value or 0.0
    tot_liab_open = bs["M16"].value
    
    gb_open = bs["M28"].value
    dep_open = bs["M29"].value or 0.0
    cwip_open = bs["M30"].value or 0.0
    inv_open = bs["M31"].value or 0.0
    inven_open = bs["M33"].value or 0.0
    rec_open = bs["M34"].value or 0.0
    cash_open = bs["M35"].value
    loans_open = bs["M36"].value or 0.0
    oth_ast_open = bs["M37"].value or 0.0
    tot_ast_open = bs["M38"].value
    open_rounding_diff = round(tot_ast_open - tot_liab_open, 2)
    
    print(f"Opening FY25: Assets={tot_ast_open}, Liab={tot_liab_open}, Rounding Diff={open_rounding_diff} Cr")
    
    # 3. Forecast Closing Balance Sheet Items (Col N - FY26E)
    eq_close = eq_open
    res_close = round(res_open + ret_earnings, 2)
    
    # Debt held flat or rolled forward:
    st_borr_close = st_borr_open
    lease_close = lease_open
    pref_close = pref_open
    oth_borr_close = oth_borr_open
    tot_borr_close = tot_borr_open
    
    # Operating liabilities (growth rate 5%)
    pay_close = round(pay_open * 1.05, 2) if pay_open > 0 else 0.0
    adv_close = adv_open
    oliab_close = round(oliab_open * 1.05, 2) if oliab_open > 0 else 0.0
    nci_close = nci_open
    
    tot_liab_close = round(eq_close + res_close + st_borr_close + lease_close + pref_close + oth_borr_close + nci_close + pay_close + adv_close + oliab_close, 2)
    
    # Fixed Assets
    if t == "MM":
        capex = round(depr * 1.25, 2)
        gb_close = round(gb_open + capex - depr, 2)
        dep_close = 0.0
        net_fa_close = gb_close
    else:
        capex = round(depr * 1.20, 2)
        gb_close = round(gb_open + capex, 2)
        dep_close = round(dep_open + depr, 2)
        net_fa_close = round(gb_close - dep_close, 2)
        
    cwip_close = cwip_open
    inv_close = inv_open
    
    inven_close = round(inven_open * 1.05, 2) if inven_open > 0 else 0.0
    rec_close = round(rec_open * 1.05, 2) if rec_open > 0 else 0.0
    loans_close = loans_open
    oth_ast_close = round(oth_ast_open * 1.05, 2) if oth_ast_open > 0 else 0.0
    
    # 4. Cash Flow Statement Integration (Col N)
    delta_inven = inven_close - inven_open
    delta_rec = rec_close - rec_open
    delta_loans = loans_close - loans_open
    delta_oth_ast = oth_ast_close - oth_ast_open
    delta_pay = pay_close - pay_open
    delta_oliab = oliab_close - oliab_open
    
    wc_changes = round(-delta_inven - delta_rec - delta_loans - delta_oth_ast + delta_pay + delta_oliab, 2)
    cfo = round(pat + depr + wc_changes, 2)
    cfi = round(-capex, 2)
    cff = round(-div, 2)
    net_cf = round(cfo + cfi + cff, 2)
    
    cash_close = round(cash_open + net_cf, 2)
    
    tot_ast_close = round(net_fa_close + cwip_close + inv_close + inven_close + rec_close + cash_close + loans_close + oth_ast_close, 2)
    
    final_diff = round(tot_ast_close - tot_liab_close, 2)
    
    print(f"Closing FY26E:")
    print(f"  Liabilities: Eq={eq_close}, Res={res_close}, Borr={tot_borr_close}, Pay={pay_close}, OLiab={oliab_close} => Total Liab={tot_liab_close:,.2f}")
    print(f"  Assets: NetFA={net_fa_close}, CWIP={cwip_close}, Inv={inv_close}, Inven={inven_close}, Rec={rec_close}, Cash={cash_close:,.2f}, Loans={loans_close}, OthAst={oth_ast_close} => Total Ast={tot_ast_close:,.2f}")
    print(f"  Balance Check (Total Assets - Total Liabilities): {final_diff:+.2f} Cr (Opening Screener Rounding: {open_rounding_diff:+.2f} Cr)")
