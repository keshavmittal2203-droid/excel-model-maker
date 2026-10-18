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
    
    print(f"\n{'='*70}\nSIMULATION: 3-STATEMENT INTEGRATION FOR {t}\n{'='*70}")
    
    # 1. P&L Forecast (Data Sheet Col K)
    sales = ds["K17"].value
    depr = ds["K26"].value
    interest = ds["K27"].value
    pbt = ds["K28"].value
    tax = ds["K29"].value
    pat = ds["K30"].value
    div = ds["K31"].value
    ret_earnings = pat - div
    print(f"P&L: Sales={sales:,.2f}, Depr={depr:,.2f}, PBT={pbt:,.2f}, Tax={tax:,.2f}, PAT={pat:,.2f}, Div={div:,.2f} => Retained={ret_earnings:,.2f}")
    
    # 2. Opening Balance Sheet (Col M - FY25)
    eq_open = bs["M4"].value
    res_open = bs["M5"].value
    borr_open = bs["M10"].value or 0.0
    nci_open = bs["M12"].value or 0.0
    pay_open = bs["M13"].value or 0.0
    oliab_open = bs["M15"].value or 0.0
    
    gb_open = bs["M28"].value
    dep_open = bs["M29"].value or 0.0
    cwip_open = bs["M30"].value or 0.0
    inv_open = bs["M31"].value or 0.0
    inven_open = bs["M33"].value or 0.0
    rec_open = bs["M34"].value or 0.0
    cash_open = bs["M35"].value
    loans_open = bs["M36"].value or 0.0
    oth_ast_open = bs["M37"].value or 0.0
    
    # 3. Forecast Operating Drivers for FY26E
    # CapEx:
    if t == "MM":
        # MM has Net Block in Row 28 (no separate dep row 29)
        # CapEx = Depr + growth
        capex = round(depr * 1.25, 2)
        gb_close = round(gb_open + capex - depr, 2)
        dep_close = 0.0
    else:
        capex = round(depr * 1.20, 2)
        gb_close = round(gb_open + capex, 2)
        dep_close = round(dep_open + depr, 2)
        
    net_fa_close = round(gb_close - dep_close, 2)
    cwip_close = cwip_open
    inv_close = inv_open
    
    # Working capital growth (5% or tied to sales growth)
    inven_close = round(inven_open * 1.05, 2) if inven_open > 0 else 0.0
    rec_close = round(rec_open * 1.05, 2) if rec_open > 0 else 0.0
    loans_close = round(loans_open, 2)
    oth_ast_close = round(oth_ast_open * 1.05, 2) if oth_ast_open > 0 else 0.0
    
    pay_close = round(pay_open * 1.05, 2) if pay_open > 0 else 0.0
    oliab_close = round(oliab_open * 1.05, 2) if oliab_open > 0 else 0.0
    nci_close = nci_open
    borr_close = borr_open
    eq_close = eq_open
    res_close = round(res_open + ret_earnings, 2)
    
    # 4. Cash Flow Statement Integration (Col N)
    # Delta working capital:
    delta_inven = inven_close - inven_open
    delta_rec = rec_close - rec_open
    delta_loans = loans_close - loans_open
    delta_oth_ast = oth_ast_close - oth_ast_open
    delta_pay = pay_close - pay_open
    delta_oliab = oliab_close - oliab_open
    
    # Cash effect of WC changes:
    # Increase in assets is cash outflow (-); increase in liabilities is cash inflow (+)
    wc_cash_effect = round(-delta_inven - delta_rec - delta_loans - delta_oth_ast + delta_pay + delta_oliab, 2)
    
    # CFO = PAT + Depr + WC_cash_effect (or PBT adjusted)
    cfo = round(pat + depr + wc_cash_effect, 2)
    
    # CFI = -CapEx - delta_cwip - delta_inv
    cfi = round(-capex, 2)
    
    # CFF = delta_debt - Dividends
    cff = round((borr_close - borr_open) - div, 2)
    
    net_cf = round(cfo + cfi + cff, 2)
    cash_close = round(cash_open + net_cf, 2)
    
    # 5. Check Balance Sheet Totals
    tot_liab_close = round(eq_close + res_close + borr_close + nci_close + pay_close + oliab_close, 2)
    tot_ast_close = round(net_fa_close + cwip_close + inv_close + inven_close + rec_close + cash_close + loans_close + oth_ast_close, 2)
    
    diff = round(tot_ast_close - tot_liab_close, 2)
    
    print(f"CFO={cfo:,.2f} (PAT {pat} + Depr {depr} + WC {wc_cash_effect})")
    print(f"CFI={cfi:,.2f} (-CapEx {capex})")
    print(f"CFF={cff:,.2f} (-Div {div})")
    print(f"Net Cash Flow = {net_cf:,.2f}")
    print(f"Opening Cash = {cash_open:,.2f} => Closing Cash = {cash_close:,.2f}")
    print(f"Closing Reserves = {res_close:,.2f} (Open {res_open:,.2f} + Retained {ret_earnings:,.2f})")
    print(f"Total Liabilities & Equity = {tot_liab_close:,.2f}")
    print(f"Total Assets = {tot_ast_close:,.2f}")
    print(f"BALANCE CHECK (Assets - Liabilities): {diff:,.2f} Cr")
