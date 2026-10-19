"""Test the exact impact of updating Rf from 6.78% to 6.51% across all 4 companies.
"""

import openpyxl

targets = ['MM', 'LTM', 'CLEAN', 'EDELWEISS']

for t in targets:
    wb = openpyxl.load_workbook(f'companies/{t}/template_v6/{t}-FullTemplate.xlsx', data_only=True)
    wacc_ws = wb['WACC-DCF']
    dcf1 = wb['DCF 1st way']
    
    rf_old = wacc_ws['J22'].value
    erp = wacc_ws['J23'].value
    beta = wacc_ws['J24'].value
    wacc_old = wacc_ws['J39'].value
    val_old = dcf1['E42'].value
    
    # Calculate with Rf = 6.51%
    rf_new = 0.0651
    ke_new = rf_new + beta * erp
    
    we = wacc_ws['E31'].value
    wd = wacc_ws['E30'].value
    kd_at = wacc_ws['E24'].value
    
    wacc_new = round(we * ke_new + wd * kd_at, 4)
    
    print(f"=== {t} ===")
    print(f"  Old Rf: {rf_old*100:.2f}%, New Rf: {rf_new*100:.2f}%")
    print(f"  Beta: {beta:.4f}, ERP: {erp*100:.2f}%")
    print(f"  Old Ke: {(rf_old + beta*erp)*100:.4f}%, New Ke: {ke_new*100:.4f}%")
    print(f"  Old WACC: {wacc_old*100:.2f}%, New WACC: {wacc_new*100:.2f}%")
    print(f"  Current Valuation E42: {val_old}")
    
    if t in ['MM', 'EDELWEISS']:
        print(f"  Valuation Methodology: SOTP (independent of WACC). Implied per share: {val_old:.2f}\n")
    else:
        # Re-evaluate DCF 1st way with new WACC
        k13 = dcf1['K13'].value
        g = dcf1['E27'].value
        e25 = k13 * (1.0 + g)
        e29_new = e25 / (wacc_new - g)
        
        # Explicit PVs
        pv_explicit = 0.0
        for col, yr in [('G', 1), ('H', 2), ('I', 3), ('J', 4), ('K', 5)]:
            fcff = dcf1[f'{col}13'].value
            df = 1.0 / ((1.0 + wacc_new) ** yr)
            pv_explicit += fcff * df
            
        df_tv = 1.0 / ((1.0 + wacc_new) ** 5)
        pv_tv_new = e29_new * df_tv
        ev_new = pv_explicit + pv_tv_new
        cash = dcf1['E38'].value
        debt = dcf1['E39'].value
        eq_new = ev_new + cash - debt
        shares = dcf1['E41'].value
        val_new = eq_new / shares
        
        print(f"  PV of Explicit FCFF: Old = {dcf1['E34'].value:,.2f} -> New = {pv_explicit:,.2f} Cr")
        print(f"  Terminal Value (E29): Old = {dcf1['E29'].value:,.2f} -> New = {e29_new:,.2f} Cr")
        print(f"  PV of Terminal Value (E35): Old = {dcf1['E35'].value:,.2f} -> New = {pv_tv_new:,.2f} Cr")
        print(f"  EV (E36): Old = {dcf1['E36'].value:,.2f} -> New = {ev_new:,.2f} Cr")
        print(f"  Equity Value (E40): Old = {dcf1['E40'].value:,.2f} -> New = {eq_new:,.2f} Cr")
        print(f"  Valuation E42: Old = INR {val_old:.2f} -> New = INR {val_new:.2f} / share (delta: {val_new - val_old:+.2f})\n")
