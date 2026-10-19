import json

# Check MM
consolidated_mm = json.load(open('work/sources/MM_consolidated.json', encoding='utf-8'))
bs_tbl_mm = consolidated_mm['sections']['balance-sheet']['tables'][0]
hdrs_mm = bs_tbl_mm['headers']
rows_mm = bs_tbl_mm['rows']

from work.build_template_v6_models import MM_HIST_SUBS, clean_num

print("=== MM Historical Component Sums ===")
for idx in range(11):
    p_str = hdrs_mm[idx]
    eq_v = clean_num(rows_mm.get("Equity Capital", [])[idx])
    res_v = clean_num(rows_mm.get("Reserves", [])[idx])
    borr_v = clean_num(rows_mm.get("Borrowings", [])[idx])
    fa_v = clean_num(rows_mm.get("Fixed Assets", [])[idx])
    cwip_v = clean_num(rows_mm.get("CWIP", [])[idx])
    inv_v = clean_num(rows_mm.get("Investments", [])[idx])
    
    if idx < 4:
        oliab_v = clean_num(rows_mm.get("Other Liabilities", [])[idx])
        tot_l = round(eq_v + res_v + borr_v + oliab_v, 2)
        oth_ast = clean_num(rows_mm.get("Other Assets", [])[idx])
        tot_a = round(fa_v + cwip_v + inv_v + oth_ast, 2)
    else:
        k = idx - 4
        pay = MM_HIST_SUBS["pay"][k]
        nci = MM_HIST_SUBS["nci"][k]
        oth_liab = MM_HIST_SUBS["oliab"][k]
        tot_l = round(eq_v + res_v + borr_v + pay + nci + oth_liab, 2)
        
        inven = MM_HIST_SUBS["inven"][k]
        rec = MM_HIST_SUBS["rec"][k]
        cash = MM_HIST_SUBS["cash"][k]
        loans = MM_HIST_SUBS["loans"][k]
        oth_ast = 0.0
        tot_a = round(fa_v + cwip_v + inv_v + inven + rec + cash + loans + oth_ast, 2)
        
    scr_tot_l = clean_num(rows_mm.get("Total Liabilities", [])[idx])
    scr_tot_a = clean_num(rows_mm.get("Total Assets", [])[idx])
    print(f"{p_str}: Sourced Liab Sum={tot_l} (Scr={scr_tot_l}, Diff={round(tot_l-scr_tot_l,2)}) | Sourced Ast Sum={tot_a} (Scr={scr_tot_a}, Diff={round(tot_a-scr_tot_a,2)}) | Ast-Liab={round(tot_a-tot_l,2)}")
