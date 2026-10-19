import json

def to_f(v):
    if v is None: return 0.0
    try: return float(str(v).replace(',', ''))
    except: return 0.0

targets = ["MM", "LTM", "CLEAN", "EDELWEISS"]

for t in targets:
    consolidated = json.load(open(f"work/sources/{t}_consolidated.json", encoding="utf-8"))
    schedules = json.load(open(f"work/sources/{t}_schedules.json", encoding="utf-8"))
    c_sch = schedules.get("consolidated_schedules", {})
    s_sch = schedules.get("standalone_schedules", {})
    
    bs_tbl = consolidated["sections"]["balance-sheet"]["tables"][0]
    hdrs = bs_tbl["headers"]
    rows = bs_tbl["rows"]
    
    print(f"\n{'='*70}\nTARGET: {t} SCREENER REPORTED BALANCE SHEET RECONCILIATION\n{'='*70}")
    print(f"Periods: {hdrs[:11]}")
    
    for idx, p in enumerate(hdrs[:11]):
        eq = to_f(rows.get("Equity Capital", [])[idx])
        res = to_f(rows.get("Reserves", [])[idx])
        borr = to_f(rows.get("Borrowings", [])[idx])
        oliab = to_f(rows.get("Other Liabilities", [])[idx])
        tot_liab_stated = to_f(rows.get("Total Liabilities", [])[idx])
        calc_liab = round(eq + res + borr + oliab, 2)
        diff_liab = round(calc_liab - tot_liab_stated, 2)
        
        fa = to_f(rows.get("Fixed Assets", [])[idx])
        cwip = to_f(rows.get("CWIP", [])[idx])
        inv = to_f(rows.get("Investments", [])[idx])
        oast = to_f(rows.get("Other Assets", [])[idx])
        tot_ast_stated = to_f(rows.get("Total Assets", [])[idx])
        calc_ast = round(fa + cwip + inv + oast, 2)
        diff_ast = round(calc_ast - tot_ast_stated, 2)
        
        # Sub-schedules for Other Assets
        sch_oast = c_sch.get("balance-sheet:Other Assets", {}) or s_sch.get("balance-sheet:Other Assets", {})
        sub_inven = to_f(sch_oast.get("Inventories", {}).get(p))
        sub_rec = to_f(sch_oast.get("Trade receivables", {}).get(p))
        sub_cash = to_f(sch_oast.get("Cash Equivalents", {}).get(p))
        sub_loans = to_f(sch_oast.get("Loans n Advances", {}).get(p))
        sub_oth = to_f(sch_oast.get("Other asset items", {}).get(p))
        sum_oast = round(sub_inven + sub_rec + sub_cash + sub_loans + sub_oth, 2)
        diff_oast = round(sum_oast - oast, 2)
        
        print(f"\n-- {t} | Period: {p} --")
        print(f"  Liabilities: Eq={eq}, Res={res}, Borr={borr}, OLiab={oliab} | Stated={tot_liab_stated}, Calc={calc_liab}, Diff={diff_liab}")
        print(f"  Assets:      FA={fa}, CWIP={cwip}, Inv={inv}, OAssets={oast} | Stated={tot_ast_stated}, Calc={calc_ast}, Diff={diff_ast}")
        print(f"  OAssets Sub: Inven={sub_inven}, Rec={sub_rec}, Cash={sub_cash}, Loans={sub_loans}, Other={sub_oth} | Sum={sum_oast}, Screener OAssets={oast}, Diff={diff_oast}")
