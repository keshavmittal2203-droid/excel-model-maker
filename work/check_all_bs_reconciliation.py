import sys
sys.path.insert(0, '.')
import json
from work.build_template_v6_models import clean_num, MM_HIST_SUBS, CLEAN_STANDALONE

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
    
    print(f"\n{'='*75}\nTARGET: {t} DETAILED BALANCE SHEET AUDIT\n{'='*75}")
    
    bs_periods = ["Mar 2015", "Mar 2016", "Mar 2017", "Mar 2018", "Mar 2019", "Mar 2020", "Mar 2021", "Mar 2022", "Mar 2023", "Mar 2024", "Mar 2025"]
    
    for idx, p in enumerate(bs_periods):
        # Case CLEAN pre-FY18 standalone
        if t == "CLEAN" and p in CLEAN_STANDALONE:
            cd = CLEAN_STANDALONE[p]
            eq = cd["eq"]
            res = cd["res"]
            borr = cd["borr"]
            oliab = cd["oliab"]
            tot_l_calc = round(eq + res + borr + oliab, 2)
            tot_l_stated = cd["tot_liab"]
            
            gb = cd["gb"]
            dep = cd["dep"]
            fa = round(gb - dep, 2)
            cwip = cd["cwip"]
            inv = cd["inv"]
            inven = cd["inven"]
            rec = cd["rec"]
            cash = cd["cash"]
            loans = cd["loans"]
            oth_ast = cd["oth_ast"]
            tot_a_calc = round(fa + cwip + inv + inven + rec + cash + loans + oth_ast, 2)
            tot_a_stated = cd["tot_ast"]
            
            print(f"{p} (Standalone):")
            print(f"  Liabilities: Eq={eq}, Res={res}, Borr={borr}, OLiab={oliab} => Calc={tot_l_calc}, Stated={tot_l_stated}, Diff={round(tot_l_calc-tot_l_stated,2)}")
            print(f"  Assets: NetFA={fa}, CWIP={cwip}, Inv={inv}, Inven={inven}, Rec={rec}, Cash={cash}, Loans={loans}, OthAst={oth_ast} => Calc={tot_a_calc}, Stated={tot_a_stated}, Diff={round(tot_a_calc-tot_a_stated,2)}")
            print(f"  Balance Check (Calc Assets - Calc Liabilities): {round(tot_a_calc - tot_l_calc, 2)}")
            continue
            
        # Case MM
        if t == "MM":
            eq = clean_num(rows.get("Equity Capital", [])[idx])
            res = clean_num(rows.get("Reserves", [])[idx])
            borr = clean_num(rows.get("Borrowings", [])[idx])
            fa = clean_num(rows.get("Fixed Assets", [])[idx])
            cwip = clean_num(rows.get("CWIP", [])[idx])
            inv = clean_num(rows.get("Investments", [])[idx])
            tot_l_stated = clean_num(rows.get("Total Liabilities", [])[idx])
            tot_a_stated = clean_num(rows.get("Total Assets", [])[idx])
            
            if idx < 4:
                oliab = clean_num(rows.get("Other Liabilities", [])[idx])
                tot_l_calc = round(eq + res + borr + oliab, 2)
                oth_ast = clean_num(rows.get("Other Assets", [])[idx])
                tot_a_calc = round(fa + cwip + inv + oth_ast, 2)
                print(f"{p}:")
                print(f"  Liabilities: Eq={eq}, Res={res}, Borr={borr}, OLiab={oliab} => Calc={tot_l_calc}, Stated={tot_l_stated}, Diff={round(tot_l_calc-tot_l_stated,2)}")
                print(f"  Assets: FA={fa}, CWIP={cwip}, Inv={inv}, ScreenerOtherAssets={oth_ast} => Calc={tot_a_calc}, Stated={tot_a_stated}, Diff={round(tot_a_calc-tot_a_stated,2)}")
                print(f"  Balance Check (Calc Assets - Calc Liabilities): {round(tot_a_calc - tot_l_calc, 2)}")
            else:
                k = idx - 4
                pay = MM_HIST_SUBS["pay"][k]
                nci = MM_HIST_SUBS["nci"][k]
                oth_liab = MM_HIST_SUBS["oliab"][k]
                tot_l_calc = round(eq + res + borr + pay + nci + oth_liab, 2)
                
                inven = MM_HIST_SUBS["inven"][k]
                rec = MM_HIST_SUBS["rec"][k]
                cash = MM_HIST_SUBS["cash"][k]
                loans = MM_HIST_SUBS["loans"][k]
                oth_ast = 0.0
                tot_a_calc = round(fa + cwip + inv + inven + rec + cash + loans + oth_ast, 2)
                print(f"{p}:")
                print(f"  Liabilities: Eq={eq}, Res={res}, Borr={borr}, Pay={pay}, NCI={nci}, OLiab={oth_liab} => Calc={tot_l_calc}, Stated={tot_l_stated}, Diff={round(tot_l_calc-tot_l_stated,2)}")
                print(f"  Assets: FA={fa}, CWIP={cwip}, Inv={inv}, Inven={inven}, Rec={rec}, Cash={cash}, Loans={loans}, OthAst={oth_ast} => Calc={tot_a_calc}, Stated={tot_a_stated}, Diff={round(tot_a_calc-tot_a_stated,2)}")
                print(f"  Balance Check (Calc Assets - Calc Liabilities): {round(tot_a_calc - tot_l_calc, 2)}")
            continue
            
        # Case LTM, CLEAN (>=FY18), EDELWEISS
        h_idx = hdrs.index(p) if p in hdrs else -1
        eq = clean_num(rows.get("Equity Capital", [])[h_idx]) if h_idx >= 0 else 0.0
        res = clean_num(rows.get("Reserves", [])[h_idx]) if h_idx >= 0 else 0.0
        borr = clean_num(rows.get("Borrowings", [])[h_idx]) if h_idx >= 0 else 0.0
        oliab = clean_num(rows.get("Other Liabilities", [])[h_idx]) if h_idx >= 0 else 0.0
        tot_l_stated = clean_num(rows.get("Total Liabilities", [])[h_idx]) if h_idx >= 0 else 0.0
        tot_l_calc = round(eq + res + borr + oliab, 2)
        
        fa = clean_num(rows.get("Fixed Assets", [])[h_idx]) if h_idx >= 0 else 0.0
        cwip = clean_num(rows.get("CWIP", [])[h_idx]) if h_idx >= 0 else 0.0
        inv = clean_num(rows.get("Investments", [])[h_idx]) if h_idx >= 0 else 0.0
        tot_a_stated = clean_num(rows.get("Total Assets", [])[h_idx]) if h_idx >= 0 else 0.0
        
        # Sub-schedules
        sch_fa = c_sch.get("balance-sheet:Fixed Assets", {}) or s_sch.get("balance-sheet:Fixed Assets", {})
        gb = to_f(sch_fa.get("Gross Block", {}).get(p)) or fa
        dep = to_f(sch_fa.get("Accumulated Depreciation", {}).get(p)) or 0.0
        net_fa = round(gb - dep, 2) if gb else fa
        
        sch_oast = c_sch.get("balance-sheet:Other Assets", {}) or s_sch.get("balance-sheet:Other Assets", {})
        inven = to_f(sch_oast.get("Inventories", {}).get(p))
        rec = to_f(sch_oast.get("Trade receivables", {}).get(p))
        cash = to_f(sch_oast.get("Cash Equivalents", {}).get(p))
        loans = to_f(sch_oast.get("Loans n Advances", {}).get(p))
        oth_ast = to_f(sch_oast.get("Other asset items", {}).get(p))
        
        tot_a_calc = round(net_fa + cwip + inv + inven + rec + cash + loans + oth_ast, 2)
        
        print(f"{p}:")
        print(f"  Liabilities: Eq={eq}, Res={res}, Borr={borr}, OLiab={oliab} => Calc={tot_l_calc}, Stated={tot_l_stated}, Diff={round(tot_l_calc-tot_l_stated,2)}")
        print(f"  Assets: NetFA={net_fa} (GB={gb}, Dep={dep}), CWIP={cwip}, Inv={inv}, Inven={inven}, Rec={rec}, Cash={cash}, Loans={loans}, OthAst={oth_ast} => Calc={tot_a_calc}, Stated={tot_a_stated}, Diff={round(tot_a_calc-tot_a_stated,2)}")
        print(f"  Balance Check (Calc Assets - Calc Liabilities): {round(tot_a_calc - tot_l_calc, 2)}")
