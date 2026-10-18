import sys
sys.path.insert(0, '.')
import json

from work.build_template_v6_models import TARGET_CONFIGS, MM_HIST_SUBS, CLEAN_STANDALONE, clean_num

bs_cols = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"]
bs_periods = ["Mar 2015", "Mar 2016", "Mar 2017", "Mar 2018", "Mar 2019", "Mar 2020", "Mar 2021", "Mar 2022", "Mar 2023", "Mar 2024", "Mar 2025", "Mar 2026"]

for target_key in ["MM", "LTM", "CLEAN", "EDELWEISS"]:
    consolidated = json.load(open(f"work/sources/{target_key}_consolidated.json", encoding="utf-8"))
    schedules = json.load(open(f"work/sources/{target_key}_schedules.json", encoding="utf-8"))
    c_sch = schedules.get("consolidated_schedules", {})
    s_sch = schedules.get("standalone_schedules", {})
    
    bs_tbl = consolidated["sections"]["balance-sheet"]["tables"][0]
    bs_hdrs = bs_tbl["headers"]
    bs_rows = bs_tbl["rows"]
    
    def get_sch(tbl_key, row_key, p):
        val = (c_sch.get(tbl_key, {}).get(row_key, {}).get(p) or
               s_sch.get(tbl_key, {}).get(row_key, {}).get(p))
        if val is None or str(val).strip() in ("", "-", "None"):
            return None
        return clean_num(val)
        
    # Simulate BS sheet cells
    bs_cells = {}
    
    for idx, c_let in enumerate(bs_cols):
        p_str = bs_periods[idx]
        is_fy26 = (idx == 11)
        
        if is_fy26:
            prev_c = bs_cols[idx-1]
            eq = bs_cells[f"{prev_c}4"]
            res = round(bs_cells[f"{prev_c}5"] * 1.08, 2)
            borr = bs_cells[f"{prev_c}10"]
            nci = bs_cells.get(f"{prev_c}12")
            pay = round(bs_cells[f"{prev_c}13"] * 1.05, 2) if bs_cells.get(f"{prev_c}13") is not None else None
            oliab = round(bs_cells[f"{prev_c}15"] * 1.05, 2) if bs_cells.get(f"{prev_c}15") is not None else None
            
            liab_items = [eq, res, borr]
            if nci is not None: liab_items.append(nci)
            if pay is not None: liab_items.append(pay)
            if oliab is not None: liab_items.append(oliab)
            tot_liab = round(sum(liab_items), 2)
            
            gb = round(bs_cells[f"{prev_c}28"] * 1.04, 2)
            dep = round(bs_cells[f"{prev_c}29"] * 1.05, 2) if (bs_cells.get(f"{prev_c}29") is not None and bs_cells.get(f"{prev_c}29") != 0) else None
            net_fa = round(gb - dep, 2) if dep is not None else gb
            cwip = bs_cells[f"{prev_c}30"]
            inv = bs_cells[f"{prev_c}31"]
            inven = round(bs_cells[f"{prev_c}33"] * 1.05, 2) if (bs_cells.get(f"{prev_c}33") is not None and bs_cells.get(f"{prev_c}33") != 0) else None
            rec = round(bs_cells[f"{prev_c}34"] * 1.05, 2) if (bs_cells.get(f"{prev_c}34") is not None and bs_cells.get(f"{prev_c}34") != 0) else None
            cash = bs_cells[f"{prev_c}35"]
            loans = bs_cells.get(f"{prev_c}36")
            
            prev_oth = bs_cells.get(f"{prev_c}37")
            if prev_oth is not None and prev_oth != 0:
                oth_ast = round(prev_oth * 1.05, 2)
            else:
                oth_ast = 0.0 if target_key == "MM" else None
                
            ast_items = [net_fa, cwip, inv]
            if inven is not None: ast_items.append(inven)
            if rec is not None: ast_items.append(rec)
            if cash is not None: ast_items.append(cash)
            if loans is not None: ast_items.append(loans)
            if oth_ast is not None and oth_ast != 0: ast_items.append(oth_ast)
            tot_ast = round(sum(ast_items), 2)
            
            bs_cells[f"{c_let}4"] = eq
            bs_cells[f"{c_let}5"] = res
            bs_cells[f"{c_let}10"] = borr
            bs_cells[f"{c_let}12"] = nci
            bs_cells[f"{c_let}13"] = pay
            bs_cells[f"{c_let}15"] = oliab
            bs_cells[f"{c_let}16"] = tot_liab
            
            bs_cells[f"{c_let}28"] = gb
            bs_cells[f"{c_let}29"] = dep
            bs_cells[f"{c_let}30"] = cwip
            bs_cells[f"{c_let}31"] = inv
            bs_cells[f"{c_let}33"] = inven
            bs_cells[f"{c_let}34"] = rec
            bs_cells[f"{c_let}35"] = cash
            bs_cells[f"{c_let}36"] = loans
            bs_cells[f"{c_let}37"] = oth_ast
            bs_cells[f"{c_let}38"] = tot_ast
            
            print(f"=== {target_key} FY26E Forecast (Col N) Non-Plug Simulation ===")
            print(f"  Liabilities: Eq={eq}, Res={res}, Borr={borr}, NCI={nci}, Pay={pay}, OLiab={oliab} => Total Liabilities (N16) = {tot_liab}")
            print(f"  Assets: NetFA={net_fa} (GB={gb}, Dep={dep}), CWIP={cwip}, Inv={inv}, Inven={inven}, Rec={rec}, Cash={cash}, Loans={loans}, OthAst={oth_ast} => Total Assets (N38) = {tot_ast}")
            print(f"  Projected Discrepancy (Total Assets - Total Liabilities): {round(tot_ast - tot_liab, 2)} Cr\n")
            
        else:
            # Historical years
            if target_key == "CLEAN" and p_str in CLEAN_STANDALONE:
                cd = CLEAN_STANDALONE[p_str]
                bs_cells[f"{c_let}4"] = cd["eq"]
                bs_cells[f"{c_let}5"] = cd["res"]
                bs_cells[f"{c_let}10"] = cd.get("oth_borr", cd["borr"])
                bs_cells[f"{c_let}12"] = None
                bs_cells[f"{c_let}13"] = cd.get("pay")
                bs_cells[f"{c_let}15"] = cd.get("oth_liab")
                bs_cells[f"{c_let}16"] = cd["tot_liab"]
                
                bs_cells[f"{c_let}28"] = cd["gb"]
                bs_cells[f"{c_let}29"] = cd["dep"] if cd["dep"] > 0 else None
                bs_cells[f"{c_let}30"] = cd["cwip"]
                bs_cells[f"{c_let}31"] = cd["inv"]
                bs_cells[f"{c_let}33"] = cd["inven"]
                bs_cells[f"{c_let}34"] = cd["rec"]
                bs_cells[f"{c_let}35"] = cd["cash"]
                bs_cells[f"{c_let}36"] = cd["loans"] if cd["loans"] > 0 else None
                bs_cells[f"{c_let}37"] = cd["oth_ast"]
                bs_cells[f"{c_let}38"] = cd["tot_ast"]
            elif target_key == "MM":
                eq_v = clean_num(bs_rows.get("Equity Capital", [])[idx])
                res_v = clean_num(bs_rows.get("Reserves", [])[idx])
                borr_v = clean_num(bs_rows.get("Borrowings", [])[idx])
                fa_v = clean_num(bs_rows.get("Fixed Assets", [])[idx])
                cwip_v = clean_num(bs_rows.get("CWIP", [])[idx])
                inv_v = clean_num(bs_rows.get("Investments", [])[idx])
                
                bs_cells[f"{c_let}4"] = eq_v
                bs_cells[f"{c_let}5"] = res_v
                bs_cells[f"{c_let}10"] = borr_v
                bs_cells[f"{c_let}28"] = fa_v
                bs_cells[f"{c_let}29"] = None
                bs_cells[f"{c_let}30"] = cwip_v
                bs_cells[f"{c_let}31"] = inv_v
                
                if idx < 4:
                    oliab_v = clean_num(bs_rows.get("Other Liabilities", [])[idx])
                    tot_liab = round(eq_v + res_v + borr_v + oliab_v, 2)
                    bs_cells[f"{c_let}12"] = None
                    bs_cells[f"{c_let}13"] = None
                    bs_cells[f"{c_let}15"] = oliab_v
                    bs_cells[f"{c_let}16"] = tot_liab
                    
                    bs_cells[f"{c_let}33"] = None
                    bs_cells[f"{c_let}34"] = None
                    bs_cells[f"{c_let}35"] = None
                    bs_cells[f"{c_let}36"] = None
                    oth_v = clean_num(bs_rows.get("Other Assets", [])[idx])
                    bs_cells[f"{c_let}37"] = oth_v
                    bs_cells[f"{c_let}38"] = round(fa_v + cwip_v + inv_v + oth_v, 2)
                else:
                    k = idx - 4
                    pay = MM_HIST_SUBS["pay"][k]
                    nci = MM_HIST_SUBS["nci"][k]
                    oth_liab = MM_HIST_SUBS["oliab"][k]
                    tot_liab = round(eq_v + res_v + borr_v + pay + nci + oth_liab, 2)
                    bs_cells[f"{c_let}12"] = nci
                    bs_cells[f"{c_let}13"] = pay
                    bs_cells[f"{c_let}15"] = oth_liab
                    bs_cells[f"{c_let}16"] = tot_liab
                    
                    inven = MM_HIST_SUBS["inven"][k]
                    rec = MM_HIST_SUBS["rec"][k]
                    cash = MM_HIST_SUBS["cash"][k]
                    loans = MM_HIST_SUBS["loans"][k]
                    bs_cells[f"{c_let}33"] = inven
                    bs_cells[f"{c_let}34"] = rec
                    bs_cells[f"{c_let}35"] = cash
                    bs_cells[f"{c_let}36"] = loans
                    bs_cells[f"{c_let}37"] = 0.0
                    bs_cells[f"{c_let}38"] = round(fa_v + cwip_v + inv_v + inven + rec + cash + loans, 2)
            else:
                h_idx = bs_hdrs.index(p_str) if p_str in bs_hdrs else -1
                eq_v = clean_num(bs_rows.get("Equity Capital", [])[h_idx]) if h_idx >= 0 else 0.0
                res_v = clean_num(bs_rows.get("Reserves", [])[h_idx]) if h_idx >= 0 else 0.0
                borr_v = clean_num(bs_rows.get("Borrowings", [])[h_idx]) if h_idx >= 0 else 0.0
                oliab_v = clean_num(bs_rows.get("Other Liabilities", [])[h_idx]) if h_idx >= 0 else 0.0
                tot_liab = round(eq_v + res_v + borr_v + oliab_v, 2)
                
                bs_cells[f"{c_let}4"] = eq_v
                bs_cells[f"{c_let}5"] = res_v
                bs_cells[f"{c_let}10"] = borr_v
                bs_cells[f"{c_let}12"] = get_sch("balance-sheet:Other Liabilities", "Non controlling int", p_str)
                bs_cells[f"{c_let}13"] = get_sch("balance-sheet:Other Liabilities", "Trade Payables", p_str)
                bs_cells[f"{c_let}15"] = get_sch("balance-sheet:Other Liabilities", "Other liability items", p_str) or oliab_v
                bs_cells[f"{c_let}16"] = tot_liab
                
                fa_v = clean_num(bs_rows.get("Fixed Assets", [])[h_idx]) if h_idx >= 0 else 0.0
                cwip_v = clean_num(bs_rows.get("CWIP", [])[h_idx]) if h_idx >= 0 else 0.0
                inv_v = clean_num(bs_rows.get("Investments", [])[h_idx]) if h_idx >= 0 else 0.0
                gross_block_v = get_sch("balance-sheet:Fixed Assets", "Gross Block", p_str) or fa_v
                accum_dep_v = get_sch("balance-sheet:Fixed Assets", "Accumulated Depreciation", p_str)
                net_fixed = (gross_block_v - (accum_dep_v or 0.0)) if gross_block_v else fa_v
                
                inv_bs_v = get_sch("balance-sheet:Other Assets", "Inventories", p_str) or 0.0
                rec_bs_v = get_sch("balance-sheet:Other Assets", "Trade receivables", p_str) or 0.0
                cash_bs_v = get_sch("balance-sheet:Other Assets", "Cash Equivalents", p_str) or 0.0
                loans_bs_v = get_sch("balance-sheet:Other Assets", "Loans n Advances", p_str) or 0.0
                oth_asset_sch_v = get_sch("balance-sheet:Other Assets", "Other asset items", p_str) or 0.0
                
                tot_ast = round(net_fixed + cwip_v + inv_v + inv_bs_v + rec_bs_v + cash_bs_v + loans_bs_v + oth_asset_sch_v, 2)
                
                bs_cells[f"{c_let}28"] = gross_block_v
                bs_cells[f"{c_let}29"] = accum_dep_v
                bs_cells[f"{c_let}30"] = cwip_v
                bs_cells[f"{c_let}31"] = inv_v
                bs_cells[f"{c_let}33"] = inv_bs_v if inv_bs_v > 0 else None
                bs_cells[f"{c_let}34"] = rec_bs_v if rec_bs_v > 0 else None
                bs_cells[f"{c_let}35"] = cash_bs_v if cash_bs_v > 0 else None
                bs_cells[f"{c_let}36"] = loans_bs_v if loans_bs_v > 0 else None
                bs_cells[f"{c_let}37"] = oth_asset_sch_v
                bs_cells[f"{c_let}38"] = tot_ast
