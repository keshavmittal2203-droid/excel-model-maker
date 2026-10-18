import sys
sys.path.insert(0, '.')
import json
import openpyxl

from work.build_template_v6_models import clean_num, MM_HIST_SUBS, CLEAN_STANDALONE

targets = ["MM", "LTM", "CLEAN", "EDELWEISS"]
bs_cols = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"]
bs_periods = ["FY15", "FY16", "FY17", "FY18", "FY19", "FY20", "FY21", "FY22", "FY23", "FY24", "FY25", "FY26E"]

with open("work/audit_evidence_output.json", "w", encoding="utf-8") as out_f:
    evidence = {}
    
    for t in targets:
        wb_path = f"companies/{t}/template_v6/{t}-FullTemplate.xlsx"
        wb_form = openpyxl.load_workbook(wb_path, data_only=False)
        wb_val = openpyxl.load_workbook(wb_path, data_only=True)
        bs_f = wb_form["BS"]
        bs_v = wb_val["BS"]
        
        consolidated = json.load(open(f"work/sources/{t}_consolidated.json", encoding="utf-8"))
        schedules = json.load(open(f"work/sources/{t}_schedules.json", encoding="utf-8"))
        bs_tbl = consolidated["sections"]["balance-sheet"]["tables"][0]
        hdrs = bs_tbl["headers"]
        rows = bs_tbl["rows"]
        
        target_ev = {
            "historical_reconciliation": [],
            "fy26e_components": {},
            "borrowings_subschedule": []
        }
        
        # Historical & FY26E
        for idx, col in enumerate(bs_cols):
            p = bs_periods[idx]
            
            liab_f = bs_f[f"{col}16"].value
            liab_v = bs_v[f"{col}16"].value
            ast_f = bs_f[f"{col}38"].value
            ast_v = bs_v[f"{col}38"].value
            oth_ast_f = bs_f[f"{col}37"].value
            oth_ast_v = bs_v[f"{col}37"].value
            
            if col != "N":
                # Screener stated
                p_hdr = f"Mar {2015+idx}"
                h_idx = hdrs.index(p_hdr) if p_hdr in hdrs else -1
                
                scr_tot_l = clean_num(rows.get("Total Liabilities", [])[h_idx]) if h_idx >= 0 else None
                scr_tot_a = clean_num(rows.get("Total Assets", [])[h_idx]) if h_idx >= 0 else None
                
                target_ev["historical_reconciliation"].append({
                    "col": col,
                    "period": p,
                    "sourced_liab_sum": liab_v,
                    "screener_stated_liab": scr_tot_l,
                    "liab_diff": round(float(liab_v) - float(scr_tot_l), 2) if scr_tot_l is not None else 0.0,
                    "sourced_ast_sum": ast_v,
                    "screener_stated_ast": scr_tot_a,
                    "ast_diff": round(float(ast_v) - float(scr_tot_a), 2) if scr_tot_a is not None else 0.0,
                    "ast_minus_liab": round(float(ast_v) - float(liab_v), 2),
                    "other_asset_items_val": oth_ast_v
                })
            else:
                target_ev["fy26e_components"] = {
                    "total_liab_formula": str(liab_f),
                    "total_liab_value": liab_v,
                    "total_ast_formula": str(ast_f),
                    "total_ast_value": ast_v,
                    "other_asset_formula": str(oth_ast_f),
                    "other_asset_value": oth_ast_v,
                    "unplugged_diff": round(float(ast_v) - float(liab_v), 2)
                }
                
        # Borrowings subschedule (for LTM & EDELWEISS)
        if t in ["LTM", "EDELWEISS"]:
            c_sch = schedules.get("consolidated_schedules", {})
            s_sch = schedules.get("standalone_schedules", {})
            def get_sch(tbl_key, row_key, p_str):
                val = (c_sch.get(tbl_key, {}).get(row_key, {}).get(p_str) or
                       s_sch.get(tbl_key, {}).get(row_key, {}).get(p_str))
                if val is None or str(val).strip() in ("", "-", "None"): return 0.0
                return clean_num(val)
                
            for idx in range(11):
                p_str = f"Mar {2015+idx}"
                h_idx = hdrs.index(p_str) if p_str in hdrs else -1
                scr_borr = clean_num(rows.get("Borrowings", [])[h_idx]) if h_idx >= 0 else 0.0
                st = get_sch("balance-sheet:Borrowings", "Short term Borrowings", p_str)
                lease = get_sch("balance-sheet:Borrowings", "Lease Liabilities", p_str)
                pref = get_sch("balance-sheet:Borrowings", "Preference Capital", p_str)
                lt = get_sch("balance-sheet:Borrowings", "Long term Borrowings", p_str)
                oth = get_sch("balance-sheet:Borrowings", "Other Borrowings", p_str)
                sub_sum = round(st + lease + pref + lt + oth, 2)
                
                target_ev["borrowings_subschedule"].append({
                    "period": p_str,
                    "screener_borr": scr_borr,
                    "st_borr": st,
                    "lease": lease,
                    "pref": pref,
                    "lt_borr": lt,
                    "oth_borr": oth,
                    "sub_sum": sub_sum,
                    "diff": round(sub_sum - scr_borr, 2)
                })
                
        evidence[t] = target_ev
        
    json.dump(evidence, out_f, indent=2)
print("Audit evidence extracted to work/audit_evidence_output.json")
