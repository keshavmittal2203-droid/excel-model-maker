import json

evidence = json.load(open('work/audit_evidence_output.json', encoding='utf-8'))

for t in ['MM', 'LTM', 'CLEAN', 'EDELWEISS']:
    print(f"\n{'='*80}\nTARGET: {t}\n{'='*80}")
    print(f"{'Period':8s} | {'Col':3s} | {'Sourced Liab':12s} | {'Scr Stated L':12s} | {'Diff L':6s} | {'Sourced Ast':12s} | {'Scr Stated A':12s} | {'Diff A':6s} | {'Ast-Liab':8s} | {'Row 37 OthAst':12s}")
    print("-" * 105)
    for r in evidence[t]['historical_reconciliation']:
        p = r['period']
        col = r['col']
        sl = f"{r['sourced_liab_sum']:,.2f}"
        ssl = f"{r['screener_stated_liab']:,.2f}" if r['screener_stated_liab'] is not None else "N/A"
        dl = f"{r['liab_diff']:+.2f}"
        sa = f"{r['sourced_ast_sum']:,.2f}"
        ssa = f"{r['screener_stated_ast']:,.2f}" if r['screener_stated_ast'] is not None else "N/A"
        da = f"{r['ast_diff']:+.2f}"
        al = f"{r['ast_minus_liab']:+.2f}"
        oa = f"{r['other_asset_items_val']:,.2f}" if r['other_asset_items_val'] is not None else "None"
        print(f"{p:8s} | {col:3s} | {sl:12s} | {ssl:12s} | {dl:6s} | {sa:12s} | {ssa:12s} | {da:6s} | {al:8s} | {oa:12s}")
        
    fy26 = evidence[t]['fy26e_components']
    print("\nFY26E Forecast (Col N):")
    print(f"  Total Liabilities (N16) Formula : {fy26['total_liab_formula']} = INR {fy26['total_liab_value']:,.2f} Cr")
    print(f"  Row 37 Other Assets Formula     : {fy26['other_asset_formula']} = INR {fy26['other_asset_value']} Cr")
    print(f"  Total Assets (N38) Formula      : {fy26['total_ast_formula']} = INR {fy26['total_ast_value']:,.2f} Cr")
    print(f"  Un-plugged Forecast Variance    : Total Assets - Total Liabilities = INR {fy26['unplugged_diff']:+,.2f} Cr")

    if t in ["LTM", "EDELWEISS"]:
        print(f"\n{t} Borrowings Sub-Schedule Reconciliation:")
        print(f"{'Period':10s} | {'Screener Total':14s} | {'ST Borr (R7)':12s} | {'Lease (R8)':10s} | {'Pref (R9)':10s} | {'LT Borr':10s} | {'Oth Borr':10s} | {'Sub Sum':10s} | {'Diff':6s}")
        print("-" * 105)
        for b in evidence[t]['borrowings_subschedule']:
            print(f"{b['period']:10s} | {b['screener_borr']:14,.2f} | {b['st_borr']:12,.2f} | {b['lease']:10,.2f} | {b['pref']:10,.2f} | {b['lt_borr']:10,.2f} | {b['oth_borr']:10,.2f} | {b['sub_sum']:10,.2f} | {b['diff']:+6.2f}")
