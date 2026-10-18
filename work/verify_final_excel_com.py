import os, json, win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

targets = ["MM", "LTM", "CLEAN", "EDELWEISS"]
results = {}

for t in targets:
    path = os.path.abspath(f"companies/{t}/template_v16/{t}-FullTemplate.xlsx")
    wb = excel.Workbooks.Open(path)
    excel.CalculateFullRebuild()

    bs = wb.Sheets("BS")
    c40_n40_f = [bs.Cells(40, c).Formula for c in range(3, 15)]
    c40_n40_v = [bs.Cells(40, c).Value for c in range(3, 15)]
    c16_n16_v = [bs.Cells(16, c).Value for c in range(3, 15)]
    c38_n38_v = [bs.Cells(38, c).Value for c in range(3, 15)]

    # Historical values in column M (FY25)
    m37_val = bs.Range("M37").Value
    m15_val = bs.Range("M15").Value
    m16_val = bs.Range("M16").Value
    m38_val = bs.Range("M38").Value

    # Formula errors across all 16 sheets
    err_count = 0
    err_list = []
    for s_idx in range(1, wb.Sheets.Count + 1):
        s = wb.Sheets(s_idx)
        used = s.UsedRange.Value
        if used:
            rows = used if isinstance(used, tuple) else (used,)
            for r_idx, row in enumerate(rows):
                if not isinstance(row, tuple):
                    row = (row,)
                for c_idx, cell in enumerate(row):
                    if isinstance(cell, int) and cell in [
                        -2146826281, # #DIV/0!
                        -2146826273, # #VALUE!
                        -2146826265, # #REF!
                        -2146826259, # #NAME?
                        -2146826252, # #NUM!
                        -2146826288  # #NULL!
                    ]:
                        err_count += 1
                        err_list.append((s.Name, r_idx + 1, c_idx + 1, cell))

    dcf = wb.Sheets("DCF 1st way")
    val_e42 = dcf.Range("E42").Value
    cov = wb.Sheets("Model ->")
    val_c9 = cov.Range("C9").Value

    results[t] = {
        "row40_formulas": c40_n40_f,
        "row40_values": [round(x, 4) if x is not None else None for x in c40_n40_v],
        "row16_values": [round(x, 2) if x is not None else None for x in c16_n16_v],
        "row38_values": [round(x, 2) if x is not None else None for x in c38_n38_v],
        "m37": m37_val,
        "m15": m15_val,
        "m16": m16_val,
        "m38": m38_val,
        "err_count": err_count,
        "err_list": err_list,
        "val_e42": val_e42,
        "val_c9": val_c9
    }
    wb.Close(False)

excel.Quit()

for t, r in results.items():
    print(f"=== {t} ===")
    print("Row 40 Formulas:", r["row40_formulas"])
    print("Row 40 Values (C to N):", r["row40_values"])
    print(f"FY25 (Col M): M37={r['m37']}, M15={r['m15']}, M38={r['m38']}, M16={r['m16']}")
    print(f"FY26E (Col N): Assets={r['row38_values'][-1]}, Liab={r['row16_values'][-1]}, Diff={r['row40_values'][-1]}")
    print(f"Formula Errors: {r['err_count']}, DCF E42={r['val_e42']}, Cover C9={r['val_c9']}")
