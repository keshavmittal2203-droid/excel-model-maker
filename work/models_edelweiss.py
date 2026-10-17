"""Native OpenPyXL Model Builder for Edelweiss Financial Services Limited (EDELWEISS).
Strictly blocks industrial FCFF DCF and enforces Financial-Services Sum-of-the-Parts (SOTP).
"""
import sys
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
for p in (str(ROOT), str(Path(__file__).resolve().parent)):
    if p not in sys.path:
        sys.path.insert(0, p)
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl_styles import (
    FONT_TITLE, FONT_SUBTITLE, FONT_SEC_HDR, FONT_TBL_HDR, FONT_REGULAR, FONT_BOLD,
    FONT_FORMULA, FONT_LINK, FONT_SMALL_ITALIC,
    FILL_NAVY, FILL_SEC_HDR, FILL_SUB_HDR, FILL_ZEBRA, FILL_TOTAL, FILL_ACCENT,
    BORDER_DATA, BORDER_TOTAL, BORDER_HEADER,
    ALIGN_LEFT, ALIGN_RIGHT, ALIGN_CENTER,
    FMT_INR_CR, FMT_DEC2, FMT_PCT, FMT_INT, FMT_MULT,
    style_range, set_col_widths, setup_page, apply_cell_color_coding
)


def build_edelweiss_workbook(out_path: Path) -> Path:
    return build_edelweiss_model(out_path)


def build_edelweiss_model(out_path: Path) -> Path:
    wb = openpyxl.Workbook()
    ws_exec = wb.active
    ws_exec.title = "Cover & Notes"
    ws_is = wb.create_sheet("Income Statement")
    ws_bs = wb.create_sheet("Balance Sheet")
    ws_cf = wb.create_sheet("Cash Flow")
    ws_ops = wb.create_sheet("Segment Operations")
    ws_val = wb.create_sheet("SOTP Valuation")

    # 7 audited historical years (FY19-FY25, cols B-H) + 5 connected forecast years (FY26E-FY30E, cols I-M)
    periods = ["FY19", "FY20", "FY21", "FY22", "FY23", "FY24", "FY25", "FY26E", "FY27E", "FY28E", "FY29E", "FY30E"]

    # =========================================================================
    # 1. SEGMENT OPERATIONS SCHEDULE (Sheet 5 — Connected Multi-Vertical Engine)
    # =========================================================================
    setup_page(ws_ops, "Edelweiss Financial Services Limited - Segment Operations, AUM & Ratios")
    ws_ops["A1"] = "EDELWEISS FINANCIAL SERVICES LTD — SEGMENT OPERATIONS, AUM, IEV, BOOK VALUE & RATIOS"
    ws_ops["A1"].font = FONT_TITLE
    ws_ops["A2"] = "Multi-Vertical Operating Engine Feeding SOTP Valuation Base Metrics ('SOTP Valuation'!B6:B10) & Consolidated Revenue"
    ws_ops["A2"].font = FONT_SMALL_ITALIC

    ws_ops.cell(3, 1, "Business Vertical / Operating Metric (₹ Cr)").font = FONT_TBL_HDR
    ws_ops.cell(3, 1).fill = FILL_NAVY
    for idx, p in enumerate(periods, start=2):
        c = ws_ops.cell(3, idx, p)
        c.font = FONT_TBL_HDR
        c.fill = FILL_NAVY
        c.alignment = ALIGN_RIGHT

    ws_ops["A4"] = "--- Core Financial Services Operating Verticals (FY19–FY25 Audited / Segment Disclosures) ---"
    ws_ops["A4"].font = FONT_SEC_HDR
    ws_ops["A4"].fill = FILL_SEC_HDR

    # Historical Vertical Metrics (FY19-FY25, 7 years) -> FY25 values (Col H) feed SOTP Valuation B6:B10!
    eaaa_aum_hist  = [18500, 21200, 26500, 32400, 41800, 50200, 56000]
    eaaa_pat_hist  = [50, 75, 110, 145, 160, 175, 230]
    earcl_aum_hist = [46500, 43200, 40800, 39500, 37100, 31500, 28000]
    earcl_nw_hist  = [1850, 2010, 2180, 2350, 2520, 2680, 2800]
    mf_aaum_hist   = [12000, 28000, 55000, 85000, 105000, 127000, 141000]
    ins_ev_hist    = [950, 1050, 1150, 1250, 1350, 1420, 1500]
    cred_bk_hist   = [36100, 28400, 21500, 16800, 14200, 11800, 10200]
    cred_nw_hist   = [4500, 3800, 3600, 3500, 3400, 3300, 3200]

    # Consolidated Revenue Split (Interest & Credit Treasury vs Fee, Commission & Insurance Premium) summing to Consolidated Sales
    cons_rev_hist  = [11078, 9513, 10775, 7151, 8500, 9523, 9393]
    int_inc_hist   = [7200, 5850, 5100, 3450, 3650, 3850, 3600]
    fee_prem_hist  = [cons_rev_hist[i] - int_inc_hist[i] for i in range(7)]

    # Ratios from work/sources/EDELWEISS_consolidated.json (Mar 2019 - Mar 2025)
    debtor_days_hist = [91, 50, 17, 24, 18, 14, 15]
    roce_hist        = [0.12, 0.05, 0.10, 0.10, 0.10, 0.12, 0.13]

    ops_rows = {
        5: ("EAAA (Alternatives - 100% Stake) Fee-Paying AUM (₹ Cr)", FMT_INR_CR, False),
        6: ("EAAA (Alternatives - 100% Stake) Operating PAT (₹ Cr) [Feeds SOTP B6]", FMT_INR_CR, True),
        7: ("EARCL (Asset Reconstruction) Distressed Assets AUM (₹ Cr)", FMT_INR_CR, False),
        8: ("EARCL Attributable Net Worth (₹ Cr, Net of CDPQ 20.2% NCI) [Feeds SOTP B7]", FMT_INR_CR, True),
        9: ("Edelweiss Mutual Fund (EAML - 85% Post-WestBridge) Total AAUM (₹ Cr) [Feeds SOTP B8]", FMT_INR_CR, True),
        10: ("Insurance Attributable Actuarial Indian Embedded Value (IEV, NOT EV - ₹ Cr) [Feeds SOTP B9]", FMT_INR_CR, True),
        11: ("NBFC & Housing Credit Book (Wholesale Rundown + Retail MSME - ₹ Cr)", FMT_INR_CR, False),
        12: ("NBFC & Housing Credit Allocated Net Worth (100% Stake - ₹ Cr) [Feeds SOTP B10]", FMT_INR_CR, True),
        13: ("--- Consolidated Operating Revenue Build-Up & Vertical Drivers ---", None, True),
        14: ("Interest & Credit Book Yield Income (₹ Cr) [Driven by Credit Book & Yield]", FMT_INR_CR, False),
        15: ("Fee, AMC/EAAA Commission & Insurance Premium Income (₹ Cr)", FMT_INR_CR, False),
        16: ("Total Consolidated Operating Income (₹ Cr) [= Row 14 + Row 15]", FMT_INR_CR, True),
        17: ("Fee, Asset Management & Insurance GWP Growth Driver (%)", FMT_PCT, False),
        18: ("Operating, Credit Provision & Policyholder Expense Ratio (% of Income)", FMT_PCT, False),
        19: ("Effective Tax Rate (% of PBT)", FMT_PCT, False),
        20: ("Dividend Payout Ratio (% of PAT)", FMT_PCT, False),
        21: ("--- Leverage & Return Ratios (Primary Filing / Derived) ---", None, True),
        22: ("Consolidated Debt-to-Equity Ratio (x)", FMT_MULT, False),
        23: ("Debtor / Fee Collection Days", FMT_INT, False),
        24: ("Consolidated ROCE (%)", FMT_PCT, False),
        25: ("Consolidated ROE (%)", FMT_PCT, True),
    }

    for r, (lbl, _, is_tot) in ops_rows.items():
        c = ws_ops.cell(r, 1, lbl)
        if lbl.startswith("---"):
            c.font = FONT_SEC_HDR
            c.fill = FILL_SEC_HDR
        else:
            c.font = FONT_BOLD if is_tot else FONT_REGULAR
            c.fill = FILL_TOTAL if is_tot else (FILL_ZEBRA if r % 2 == 1 else openpyxl.styles.PatternFill(fill_type=None))
            c.border = BORDER_TOTAL if is_tot else BORDER_DATA

    for i in range(7):
        col = i + 2
        col_let = get_column_letter(col)
        prev_let = get_column_letter(col - 1) if col > 2 else None
        ws_ops.cell(5, col, eaaa_aum_hist[i])
        ws_ops.cell(6, col, eaaa_pat_hist[i])
        ws_ops.cell(7, col, earcl_aum_hist[i])
        ws_ops.cell(8, col, earcl_nw_hist[i])
        ws_ops.cell(9, col, mf_aaum_hist[i])
        ws_ops.cell(10, col, ins_ev_hist[i])
        ws_ops.cell(11, col, cred_bk_hist[i])
        ws_ops.cell(12, col, cred_nw_hist[i])

        ws_ops.cell(14, col, int_inc_hist[i])
        ws_ops.cell(15, col, fee_prem_hist[i])
        ws_ops.cell(16, col, f"={col_let}14+{col_let}15")
        ws_ops.cell(17, col, 0.120 if i == 0 else f"=({col_let}15-{prev_let}15)/{prev_let}15")
        ws_ops.cell(18, col, f"='Income Statement'!{col_let}6/'Income Statement'!{col_let}5")
        ws_ops.cell(19, col, 0.250)
        ws_ops.cell(20, col, 0.250)

        ws_ops.cell(22, col, f"='Balance Sheet'!{col_let}6/'Balance Sheet'!{col_let}16")
        ws_ops.cell(23, col, debtor_days_hist[i])
        ws_ops.cell(24, col, roce_hist[i])
        ws_ops.cell(25, col, f"='Income Statement'!{col_let}14/'Balance Sheet'!{col_let}16")

    # Forecast Columns I..M (idx 9..13, FY26E..FY30E) — Vertical-Specific Drivers
    fee_g_fcst  = [0.115, 0.115, 0.110, 0.110, 0.105]
    opex_r_fcst = [0.640, 0.635, 0.630, 0.625, 0.620]

    for j in range(5):
        col = j + 9
        col_let = get_column_letter(col)
        prev_let = get_column_letter(col - 1)
        ws_ops.cell(5, col, f"=ROUND({prev_let}5*1.15,0)")
        ws_ops.cell(6, col, f"=ROUND({col_let}5*0.0042,0)")
        ws_ops.cell(7, col, f"=ROUND({prev_let}7*0.95,0)")
        ws_ops.cell(8, col, f"=ROUND({prev_let}8*1.07,0)")
        ws_ops.cell(9, col, f"=ROUND({prev_let}9*1.16,0)")
        ws_ops.cell(10, col, f"=ROUND({prev_let}10*1.10,0)")
        ws_ops.cell(11, col, f"=ROUND({prev_let}11*1.04,0)")
        ws_ops.cell(12, col, f"=ROUND({prev_let}12*1.05,0)")

        ws_ops.cell(17, col, fee_g_fcst[j])
        ws_ops.cell(18, col, opex_r_fcst[j])
        ws_ops.cell(19, col, 0.250)
        ws_ops.cell(20, col, 0.250)

        # Interest Income linked to Credit Book (Row 11) + Treasury Yield
        ws_ops.cell(14, col, f"=ROUND({prev_let}14*(1+({col_let}11-{prev_let}11)/{prev_let}11*0.75),0)")
        # Fee & Insurance Premium Income linked to EAAA/AMC/Insurance growth driver
        ws_ops.cell(15, col, f"=ROUND({prev_let}15*(1+{col_let}17),0)")
        ws_ops.cell(16, col, f"={col_let}14+{col_let}15")

        ws_ops.cell(22, col, f"='Balance Sheet'!{col_let}6/'Balance Sheet'!{col_let}16")
        ws_ops.cell(23, col, 15)
        ws_ops.cell(24, col, f"=('Income Statement'!{col_let}7-'Income Statement'!{col_let}9)/('Balance Sheet'!{col_let}16+'Balance Sheet'!{col_let}6)")
        ws_ops.cell(25, col, f"='Income Statement'!{col_let}14/'Balance Sheet'!{col_let}16")

    for r, (_, fmt, is_tot) in ops_rows.items():
        if fmt is None:
            for c_idx in range(2, 14):
                ws_ops.cell(r, c_idx).fill = FILL_SEC_HDR
            continue
        for c_idx in range(2, 14):
            cell = ws_ops.cell(r, c_idx)
            cell.number_format = fmt
            cell.alignment = ALIGN_RIGHT
            cell.font = FONT_BOLD if is_tot else FONT_REGULAR
            cell.fill = FILL_TOTAL if is_tot else (FILL_ZEBRA if r % 2 == 1 else openpyxl.styles.PatternFill(fill_type=None))
            cell.border = BORDER_TOTAL if is_tot else BORDER_DATA

    set_col_widths(ws_ops, {"A": 62, **{get_column_letter(i): 13 for i in range(2, 14)}})

    # =========================================================================
    # 2. INCOME STATEMENT (Consolidated Ind AS, FY19-FY25 Audited + FY26E-FY30E Connected)
    # =========================================================================
    setup_page(ws_is, "Edelweiss Financial Services Limited - Consolidated Income Statement")
    ws_is["A1"] = "EDELWEISS FINANCIAL SERVICES LTD — CONSOLIDATED STATEMENT OF PROFIT & LOSS (₹ CRORES)"
    ws_is["A1"].font = FONT_TITLE
    ws_is["A2"] = "Audited Consolidated FY19–FY25 (work/sources/EDELWEISS_consolidated.json) | FY26E–FY30E Linked to Segment Operations"
    ws_is["A2"].font = FONT_SMALL_ITALIC

    ws_is.cell(3, 1, "Line Item (₹ Cr)").font = FONT_TBL_HDR
    ws_is.cell(3, 1).fill = FILL_NAVY
    for idx, p in enumerate(periods, start=2):
        c = ws_is.cell(3, idx, p)
        c.font = FONT_TBL_HDR
        c.fill = FILL_NAVY
        c.alignment = ALIGN_RIGHT

    # Exact FY19-FY25 (Mar 2019 - Mar 2025) from work/sources/EDELWEISS_consolidated.json (NO Mar 2026 look-ahead!)
    sales_hist = [11078, 9513, 10775, 7151, 8500, 9523, 9393]
    exp_hist   = [4507, 7034, 6609, 3947, 5667, 6252, 6033]
    dep_hist   = [132, 232, 260, 151, 138, 126, 147]
    int_hist   = [4783, 4793, 3834, 2984, 2575, 2786, 2537]
    oth_hist   = [88, 89, 74, 158, 264, 79, 126]
    # Exact Tax Expense = PBT - Reported Net Profit ([1044, -2044, 254, 212, 406, 528, 536])
    tax_hist   = [700, -413, -108, 15, -22, -90, 266]

    is_rows = [
        (4, "Total Income Growth YoY (%)", FMT_PCT, False),
        (5, "Total Consolidated Revenue from Operations", FMT_INR_CR, True),
        (6, "Employee, Insurance Policyholder & Operating Expenses", FMT_INR_CR, False),
        (7, "Operating Profit Before Finance Costs (PBI)", FMT_INR_CR, True),
        (8, "Operating Margin (%)", FMT_PCT, False),
        (9, "Depreciation & Amortization", FMT_INR_CR, False),
        (10, "Finance Costs (Borrowing Interest Expense)", FMT_INR_CR, False),
        (11, "Other Income", FMT_INR_CR, False),
        (12, "Profit Before Tax (PBT)", FMT_INR_CR, True),
        (13, "Tax Expense / (Deferred Tax Credit)", FMT_INR_CR, False),
        (14, "Reported Consolidated Net Profit (PAT)", FMT_INR_CR, True),
        (15, "PAT Margin (%)", FMT_PCT, False),
        (16, "Diluted EPS (₹ / Share)", FMT_DEC2, False),
    ]

    for r, label, fmt, is_tot in is_rows:
        ws_is.cell(r, 1, label).font = FONT_BOLD if is_tot else FONT_REGULAR
        ws_is.cell(r, 1).fill = FILL_TOTAL if is_tot else (FILL_ZEBRA if r % 2 == 1 else openpyxl.styles.PatternFill(fill_type=None))
        ws_is.cell(r, 1).border = BORDER_TOTAL if is_tot else BORDER_DATA

    for i in range(7):
        col = i + 2
        col_let = get_column_letter(col)
        prev_let = get_column_letter(col - 1) if col > 2 else None
        ws_is.cell(5, col, sales_hist[i])
        ws_is.cell(6, col, exp_hist[i])
        ws_is.cell(9, col, dep_hist[i])
        ws_is.cell(10, col, int_hist[i])
        ws_is.cell(11, col, oth_hist[i])
        ws_is.cell(13, col, tax_hist[i])
        if prev_let:
            ws_is.cell(4, col, f"=({col_let}5-{prev_let}5)/{prev_let}5")
        else:
            ws_is.cell(4, col, 0.220)

    for col in range(9, 14):
        col_let = get_column_letter(col)
        prev_let = get_column_letter(col - 1)
        ws_is.cell(4, col, f"=({col_let}5-{prev_let}5)/{prev_let}5")
        ws_is.cell(5, col, f"='Segment Operations'!{col_let}16")
        ws_is.cell(6, col, f"=ROUND({col_let}5*'Segment Operations'!{col_let}18,0)")
        ws_is.cell(9, col, f"=ROUND('Balance Sheet'!{prev_let}10*0.115,0)")
        ws_is.cell(10, col, f"=ROUND('Balance Sheet'!{prev_let}6*0.135,0)")
        ws_is.cell(11, col, f"=ROUND({prev_let}11*1.04,0)")
        ws_is.cell(13, col, f"=ROUND({col_let}12*'Segment Operations'!{col_let}19,0)")

    for col in range(2, 14):
        col_let = get_column_letter(col)
        ws_is.cell(7, col, f"={col_let}5-{col_let}6")
        ws_is.cell(8, col, f"={col_let}7/{col_let}5")
        ws_is.cell(12, col, f"={col_let}7-{col_let}9-{col_let}10+{col_let}11")
        ws_is.cell(14, col, f"={col_let}12-{col_let}13")
        ws_is.cell(15, col, f"={col_let}14/{col_let}5")
        ws_is.cell(16, col, f"=ROUND({col_let}14/'Cover & Notes'!$B$12,2)")

    for r, _, fmt, is_tot in is_rows:
        for c in range(2, 14):
            cell = ws_is.cell(r, c)
            cell.number_format = fmt
            cell.alignment = ALIGN_RIGHT
            cell.font = FONT_BOLD if is_tot else (FONT_FORMULA if str(cell.value).startswith("=") else FONT_REGULAR)
            cell.fill = FILL_TOTAL if is_tot else (FILL_ZEBRA if r % 2 == 1 else openpyxl.styles.PatternFill(fill_type=None))
            cell.border = BORDER_TOTAL if is_tot else BORDER_DATA

    set_col_widths(ws_is, {"A": 48, **{get_column_letter(i): 14 for i in range(2, 14)}})

    # =========================================================================
    # 3. BALANCE SHEET (Exact FY19-FY25 from EDELWEISS_schedules.json, Zero Plugs)
    # =========================================================================
    setup_page(ws_bs, "Edelweiss Financial Services Limited - Consolidated Balance Sheet")
    ws_bs["A1"] = "EDELWEISS FINANCIAL SERVICES LTD — CONSOLIDATED BALANCE SHEET & SUB-SCHEDULES (₹ CRORES)"
    ws_bs["A1"].font = FONT_TITLE
    ws_bs["A2"] = "Audited FY19–FY25 (work/sources/EDELWEISS_schedules.json, Zero Plugs) | FY26E–FY30E Closed via Double-Entry Cash Flow"
    ws_bs["A2"].font = FONT_SMALL_ITALIC

    ws_bs.cell(3, 1, "Balance Sheet Component (₹ Cr)").font = FONT_TBL_HDR
    ws_bs.cell(3, 1).fill = FILL_NAVY
    for idx, p in enumerate(periods, start=2):
        c = ws_bs.cell(3, idx, p)
        c.font = FONT_TBL_HDR
        c.fill = FILL_NAVY
        c.alignment = ALIGN_RIGHT

    # Exact FY19-FY25 (Mar 2019 - Mar 2025) from work/sources/EDELWEISS_consolidated.json & EDELWEISS_schedules.json
    eq_hist      = [89, 89, 89, 90, 90, 90, 92]
    res_hist     = [7588, 6040, 6488, 6448, 6654, 4672, 4333]
    bor_hist     = [46148, 36657, 28436, 22768, 21791, 20358, 18004]
    fa_hist      = [1265, 2345, 1781, 1597, 1367, 1290, 1271]
    cwip_hist    = [44, 43, 13, 20, 25, 36, 18]
    inv_hist     = [12713, 10012, 12882, 14139, 17162, 18491, 17574]

    # Granular Sub-Schedule Items from work/sources/EDELWEISS_schedules.json (Mar 2019 - Mar 2025)
    invent_hist  = [169, 44, 0, 0, 0, 0, 0]
    rec_hist     = [2754, 1305, 506, 469, 413, 358, 394]
    cash_hist    = [6455, 8610, 4760, 3039, 3699, 3331, 5241]
    loans_hist   = [433, 584, 1070, 1281, 1515, 1230, 1237]
    oast_items   = [40068, 31073, 25122, 22519, 19700, 17978, 15547]

    pay_hist     = [1975, 1283, 489, 1290, 1412, 1674, 1260]
    nci_hist     = [1038, 1078, 1100, 1055, 1102, 1287, 1493]
    oliab_items  = [7063, 8869, 9532, 11413, 12832, 14633, 16100]

    bs_rows = [
        (4, "Equity Share Capital", FMT_INR_CR, False),
        (5, "Reserves & Surplus", FMT_INR_CR, False),
        (6, "Total Borrowings (Operating NBFC/HFC/ARC + HoldCo Debt)", FMT_INR_CR, False),
        (7, "Other Liabilities, NCI & Insurance Policyholder Funds (Rows 26–28)", FMT_INR_CR, False),
        (8, "Total Equity & Liabilities", FMT_INR_CR, True),
        (10, "Net Block (Property, Plant & Intangibles)", FMT_INR_CR, False),
        (11, "Capital Work in Progress (CWIP)", FMT_INR_CR, False),
        (12, "Financial Investments (AIF/ARC Security Receipts & Insurance)", FMT_INR_CR, False),
        (13, "Loans, Credit Book & Receivables (Excl. Cash, Rows 20,21,23,24)", FMT_INR_CR, False),
        (14, "Cash & Cash Equivalents (Schedule Row 22)", FMT_INR_CR, False),
        (15, "Total Assets", FMT_INR_CR, True),
        (16, "Consolidated Net Worth (Equity + Reserves)", FMT_INR_CR, True),
        (17, "Balance Check (Total Assets - Total Liabilities)", FMT_DEC2, True),
    ]

    for r, lbl, fmt, is_tot in bs_rows:
        ws_bs.cell(r, 1, lbl).font = FONT_BOLD if is_tot else FONT_REGULAR
        ws_bs.cell(r, 1).fill = FILL_TOTAL if is_tot else (FILL_ZEBRA if r % 2 == 1 else openpyxl.styles.PatternFill(fill_type=None))
        ws_bs.cell(r, 1).border = BORDER_TOTAL if is_tot else BORDER_DATA

    # Statutory Sub-Schedule Breakdown (Rows 19-29)
    ws_bs.cell(19, 1, "STATUTORY BALANCE SHEET SUB-SCHEDULES (CREDIT ASSETS, CASH, NCI & POLICYHOLDER FUNDS)").font = FONT_SEC_HDR
    ws_bs.cell(19, 1).fill = FILL_SEC_HDR
    for c_i in range(2, 14):
        ws_bs.cell(19, c_i).fill = FILL_SEC_HDR

    sub_rows = [
        (20, "Inventories / Held-for-Trading Securities", FMT_INR_CR, False),
        (21, "Trade & Fee Management Receivables", FMT_INR_CR, False),
        (22, "Cash & Cash Equivalents (Cash + Bank & Treasury Liquidity)", FMT_INR_CR, True),
        (23, "Short-Term Loans & Advances", FMT_INR_CR, False),
        (24, "Wholesale & Retail Credit Book & Financial Assets", FMT_INR_CR, False),
        (25, "Total Credit Book, Receivables & Cash (Rows 20–24 = Row 13 + Row 14)", FMT_INR_CR, True),
        (26, "Trade & Operating Payables", FMT_INR_CR, False),
        (27, "Non-Controlling Interest (CDPQ 20.2% in EARCL & Tokio Marine 25% in Life)", FMT_INR_CR, False),
        (28, "Insurance Policyholder Liabilities & Other Provisions", FMT_INR_CR, False),
        (29, "Total Other Liabilities, NCI & Policyholder Funds (Rows 26–28 = Row 7)", FMT_INR_CR, True),
    ]
    for r, lbl, fmt, is_tot in sub_rows:
        ws_bs.cell(r, 1, lbl).font = FONT_BOLD if is_tot else FONT_REGULAR
        ws_bs.cell(r, 1).fill = FILL_TOTAL if is_tot else (FILL_ZEBRA if r % 2 == 1 else openpyxl.styles.PatternFill(fill_type=None))
        ws_bs.cell(r, 1).border = BORDER_TOTAL if is_tot else BORDER_DATA

    for i in range(7):
        col = i + 2
        col_let = get_column_letter(col)
        ws_bs.cell(20, col, invent_hist[i])
        ws_bs.cell(21, col, rec_hist[i])
        ws_bs.cell(22, col, cash_hist[i])
        ws_bs.cell(23, col, loans_hist[i])
        ws_bs.cell(24, col, oast_items[i])
        ws_bs.cell(25, col, f"=SUM({col_let}20:{col_let}24)")

        ws_bs.cell(26, col, pay_hist[i])
        ws_bs.cell(27, col, nci_hist[i])
        ws_bs.cell(28, col, oliab_items[i])
        ws_bs.cell(29, col, f"=SUM({col_let}26:{col_let}28)")

        ws_bs.cell(4, col, eq_hist[i])
        ws_bs.cell(5, col, res_hist[i])
        ws_bs.cell(6, col, bor_hist[i])
        ws_bs.cell(7, col, f"={col_let}29")
        ws_bs.cell(8, col, f"=SUM({col_let}4:{col_let}7)")

        ws_bs.cell(10, col, fa_hist[i])
        ws_bs.cell(11, col, cwip_hist[i])
        ws_bs.cell(12, col, inv_hist[i])
        ws_bs.cell(13, col, f"={col_let}20+{col_let}21+{col_let}23+{col_let}24")
        ws_bs.cell(14, col, f"={col_let}22")
        ws_bs.cell(15, col, f"=SUM({col_let}10:{col_let}14)")
        ws_bs.cell(16, col, f"={col_let}4+{col_let}5")
        ws_bs.cell(17, col, f"={col_let}15-{col_let}8")

    # Forecast Columns I..M (idx 9..13, FY26E..FY30E) — True Double-Entry Roll-Forward (Zero Plugs!)
    for col in range(9, 14):
        col_let = get_column_letter(col)
        prev_let = get_column_letter(col - 1)
        ws_bs.cell(20, col, f"={prev_let}20")
        ws_bs.cell(21, col, f"=ROUND('Income Statement'!{col_let}5*'Segment Operations'!{col_let}23/365,0)")
        ws_bs.cell(22, col, f"={prev_let}22+'Cash Flow'!{col_let}7")
        ws_bs.cell(23, col, f"=ROUND({prev_let}23*1.03,0)")
        ws_bs.cell(24, col, f"=ROUND({prev_let}24*0.97,0)")
        ws_bs.cell(25, col, f"=SUM({col_let}20:{col_let}24)")

        ws_bs.cell(26, col, f"=ROUND({prev_let}26*1.05,0)")
        ws_bs.cell(27, col, f"=ROUND({prev_let}27*1.05,0)")
        ws_bs.cell(28, col, f"=ROUND({prev_let}28*1.08,0)")
        ws_bs.cell(29, col, f"=SUM({col_let}26:{col_let}28)")

        ws_bs.cell(4, col, f"={prev_let}4")
        ws_bs.cell(5, col, f"={prev_let}5+'Income Statement'!{col_let}14-ROUND('Income Statement'!{col_let}14*'Segment Operations'!{col_let}20,0)")
        ws_bs.cell(6, col, f"=ROUND({prev_let}6*0.95,0)")
        ws_bs.cell(7, col, f"={col_let}29")
        ws_bs.cell(8, col, f"=SUM({col_let}4:{col_let}7)")

        ws_bs.cell(10, col, f"={prev_let}10+ROUND('Income Statement'!{col_let}5*0.015,0)-'Income Statement'!{col_let}9")
        ws_bs.cell(11, col, f"={prev_let}11")
        ws_bs.cell(12, col, f"=ROUND({prev_let}12*1.05,0)")
        ws_bs.cell(13, col, f"={col_let}20+{col_let}21+{col_let}23+{col_let}24")
        ws_bs.cell(14, col, f"={col_let}22")
        ws_bs.cell(15, col, f"=SUM({col_let}10:{col_let}14)")
        ws_bs.cell(16, col, f"={col_let}4+{col_let}5")
        ws_bs.cell(17, col, f"={col_let}15-{col_let}8")

    for r, _, fmt, is_tot in bs_rows + sub_rows:
        for c in range(2, 14):
            cell = ws_bs.cell(r, c)
            cell.number_format = fmt
            cell.alignment = ALIGN_RIGHT
            cell.font = FONT_BOLD if is_tot else (FONT_FORMULA if str(cell.value).startswith("=") else FONT_REGULAR)
            cell.fill = FILL_TOTAL if is_tot else (FILL_ZEBRA if r % 2 == 1 else openpyxl.styles.PatternFill(fill_type=None))
            cell.border = BORDER_TOTAL if is_tot else BORDER_DATA

    set_col_widths(ws_bs, {"A": 58, **{get_column_letter(i): 14 for i in range(2, 14)}})

    # =========================================================================
    # 4. CASH FLOW STATEMENT (Exact FY19-FY25 + Connected FY26E-FY30E)
    # =========================================================================
    setup_page(ws_cf, "Edelweiss Financial Services Limited - Consolidated Cash Flow")
    ws_cf["A1"] = "EDELWEISS FINANCIAL SERVICES LTD — CONSOLIDATED CASH FLOW STATEMENT (₹ CRORES)"
    ws_cf["A1"].font = FONT_TITLE
    ws_cf["A2"] = "Note: Positive CFO reflects wholesale loan book repayments funding debt reduction (CFF); Industrial FCFF is not applicable."
    ws_cf["A2"].font = FONT_SMALL_ITALIC

    ws_cf.cell(3, 1, "Cash Flow Component (₹ Cr)").font = FONT_TBL_HDR
    ws_cf.cell(3, 1).fill = FILL_NAVY
    for idx, p in enumerate(periods, start=2):
        c = ws_cf.cell(3, idx, p)
        c.font = FONT_TBL_HDR
        c.fill = FILL_NAVY
        c.alignment = ALIGN_RIGHT

    cfo_hist = [5685, 12098, 3458, 5592, 2891, 2894, 3384]
    cfi_hist = [-1312, 199, 4085, -1058, -425, -482, 3794]
    cff_hist = [-3665, -10471, -8587, -6444, -1708, -2785, -4828]

    cf_rows = [
        (4, "Cash from Operating Activities (incl. Loan Book Repayments & Policyholder Inflows)", FMT_INR_CR, True),
        (5, "Cash from Investing Activities (CFI - CapEx & Financial Investments)", FMT_INR_CR, False),
        (6, "Cash from Financing Activities (CFF - Debt Repayment & Dividends)", FMT_INR_CR, False),
        (7, "Net Cash Flow", FMT_INR_CR, True),
    ]
    for r, lbl, fmt, is_tot in cf_rows:
        ws_cf.cell(r, 1, lbl).font = FONT_BOLD if is_tot else FONT_REGULAR
        ws_cf.cell(r, 1).fill = FILL_TOTAL if is_tot else FILL_ZEBRA
        ws_cf.cell(r, 1).border = BORDER_TOTAL if is_tot else BORDER_DATA

    for i in range(7):
        col = i + 2
        col_let = get_column_letter(col)
        ws_cf.cell(4, col, cfo_hist[i])
        ws_cf.cell(5, col, cfi_hist[i])
        ws_cf.cell(6, col, cff_hist[i])
        ws_cf.cell(7, col, f"=SUM({col_let}4:{col_let}6)")

    for col in range(9, 14):
        col_let = get_column_letter(col)
        prev_let = get_column_letter(col - 1)
        # Double-entry CFO = PAT + D&A - (ΔCredit/Operating Assets Excl. Cash - ΔPolicyholder/Other Liabilities)
        ws_cf.cell(4, col, f"='Income Statement'!{col_let}14+'Income Statement'!{col_let}9-(('Balance Sheet'!{col_let}13-'Balance Sheet'!{prev_let}13)-('Balance Sheet'!{col_let}7-'Balance Sheet'!{prev_let}7))")
        # Double-entry CFI = -CapEx - ΔCWIP - ΔFinancial Investments
        ws_cf.cell(5, col, f"=-ROUND('Income Statement'!{col_let}5*0.015,0)-('Balance Sheet'!{col_let}11-'Balance Sheet'!{prev_let}11)-('Balance Sheet'!{col_let}12-'Balance Sheet'!{prev_let}12)")
        # Double-entry CFF = ΔBorrowings + ΔEquity - Dividends
        ws_cf.cell(6, col, f"=('Balance Sheet'!{col_let}6-'Balance Sheet'!{prev_let}6)+('Balance Sheet'!{col_let}4-'Balance Sheet'!{prev_let}4)-ROUND('Income Statement'!{col_let}14*'Segment Operations'!{col_let}20,0)")
        ws_cf.cell(7, col, f"=SUM({col_let}4:{col_let}6)")

    for r, _, fmt, is_tot in cf_rows:
        for c in range(2, 14):
            cell = ws_cf.cell(r, c)
            cell.number_format = fmt
            cell.alignment = ALIGN_RIGHT
            cell.font = FONT_BOLD if is_tot else FONT_REGULAR
            cell.fill = FILL_TOTAL if is_tot else FILL_ZEBRA
            cell.border = BORDER_TOTAL if is_tot else BORDER_DATA

    set_col_widths(ws_cf, {"A": 58, **{get_column_letter(i): 14 for i in range(2, 14)}})

    # =========================================================================
    # 5. SOTP VALUATION, AUDITED STANDALONE HOLDCO DEBT BRIDGE & SENSITIVITY
    # =========================================================================
    setup_page(ws_val, "Edelweiss Financial Services Limited - SOTP Valuation & Comps")
    ws_val["A1"] = "EDELWEISS FINANCIAL SERVICES LTD — FINANCIAL SERVICES SUM-OF-THE-PARTS (SOTP) VALUATION"
    ws_val["A1"].font = FONT_TITLE
    ws_val["A2"] = "Fair Value of Operating Verticals (Explicit Ownership Stakes & IEV) Less Audited FY25 Standalone HoldCo Net Debt (₹ Crores)"
    ws_val["A2"].font = FONT_SMALL_ITALIC

    ws_val["A3"] = "METHODOLOGICAL EXCLUSION NOTICE: Industrial FCFF DCF is strictly prohibited for Edelweiss because borrowings (₹18,004 Cr in FY25) are operating raw material across NBFC, Housing, ARC & Insurance. Note: IEV in Row 9 denotes Actuarial Indian Embedded Value (NOT Enterprise Value)."
    ws_val["A3"].font = FONT_BOLD
    ws_val["A3"].fill = FILL_ACCENT

    sotp_headers = [
        "Financial Services Vertical & EFSL Stake",
        "EFSL Attributable Base Metric (₹ Cr)",
        "Target Multiple (P/E, P/B, % AAUM, P/IEV)",
        "Conglomerate Discount (%)",
        "Attributed Equity Value (₹ Cr)",
        "Per Share Value (₹)",
        "Ownership Stake, NCI & Peer Benchmark Evidence",
    ]
    for idx, h in enumerate(sotp_headers, start=1):
        c = ws_val.cell(5, idx, h)
        c.font = FONT_TBL_HDR
        c.fill = FILL_NAVY
        c.alignment = ALIGN_CENTER if idx > 1 else ALIGN_LEFT

    # Note: B6:B10 link directly to 'Segment Operations'!H6, H8, H9, H10, H12!
    # E12 links directly to Audited FY25 EFSL Standalone HoldCo Net Debt Bridge (=-B39 = -₹3,067 Cr)!
    # Applies 15.0% Conglomerate/HoldCo Discount (D6:D10 = 0.15) across operating verticals:
    # Gross Vertical Value after 15% discount = 18,120 * 0.85 = 15,402 Cr less 3,067 Cr Standalone HoldCo Net Debt = ₹12,335 Cr (₹130.53/sh)
    sotp_rows = [
        (6, "Alternative Asset Management (EAAA — 100% EFSL Stake)", "='Segment Operations'!H6", 22.0, 0.15, "=B6*C6*(1-D6)", "=E6/$E$14", "22.0x FY25 PAT (₹230 Cr) × 100% Stake × (1 - 15% HoldCo Disc.)"),
        (7, "Asset Reconstruction (EARCL — 79.8% EFSL Stake; CDPQ 20.2% NCI)", "='Segment Operations'!H8", 1.20, 0.15, "=B7*C7*(1-D7)", "=E7/$E$14", "1.20x EFSL Attributable Net Worth (₹2,800 Cr net of CDPQ NCI)"),
        (8, "Mutual Fund (EAML — 85% Post-WestBridge 15% Deal Benchmark)", "='Segment Operations'!H9", 0.040, 0.15, "=B8*C8*(1-D8)", "=E8/$E$14", "4.0% of Total AAUM (₹1,41,000 Cr; WestBridge 15% @ ₹450 Cr)"),
        (9, "Insurance (75% Life [Tokio Marine 25% NCI] + 100% General IEV)", "='Segment Operations'!H10", 1.00, 0.15, "=B9*C9*(1-D9)", "=E9/$E$14", "1.00x Attributable Actuarial Indian Embedded Value (IEV = ₹1,500 Cr)"),
        (10, "Credit & Housing Finance (ECLF & Nido — 100% EFSL Stake)", "='Segment Operations'!H12", 0.80, 0.15, "=B10*C10*(1-D10)", "=E10/$E$14", "0.80x Allocated Book Value (₹3,200 Cr) × (1 - 15% HoldCo Disc.)"),
        (11, "Total Net Business Value (Post-15% Conglomerate Discount)", "", "", "", "=SUM(E6:E10)", "=E11/$E$14", "Sum of Discounted Operating Verticals (₹15,402 Cr; Gross = ₹18,120 Cr)"),
        (12, "Less: Audited FY25 EFSL Standalone Corporate HoldCo Net Debt", "", "", "", "=-B39", "=E12/$E$14", "Audited FY25 Standalone Borrowings (₹3,072 Cr) less Cash (₹5 Cr) [Row 39]"),
        (13, "Implied Consolidated Equity Fair Value (₹ Cr)", "", "", "", "=E11+E12", "=E13/$E$14", "Net Attributable SOTP Equity Value (₹12,335 Cr)"),
        (14, "Diluted Shares Outstanding (Cr)", "", "", "", 94.50, "", "Face Value ₹1.00 | BSE/NSE Shareholding Filing"),
        (15, "Implied SOTP Fair Value Per Share (₹)", "", "", "", "=E13/E14", "=E15", "Compare vs Official NSE Close ₹108.85 (15-Oct-2025)"),
    ]

    for r, comp, base_m, mult_m, disc, val_f, ps_f, meth in sotp_rows:
        is_tot = r in (11, 13, 15)
        ws_val.cell(r, 1, comp).font = FONT_BOLD if is_tot else FONT_REGULAR
        ws_val.cell(r, 2, base_m).number_format = FMT_INR_CR
        ws_val.cell(r, 3, mult_m).number_format = FMT_PCT if isinstance(mult_m, float) and mult_m < 0.1 else FMT_MULT
        ws_val.cell(r, 4, disc).number_format = FMT_PCT
        ws_val.cell(r, 5, val_f).number_format = FMT_DEC2 if r in (14, 15) else FMT_INR_CR
        ws_val.cell(r, 6, ps_f).number_format = FMT_DEC2
        ws_val.cell(r, 7, meth).font = FONT_SMALL_ITALIC
        for c_i in range(1, 8):
            cell = ws_val.cell(r, c_i)
            if c_i in (2, 3, 4, 5, 6):
                cell.alignment = ALIGN_RIGHT
            if is_tot:
                cell.font = FONT_BOLD
                cell.fill = FILL_TOTAL if r != 15 else FILL_ACCENT
                cell.border = BORDER_TOTAL
            else:
                cell.border = BORDER_DATA

    # Independent 2-Way Sensitivity Matrix (Rows 19-26) — Centered on C6 (22.0x) and D6 (15.0% Conglomerate Discount)
    ws_val["A19"] = "SENSITIVITY ANALYSIS: IMPLIED PER-SHARE VALUE (₹) — EAAA P/E MULTIPLE VS. CONGLOMERATE DISCOUNT"
    ws_val["A19"].font = FONT_SEC_HDR
    ws_val["A19"].fill = FILL_SEC_HDR

    eaaa_mults = [18.0, 22.0, 24.0, 26.0, 28.0]
    disc_steps = [0.25, 0.20, 0.15, 0.10, 0.00]
    ws_val.cell(20, 1, "HoldCo Disc \\ EAAA P/E").font = FONT_TBL_HDR
    ws_val.cell(20, 1).fill = FILL_NAVY
    for j, m in enumerate(eaaa_mults, start=2):
        c = ws_val.cell(20, j, m)
        c.font = FONT_TBL_HDR
        c.fill = FILL_NAVY
        c.number_format = FMT_MULT
        c.alignment = ALIGN_RIGHT

    for i, d in enumerate(disc_steps, start=21):
        hdr = ws_val.cell(i, 1, d)
        hdr.font = FONT_BOLD
        hdr.number_format = FMT_PCT
        hdr.fill = FILL_SUB_HDR
        hdr.border = BORDER_DATA
        for j, m in enumerate(eaaa_mults, start=2):
            col_let = get_column_letter(j)
            cell = ws_val.cell(i, j)
            # Independent SOTP equation across ALL 25 cells (including Base Case C23!)
            cell.value = f"=($B$6*{col_let}$20*(1-$A{i}) + $B$7*$C$7*(1-$A{i}) + $B$8*$C$8*(1-$A{i}) + $B$9*$C$9*(1-$A{i}) + $B$10*$C$10*(1-$A{i}) + $E$12)/$E$14"
            if i == 23 and j == 3:
                cell.fill = FILL_ACCENT
                cell.font = FONT_BOLD
            else:
                cell.font = FONT_FORMULA
            cell.number_format = FMT_DEC2
            cell.alignment = ALIGN_RIGHT
            cell.border = BORDER_DATA

    ws_val.cell(26, 1, "Sensitivity Base Case Reconciliation Check (Cell C23 - Main Output E15)").font = FONT_BOLD
    ws_val.cell(26, 2, "=C23-E15").number_format = FMT_DEC2
    ws_val.cell(26, 2).font = FONT_BOLD
    ws_val.cell(26, 3, "=IF(ABS(B26)<0.01,\"RECONCILED (0.00)\",\"MISMATCH\")").font = FONT_BOLD

    # Operating Scenario Analysis (Bear / Base / Bull) — Rows 28-32
    ws_val["A28"] = "OPERATING SCENARIO ANALYSIS (BEAR / BASE / BULL CASE)"
    ws_val["A28"].font = FONT_SEC_HDR
    ws_val["A28"].fill = FILL_SEC_HDR
    scen_headers = ["Scenario Case", "EAAA PAT (₹ Cr)", "EAAA P/E Multiple (x)", "Conglomerate Discount (%)", "HoldCo Net Debt (₹ Cr)", "Implied Fair Value / Share (₹)", "Implied vs CMP (%)"]
    for idx, h in enumerate(scen_headers, start=1):
        c = ws_val.cell(29, idx, h)
        c.font = FONT_TBL_HDR
        c.fill = FILL_NAVY
        c.alignment = ALIGN_RIGHT if idx > 1 else ALIGN_LEFT

    scenarios = [
        (30, "Bear Case (25% Conglomerate Discount & Intragroup Guarantee Buffer)", "=ROUND(B6*0.90,0)", 18.0, 0.25, -4500),
        (31, "Base Case (Audited FY25 Vertical Baseline & Standalone HoldCo Debt)", "=B6", "=C6", "=D6", "=E12"),
        (32, "Bull Case (EAAA Stake Unbundling, 5% HoldCo Disc & Debt Paydown)", "=ROUND(B6*1.15,0)", 25.0, 0.05, -2200),
    ]
    for r, s_name, e_pat, e_mult, disc_v, hc_debt in scenarios:
        is_base = r == 31
        ws_val.cell(r, 1, s_name).font = FONT_BOLD if is_base else FONT_REGULAR
        ws_val.cell(r, 2, e_pat).number_format = FMT_INR_CR
        ws_val.cell(r, 3, e_mult).number_format = FMT_MULT
        ws_val.cell(r, 4, disc_v).number_format = FMT_PCT
        ws_val.cell(r, 5, hc_debt).number_format = FMT_INR_CR
        ws_val.cell(r, 6, f"=((B{r}*C{r}+$B$7*$C$7+$B$8*$C$8+$B$9*$C$9+$B$10*$C$10)*(1-D{r})+E{r})/$E$14").number_format = FMT_DEC2
        ws_val.cell(r, 7, f"=(F{r}-'Cover & Notes'!$B$10)/'Cover & Notes'!$B$10").number_format = FMT_PCT
        for c_i in range(1, 8):
            cell = ws_val.cell(r, c_i)
            if c_i > 1:
                cell.alignment = ALIGN_RIGHT
            cell.fill = FILL_ACCENT if is_base else FILL_ZEBRA
            cell.border = BORDER_TOTAL if is_base else BORDER_DATA

    # Audited FY25 EFSL Standalone HoldCo Net Debt vs Consolidated Borrowings Bridge — Rows 34-40
    ws_val["A34"] = "AUDITED FY25 CONSOLIDATED BORROWINGS TO EFSL STANDALONE HOLDCO NET DEBT BRIDGE"
    ws_val["A34"].font = FONT_SEC_HDR
    ws_val["A34"].fill = FILL_SEC_HDR
    debt_rows = [
        (35, "Total Consolidated Borrowings on Balance Sheet ('Balance Sheet'!H6)", "='Balance Sheet'!H6", "Audited FY25 Consolidated Borrowings across NBFC, HFC, EARCL & HoldCo (₹18,004 Cr)"),
        (36, "Less: Ring-Fenced Operating Subsidiary Borrowings (ECLF, Nido, EARCL, EAAA)", -14932, "Operating credit & ARC liabilities already deducted in vertical Net Worth (P/B)"),
        (37, "Audited FY25 EFSL Standalone (Parent HoldCo) Gross Borrowings", "=B35+B36", "Matches work/sources/EDELWEISS_standalone.json Mar 2025 Borrowings (₹3,072 Cr)"),
        (38, "Less: Audited FY25 EFSL Standalone Unencumbered Cash & Bank Equivalents", -5, "Matches work/sources/EDELWEISS_schedules.json Standalone Mar 2025 Cash (₹5 Cr)"),
        (39, "Audited FY25 EFSL Standalone Corporate HoldCo Net Debt (Feeds -E12)", "=B37+B38", "Exact Audited Parent HoldCo Net Debt Deducted in SOTP Row 12 (₹3,067 Cr)"),
        (40, "Memo: Consolidated Non-Controlling Interest ('Balance Sheet'!H27)", "='Balance Sheet'!H27", "CDPQ (20.2% EARCL) & Tokio Marine (25% Life) excluded via Attributable Vertical Basis"),
    ]
    for r, d_lbl, d_val, d_note in debt_rows:
        is_tot = r in (37, 39)
        ws_val.cell(r, 1, d_lbl).font = FONT_BOLD if is_tot else FONT_REGULAR
        ws_val.cell(r, 1).border = BORDER_TOTAL if is_tot else BORDER_DATA
        c = ws_val.cell(r, 2, d_val)
        c.font = FONT_BOLD if is_tot else FONT_REGULAR
        c.number_format = FMT_INR_CR
        c.alignment = ALIGN_RIGHT
        c.fill = FILL_ACCENT if r == 39 else (FILL_TOTAL if is_tot else FILL_ZEBRA)
        c.border = BORDER_TOTAL if is_tot else BORDER_DATA
        n_c = ws_val.cell(r, 3, d_note)
        n_c.font = FONT_SMALL_ITALIC
        n_c.border = BORDER_DATA

    # Indian Financial Services Peer Multiples Table — Rows 42-47
    ws_val["A42"] = "INDIAN WEALTH, AMC & FINANCIAL SERVICES PEER MULTIPLES (15-OCT-2025 OFFICIAL NSE SNAPSHOT)"
    ws_val["A42"].font = FONT_SEC_HDR
    ws_val["A42"].fill = FILL_SEC_HDR
    peer_headers = ["Comparable Financial Services Peer", "Ticker (NSE)", "Market Cap (₹ Cr)", "FY25 PAT / Book (₹ Cr)", "Valuation Multiple", "ROE (%)", "Benchmark Relevance"]
    for idx, h in enumerate(peer_headers, start=1):
        c = ws_val.cell(43, idx, h)
        c.font = FONT_TBL_HDR
        c.fill = FILL_NAVY
        c.alignment = ALIGN_RIGHT if idx in (3, 4, 5, 6) else ALIGN_LEFT

    peers = [
        (44, "360 ONE WAM Ltd (Wealth & Alternatives)", "360ONE", 38500, 1015, "=C44/D44", 0.245, "37.9x P/E vs 22.0x Applied to EAAA (18.7x post-15% HoldCo Disc.)"),
        (45, "HDFC Asset Management Co Ltd", "HDFCAMC", 94200, 2460, "=C45/D45", 0.324, "12.5% of AAUM vs 4.0% Applied to Edelweiss MF"),
        (46, "UTI Asset Management Co Ltd", "UTIAMC", 15800, 765, "=C46/D46", 0.185, "5.2% of AAUM Benchmark for Mid-Sized AMC"),
        (47, "PNB Housing Finance Ltd (P/B Basis)", "PNBHOUSING", 24600, 16850, "=C47/D47", 0.118, "1.46x P/B vs 0.80x P/B Applied to Edelweiss Credit"),
    ]
    for r, p_name, tkr, mc_v, pat_v, m_f, roe_v, src_txt in peers:
        ws_val.cell(r, 1, p_name).font = FONT_REGULAR
        ws_val.cell(r, 2, tkr).font = FONT_REGULAR
        ws_val.cell(r, 3, mc_v).number_format = FMT_INR_CR
        ws_val.cell(r, 4, pat_v).number_format = FMT_INR_CR
        ws_val.cell(r, 5, m_f).number_format = FMT_MULT
        ws_val.cell(r, 6, roe_v).number_format = FMT_PCT
        ws_val.cell(r, 7, src_txt).font = FONT_SMALL_ITALIC
        for c_i in range(1, 8):
            cell = ws_val.cell(r, c_i)
            if c_i in (3, 4, 5, 6):
                cell.alignment = ALIGN_RIGHT
            cell.fill = FILL_ZEBRA
            cell.border = BORDER_DATA

    set_col_widths(ws_val, {"A": 52, "B": 24, "C": 26, "D": 22, "E": 24, "F": 20, "G": 50})

    # =========================================================================
    # 6. COVER & NOTES (Linked to Statements and SOTP Valuation)
    # =========================================================================
    setup_page(ws_exec, "Edelweiss Financial Services Limited - Cover & Notes")
    ws_exec["A1"] = "EDELWEISS FINANCIAL SERVICES LIMITED (BSE: 532922 | NSE: EDELWEISS)"
    ws_exec["A1"].font = FONT_TITLE
    ws_exec["A2"] = "Institutional Equity Research Model — Diversified Financial Services SOTP Valuation (₹ Crores)"
    ws_exec["A2"].font = FONT_SUBTITLE

    ws_exec["A4"] = "1. ENTITY METADATA & FINANCIAL SERVICES METHODOLOGY BOUNDARY"
    style_range(ws_exec, 4, 1, 4, 6, font=FONT_SEC_HDR, fill=FILL_SEC_HDR)

    meta_rows = [
        ("Legal Entity Name", "Edelweiss Financial Services Limited", "Primary Valuation Methodology", "Financial Services Sum-of-the-Parts (SOTP)"),
        ("Ticker / Scrip Code", "NSE: EDELWEISS | BSE: 532922", "Industrial FCFF DCF Status", "STRICTLY WITHHELD (Financial Institution)"),
        ("Sector / Industry", "Diversified Financial Services (AIF, ARC, AMC, Ins., Credit)", "Latest Audited Period", "FY25 (Year Ended March 31, 2025)"),
        ("Face Value per Share (₹)", 1.00, "Implied SOTP Value / Share (₹)", "='SOTP Valuation'!E15"),
        ("Current Market Price (₹) [15-Oct-2025 Official NSE Close]", 108.85, "Implied Upside / (Downside) (%)", "=(E8-B10)/B10"),
        ("Market Capitalization (₹ Cr)", "=B10*B12", "Consolidated FY25 Total Income (₹ Cr)", "='Income Statement'!H5"),
        ("Diluted Shares Outstanding (Cr)", 94.50, "Consolidated FY25 PAT (₹ Cr)", "='Income Statement'!H14"),
    ]
    for offset, (k1, v1, k2, v2) in enumerate(meta_rows, start=5):
        r = offset + 1 if offset >= 9 else offset
        ws_exec.cell(r, 1, k1).font = FONT_BOLD
        c_v1 = ws_exec.cell(r, 2, v1)
        ws_exec.cell(r, 4, k2).font = FONT_BOLD
        c_v2 = ws_exec.cell(r, 5, v2)
        c_v1.font = FONT_FORMULA if str(v1).startswith("=") else FONT_REGULAR
        c_v2.font = FONT_LINK if str(v2).startswith("='") else (FONT_FORMULA if str(v2).startswith("=") else FONT_REGULAR)
        style_range(ws_exec, r, 1, r, 6, border=BORDER_DATA)

    ws_exec["B8"].number_format = FMT_DEC2
    ws_exec["B10"].number_format = FMT_DEC2
    ws_exec["B11"].number_format = FMT_INR_CR
    ws_exec["B12"].number_format = FMT_DEC2
    ws_exec["E8"].number_format = FMT_DEC2
    ws_exec["E8"].font = FONT_BOLD
    ws_exec["E8"].fill = FILL_ACCENT
    ws_exec["E10"].number_format = FMT_PCT
    ws_exec["E10"].font = FONT_BOLD
    ws_exec["E11"].number_format = FMT_INR_CR
    ws_exec["E12"].number_format = FMT_INR_CR

    ws_exec["A14"] = "2. WHY INDUSTRIAL FCFF IS STRICTLY WITHHELD FOR EDELWEISS"
    style_range(ws_exec, 14, 1, 14, 6, font=FONT_SEC_HDR, fill=FILL_SEC_HDR)
    notes = [
        "• Financial Institution Accounting Structure: For a multi-vertical financial services group (Alternatives EAAA, ARC EARCL, Mutual Fund, Life/General Insurance, NBFC/HFC), debt is operating raw material rather than corporate capital structure.",
        "• Cash Flow Distortion in Industrial FCFF: As Edelweiss runs down its wholesale loan book (borrowings reduced from ₹46,148 Cr in FY19 to ₹18,004 Cr in FY25), principal repayments appear as positive operating cash inflows (+₹3,384 Cr CFO in FY25). An industrial FCFF model would misinterpret loan book liquidation as perpetual free cash flow.",
        "• SOTP Valuation & Audited HoldCo Debt Discipline: Each vertical is valued on its EFSL-attributable metric (net of CDPQ 20.2% NCI in EARCL and Tokio Marine 25% NCI in Life Insurance, with IEV explicitly defined as Actuarial Indian Embedded Value) less a 15% Conglomerate Discount and Audited FY25 EFSL Standalone HoldCo Net Debt (-₹3,067 Cr = ₹3,072 Cr Standalone Borrowings less ₹5 Cr Standalone Cash).",
    ]
    for idx, note in enumerate(notes, start=15):
        ws_exec.cell(idx, 1, note).font = FONT_REGULAR

    ws_exec["A19"] = "3. MULTI-YEAR FINANCIAL PERFORMANCE SUMMARY (FY21–FY30E)"
    style_range(ws_exec, 19, 1, 19, 12, font=FONT_SEC_HDR, fill=FILL_SEC_HDR)
    sum_cols = ["FY21", "FY22", "FY23", "FY24", "FY25", "FY26E", "FY27E", "FY28E", "FY29E", "FY30E"]
    ws_exec.cell(20, 1, "Key Financial Metric (₹ Cr)").font = FONT_TBL_HDR
    ws_exec.cell(20, 1).fill = FILL_NAVY
    for i, p in enumerate(sum_cols, start=2):
        c = ws_exec.cell(20, i, p)
        c.font = FONT_TBL_HDR
        c.fill = FILL_NAVY
        c.alignment = ALIGN_RIGHT

    summary_links = [
        ("Total Consolidated Operating Income", "Income Statement", 5, FMT_INR_CR),
        ("Operating Profit Before Interest (PBI)", "Income Statement", 7, FMT_INR_CR),
        ("Finance Costs (Borrowing Interest)", "Income Statement", 10, FMT_INR_CR),
        ("Reported Consolidated PAT", "Income Statement", 14, FMT_INR_CR),
        ("Diluted EPS (₹)", "Income Statement", 16, FMT_DEC2),
        ("Consolidated Net Worth (Equity + Reserves)", "Balance Sheet", 16, FMT_INR_CR),
    ]
    for r_idx, (label, sheet_ref, src_row, fmt) in enumerate(summary_links, start=21):
        ws_exec.cell(r_idx, 1, label).font = FONT_BOLD if "PAT" in label or "Income" in label else FONT_REGULAR
        ws_exec.cell(r_idx, 1).border = BORDER_DATA
        for c_idx in range(2, 12):
            src_col_let = get_column_letter(c_idx + 2)
            cell = ws_exec.cell(r_idx, c_idx, f"='{sheet_ref}'!{src_col_let}{src_row}")
            cell.font = FONT_LINK
            cell.number_format = fmt
            cell.alignment = ALIGN_RIGHT
            cell.border = BORDER_DATA
            if r_idx % 2 == 1:
                cell.fill = FILL_ZEBRA

    set_col_widths(ws_exec, {"A": 44, "B": 28, "C": 14, "D": 32, "E": 28, "F": 14, "G": 14, "H": 14, "I": 14, "J": 14, "K": 14, "L": 14})

    for ws in wb.worksheets:
        apply_cell_color_coding(ws)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path
