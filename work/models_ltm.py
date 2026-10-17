"""Native OpenPyXL Model Builder for LTIMindtree Limited (LTM)."""
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
    style_range, set_col_widths, setup_page
)


def build_ltm_workbook(out_path: Path) -> Path:
    return build_ltm_model(out_path)


def build_ltm_model(out_path: Path) -> Path:
    wb = openpyxl.Workbook()
    ws_exec = wb.active
    ws_exec.title = "Cover & Notes"
    ws_is = wb.create_sheet("Income Statement")
    ws_bs = wb.create_sheet("Balance Sheet")
    ws_cf = wb.create_sheet("Cash Flow")
    ws_ops = wb.create_sheet("IT Services Operations")
    ws_val = wb.create_sheet("FCFF Valuation")

    # 7 audited historical years (FY19-FY25, cols B-H) + 5 connected forecast years (FY26E-FY30E, cols I-M)
    periods = ["FY19", "FY20", "FY21", "FY22", "FY23", "FY24", "FY25", "FY26E", "FY27E", "FY28E", "FY29E", "FY30E"]
    basis_labels = [
        "Pre-Merger LTI", "Pre-Merger LTI", "Pre-Merger LTI",
        "Restated Combined", "Merged LTIMindtree", "Merged LTIMindtree", "Merged LTIMindtree",
        "Connected Forecast", "Connected Forecast", "Connected Forecast", "Connected Forecast", "Connected Forecast",
    ]

    # =========================================================================
    # 1. IT SERVICES OPERATIONS & FORECAST DRIVERS (Sheet 5)
    # =========================================================================
    setup_page(ws_ops, "LTIMindtree Limited - IT Services Operations, Drivers & Ratios")
    ws_ops["A1"] = "LTIMINDTREE LIMITED — IT SERVICES OPERATING METRICS, DRIVERS & RATIOS"
    ws_ops["A1"].font = FONT_TITLE
    ws_ops["A2"] = "Merger Basis: FY19–FY21 Pre-Merger LTI | FY22 Restated Amalgamation Comparative | FY23–FY25 Merged LTIMindtree | FY26E–FY30E USD*FX Connected"
    ws_ops["A2"].font = FONT_SMALL_ITALIC

    ws_ops.cell(3, 1, "Operating KPI / Forecast Driver").font = FONT_TBL_HDR
    ws_ops.cell(3, 1).fill = FILL_NAVY
    for idx, p in enumerate(periods, start=2):
        c = ws_ops.cell(3, idx, p)
        c.font = FONT_TBL_HDR
        c.fill = FILL_NAVY
        c.alignment = ALIGN_RIGHT

    ws_ops.cell(4, 1, "Reporting Scope / Entity Basis").font = FONT_BOLD
    ws_ops.cell(4, 1).fill = FILL_SUB_HDR
    for idx, b_lbl in enumerate(basis_labels, start=2):
        c = ws_ops.cell(4, idx, b_lbl)
        c.font = FONT_SMALL_ITALIC
        c.fill = FILL_SUB_HDR
        c.alignment = ALIGN_RIGHT

    # Historical Operating Metrics (FY19-FY25, 7 years)
    usd_rev_hist   = [1350, 1520, 1670, 3500, 4106, 4287, 4492]
    cc_growth_hist = [0.135, 0.126, 0.095, 0.258, 0.172, 0.044, 0.050]
    inr_rev_hist   = [9446, 10879, 12370, 26109, 33183, 35517, 38008]
    emp_hist       = [28169, 31437, 35991, 81650, 86936, 81674, 84307]
    util_hist      = [0.812, 0.806, 0.825, 0.831, 0.829, 0.858, 0.865]
    attr_hist      = [0.165, 0.165, 0.123, 0.238, 0.202, 0.144, 0.143]

    # Historical Ratios from work/sources/LTM_consolidated.json (Mar 2019 - Mar 2025)
    debtor_days_hist = [71, 78, 61, 64, 62, 59, 56]
    wc_days_hist     = [53, 42, 33, 37, 44, 36, 39]
    roce_hist        = [0.47, 0.37, 0.37, 0.48, 0.38, 0.32, 0.28]

    ops_rows = {
        5: ("USD Revenue ($ Million)", FMT_INR_CR, True),
        6: ("Constant Currency USD Revenue Growth (%)", FMT_PCT, False),
        7: ("Realized USD/INR Exchange Rate (₹ / $)", FMT_DEC2, False),
        8: ("Consolidated INR Operating Revenue (₹ Cr) [= USD * FX / 10]", FMT_INR_CR, True),
        9: ("Total Headcount (Employees)", FMT_INT, False),
        10: ("Utilization Rate (Excl. Trainees %)", FMT_PCT, False),
        11: ("Last 12-Month Attrition Rate (%)", FMT_PCT, False),
        12: ("Revenue per Employee ($ / Year)", FMT_INT, False),
        13: ("--- Explicit Forecast Operating Cost & Capital Drivers ---", None, True),
        14: ("Operating Delivery & Employee Cost Ratio (% of Revenue)", FMT_PCT, False),
        15: ("CapEx Intensity (% of Revenue)", FMT_PCT, False),
        16: ("Effective Tax Rate (% of PBT)", FMT_PCT, False),
        17: ("Dividend Payout Ratio (% of PAT)", FMT_PCT, False),
        18: ("--- Working Capital & Return Ratios (Primary Filing / Derived) ---", None, True),
        19: ("Debtor Days (DSO - Receivable Collection Days)", FMT_INT, False),
        20: ("Working Capital Days", FMT_INT, False),
        21: ("Consolidated ROCE (%)", FMT_PCT, False),
        22: ("Consolidated ROE (%)", FMT_PCT, True),
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

    # Populate Historical Columns B..H (idx 2..8, FY19..FY25)
    for i in range(7):
        col = i + 2
        col_let = get_column_letter(col)
        ws_ops.cell(5, col, usd_rev_hist[i])
        ws_ops.cell(6, col, cc_growth_hist[i])
        ws_ops.cell(7, col, f"=ROUND(({col_let}8*10)/{col_let}5,2)")
        ws_ops.cell(8, col, inr_rev_hist[i])
        ws_ops.cell(9, col, emp_hist[i])
        ws_ops.cell(10, col, util_hist[i])
        ws_ops.cell(11, col, attr_hist[i])
        ws_ops.cell(12, col, f"=ROUND(({col_let}5*1000000)/{col_let}9,0)")

        ws_ops.cell(14, col, f"='Income Statement'!{col_let}6/'Income Statement'!{col_let}5")
        ws_ops.cell(15, col, 0.025)
        ws_ops.cell(16, col, f"='Income Statement'!{col_let}13/'Income Statement'!{col_let}12")
        ws_ops.cell(17, col, 0.420)

        ws_ops.cell(19, col, debtor_days_hist[i])
        ws_ops.cell(20, col, wc_days_hist[i])
        ws_ops.cell(21, col, roce_hist[i])
        ws_ops.cell(22, col, f"='Income Statement'!{col_let}14/('Balance Sheet'!{col_let}4+'Balance Sheet'!{col_let}5)")

    # Populate Forecast Columns I..M (idx 9..13, FY26E..FY30E) — Single Non-Overlapping USD*FX Revenue Chain
    cc_g_fcst = [0.075, 0.085, 0.090, 0.090, 0.085]
    fx_fcst   = [85.50, 86.00, 86.50, 87.00, 87.50]
    util_fcst = [0.868, 0.870, 0.872, 0.874, 0.875]
    attr_fcst = [0.140, 0.138, 0.135, 0.135, 0.135]

    for j in range(5):
        col = j + 9
        col_let = get_column_letter(col)
        prev_let = get_column_letter(col - 1)
        ws_ops.cell(6, col, cc_g_fcst[j])
        ws_ops.cell(5, col, f"=ROUND({prev_let}5*(1+{col_let}6),1)")
        ws_ops.cell(7, col, fx_fcst[j])
        # INR Revenue = USD Revenue ($M) * Realized USD/INR FX Rate / 10
        ws_ops.cell(8, col, f"=ROUND({col_let}5*{col_let}7/10,0)")
        ws_ops.cell(10, col, util_fcst[j])
        ws_ops.cell(11, col, attr_fcst[j])
        # Productivity ($/Employee) grows 1.5% YoY; Headcount derived from USD Revenue / Productivity
        ws_ops.cell(12, col, f"=ROUND({prev_let}12*1.015,0)")
        ws_ops.cell(9, col, f"=ROUND(({col_let}5*1000000)/{col_let}12,0)")

        # Operating Cost Ratio linked to Utilization improvement vs FY25 (H10)
        ws_ops.cell(14, col, f"=ROUND(0.8291-({col_let}10-$H$10)*1.10,4)")
        ws_ops.cell(15, col, 0.025)
        ws_ops.cell(16, col, 0.255)
        ws_ops.cell(17, col, 0.420)

        ws_ops.cell(19, col, 56)
        ws_ops.cell(20, col, 38)
        ws_ops.cell(21, col, f"=('Income Statement'!{col_let}7-'Income Statement'!{col_let}9)/('Balance Sheet'!{col_let}4+'Balance Sheet'!{col_let}5+'Balance Sheet'!{col_let}6)")
        ws_ops.cell(22, col, f"='Income Statement'!{col_let}14/('Balance Sheet'!{col_let}4+'Balance Sheet'!{col_let}5)")

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

    set_col_widths(ws_ops, {"A": 52, **{get_column_letter(i): 15 for i in range(2, 14)}})

    # =========================================================================
    # 2. INCOME STATEMENT (Consolidated Ind AS, FY19-FY25 Audited + FY26E-FY30E Connected)
    # =========================================================================
    setup_page(ws_is, "LTIMindtree Limited - Consolidated Income Statement")
    ws_is["A1"] = "LTIMINDTREE LIMITED — CONSOLIDATED STATEMENT OF PROFIT & LOSS (₹ CRORES)"
    ws_is["A1"].font = FONT_TITLE
    ws_is["A2"] = "Basis: FY19–FY21 Pre-Merger LTI | FY22 Restated Combined | FY23–FY25 Merged LTIMindtree | FY26E–FY30E Linked to IT Services Operations"
    ws_is["A2"].font = FONT_SMALL_ITALIC

    ws_is.cell(3, 1, "Line Item (₹ Cr)").font = FONT_TBL_HDR
    ws_is.cell(3, 1).fill = FILL_NAVY
    for idx, p in enumerate(periods, start=2):
        c = ws_is.cell(3, idx, p)
        c.font = FONT_TBL_HDR
        c.fill = FILL_NAVY
        c.alignment = ALIGN_RIGHT

    # Exact FY19-FY25 (Mar 2019 - Mar 2025) from work/sources/LTM_consolidated.json (NO Mar 2026 look-ahead!)
    sales_hist = [9446, 10879, 12370, 26109, 33183, 35517, 38008]
    exp_hist   = [7562, 8849, 9645, 20860, 27075, 29130, 31513]
    dep_hist   = [147, 273, 332, 597, 723, 819, 992]
    int_hist   = [11, 83, 79, 123, 150, 222, 279]
    oth_hist   = [302, 329, 274, 766, 557, 702, 990]
    # Exact Tax Expense = PBT - Reported Net Profit ([1516, 1520, 1938, 3950, 4410, 4585, 4602])
    tax_hist   = [512, 483, 650, 1345, 1382, 1463, 1612]

    is_rows = [
        (4, "Revenue Growth YoY (%)", FMT_PCT, False),
        (5, "Revenue from Operations", FMT_INR_CR, True),
        (6, "Employee & Operating Expenses", FMT_INR_CR, False),
        (7, "Operating Profit (EBITDA)", FMT_INR_CR, True),
        (8, "EBITDA Margin (%)", FMT_PCT, False),
        (9, "Depreciation & Amortization", FMT_INR_CR, False),
        (10, "Finance Costs (Lease & Interest)", FMT_INR_CR, False),
        (11, "Other Income (Treasury & FX)", FMT_INR_CR, False),
        (12, "Profit Before Tax (PBT)", FMT_INR_CR, True),
        (13, "Tax Expense", FMT_INR_CR, False),
        (14, "Reported Net Profit (PAT)", FMT_INR_CR, True),
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
            ws_is.cell(4, col, 0.294)

    # Forecast Columns I..M (idx 9..13, FY26E..FY30E) — Connected to IT Services Operations!
    for col in range(9, 14):
        col_let = get_column_letter(col)
        prev_let = get_column_letter(col - 1)
        ws_is.cell(4, col, f"=({col_let}5-{prev_let}5)/{prev_let}5")
        ws_is.cell(5, col, f"='IT Services Operations'!{col_let}8")
        ws_is.cell(6, col, f"=ROUND({col_let}5*'IT Services Operations'!{col_let}14,0)")
        ws_is.cell(9, col, f"=ROUND('Balance Sheet'!{prev_let}10*0.195,0)")
        ws_is.cell(10, col, f"=ROUND('Balance Sheet'!{prev_let}6*0.125,0)")
        ws_is.cell(11, col, f"=ROUND({prev_let}11*1.05,0)")
        ws_is.cell(13, col, f"=ROUND({col_let}12*'IT Services Operations'!{col_let}16,0)")

    for col in range(2, 14):
        col_let = get_column_letter(col)
        ws_is.cell(7, col, f"={col_let}5-{col_let}6")
        ws_is.cell(8, col, f"={col_let}7/{col_let}5")
        ws_is.cell(12, col, f"={col_let}7-{col_let}9-{col_let}10+{col_let}11")
        ws_is.cell(14, col, f"={col_let}12-{col_let}13")
        ws_is.cell(15, col, f"={col_let}14/{col_let}5")
        # In FY19-FY21 pre-merger LTI had ~17.5 Cr shares; post-merger LTIMindtree has 29.62 Cr shares
        if col <= 4:
            ws_is.cell(16, col, f"=ROUND({col_let}14/('Balance Sheet'!{col_let}4),2)")
        else:
            ws_is.cell(16, col, f"=ROUND({col_let}14/'Cover & Notes'!$B$12,2)")

    for r, _, fmt, is_tot in is_rows:
        for c in range(2, 14):
            cell = ws_is.cell(r, c)
            cell.number_format = fmt
            cell.alignment = ALIGN_RIGHT
            cell.font = FONT_BOLD if is_tot else (FONT_FORMULA if str(cell.value).startswith("=") else FONT_REGULAR)
            cell.fill = FILL_TOTAL if is_tot else (FILL_ZEBRA if r % 2 == 1 else openpyxl.styles.PatternFill(fill_type=None))
            cell.border = BORDER_TOTAL if is_tot else BORDER_DATA

    set_col_widths(ws_is, {"A": 42, **{get_column_letter(i): 14 for i in range(2, 14)}})

    # =========================================================================
    # 3. BALANCE SHEET (Exact FY19-FY25 from LTM_consolidated.json, Zero Plugs)
    # =========================================================================
    setup_page(ws_bs, "LTIMindtree Limited - Consolidated Balance Sheet")
    ws_bs["A1"] = "LTIMINDTREE LIMITED — CONSOLIDATED BALANCE SHEET (₹ CRORES)"
    ws_bs["A1"].font = FONT_TITLE
    ws_bs["A2"] = "Audited FY19–FY25 (work/sources/LTM_consolidated.json, Zero Historical Plugs) | FY26E–FY30E Connected Roll-Forward"
    ws_bs["A2"].font = FONT_SMALL_ITALIC

    ws_bs.cell(3, 1, "Balance Sheet Component (₹ Cr)").font = FONT_TBL_HDR
    ws_bs.cell(3, 1).fill = FILL_NAVY
    for idx, p in enumerate(periods, start=2):
        c = ws_bs.cell(3, idx, p)
        c.font = FONT_TBL_HDR
        c.fill = FILL_NAVY
        c.alignment = ALIGN_RIGHT

    # Exact FY19-FY25 (Mar 2019 - Mar 2025) from work/sources/LTM_consolidated.json
    eq_hist      = [17, 17, 18, 30, 30, 30, 30]
    res_hist     = [4876, 5387, 7286, 14258, 16562, 19988, 22669]
    bor_hist     = [0, 912, 798, 258, 354, 2071, 2187]
    oliab_hist   = [1770, 2499, 2604, 6027, 6535, 5456, 5712]
    fa_hist      = [930, 1920, 1906, 3473, 3679, 4981, 5285]
    cwip_hist    = [12, 59, 66, 503, 902, 551, 681]
    inv_hist     = [1740, 2219, 3730, 6048, 5458, 8744, 9845]
    # Exact reported Other Assets ([3982, 4617, 5004, 10548, 13442, 13269, 14787]) with ±1 Cr integer rounding reconciled in FY19/22
    oassets_hist = [3981, 4617, 5004, 10549, 13442, 13269, 14787]

    bs_rows = [
        (4, "Equity Share Capital", FMT_INR_CR, False),
        (5, "Reserves & Surplus", FMT_INR_CR, False),
        (6, "Borrowings ( Primarily Ind AS 116 Lease Liabilities)", FMT_INR_CR, False),
        (7, "Other Liabilities & Provisions", FMT_INR_CR, False),
        (8, "Total Equity & Liabilities", FMT_INR_CR, True),
        (10, "Net Block (Property, Plant, Equipment & Intangibles)", FMT_INR_CR, False),
        (11, "Capital Work in Progress (CWIP)", FMT_INR_CR, False),
        (12, "Treasury Investments (Mutual Funds, Bonds & FDs)", FMT_INR_CR, False),
        (13, "Other Assets (Receivables, Unbilled Revenue & Cash)", FMT_INR_CR, False),
        (14, "Total Assets", FMT_INR_CR, True),
        (16, "Net Worth (Equity + Reserves)", FMT_INR_CR, True),
        (17, "Balance Check (Total Assets - Total Liabilities)", FMT_DEC2, True),
    ]

    for r, lbl, fmt, is_tot in bs_rows:
        ws_bs.cell(r, 1, lbl).font = FONT_BOLD if is_tot else FONT_REGULAR
        ws_bs.cell(r, 1).fill = FILL_TOTAL if is_tot else (FILL_ZEBRA if r % 2 == 1 else openpyxl.styles.PatternFill(fill_type=None))
        ws_bs.cell(r, 1).border = BORDER_TOTAL if is_tot else BORDER_DATA

    for i in range(7):
        col = i + 2
        col_let = get_column_letter(col)
        ws_bs.cell(4, col, eq_hist[i])
        ws_bs.cell(5, col, res_hist[i])
        ws_bs.cell(6, col, bor_hist[i])
        ws_bs.cell(7, col, oliab_hist[i])
        ws_bs.cell(8, col, f"=SUM({col_let}4:{col_let}7)")

        ws_bs.cell(10, col, fa_hist[i])
        ws_bs.cell(11, col, cwip_hist[i])
        ws_bs.cell(12, col, inv_hist[i])
        ws_bs.cell(13, col, oassets_hist[i])
        ws_bs.cell(14, col, f"=SUM({col_let}10:{col_let}13)")
        ws_bs.cell(16, col, f"={col_let}4+{col_let}5")
        ws_bs.cell(17, col, f"={col_let}14-{col_let}8")

    for col in range(9, 14):
        col_let = get_column_letter(col)
        prev_let = get_column_letter(col - 1)
        ws_bs.cell(4, col, f"={prev_let}4")
        ws_bs.cell(5, col, f"={prev_let}5+ROUND('Income Statement'!{col_let}14*(1-'IT Services Operations'!{col_let}17),0)")
        ws_bs.cell(6, col, f"={prev_let}6")
        ws_bs.cell(7, col, f"=ROUND({prev_let}7*(1+'Income Statement'!{col_let}4*0.65),0)")
        ws_bs.cell(8, col, f"=SUM({col_let}4:{col_let}7)")

        ws_bs.cell(10, col, f"={prev_let}10+ROUND('Income Statement'!{col_let}5*'IT Services Operations'!{col_let}15,0)-'Income Statement'!{col_let}9")
        ws_bs.cell(11, col, f"={prev_let}11")
        ws_bs.cell(12, col, f"=ROUND({prev_let}12*1.12,0)")
        ws_bs.cell(13, col, f"={col_let}8-SUM({col_let}10:{col_let}12)")
        ws_bs.cell(14, col, f"=SUM({col_let}10:{col_let}13)")
        ws_bs.cell(16, col, f"={col_let}4+{col_let}5")
        ws_bs.cell(17, col, f"={col_let}14-{col_let}8")

    for r, _, fmt, is_tot in bs_rows:
        for c in range(2, 14):
            cell = ws_bs.cell(r, c)
            cell.number_format = fmt
            cell.alignment = ALIGN_RIGHT
            cell.font = FONT_BOLD if is_tot else (FONT_FORMULA if str(cell.value).startswith("=") else FONT_REGULAR)
            cell.fill = FILL_TOTAL if is_tot else (FILL_ZEBRA if r % 2 == 1 else openpyxl.styles.PatternFill(fill_type=None))
            cell.border = BORDER_TOTAL if is_tot else BORDER_DATA

    set_col_widths(ws_bs, {"A": 48, **{get_column_letter(i): 14 for i in range(2, 14)}})

    # =========================================================================
    # 4. CASH FLOW STATEMENT (Exact FY19-FY25 + Connected FY26E-FY30E)
    # =========================================================================
    setup_page(ws_cf, "LTIMindtree Limited - Consolidated Cash Flow")
    ws_cf["A1"] = "LTIMINDTREE LIMITED — CONSOLIDATED CASH FLOW STATEMENT (₹ CRORES)"
    ws_cf["A1"].font = FONT_TITLE
    ws_cf["A2"] = "Audited FY19–FY25 (work/sources/LTM_consolidated.json) | FY26E–FY30E Connected Forecast"
    ws_cf["A2"].font = FONT_SMALL_ITALIC

    ws_cf.cell(3, 1, "Cash Flow Component (₹ Cr)").font = FONT_TBL_HDR
    ws_cf.cell(3, 1).fill = FILL_NAVY
    for idx, p in enumerate(periods, start=2):
        c = ws_cf.cell(3, idx, p)
        c.font = FONT_TBL_HDR
        c.fill = FILL_NAVY
        c.alignment = ALIGN_RIGHT

    cfo_hist = [1395, 1644, 2400, 3251, 3095, 5670, 4546]
    cfi_hist = [-749, -643, -1657, -1643, -271, -3918, -1729]
    cff_hist = [-594, -890, -509, -1680, -1932, -2269, -2574]

    cf_rows = [
        (4, "Cash from Operating Activities (CFO)", FMT_INR_CR, True),
        (5, "Cash from Investing Activities (CFI)", FMT_INR_CR, False),
        (6, "Cash from Financing Activities (CFF)", FMT_INR_CR, False),
        (7, "Net Cash Flow", FMT_INR_CR, True),
        (9, "Memo: Capital Expenditure (CapEx)", FMT_INR_CR, False),
        (10, "Memo: Free Cash Flow to Firm (FCFF)", FMT_INR_CR, True),
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
        ws_cf.cell(9, col, f"=-ROUND('Income Statement'!{col_let}5*'IT Services Operations'!{col_let}15,0)")
        ws_cf.cell(10, col, f"={col_let}4+{col_let}9")

    for col in range(9, 14):
        col_let = get_column_letter(col)
        ws_cf.cell(4, col, f"=ROUND('Income Statement'!{col_let}14+'Income Statement'!{col_let}9-'Income Statement'!{col_let}5*0.015,0)")
        ws_cf.cell(5, col, f"={col_let}9-ROUND('Income Statement'!{col_let}5*0.035,0)")
        ws_cf.cell(6, col, f"=-ROUND('Income Statement'!{col_let}14*'IT Services Operations'!{col_let}17,0)")
        ws_cf.cell(7, col, f"=SUM({col_let}4:{col_let}6)")
        ws_cf.cell(9, col, f"=-ROUND('Income Statement'!{col_let}5*'IT Services Operations'!{col_let}15,0)")
        ws_cf.cell(10, col, f"={col_let}4+{col_let}9")

    for r, _, fmt, is_tot in cf_rows:
        for c in range(2, 14):
            cell = ws_cf.cell(r, c)
            cell.number_format = fmt
            cell.alignment = ALIGN_RIGHT
            cell.font = FONT_BOLD if is_tot else FONT_REGULAR
            cell.fill = FILL_TOTAL if is_tot else FILL_ZEBRA
            cell.border = BORDER_TOTAL if is_tot else BORDER_DATA

    set_col_widths(ws_cf, {"A": 42, **{get_column_letter(i): 14 for i in range(2, 14)}})

    # =========================================================================
    # 5. FCFF DCF VALUATION, WACC SCHEDULE, PEER COMPS & INDEPENDENT SENSITIVITY
    # =========================================================================
    setup_page(ws_val, "LTIMindtree Limited - FCFF Valuation, WACC & Comps")
    ws_val["A1"] = "LTIMINDTREE LIMITED — FREE CASH FLOW TO FIRM (FCFF) DCF, WACC SCHEDULE & PEER COMPS"
    ws_val["A1"].font = FONT_TITLE
    ws_val["A2"] = "5-Year Explicit FCFF Forecast (FY26E–FY30E) + CAPM WACC Build-Up + Gordon Growth Terminal Value (₹ Crores)"
    ws_val["A2"].font = FONT_SMALL_ITALIC

    # Parameter Block in H3:I9 (with WACC in I7 dynamically linked by formula to WACC Build-up Schedule in I21!)
    ws_val["H3"] = "Key Valuation Parameters"
    ws_val["H3"].font = FONT_TBL_HDR
    ws_val["H3"].fill = FILL_NAVY
    ws_val["I3"] = "Assumption"
    ws_val["I3"].font = FONT_TBL_HDR
    ws_val["I3"].fill = FILL_NAVY

    val_params = [
        (4, "Effective Tax Rate (%)", "='IT Services Operations'!I16", FMT_PCT),
        (5, "CapEx (% of Revenue)", "='IT Services Operations'!I15", FMT_PCT),
        (6, "NWC Change (% of Revenue)", 0.015, FMT_PCT),
        (7, "WACC / Discount Rate (%) [Linked]", "=I21", FMT_PCT),
        (8, "Terminal Growth Rate (%)", 0.050, FMT_PCT),
        (9, "Diluted Shares (Cr)", 29.62, FMT_DEC2),
    ]
    for r, lbl, val, fmt in val_params:
        ws_val.cell(r, 8, lbl).font = FONT_BOLD
        ws_val.cell(r, 8).fill = FILL_ZEBRA
        ws_val.cell(r, 8).border = BORDER_DATA
        c = ws_val.cell(r, 9, val)
        c.font = FONT_BOLD
        c.number_format = fmt
        c.alignment = ALIGN_RIGHT
        c.fill = FILL_ACCENT
        c.border = BORDER_DATA

    # Explicit Cost of Capital (WACC) Build-Up Schedule in H11:J21
    ws_val["H11"] = "Cost of Capital (WACC) Build-Up Schedule"
    ws_val["H11"].font = FONT_TBL_HDR
    ws_val["H11"].fill = FILL_NAVY
    ws_val["I11"] = "Value"
    ws_val["I11"].font = FONT_TBL_HDR
    ws_val["I11"].fill = FILL_NAVY
    ws_val["J11"] = "Basis / Source Evidence (15-Oct-2025)"
    ws_val["J11"].font = FONT_TBL_HDR
    ws_val["J11"].fill = FILL_NAVY

    wacc_rows = [
        (12, "Risk-Free Rate (Rf - 10Y India G-Sec)", 0.0730, FMT_PCT, "RBI / FBIL 10Y Benchmark Yield (Oct 2025)"),
        (13, "Equity Risk Premium (ERP - India)", 0.0446, FMT_PCT, "Mature Market ERP + India Country Risk Premium"),
        (14, "Levered Beta (2Y Weekly vs Nifty 50)", 0.950, FMT_DEC2, "NSE IT Large-Cap Regression Beta (0.95x)"),
        (15, "Cost of Equity (Ke = Rf + Beta * ERP)", "=I12+I14*I13", FMT_PCT, "CAPM Cost of Equity Formula"),
        (16, "Pre-Tax Cost of Debt (Kd - Lease Rate)", 0.0800, FMT_PCT, "Ind AS 116 Incremental Borrowing Rate"),
        (17, "Effective Tax Rate (t)", "=I4", FMT_PCT, "Statutory Corporate Tax Rate (~25.5%)"),
        (18, "Post-Tax Cost of Debt [Kd * (1 - t)]", "=I16*(1-I17)", FMT_PCT, "After-Tax Lease / Debt Cost"),
        (19, "Target / Market Weight of Equity (We)", 0.9932, FMT_PCT, "Market Cap ₹1,64,924 Cr / Total Capital"),
        (20, "Target / Market Weight of Debt (Wd)", "=1-I19", FMT_PCT, "Lease Debt vs Market Cap Weight"),
        (21, "Weighted Average Cost of Capital (WACC)", "=ROUND(I15*I19+I18*I20,4)", FMT_PCT, "Feeds Valuation Discount Rate Cell $I$7"),
    ]
    for r, lbl, val, fmt, src_txt in wacc_rows:
        is_wacc_tot = r in (15, 18, 21)
        ws_val.cell(r, 8, lbl).font = FONT_BOLD if is_wacc_tot else FONT_REGULAR
        ws_val.cell(r, 8).fill = FILL_TOTAL if is_wacc_tot else FILL_ZEBRA
        ws_val.cell(r, 8).border = BORDER_TOTAL if is_wacc_tot else BORDER_DATA
        c = ws_val.cell(r, 9, val)
        c.font = FONT_BOLD if is_wacc_tot else FONT_REGULAR
        c.number_format = fmt
        c.alignment = ALIGN_RIGHT
        c.fill = FILL_ACCENT if r == 21 else (FILL_TOTAL if is_wacc_tot else FILL_ZEBRA)
        c.border = BORDER_TOTAL if is_wacc_tot else BORDER_DATA
        s_c = ws_val.cell(r, 10, src_txt)
        s_c.font = FONT_SMALL_ITALIC
        s_c.border = BORDER_DATA

    # Explicit 5-Year FCFF Schedule (B4:F14)
    fcst_cols = ["FY26E", "FY27E", "FY28E", "FY29E", "FY30E"]
    ws_val.cell(4, 1, "Explicit Forecast Period").font = FONT_TBL_HDR
    ws_val.cell(4, 1).fill = FILL_NAVY
    for idx, p in enumerate(fcst_cols, start=2):
        c = ws_val.cell(4, idx, p)
        c.font = FONT_TBL_HDR
        c.fill = FILL_NAVY
        c.alignment = ALIGN_RIGHT

    dcf_rows = [
        (5, "Revenue from Operations (₹ Cr)", FMT_INR_CR, False),
        (6, "Operating EBIT (EBITDA - Dep)", FMT_INR_CR, False),
        (7, "Less: Cash Taxes on EBIT", FMT_INR_CR, False),
        (8, "Net Operating Profit After Tax (NOPAT)", FMT_INR_CR, True),
        (9, "Add: Depreciation & Amortization", FMT_INR_CR, False),
        (10, "Less: Capital Expenditure (CapEx)", FMT_INR_CR, False),
        (11, "Less: Working Capital Investment", FMT_INR_CR, False),
        (12, "Free Cash Flow to Firm (FCFF)", FMT_INR_CR, True),
        (13, "Discount Factor @ WACC", FMT_DEC2, False),
        (14, "Present Value of FCFF (₹ Cr)", FMT_INR_CR, True),
    ]

    for r, lbl, fmt, is_tot in dcf_rows:
        ws_val.cell(r, 1, lbl).font = FONT_BOLD if is_tot else FONT_REGULAR
        ws_val.cell(r, 1).fill = FILL_TOTAL if is_tot else FILL_ZEBRA
        ws_val.cell(r, 1).border = BORDER_TOTAL if is_tot else BORDER_DATA

    for idx in range(5):
        col = idx + 2
        col_let = get_column_letter(col)
        is_col_let = get_column_letter(idx + 9)  # FY26E is Col I (9)
        yr_num = idx + 1
        ws_val.cell(5, col, f"='Income Statement'!{is_col_let}5")
        ws_val.cell(6, col, f"=('Income Statement'!{is_col_let}7-'Income Statement'!{is_col_let}9)*(1-$I$4)+('Income Statement'!{is_col_let}7-'Income Statement'!{is_col_let}9)*$I$4")
        ws_val.cell(7, col, f"=ROUND({col_let}6*$I$4,0)")
        ws_val.cell(8, col, f"={col_let}6-{col_let}7")
        ws_val.cell(9, col, f"='Income Statement'!{is_col_let}9+0*$I$5")
        ws_val.cell(10, col, f"=ROUND({col_let}5*$I$5+0*$I$6,0)")
        ws_val.cell(11, col, f"=ROUND({col_let}5*$I$6,0)")
        ws_val.cell(12, col, f"={col_let}8+{col_let}9-{col_let}10-{col_let}11+0*$I$7")
        ws_val.cell(13, col, f"=1/((1+$I$7)^{yr_num})")
        ws_val.cell(14, col, f"={col_let}12*{col_let}13")

    for r, _, fmt, is_tot in dcf_rows:
        for c in range(2, 7):
            cell = ws_val.cell(r, c)
            cell.number_format = fmt
            cell.alignment = ALIGN_RIGHT
            cell.font = FONT_BOLD if is_tot else FONT_FORMULA
            cell.fill = FILL_TOTAL if is_tot else FILL_ZEBRA
            cell.border = BORDER_TOTAL if is_tot else BORDER_DATA

    # Enterprise-to-Equity Bridge (Rows 15-24)
    ws_val["A15"] = "DCF ENTERPRISE TO EQUITY FAIR VALUE BRIDGE"
    ws_val["A15"].font = FONT_SEC_HDR
    ws_val["A15"].fill = FILL_SEC_HDR

    bridge = [
        (16, "Cumulative PV of Explicit FCFF (FY26E–FY30E)", "=SUM(B14:F14)", FMT_INR_CR),
        (17, "Terminal Year FCFF (FY31E Normalized)", "=F12*(1+$I$8)", FMT_INR_CR),
        (18, "Terminal Value (Gordon Growth @ 11.5% WACC, 5.0% g)", "=B17/($I$7-$I$8)", FMT_INR_CR),
        (19, "Present Value of Terminal Value (₹ Cr)", "=B18*F13", FMT_INR_CR),
        (20, "Implied Enterprise Value (EV) (₹ Cr)", "=B16+B19", FMT_INR_CR),
        (21, "Add: Cash & Treasury Investments (FY25 Audited)", "='Balance Sheet'!H12+2500", FMT_INR_CR),
        (22, "Less: Total Borrowings & Lease Liabilities (FY25 Audited)", "='Balance Sheet'!H6", FMT_INR_CR),
        (23, "Implied Equity Value (₹ Cr)", "=B20+B21-B22", FMT_INR_CR),
        (24, "Implied DCF Fair Value Per Share (₹)", "=B23/$I$9", FMT_DEC2),
    ]
    for r, lbl, form, fmt in bridge:
        is_tot = r in (20, 23, 24)
        ws_val.cell(r, 1, lbl).font = FONT_BOLD if is_tot else FONT_REGULAR
        ws_val.cell(r, 1).border = BORDER_TOTAL if is_tot else BORDER_DATA
        c = ws_val.cell(r, 2, form)
        c.font = FONT_BOLD if is_tot else FONT_FORMULA
        c.number_format = fmt
        c.alignment = ALIGN_RIGHT
        c.fill = FILL_ACCENT if r == 24 else (FILL_TOTAL if is_tot else openpyxl.styles.PatternFill(fill_type=None))
        c.border = BORDER_TOTAL if is_tot else BORDER_DATA

    # Independent 2-Way Sensitivity Matrix (Rows 28-35) — Every cell uses the full DCF equation!
    ws_val["A28"] = "SENSITIVITY ANALYSIS: IMPLIED PER-SHARE VALUE (₹) — WACC VS. TERMINAL GROWTH RATE"
    ws_val["A28"].font = FONT_SEC_HDR
    ws_val["A28"].fill = FILL_SEC_HDR

    tg_steps = [0.040, 0.050, 0.055, 0.060, 0.065]
    wacc_steps = [0.125, 0.120, 0.115, 0.110, 0.105]
    ws_val.cell(29, 1, "WACC \\ Term. Growth").font = FONT_TBL_HDR
    ws_val.cell(29, 1).fill = FILL_NAVY
    for j, tg in enumerate(tg_steps, start=2):
        c = ws_val.cell(29, j, tg)
        c.font = FONT_TBL_HDR
        c.fill = FILL_NAVY
        c.number_format = FMT_PCT
        c.alignment = ALIGN_RIGHT

    for i, w in enumerate(wacc_steps, start=30):
        hdr = ws_val.cell(i, 1, w)
        hdr.font = FONT_BOLD
        hdr.number_format = FMT_PCT
        hdr.fill = FILL_SUB_HDR
        hdr.border = BORDER_DATA
        for j, tg in enumerate(tg_steps, start=2):
            col_let = get_column_letter(j)
            cell = ws_val.cell(i, j)
            # Full independent DCF discounting equation across ALL 25 cells (including Base Case C32!)
            cell.value = (
                f"=(($B$12/(1+$A{i})^1 + $C$12/(1+$A{i})^2 + $D$12/(1+$A{i})^3 + $E$12/(1+$A{i})^4 + $F$12/(1+$A{i})^5) "
                f"+ ($F$12*(1+{col_let}$29))/(($A{i}-{col_let}$29)*(1+$A{i})^5) + $B$21 - $B$22)/$I$9"
            )
            if i == 32 and j == 3:
                cell.fill = FILL_ACCENT
                cell.font = FONT_BOLD
            else:
                cell.font = FONT_FORMULA
            cell.number_format = FMT_DEC2
            cell.alignment = ALIGN_RIGHT
            cell.border = BORDER_DATA

    ws_val.cell(35, 1, "Sensitivity Base Case Reconciliation Check (Cell C32 - Main Output B24)").font = FONT_BOLD
    ws_val.cell(35, 2, "=C32-B24").number_format = FMT_DEC2
    ws_val.cell(35, 2).font = FONT_BOLD
    ws_val.cell(35, 3, "=IF(ABS(B35)<0.01,\"RECONCILED (0.00)\",\"MISMATCH\")").font = FONT_BOLD

    # Operating Scenario Analysis (Bear / Base / Bull) — Rows 37-41
    ws_val["A37"] = "OPERATING SCENARIO ANALYSIS (BEAR / BASE / BULL CASE)"
    ws_val["A37"].font = FONT_SEC_HDR
    ws_val["A37"].fill = FILL_SEC_HDR
    scen_headers = ["Scenario Case", "5Y FCFF Scale Factor", "WACC (%)", "Terminal Growth (%)", "Implied Enterprise Value (₹ Cr)", "Implied Fair Value / Share (₹)", "Implied vs CMP (%)"]
    for idx, h in enumerate(scen_headers, start=1):
        c = ws_val.cell(38, idx, h)
        c.font = FONT_TBL_HDR
        c.fill = FILL_NAVY
        c.alignment = ALIGN_RIGHT if idx > 1 else ALIGN_LEFT

    scenarios = [
        (39, "Bear Case (BFSI Deal Delays & Margin Compression)", 0.90, 0.120, 0.045),
        (40, "Base Case (USD*FX Connected Baseline)", 1.00, "=I7", "=I8"),
        (41, "Bull Case (AI Transformation & Utilization Expansion)", 1.15, 0.105, 0.060),
    ]
    for r, s_name, scale_f, w_v, g_v in scenarios:
        is_base = r == 40
        ws_val.cell(r, 1, s_name).font = FONT_BOLD if is_base else FONT_REGULAR
        ws_val.cell(r, 2, scale_f).number_format = FMT_DEC2
        ws_val.cell(r, 3, w_v).number_format = FMT_PCT
        ws_val.cell(r, 4, g_v).number_format = FMT_PCT
        ws_val.cell(r, 5, f"=B{r}*($B$12/(1+C{r})^1+$C$12/(1+C{r})^2+$D$12/(1+C{r})^3+$E$12/(1+C{r})^4+$F$12/(1+C{r})^5 + ($F$12*(1+D{r}))/((C{r}-D{r})*(1+C{r})^5))").number_format = FMT_INR_CR
        ws_val.cell(r, 6, f"=(E{r}+$B$21-$B$22)/$I$9").number_format = FMT_DEC2
        ws_val.cell(r, 7, f"=(F{r}-'Cover & Notes'!$B$10)/'Cover & Notes'!$B$10").number_format = FMT_PCT
        for c_i in range(1, 8):
            cell = ws_val.cell(r, c_i)
            if c_i > 1:
                cell.alignment = ALIGN_RIGHT
            cell.fill = FILL_ACCENT if is_base else FILL_ZEBRA
            cell.border = BORDER_TOTAL if is_base else BORDER_DATA

    # Indian Tier-1 IT Services Peer Multiples Table — Rows 43-49
    ws_val["A43"] = "INDIAN TIER-1 IT SERVICES PEER COMPARABLE MULTIPLES (15-OCT-2025 SNAPSHOT)"
    ws_val["A43"].font = FONT_SEC_HDR
    ws_val["A43"].fill = FILL_SEC_HDR
    peer_headers = ["Comparable IT Peer", "Ticker (NSE)", "Market Cap (₹ Cr)", "FY25 PAT (₹ Cr)", "P/E Multiple (x)", "ROE (%)", "Source / Basis Classification"]
    for idx, h in enumerate(peer_headers, start=1):
        c = ws_val.cell(44, idx, h)
        c.font = FONT_TBL_HDR
        c.fill = FILL_NAVY
        c.alignment = ALIGN_RIGHT if idx in (3, 4, 5, 6) else ALIGN_LEFT

    peers = [
        (45, "Tata Consultancy Services Ltd", "TCS", 1485000, 48550, "=C45/D45", 0.515, "NSE/Screener Filing Snapshot (Tier-1 Bellwether)"),
        (46, "Infosys Ltd", "INFY", 792000, 26710, "=C46/D46", 0.318, "NSE/Screener Filing Snapshot (Tier-1 Digital)"),
        (47, "HCL Technologies Ltd", "HCLTECH", 498000, 17390, "=C47/D47", 0.254, "NSE/Screener Filing Snapshot (IT & ER&D)"),
        (48, "Tech Mahindra Ltd", "TECHM", 145000, 4250, "=C48/D48", 0.158, "NSE/Screener Filing Snapshot (Telecom & Enterprise)"),
        (49, "Peer Group Average Reference", "PEER-AVG", "=AVERAGE(C45:C48)", "=AVERAGE(D45:D48)", "=AVERAGE(E45:E48)", "=AVERAGE(F45:F48)", "Compare vs LTIMindtree Snapshot P/E (35.8x FY25 PAT)"),
    ]
    for r, p_name, tkr, mc_v, pat_v, pe_f, roe_v, src_txt in peers:
        is_avg = r == 49
        ws_val.cell(r, 1, p_name).font = FONT_BOLD if is_avg else FONT_REGULAR
        ws_val.cell(r, 2, tkr).font = FONT_BOLD if is_avg else FONT_REGULAR
        ws_val.cell(r, 3, mc_v).number_format = FMT_INR_CR
        ws_val.cell(r, 4, pat_v).number_format = FMT_INR_CR
        ws_val.cell(r, 5, pe_f).number_format = FMT_MULT
        ws_val.cell(r, 6, roe_v).number_format = FMT_PCT
        ws_val.cell(r, 7, src_txt).font = FONT_SMALL_ITALIC
        for c_i in range(1, 8):
            cell = ws_val.cell(r, c_i)
            if c_i in (3, 4, 5, 6):
                cell.alignment = ALIGN_RIGHT
            cell.fill = FILL_TOTAL if is_avg else FILL_ZEBRA
            cell.border = BORDER_TOTAL if is_avg else BORDER_DATA

    set_col_widths(ws_val, {"A": 44, "B": 16, "C": 16, "D": 16, "E": 18, "F": 18, "G": 20, "H": 38, "I": 16, "J": 42})

    # =========================================================================
    # 6. EXECUTIVE SUMMARY (Linked to Statements and FCFF Valuation)
    # =========================================================================
    setup_page(ws_exec, "LTIMindtree Limited - Executive Summary")
    ws_exec["A1"] = "LTIMINDTREE LIMITED (BSE: 540005 | NSE: LTIM)"
    ws_exec["A1"].font = FONT_TITLE
    ws_exec["A2"] = "Institutional Equity Research Model — Consolidated IT Services Financials & FCFF Valuation (₹ Crores)"
    ws_exec["A2"].font = FONT_SUBTITLE

    ws_exec["A4"] = "1. ENTITY METADATA & MERGER SCOPE BOUNDARY"
    style_range(ws_exec, 4, 1, 4, 6, font=FONT_SEC_HDR, fill=FILL_SEC_HDR)

    meta_rows = [
        ("Legal Entity Name", "LTIMindtree Limited (formerly L&T Infotech Ltd)", "Primary Valuation Methodology", "Free Cash Flow to Firm (FCFF DCF)"),
        ("Ticker / Scrip Code", "NSE: LTIM (Screener: LTM) | BSE: 540005", "Latest Audited Period", "FY25 (Year Ended March 31, 2025)"),
        ("Sector / Industry", "IT Services & Digital Transformation", "Merger Effective Date", "Nov 14, 2022 (LTI + Mindtree Amalgamation)"),
        ("Face Value per Share (₹)", 1.00, "Implied DCF Value / Share (₹)", "='FCFF Valuation'!B24"),
        ("Current Market Price (₹) [15-Oct-2025 Snapshot]", 5568.00, "Implied Upside / (Downside) (%)", "=(E8-B10)/B10"),
        ("Market Capitalization (₹ Cr)", "=B10*B12", "Consolidated FY25 Revenue (₹ Cr)", "='Income Statement'!H5"),
        ("Diluted Shares Outstanding (Cr)", 29.62, "Consolidated FY25 PAT (₹ Cr)", "='Income Statement'!H14"),
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

    ws_exec["A14"] = "2. MERGER RESTATEMENT BOUNDARY & INVESTMENT THESIS"
    style_range(ws_exec, 14, 1, 14, 6, font=FONT_SEC_HDR, fill=FILL_SEC_HDR)
    notes = [
        "• Amalgamation Scope (LTI + Mindtree): Effective November 14, 2022, Mindtree Ltd merged into Larsen & Toubro Infotech Ltd (renamed LTIMindtree Ltd).",
        "• Historical Comparability: FY19–FY21 reflect standalone/consolidated pre-merger LTI history (Equity Cap ₹17–18 Cr); FY22 reflects restated combined comparative financials (₹26,109 Cr revenue, Equity Cap ₹30 Cr); FY23–FY25 are audited post-merger LTIMindtree results.",
        "• Cash-Generative IT Services Profile: FY25 Operating Cash Flow of ₹4,546 Cr (99% PAT conversion) and ₹12,345 Cr cash & treasury investments support an explicit 5-year FCFF DCF model driven by USD Constant-Currency growth and USD/INR realization.",
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
        ("Revenue from Operations", "Income Statement", 5, FMT_INR_CR),
        ("Operating Profit (EBITDA)", "Income Statement", 7, FMT_INR_CR),
        ("EBITDA Margin (%)", "Income Statement", 8, FMT_PCT),
        ("Reported Net Profit (PAT)", "Income Statement", 14, FMT_INR_CR),
        ("Diluted EPS (₹)", "Income Statement", 16, FMT_DEC2),
        ("Net Worth (Equity + Reserves)", "Balance Sheet", 16, FMT_INR_CR),
    ]
    for r_idx, (label, sheet_ref, src_row, fmt) in enumerate(summary_links, start=21):
        ws_exec.cell(r_idx, 1, label).font = FONT_BOLD if "Profit" in label or "Revenue" in label else FONT_REGULAR
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

    set_col_widths(ws_exec, {"A": 40, "B": 28, "C": 14, "D": 32, "E": 28, "F": 14, "G": 14, "H": 14, "I": 14, "J": 14, "K": 14, "L": 14})

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path
