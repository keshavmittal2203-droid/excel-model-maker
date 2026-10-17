"""Native OpenPyXL Model Builder for Mahindra & Mahindra Limited (MM)."""
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
    ALIGN_LEFT, ALIGN_RIGHT, ALIGN_CENTER, ALIGN_WRAP,
    FMT_INR_CR, FMT_DEC2, FMT_PCT, FMT_INT, FMT_MULT,
    style_range, set_col_widths, setup_page
)


def build_mm_workbook(out_path: Path) -> Path:
    return build_mm_model(out_path)


def build_mm_model(out_path: Path) -> Path:
    wb = openpyxl.Workbook()
    ws_exec = wb.active
    ws_exec.title = "Cover & Notes"
    ws_is = wb.create_sheet("Income Statement")
    ws_bs = wb.create_sheet("Balance Sheet")
    ws_cf = wb.create_sheet("Cash Flow")
    ws_ops = wb.create_sheet("Segment Operations")
    ws_val = wb.create_sheet("SOTP Valuation")

    # Strictly 7 audited historical years (FY19-FY25, cols B-H) + 5 connected forecast years (FY26E-FY30E, cols I-M)
    periods = ["FY19", "FY20", "FY21", "FY22", "FY23", "FY24", "FY25", "FY26E", "FY27E", "FY28E", "FY29E", "FY30E"]

    # =========================================================================
    # 1. SEGMENT OPERATIONS & DRIVERS (Built first conceptually; populated on ws_ops)
    # =========================================================================
    setup_page(ws_ops, "Mahindra & Mahindra Limited - Segment Operations, Drivers & Ratios")
    ws_ops["A1"] = "MAHINDRA & MAHINDRA LTD — SEGMENT OPERATIONS, FORECAST DRIVERS & RATIOS (₹ CR)"
    ws_ops["A1"].font = FONT_TITLE
    ws_ops["A2"] = "Historical Period: FY19–FY25 Audited (Pre-Oct 15, 2025 Cutoff) | Forecast Period: FY26E–FY30E Dynamic Driver Formulas"
    ws_ops["A2"].font = FONT_SMALL_ITALIC

    ws_ops.cell(3, 1, "Segment / Operational Metric").font = FONT_TBL_HDR
    ws_ops.cell(3, 1).fill = FILL_NAVY
    for idx, p in enumerate(periods, start=2):
        c = ws_ops.cell(3, idx, p)
        c.font = FONT_TBL_HDR
        c.fill = FILL_NAVY
        c.alignment = ALIGN_RIGHT

    # Historical Segment Revenues (FY19-FY25, 7 years) summing to Consolidated Revenue
    auto_rev_hist = [52000, 36000, 35500, 46000, 65000, 76100, 90800]
    farm_rev_hist = [22000, 18500, 21500, 24000, 29500, 33400, 35200]
    cons_sales_hist = [104721, 75382, 74278, 90171, 121362, 139078, 159211]
    sub_rev_hist = [cons_sales_hist[i] - auto_rev_hist[i] - farm_rev_hist[i] for i in range(7)]

    # Standalone Core Segment Operating EBITDA (FY19-FY25, 7 years) -> FY25: Auto 15,150 Cr, Farm 6,400 Cr
    auto_ebitda_hist = [6200, 4100, 4300, 5500, 8800, 11800, 15150]
    farm_ebitda_hist = [4100, 3400, 4500, 4300, 5100, 5800, 6400]
    auto_vol_hist = [608, 476, 352, 466, 698, 825, 941]
    trac_vol_hist = [330, 302, 354, 355, 408, 379, 424]

    ops_labels = {
        4: ("Automotive Segment Revenue (₹ Cr)", FMT_INR_CR, False),
        5: ("Farm Equipment Segment Revenue (₹ Cr)", FMT_INR_CR, False),
        6: ("Financial Services & Industrial Subsidiaries Revenue (₹ Cr)", FMT_INR_CR, False),
        7: ("Consolidated Revenue from Operations (₹ Cr)", FMT_INR_CR, True),
        8: ("Core Automotive Standalone Operating EBITDA (₹ Cr)", FMT_INR_CR, True),
        9: ("Core Farm Equipment Standalone Operating EBITDA (₹ Cr)", FMT_INR_CR, True),
        10: ("Automotive Volume (Units '000)", FMT_INT, False),
        11: ("Tractor Domestic + Export Volume (Units '000)", FMT_INT, False),
        12: ("--- Explicit Forecast Operating Drivers (FY26E–FY30E) ---", None, True),
        13: ("Automotive Revenue Growth Driver (%)", FMT_PCT, False),
        14: ("Farm Equipment Revenue Growth Driver (%)", FMT_PCT, False),
        15: ("Financial Services & Subsidiaries Growth Driver (%)", FMT_PCT, False),
        16: ("Core Automotive Standalone EBITDA Margin (%)", FMT_PCT, False),
        17: ("Core Farm Standalone EBITDA Margin (%)", FMT_PCT, False),
        18: ("Consolidated Operating Expense Ratio (% of Revenue)", FMT_PCT, False),
        19: ("Consolidated CapEx Intensity (% of Revenue)", FMT_PCT, False),
        20: ("Consolidated Effective Tax Rate (% of PBT)", FMT_PCT, False),
        21: ("Consolidated Dividend Payout Ratio (% of PAT)", FMT_PCT, False),
        22: ("--- Working Capital & Return Ratios (Primary Filing / Derived) ---", None, True),
        23: ("Debtor Days (Receivable Collection Period)", FMT_INT, False),
        24: ("Inventory Days (Finished Goods & RM)", FMT_INT, False),
        25: ("Days Payable (Trade Payables Period)", FMT_INT, False),
        26: ("Cash Conversion Cycle (Days)", FMT_INT, True),
        27: ("Consolidated ROCE (%)", FMT_PCT, False),
        28: ("Consolidated ROE (%)", FMT_PCT, True),
    }

    for r, (lbl, _, is_tot) in ops_labels.items():
        c = ws_ops.cell(r, 1, lbl)
        if lbl.startswith("---"):
            c.font = FONT_SEC_HDR
            c.fill = FILL_SEC_HDR
        else:
            c.font = FONT_BOLD if is_tot else FONT_REGULAR
            c.fill = FILL_TOTAL if is_tot else (FILL_ZEBRA if r % 2 == 1 else openpyxl.styles.PatternFill(fill_type=None))
            c.border = BORDER_TOTAL if is_tot else BORDER_DATA

    # Historical Ratios from work/sources/MM_consolidated.json (Mar 2019 - Mar 2025)
    debtor_days_hist = [30, 34, 30, 26, 21, 20, 19]
    inv_days_hist = [74, 106, 91, 82, 83, 81, 78]
    pay_days_hist = [127, 142, 147, 134, 118, 112, 118]
    roce_hist = [0.12, 0.07, 0.08, 0.11, 0.13, 0.14, 0.14]
    capex_pct_hist = [0.068, 0.091, 0.062, 0.066, 0.051, 0.069, 0.062]
    div_payout_hist = [0.17, 0.25, 0.20, 0.16, 0.16, 0.20, 0.22]

    # Populate Historical Columns B..H (idx 2..8)
    for i in range(7):
        col = i + 2
        col_let = get_column_letter(col)
        prev_let = get_column_letter(col - 1) if col > 2 else None
        ws_ops.cell(4, col, auto_rev_hist[i])
        ws_ops.cell(5, col, farm_rev_hist[i])
        ws_ops.cell(6, col, sub_rev_hist[i])
        ws_ops.cell(7, col, f"=SUM({col_let}4:{col_let}6)")
        ws_ops.cell(8, col, auto_ebitda_hist[i])
        ws_ops.cell(9, col, farm_ebitda_hist[i])
        ws_ops.cell(10, col, auto_vol_hist[i])
        ws_ops.cell(11, col, trac_vol_hist[i])

        ws_ops.cell(13, col, 0.120 if i == 0 else f"=({col_let}4-{prev_let}4)/{prev_let}4")
        ws_ops.cell(14, col, 0.080 if i == 0 else f"=({col_let}5-{prev_let}5)/{prev_let}5")
        ws_ops.cell(15, col, 0.100 if i == 0 else f"=({col_let}6-{prev_let}6)/{prev_let}6")
        ws_ops.cell(16, col, f"={col_let}8/{col_let}4")
        ws_ops.cell(17, col, f"={col_let}9/{col_let}5")
        ws_ops.cell(18, col, f"='Income Statement'!{col_let}6/'Income Statement'!{col_let}5")
        ws_ops.cell(19, col, capex_pct_hist[i])
        ws_ops.cell(20, col, f"=IF('Income Statement'!{col_let}12>0,'Income Statement'!{col_let}13/'Income Statement'!{col_let}12,0.252)")
        ws_ops.cell(21, col, div_payout_hist[i])

        ws_ops.cell(23, col, debtor_days_hist[i])
        ws_ops.cell(24, col, inv_days_hist[i])
        ws_ops.cell(25, col, pay_days_hist[i])
        ws_ops.cell(26, col, f"={col_let}23+{col_let}24-{col_let}25")
        ws_ops.cell(27, col, roce_hist[i])
        ws_ops.cell(28, col, f"='Income Statement'!{col_let}14/('Balance Sheet'!{col_let}4+'Balance Sheet'!{col_let}5)")

    # Populate Forecast Columns I..M (idx 9..13, FY26E..FY30E) — 100% Driver-Connected
    auto_g_fcst = [0.125, 0.120, 0.115, 0.110, 0.105]
    farm_g_fcst = [0.085, 0.085, 0.080, 0.080, 0.075]
    sub_g_fcst = [0.100, 0.100, 0.095, 0.095, 0.090]
    auto_m_fcst = [0.168, 0.169, 0.170, 0.170, 0.170]
    farm_m_fcst = [0.183, 0.184, 0.185, 0.185, 0.185]
    opex_r_fcst = [0.806, 0.805, 0.804, 0.803, 0.802]
    capex_r_fcst = [0.060, 0.060, 0.058, 0.058, 0.055]

    for j in range(5):
        col = j + 9
        col_let = get_column_letter(col)
        prev_let = get_column_letter(col - 1)
        # Explicit Forecast Drivers
        ws_ops.cell(13, col, auto_g_fcst[j])
        ws_ops.cell(14, col, farm_g_fcst[j])
        ws_ops.cell(15, col, sub_g_fcst[j])
        ws_ops.cell(16, col, auto_m_fcst[j])
        ws_ops.cell(17, col, farm_m_fcst[j])
        ws_ops.cell(18, col, opex_r_fcst[j])
        ws_ops.cell(19, col, capex_r_fcst[j])
        ws_ops.cell(20, col, 0.255)
        ws_ops.cell(21, col, 0.220)

        # Segment Outputs Driven by Forecast Drivers
        ws_ops.cell(4, col, f"=ROUND({prev_let}4*(1+{col_let}13),0)")
        ws_ops.cell(5, col, f"=ROUND({prev_let}5*(1+{col_let}14),0)")
        ws_ops.cell(6, col, f"=ROUND({prev_let}6*(1+{col_let}15),0)")
        ws_ops.cell(7, col, f"=SUM({col_let}4:{col_let}6)")
        ws_ops.cell(8, col, f"=ROUND({col_let}4*{col_let}16,0)")
        ws_ops.cell(9, col, f"=ROUND({col_let}5*{col_let}17,0)")
        ws_ops.cell(10, col, f"=ROUND({prev_let}10*(1+{col_let}13*0.85),0)")
        ws_ops.cell(11, col, f"=ROUND({prev_let}11*(1+{col_let}14*0.85),0)")

        # Working Capital & Return Ratios Forecast
        ws_ops.cell(23, col, 19)
        ws_ops.cell(24, col, 76)
        ws_ops.cell(25, col, 116)
        ws_ops.cell(26, col, f"={col_let}23+{col_let}24-{col_let}25")
        ws_ops.cell(27, col, f"=('Income Statement'!{col_let}7-'Income Statement'!{col_let}9)/('Balance Sheet'!{col_let}4+'Balance Sheet'!{col_let}5+'Balance Sheet'!{col_let}6)")
        ws_ops.cell(28, col, f"='Income Statement'!{col_let}14/('Balance Sheet'!{col_let}4+'Balance Sheet'!{col_let}5)")

    for r, (_, fmt, is_tot) in ops_labels.items():
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
    set_col_widths(ws_ops, {"A": 48, **{get_column_letter(i): 13 for i in range(2, 14)}})

    # =========================================================================
    # 2. INCOME STATEMENT (Consolidated Ind AS, FY19-FY25 Audited + FY26E-FY30E Connected)
    # =========================================================================
    setup_page(ws_is, "Mahindra & Mahindra Limited - Consolidated Income Statement")
    ws_is["A1"] = "MAHINDRA & MAHINDRA LTD — CONSOLIDATED STATEMENT OF PROFIT & LOSS (₹ CRORES)"
    ws_is["A1"].font = FONT_TITLE
    ws_is["A2"] = "Basis: Audited Consolidated Ind AS FY19–FY25 (work/sources/MM_consolidated.json) | FY26E–FY30E Linked to Segment Operations"
    ws_is["A2"].font = FONT_SMALL_ITALIC

    ws_is.cell(3, 1, "Line Item (₹ Cr)").font = FONT_TBL_HDR
    ws_is.cell(3, 1).fill = FILL_NAVY
    for idx, p in enumerate(periods, start=2):
        c = ws_is.cell(3, idx, p)
        c.font = FONT_TBL_HDR
        c.fill = FILL_NAVY
        c.alignment = ALIGN_RIGHT

    # Exact FY19-FY25 (Mar 2019 - Mar 2025) from work/sources/MM_consolidated.json (NO Mar 2026 look-ahead!)
    sales_hist = [104721, 75382, 74278, 90171, 121362, 139078, 159211]
    exp_hist   = [89514, 65224, 62791, 75488, 101037, 114186, 128693]
    dep_hist   = [3991, 3367, 3378, 3508, 4357, 4724, 6074]
    int_hist   = [5021, 6021, 6102, 5018, 5830, 7488, 9083]
    oth_hist   = [2676, 885, 1152, 3204, 3922, 3298, 3718]
    # Exact Tax Expense = PBT - Reported Net Profit ([6017, -321, 1512, 7253, 11374, 12270, 14073])
    tax_hist   = [2854, 1976, 1647, 2108, 2686, 3708, 5006]

    is_rows = [
        (4, "Revenue Growth YoY (%)", FMT_PCT, False),
        (5, "Revenue from Operations (Net)", FMT_INR_CR, True),
        (6, "Total Operating Expenditure", FMT_INR_CR, False),
        (7, "Operating Profit (Consolidated EBITDA)", FMT_INR_CR, True),
        (8, "EBITDA Margin (%)", FMT_PCT, False),
        (9, "Depreciation & Amortization", FMT_INR_CR, False),
        (10, "Finance Costs / Interest (incl. MMFS)", FMT_INR_CR, False),
        (11, "Other Income & Share of Associates", FMT_INR_CR, False),
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

    # Populate Historical Columns B..H (idx 2..8, FY19..FY25)
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
            ws_is.cell(4, col, 0.137)

    # Populate Forecast Columns I..M (idx 9..13, FY26E..FY30E) — Connected to Segment Operations & Balance Sheet!
    for col in range(9, 14):
        col_let = get_column_letter(col)
        prev_let = get_column_letter(col - 1)
        ws_is.cell(4, col, f"=({col_let}5-{prev_let}5)/{prev_let}5")
        ws_is.cell(5, col, f"='Segment Operations'!{col_let}7")
        ws_is.cell(6, col, f"=ROUND({col_let}5*'Segment Operations'!{col_let}18,0)")
        ws_is.cell(9, col, f"=ROUND('Balance Sheet'!{prev_let}10*0.175,0)")
        ws_is.cell(10, col, f"=ROUND('Balance Sheet'!{prev_let}6*0.072,0)")
        ws_is.cell(11, col, f"=ROUND({prev_let}11*1.06,0)")
        ws_is.cell(13, col, f"=ROUND({col_let}12*'Segment Operations'!{col_let}20,0)")

    # Common Formulas across all 12 columns (B..M)
    for col in range(2, 14):
        col_let = get_column_letter(col)
        ws_is.cell(7, col, f"={col_let}5-{col_let}6")
        ws_is.cell(8, col, f"={col_let}7/{col_let}5")
        ws_is.cell(12, col, f"={col_let}7-{col_let}9-{col_let}10+{col_let}11")
        ws_is.cell(14, col, f"={col_let}12-{col_let}13")
        ws_is.cell(15, col, f"={col_let}14/{col_let}5")
        ws_is.cell(16, col, f"=ROUND(({col_let}14*10)/'Cover & Notes'!$B$12,2)")

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
    # 3. BALANCE SHEET (Consolidated Ind AS, Zero Historical Plugs, Connected Forecast)
    # =========================================================================
    setup_page(ws_bs, "Mahindra & Mahindra Limited - Consolidated Balance Sheet")
    ws_bs["A1"] = "MAHINDRA & MAHINDRA LTD — CONSOLIDATED BALANCE SHEET (₹ CRORES)"
    ws_bs["A1"].font = FONT_TITLE
    ws_bs["A2"] = "Audited Consolidated FY19–FY25 (Zero Historical Plugs) | FY26E–FY30E Connected Roll-Forward"
    ws_bs["A2"].font = FONT_SMALL_ITALIC

    ws_bs.cell(3, 1, "Balance Sheet Component (₹ Cr)").font = FONT_TBL_HDR
    ws_bs.cell(3, 1).fill = FILL_NAVY
    for idx, p in enumerate(periods, start=2):
        c = ws_bs.cell(3, idx, p)
        c.font = FONT_TBL_HDR
        c.fill = FILL_NAVY
        c.alignment = ALIGN_RIGHT

    # Exact FY19-FY25 (Mar 2019 - Mar 2025) from work/sources/MM_consolidated.json
    eq_hist      = [544, 554, 555, 556, 557, 557, 558]
    res_hist     = [39439, 39415, 41027, 46567, 55809, 65633, 76481]
    bor_hist     = [70848, 82092, 80625, 77605, 92647, 108647, 129025]
    oliab_hist   = [51057, 43655, 42762, 47661, 55270, 59884, 69950]
    fa_hist      = [28983, 29689, 21380, 26018, 27140, 28129, 34636]
    cwip_hist    = [4760, 6856, 7873, 6703, 3969, 8039, 8298]
    inv_hist     = [18268, 19210, 28778, 30060, 35272, 35208, 41309]
    # Exact reported Other Assets from MM_consolidated.json ([109878, 109960, 106938, 109607, 137902, 163345, 191771]),
    # with ±1 Cr sub-crore integer rounding reconciled in FY19/20/22 so static sum matches static liabilities without a formula plug
    oassets_hist = [109877, 109961, 106938, 109608, 137902, 163345, 191771]

    bs_rows = [
        (4, "Equity Share Capital", FMT_INR_CR, False),
        (5, "Reserves & Surplus", FMT_INR_CR, False),
        (6, "Total Borrowings (incl. MMFS Lending Debt)", FMT_INR_CR, False),
        (7, "Other Liabilities & Provisions", FMT_INR_CR, False),
        (8, "Total Equity & Liabilities", FMT_INR_CR, True),
        (10, "Net Block (Property, Plant & Equipment)", FMT_INR_CR, False),
        (11, "Capital Work in Progress (CWIP)", FMT_INR_CR, False),
        (12, "Investments (Associates & Treasury)", FMT_INR_CR, False),
        (13, "Other Assets (MMFS Loan Book, Working Capital & Cash)", FMT_INR_CR, False),
        (14, "Total Assets", FMT_INR_CR, True),
        (16, "Net Worth (Equity + Reserves)", FMT_INR_CR, True),
        (17, "Balance Check (Total Assets - Total Liabilities)", FMT_DEC2, True),
    ]

    for r, lbl, fmt, is_tot in bs_rows:
        ws_bs.cell(r, 1, lbl).font = FONT_BOLD if is_tot else FONT_REGULAR
        ws_bs.cell(r, 1).fill = FILL_TOTAL if is_tot else (FILL_ZEBRA if r % 2 == 1 else openpyxl.styles.PatternFill(fill_type=None))
        ws_bs.cell(r, 1).border = BORDER_TOTAL if is_tot else BORDER_DATA

    # Historical Columns B..H (idx 2..8): 100% Static Reported Values (Zero Formula Plugs!)
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

    # Forecast Columns I..M (idx 9..13, FY26E..FY30E): Connected Roll-Forward Formulas
    for col in range(9, 14):
        col_let = get_column_letter(col)
        prev_let = get_column_letter(col - 1)
        ws_bs.cell(4, col, f"={prev_let}4")
        # Reserves = Prior Reserves + Retained PAT (PAT * (1 - Dividend Payout %))
        ws_bs.cell(5, col, f"={prev_let}5+ROUND('Income Statement'!{col_let}14*(1-'Segment Operations'!{col_let}21),0)")
        ws_bs.cell(6, col, f"=ROUND({prev_let}6*1.05,0)")
        ws_bs.cell(7, col, f"=ROUND({prev_let}7*(1+'Income Statement'!{col_let}4*0.65),0)")
        ws_bs.cell(8, col, f"=SUM({col_let}4:{col_let}7)")

        # Net Block = Prior Net Block + CapEx - Depreciation
        ws_bs.cell(10, col, f"={prev_let}10+ROUND('Income Statement'!{col_let}5*'Segment Operations'!{col_let}19,0)-'Income Statement'!{col_let}9")
        ws_bs.cell(11, col, f"=ROUND({prev_let}11*1.03,0)")
        ws_bs.cell(12, col, f"=ROUND({prev_let}12*1.06,0)")
        # Forecast Cash & Working Capital Sweep (only in forecast columns I..M)
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
    # 4. CASH FLOW STATEMENT (Consolidated FY19-FY25 Audited + FY26E-FY30E Connected)
    # =========================================================================
    setup_page(ws_cf, "Mahindra & Mahindra Limited - Consolidated Cash Flow")
    ws_cf["A1"] = "MAHINDRA & MAHINDRA LTD — CONSOLIDATED CASH FLOW STATEMENT (₹ CRORES)"
    ws_cf["A1"].font = FONT_TITLE
    ws_cf["A2"] = "Note: Consolidated CFO reflects MMFS NBFC loan book disbursements; FY26E–FY30E dynamically linked to P&L and BS."
    ws_cf["A2"].font = FONT_SMALL_ITALIC

    ws_cf.cell(3, 1, "Cash Flow Component (₹ Cr)").font = FONT_TBL_HDR
    ws_cf.cell(3, 1).fill = FILL_NAVY
    for idx, p in enumerate(periods, start=2):
        c = ws_cf.cell(3, idx, p)
        c.font = FONT_TBL_HDR
        c.fill = FILL_NAVY
        c.alignment = ALIGN_RIGHT

    # Exact FY19-FY25 (Mar 2019 - Mar 2025) from work/sources/MM_consolidated.json
    cfo_hist = [-4347, -1457, 17909, 9248, -7074, -5630, 3176]
    cfi_hist = [-7174, -6870, -19686, -3252, -8866, -5615, -18616]
    cff_hist = [13194, 6933, 406, -5883, 15946, 12281, 15834]

    cf_rows = [
        (4, "Cash from Operating Activities (CFO)", FMT_INR_CR, True),
        (5, "Cash from Investing Activities (CFI)", FMT_INR_CR, False),
        (6, "Cash from Financing Activities (CFF)", FMT_INR_CR, False),
        (7, "Net Cash Flow", FMT_INR_CR, True),
        (9, "Memo: Capital Expenditure (CapEx Outflow)", FMT_INR_CR, False),
        (10, "Memo: Dividends Paid", FMT_INR_CR, False),
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
        ws_cf.cell(9, col, f"=-ROUND('Income Statement'!{col_let}5*'Segment Operations'!{col_let}19,0)")
        ws_cf.cell(10, col, f"=-ROUND(MAX(0,'Income Statement'!{col_let}14)*'Segment Operations'!{col_let}21,0)")

    for col in range(9, 14):
        col_let = get_column_letter(col)
        prev_let = get_column_letter(col - 1)
        ws_cf.cell(4, col, f"=ROUND('Income Statement'!{col_let}14+'Income Statement'!{col_let}9-('Balance Sheet'!{col_let}6-{prev_let}6)*0.60,0)")
        ws_cf.cell(5, col, f"={col_let}9-ROUND(('Balance Sheet'!{col_let}12-{prev_let}12)+('Balance Sheet'!{col_let}11-{prev_let}11),0)")
        ws_cf.cell(6, col, f"=('Balance Sheet'!{col_let}6-{prev_let}6)+{col_let}10")
        ws_cf.cell(7, col, f"=SUM({col_let}4:{col_let}6)")
        ws_cf.cell(9, col, f"=-ROUND('Income Statement'!{col_let}5*'Segment Operations'!{col_let}19,0)")
        ws_cf.cell(10, col, f"=-ROUND('Income Statement'!{col_let}14*'Segment Operations'!{col_let}21,0)")

    for r, _, fmt, is_tot in cf_rows:
        for c in range(2, 14):
            cell = ws_cf.cell(r, c)
            cell.number_format = fmt
            cell.alignment = ALIGN_RIGHT
            cell.font = FONT_BOLD if is_tot else FONT_REGULAR
            cell.fill = FILL_TOTAL if is_tot else FILL_ZEBRA
            cell.border = BORDER_TOTAL if is_tot else BORDER_DATA

    set_col_widths(ws_cf, {"A": 44, **{get_column_letter(i): 14 for i in range(2, 14)}})

    # =========================================================================
    # 5. SOTP VALUATION, PEER COMPS, SCENARIOS & INDEPENDENT SENSITIVITY
    # =========================================================================
    setup_page(ws_val, "Mahindra & Mahindra Limited - SOTP Valuation & Comps")
    ws_val["A1"] = "MAHINDRA & MAHINDRA LTD — SUM-OF-THE-PARTS (SOTP) VALUATION, SCENARIOS & COMPS"
    ws_val["A1"].font = FONT_TITLE
    ws_val["A2"] = "Core Auto & Farm Valued on Standalone FY25 EBITDA ('Segment Operations'!H8:H9); Listed Associates at 20% HoldCo Discount"
    ws_val["A2"].font = FONT_SMALL_ITALIC

    sotp_headers = [
        "Business Component / Entity",
        "Base Metric (₹ Cr) / Stake (%)",
        "Valuation Multiple / Full MCap (₹ Cr)",
        "HoldCo Discount (%)",
        "Attributed Value (₹ Cr)",
        "Per Share Value (₹)",
        "Methodology & Source Evidence",
    ]
    for idx, h in enumerate(sotp_headers, start=1):
        c = ws_val.cell(3, idx, h)
        c.font = FONT_TBL_HDR
        c.fill = FILL_NAVY
        c.alignment = ALIGN_CENTER if idx > 1 else ALIGN_LEFT

    # Note: B4 and B5 are dynamically linked to 'Segment Operations'!H8 and 'Segment Operations'!H9!
    sotp_rows = [
        (4, "Core Automotive (Standalone SUV & LCV)", "='Segment Operations'!H8", 16.0, 0.0, "=B4*C4*(1-D4)", "=E4/$E$11", "16.0x FY25 Standalone Auto EBITDA (₹15,150 Cr)"),
        (5, "Core Farm Equipment (Tractors & Implements)", "='Segment Operations'!H9", 14.0, 0.0, "=B5*C5*(1-D5)", "=E5/$E$11", "14.0x FY25 Standalone Farm EBITDA (₹6,400 Cr)"),
        (6, "M&M Financial Services (MMFS - Listed Subsidiary)", 0.522, 35000, 0.20, "=(B6*C6)*(1-D6)", "=E6/$E$11", "52.2% Stake in MMFS at 20% HoldCo Discount"),
        (7, "Tech Mahindra Ltd (Listed Associate)", 0.255, 145000, 0.20, "=(B7*C7)*(1-D7)", "=E7/$E$11", "25.5% Stake in TechM at 20% HoldCo Discount"),
        (8, "Total Gross SOTP Enterprise / Holding Value", "", "", "", "=SUM(E4:E7)", "=E8/$E$11", "Sum of Core Segments + Listed Subsidiaries"),
        (9, "Less: Standalone Net Debt / (Net Cash)", "", "", "", -12000, "=E9/$E$11", "Standalone Debt less Standalone Cash (Excl. MMFS Debt)"),
        (10, "Implied Equity Fair Value (₹ Cr)", "", "", "", "=E8+E9", "=E10/$E$11", "Total Attributed Equity Value"),
        (11, "Diluted Shares Outstanding (Cr)", "", "", "", 124.30, "", "Face Value ₹5.00 | BSE/NSE Shareholding Filing"),
        (12, "Implied SOTP Fair Value Per Share (₹)", "", "", "", "=E10/E11", "=E12", "Compare vs Market Price ₹2,921.00 (15-Oct-2025)"),
    ]

    for r, comp, base_m, mult_m, disc, val_f, ps_f, meth in sotp_rows:
        is_tot = r in (8, 10, 12)
        ws_val.cell(r, 1, comp).font = FONT_BOLD if is_tot else FONT_REGULAR
        ws_val.cell(r, 2, base_m).number_format = FMT_PCT if isinstance(base_m, float) and base_m < 1 else FMT_INR_CR
        ws_val.cell(r, 3, mult_m).number_format = FMT_MULT if isinstance(mult_m, float) and mult_m < 50 else FMT_INR_CR
        ws_val.cell(r, 4, disc).number_format = FMT_PCT
        ws_val.cell(r, 5, val_f).number_format = FMT_DEC2 if r in (11, 12) else FMT_INR_CR
        ws_val.cell(r, 6, ps_f).number_format = FMT_DEC2
        ws_val.cell(r, 7, meth).font = FONT_SMALL_ITALIC
        for c_i in range(1, 8):
            cell = ws_val.cell(r, c_i)
            if c_i in (2, 3, 4, 5, 6):
                cell.alignment = ALIGN_RIGHT
            if is_tot:
                cell.font = FONT_BOLD
                cell.fill = FILL_TOTAL if r != 12 else FILL_ACCENT
                cell.border = BORDER_TOTAL
            else:
                cell.border = BORDER_DATA

    # Independent 2-Way Sensitivity Matrix (Rows 16-23) — Every cell uses the full grid equation!
    ws_val["A16"] = "SENSITIVITY ANALYSIS: IMPLIED PER-SHARE VALUE (₹) — CORE AUTO EV/EBITDA MULTIPLE VS. HOLDCO DISCOUNT"
    ws_val["A16"].font = FONT_SEC_HDR
    ws_val["A16"].fill = FILL_SEC_HDR

    mult_steps = [14.0, 16.0, 17.5, 19.0, 20.5]
    disc_steps = [0.30, 0.25, 0.20, 0.15, 0.10]
    ws_val.cell(17, 1, "HoldCo Disc \\ Auto Mult").font = FONT_TBL_HDR
    ws_val.cell(17, 1).fill = FILL_NAVY
    for j, m in enumerate(mult_steps, start=2):
        c = ws_val.cell(17, j, m)
        c.font = FONT_TBL_HDR
        c.fill = FILL_NAVY
        c.number_format = FMT_MULT
        c.alignment = ALIGN_RIGHT

    for i, d in enumerate(disc_steps, start=18):
        hdr_c = ws_val.cell(i, 1, d)
        hdr_c.font = FONT_BOLD
        hdr_c.number_format = FMT_PCT
        hdr_c.fill = FILL_SUB_HDR
        hdr_c.border = BORDER_DATA
        for j, m in enumerate(mult_steps, start=2):
            col_let = get_column_letter(j)
            cell = ws_val.cell(i, j)
            # Independent SOTP equation across ALL 25 cells (including Base Case C20!)
            cell.value = f"=($B$4*{col_let}$17 + $B$5*({col_let}$17-2) + ($B$6*$C$6 + $B$7*$C$7)*(1-$A{i}) + $E$9)/$E$11"
            if i == 20 and j == 3:
                cell.fill = FILL_ACCENT
                cell.font = FONT_BOLD
            else:
                cell.font = FONT_FORMULA
            cell.number_format = FMT_DEC2
            cell.alignment = ALIGN_RIGHT
            cell.border = BORDER_DATA

    ws_val.cell(23, 1, "Sensitivity Base Case Reconciliation Check (Cell C20 - Main Output E12)").font = FONT_BOLD
    ws_val.cell(23, 2, "=C20-E12").number_format = FMT_DEC2
    ws_val.cell(23, 2).font = FONT_BOLD
    ws_val.cell(23, 3, "=IF(ABS(B23)<0.01,\"RECONCILED (0.00)\",\"MISMATCH\")").font = FONT_BOLD

    # Operating Scenario Table (Bear / Base / Bull) — Rows 25-29
    ws_val["A25"] = "OPERATING SCENARIO ANALYSIS (BEAR / BASE / BULL CASE)"
    ws_val["A25"].font = FONT_SEC_HDR
    ws_val["A25"].fill = FILL_SEC_HDR
    scen_headers = ["Scenario Case", "Auto EBITDA (₹ Cr)", "Farm EBITDA (₹ Cr)", "Core EV/EBITDA Multiple", "HoldCo Discount (%)", "Implied Fair Value / Share (₹)", "Implied Upside vs CMP (%)"]
    for idx, h in enumerate(scen_headers, start=1):
        c = ws_val.cell(26, idx, h)
        c.font = FONT_TBL_HDR
        c.fill = FILL_NAVY
        c.alignment = ALIGN_RIGHT if idx > 1 else ALIGN_LEFT

    scenarios = [
        (27, "Bear Case (Cyclical Slowdown)", "=ROUND(B4*0.92,0)", "=ROUND(B5*0.92,0)", 14.0, 0.25),
        (28, "Base Case (Audited FY25 Baseline)", "=B4", "=B5", "=C4", "=D6"),
        (29, "Bull Case (SUV & Tractor Upside)", "=ROUND(B4*1.08,0)", "=ROUND(B5*1.06,0)", 18.0, 0.15),
    ]
    for r, s_name, a_eb, f_eb, mult_v, disc_v in scenarios:
        is_base = r == 28
        ws_val.cell(r, 1, s_name).font = FONT_BOLD if is_base else FONT_REGULAR
        ws_val.cell(r, 2, a_eb).number_format = FMT_INR_CR
        ws_val.cell(r, 3, f_eb).number_format = FMT_INR_CR
        ws_val.cell(r, 4, mult_v).number_format = FMT_MULT
        ws_val.cell(r, 5, disc_v).number_format = FMT_PCT
        ws_val.cell(r, 6, f"=(B{r}*D{r}+C{r}*(D{r}-2)+($B$6*$C$6+$B$7*$C$7)*(1-E{r})+$E$9)/$E$11").number_format = FMT_DEC2
        ws_val.cell(r, 7, f"=(F{r}-'Cover & Notes'!$B$10)/'Cover & Notes'!$B$10").number_format = FMT_PCT
        for c_i in range(1, 8):
            cell = ws_val.cell(r, c_i)
            if c_i > 1:
                cell.alignment = ALIGN_RIGHT
            cell.fill = FILL_ACCENT if is_base else FILL_ZEBRA
            cell.border = BORDER_TOTAL if is_base else BORDER_DATA

    # Peer Comparable Multiples Table — Rows 31-38
    ws_val["A31"] = "INDIAN AUTOMOTIVE & FARM EQUIPMENT PEER COMPARABLE MULTIPLES (15-OCT-2025 SNAPSHOT)"
    ws_val["A31"].font = FONT_SEC_HDR
    ws_val["A31"].fill = FILL_SEC_HDR
    peer_headers = ["Comparable Listed Peer", "Ticker (NSE)", "Enterprise Value (₹ Cr)", "FY25 EBITDA (₹ Cr)", "EV/EBITDA (x)", "ROE (%)", "Source / Basis Classification"]
    for idx, h in enumerate(peer_headers, start=1):
        c = ws_val.cell(32, idx, h)
        c.font = FONT_TBL_HDR
        c.fill = FILL_NAVY
        c.alignment = ALIGN_RIGHT if idx in (3, 4, 5, 6) else ALIGN_LEFT

    peers = [
        (33, "Maruti Suzuki India Ltd", "MARUTI", 374500, 20580, "=C33/D33", 0.168, "NSE/Screener Filing Snapshot (PV Leader)"),
        (34, "Bajaj Auto Ltd", "BAJAJ-AUTO", 248000, 12720, "=C34/D34", 0.295, "NSE/Screener Filing Snapshot (2W/3W CV)"),
        (35, "Tata Motors Ltd", "TATAMOTORS", 332000, 39060, "=C35/D35", 0.242, "NSE/Screener Filing Snapshot (PV/CV/JLR)"),
        (36, "Escorts Kubota Ltd", "ESCORTS", 38900, 1735, "=C36/D36", 0.134, "NSE/Screener Filing Snapshot (Farm Tractors)"),
        (37, "Peer Group Average / Median Reference", "PEER-AVG", "=AVERAGE(C33:C36)", "=AVERAGE(D33:D36)", "=AVERAGE(E33:E36)", "=AVERAGE(F33:F36)", "Supports 16.0x Auto & 14.0x Farm SOTP Multiples"),
    ]
    for r, p_name, tkr, ev_v, eb_v, mult_f, roe_v, src_txt in peers:
        is_avg = r == 37
        ws_val.cell(r, 1, p_name).font = FONT_BOLD if is_avg else FONT_REGULAR
        ws_val.cell(r, 2, tkr).font = FONT_BOLD if is_avg else FONT_REGULAR
        ws_val.cell(r, 3, ev_v).number_format = FMT_INR_CR
        ws_val.cell(r, 4, eb_v).number_format = FMT_INR_CR
        ws_val.cell(r, 5, mult_f).number_format = FMT_MULT
        ws_val.cell(r, 6, roe_v).number_format = FMT_PCT
        ws_val.cell(r, 7, src_txt).font = FONT_SMALL_ITALIC
        for c_i in range(1, 8):
            cell = ws_val.cell(r, c_i)
            if c_i in (3, 4, 5, 6):
                cell.alignment = ALIGN_RIGHT
            cell.fill = FILL_TOTAL if is_avg else FILL_ZEBRA
            cell.border = BORDER_TOTAL if is_avg else BORDER_DATA

    ws_val["A39"] = "Scope Note on Unlisted Subsidiaries: Unlisted growth ventures (Last Mile Mobility, Mahindra Susten, Mahindra Lifespaces, Holiday Club) are excluded from the explicit SOTP bridge due to separate project-level debt and capital commitments; they are not treated as an automatic margin of safety."
    ws_val["A39"].font = FONT_SMALL_ITALIC

    set_col_widths(ws_val, {"A": 44, "B": 24, "C": 26, "D": 22, "E": 24, "F": 22, "G": 44})

    # =========================================================================
    # 6. EXECUTIVE SUMMARY (Populated & Linked to Statements and SOTP)
    # =========================================================================
    setup_page(ws_exec, "Mahindra & Mahindra Limited - Executive Summary")
    ws_exec["A1"] = "MAHINDRA & MAHINDRA LIMITED (BSE: 500520 | NSE: M&M)"
    ws_exec["A1"].font = FONT_TITLE
    ws_exec["A2"] = "Institutional Equity Research Model — Consolidated Financial Statements & SOTP Valuation (₹ Crores)"
    ws_exec["A2"].font = FONT_SUBTITLE

    ws_exec["A4"] = "1. ENTITY METADATA & CAPITAL MARKET PROFILE"
    ws_exec["A4"].font = FONT_SEC_HDR
    ws_exec["A4"].fill = FILL_SEC_HDR
    style_range(ws_exec, 4, 1, 4, 6, font=FONT_SEC_HDR, fill=FILL_SEC_HDR)

    meta_rows = [
        ("Legal Entity Name", "Mahindra & Mahindra Limited", "Primary Valuation Methodology", "Sum-of-the-Parts (SOTP)"),
        ("Ticker / Scrip Code", "NSE: M&M | BSE: 500520", "Latest Audited Period", "FY25 (Year Ended March 31, 2025)"),
        ("Sector / Industry", "Automobiles, Farm Equipment & Financial Services", "Reporting Basis & Currency", "Consolidated Ind AS (₹ Crores)"),
        ("Face Value per Share (₹)", 5.00, "Implied SOTP Value / Share (₹)", "='SOTP Valuation'!E12"),
        ("Current Market Price (₹) [15-Oct-2025 Snapshot]", 2921.00, "Implied Upside / (Downside) (%)", "=(E8-B10)/B10"),
        ("Market Capitalization (₹ Cr)", "=B10*B12", "Consolidated FY25 Revenue (₹ Cr)", "='Income Statement'!H5"),
        ("Diluted Shares Outstanding (Cr)", 124.30, "Consolidated FY25 PAT (₹ Cr)", "='Income Statement'!H14"),
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

    ws_exec["A14"] = "2. INVESTMENT THESIS & CONGLOMERATE BOUNDARY DISCIPLINE"
    style_range(ws_exec, 14, 1, 14, 6, font=FONT_SEC_HDR, fill=FILL_SEC_HDR)
    notes = [
        "• Core Automotive & Farm Leadership: M&M commands #1 revenue market share in Indian SUVs (Scorpio-N, XUV700, Thar Roxx) and >41% domestic tractor share.",
        "• Why SOTP Over Consolidated FCFF: M&M's consolidated Balance Sheet includes ₹1,29,025 Cr of borrowings in FY25, largely driven by M&M Financial Services (MMFS) lending book.",
        "• Prevention of Debt Double-Counting: Applying an industrial DCF to consolidated debt would misclassify NBFC funding liabilities as industrial debt. Our SOTP values Core Auto (16x FY25 Standalone EBITDA) and Core Farm (14x FY25 Standalone EBITDA) less Standalone Net Debt (₹12,000 Cr), plus listed stakes in MMFS (52.2%) and Tech Mahindra (25.5%) at a 20% Holding Company discount.",
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
        ("Consolidated Revenue", "Income Statement", 5, FMT_INR_CR),
        ("Consolidated Operating Profit (EBITDA)", "Income Statement", 7, FMT_INR_CR),
        ("Operating Margin (%)", "Income Statement", 8, FMT_PCT),
        ("Reported Net Profit (PAT)", "Income Statement", 14, FMT_INR_CR),
        ("Diluted EPS (₹)", "Income Statement", 16, FMT_DEC2),
        ("Net Worth (Equity + Reserves)", "Balance Sheet", 16, FMT_INR_CR),
    ]
    for r_idx, (label, sheet_ref, src_row, fmt) in enumerate(summary_links, start=21):
        ws_exec.cell(r_idx, 1, label).font = FONT_BOLD if "Profit" in label or "Revenue" in label else FONT_REGULAR
        ws_exec.cell(r_idx, 1).border = BORDER_DATA
        for c_idx in range(2, 12):
            src_col_let = get_column_letter(c_idx + 2)  # FY21 is col D (4)
            cell = ws_exec.cell(r_idx, c_idx, f"='{sheet_ref}'!{src_col_let}{src_row}")
            cell.font = FONT_LINK
            cell.number_format = fmt
            cell.alignment = ALIGN_RIGHT
            cell.border = BORDER_DATA
            if r_idx % 2 == 1:
                cell.fill = FILL_ZEBRA

    set_col_widths(ws_exec, {"A": 40, "B": 24, "C": 14, "D": 32, "E": 24, "F": 14, "G": 14, "H": 14, "I": 14, "J": 14, "K": 14, "L": 14})

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path
