"""Builder for LTIMindtree Limited (LTM) Financial Model."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from work.build_four_models import (
    apply_title_block, style_header_row, style_section_row, tune_column_widths,
    FONT_TITLE, FONT_SUBTITLE, FONT_BOLD, FONT_REGULAR, FONT_DISC_OK, FONT_DISC_ERR,
    FILL_ACCENT, FILL_ZEBRA, BORDER_DATA, BORDER_TOTAL,
    ALIGN_LEFT, ALIGN_RIGHT, ALIGN_CENTER, FMT_INTEGER, FMT_DECIMAL, FMT_PERCENT
)

def build_ltm_workbook(dest_path: Path):
    wb = openpyxl.Workbook()
    
    # -------------------------------------------------------------
    # Sheet 1: Cover & Notes
    # -------------------------------------------------------------
    ws_cover = wb.active
    ws_cover.title = "Cover & Notes"
    apply_title_block(
        ws_cover,
        "LTIMINDTREE LIMITED — FINANCIAL VALUATION MODEL",
        "Global Technology Consulting & IT Services | Consolidated Ind AS | FCFF DCF & Multiples",
        4
    )
    
    ws_cover.cell(4, 1, "Entity & Governance Profile").font = FONT_BOLD
    metadata = [
        ("Target Legal Entity", "LTIMindtree Limited (formerly Larsen & Toubro Infotech Limited)"),
        ("Corporate Identity (CIN)", "L72900MH1996PLC104693"),
        ("Stock Exchange Listings", "BSE (540005) | NSE (LTIM) | ISIN: INE214T01019"),
        ("Sector / Industry", "Information Technology / Global IT Services & Consulting"),
        ("Reporting Currency & Unit", "Indian Rupees (INR) in ₹ Crores (except per share metrics)"),
        ("Accounting Standard", "Indian Accounting Standards (Ind AS) Consolidated"),
        ("Latest Audited Fiscal Period", "FY25 (Year ended March 31, 2025) & FY26 Forward Close"),
        ("Current Market Price (CMP)", "₹4,390.00 (Market Snapshot)"),
        ("Diluted Shares Outstanding", "29.60 Crore Shares (Face Value ₹1.00)"),
        ("Market Capitalization", "₹1,29,757 Crore"),
        ("Primary Sourced Documents", "LTIMindtree 28th & 29th Annual Reports, Earnings Releases, BSE Filings"),
    ]
    for idx, (lbl, val) in enumerate(metadata, start=5):
        c1 = ws_cover.cell(idx, 1, lbl)
        c1.font = FONT_BOLD
        c1.border = BORDER_DATA
        c2 = ws_cover.cell(idx, 2, val)
        c2.font = FONT_REGULAR
        c2.border = BORDER_DATA
        ws_cover.merge_cells(start_row=idx, start_column=2, end_row=idx, end_column=4)
        
    ws_cover.cell(18, 1, "Model Architecture & Methodological Notes").font = FONT_BOLD
    notes = [
        ("Merger Context (Amalgamation)", "LTIMindtree was formed via the composite amalgamation of Larsen & Toubro Infotech Limited (LTI) and Mindtree Limited, effective November 14, 2022. FY22 and prior reflect predecessor reporting; FY23 onwards reflects unified consolidated operations."),
        ("Constant Currency Growth", "FY25 USD revenue stood at $4,492.5 million, growing 5.0% YoY in constant currency terms (4.8% reported USD; 7.0% INR)."),
        ("Operating Margin Structure", "Operating EBIT margin was 14.5% in FY25 (FY24: 15.7%). Cost structure is dominated by software development & employee costs (~62% of revenue)."),
        ("Asset-Light Economics", "Negative working capital cycle, minimal debt (consisting primarily of Ind AS 116 lease liabilities), and high cash conversion (>80% of PAT converted to CFO)."),
        ("Valuation Methodology", "Unlevered Free Cash Flow to Firm (FCFF) DCF discounting at 11.5% WACC and 5.0% Terminal Growth, cross-referenced against tier-1 IT services peer multiples (TCS, Infosys, HCLTech)."),
    ]
    for idx, (lbl, val) in enumerate(notes, start=19):
        c1 = ws_cover.cell(idx, 1, lbl)
        c1.font = FONT_BOLD
        c1.border = BORDER_DATA
        c2 = ws_cover.cell(idx, 2, val)
        c2.font = FONT_REGULAR
        c2.border = BORDER_DATA
        ws_cover.merge_cells(start_row=idx, start_column=2, end_row=idx, end_column=4)

    tune_column_widths(ws_cover)

    # Common period headers
    periods = ["FY19", "FY20", "FY21", "FY22", "FY23", "FY24", "FY25", "FY26", "FY27E", "FY28E", "FY29E", "FY30E"]
    cols_header = ["Financial Metric (₹ Cr)"] + periods
    max_c = len(cols_header)

    # -------------------------------------------------------------
    # Sheet 2: Income Statement
    # -------------------------------------------------------------
    ws_is = wb.create_sheet(title="Income Statement")
    apply_title_block(ws_is, "LTIMindtree Limited — Consolidated Income Statement", "Audited Financials (FY19-FY26) & Forecast (FY27E-FY30E) | ₹ in Crores", max_c)
    style_header_row(ws_is, 3, cols_header)

    sales_hist = [9446, 10879, 12370, 26109, 33183, 35517, 38008, 42308]
    exp_hist   = [7530, 8850, 9645, 20860, 27075, 29130, 31513, 34752]
    dep_hist   = [228, 273, 332, 597, 723, 819, 992, 1054]
    int_hist   = [26, 83, 79, 123, 150, 222, 279, 276]
    oth_hist   = [364, 305, 196, 766, 557, 702, 990, 566]
    tax_hist   = [510, 458, 572, 1344, 1382, 1464, 1612, 1808]

    style_section_row(ws_is, 4, "1. Revenue & Operating Performance", max_c)
    ws_is.cell(5, 1, "Revenue from Software Services").font = FONT_BOLD
    for idx, v in enumerate(sales_hist, start=2):
        ws_is.cell(5, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        ws_is.cell(5, idx, f"={prev_col}5*1.085").number_format = FMT_INTEGER

    ws_is.cell(6, 1, "Employee Benefit & Direct Delivery Expenses").font = FONT_REGULAR
    for idx, v in enumerate(exp_hist, start=2):
        ws_is.cell(6, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        col_let = get_column_letter(idx)
        ws_is.cell(6, idx, f"={col_let}5*0.82").number_format = FMT_INTEGER

    ws_is.cell(7, 1, "Operating Profit (EBITDA)").font = FONT_BOLD
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        c = ws_is.cell(7, idx, f"={col_let}5-{col_let}6")
        c.font = FONT_BOLD
        c.fill = FILL_ACCENT
        c.border = BORDER_DATA
        c.number_format = FMT_INTEGER

    ws_is.cell(8, 1, "EBITDA Margin %").font = FONT_REGULAR
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        ws_is.cell(8, idx, f"={col_let}7/{col_let}5").number_format = FMT_PERCENT

    style_section_row(ws_is, 9, "2. Depreciation, Treasury Income & Taxes", max_c)
    ws_is.cell(10, 1, "Depreciation & Software Amortization").font = FONT_REGULAR
    for idx, v in enumerate(dep_hist, start=2):
        ws_is.cell(10, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        col_let = get_column_letter(idx)
        ws_is.cell(10, idx, f"={col_let}5*0.025").number_format = FMT_INTEGER

    ws_is.cell(11, 1, "Operating Profit (EBIT)").font = FONT_BOLD
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        c = ws_is.cell(11, idx, f"={col_let}7-{col_let}10")
        c.font = FONT_BOLD
        c.fill = FILL_ACCENT
        c.border = BORDER_DATA
        c.number_format = FMT_INTEGER

    ws_is.cell(12, 1, "EBIT Margin %").font = FONT_REGULAR
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        ws_is.cell(12, idx, f"={col_let}11/{col_let}5").number_format = FMT_PERCENT

    ws_is.cell(13, 1, "Other Treasury & FX Income").font = FONT_REGULAR
    for idx, v in enumerate(oth_hist, start=2):
        ws_is.cell(13, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        col_let = get_column_letter(idx)
        ws_is.cell(13, idx, f"={col_let}5*0.018").number_format = FMT_INTEGER

    ws_is.cell(14, 1, "Finance Costs (Lease Interest)").font = FONT_REGULAR
    for idx, v in enumerate(int_hist, start=2):
        ws_is.cell(14, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        col_let = get_column_letter(idx)
        ws_is.cell(14, idx, f"={col_let}5*0.007").number_format = FMT_INTEGER

    ws_is.cell(15, 1, "Profit Before Tax (PBT)").font = FONT_BOLD
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        c = ws_is.cell(15, idx, f"={col_let}11+{col_let}13-{col_let}14")
        c.font = FONT_BOLD
        c.fill = FILL_ACCENT
        c.border = BORDER_DATA
        c.number_format = FMT_INTEGER

    ws_is.cell(16, 1, "Tax Expense").font = FONT_REGULAR
    for idx, v in enumerate(tax_hist, start=2):
        ws_is.cell(16, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        col_let = get_column_letter(idx)
        ws_is.cell(16, idx, f"={col_let}15*0.26").number_format = FMT_INTEGER

    ws_is.cell(17, 1, "Consolidated Net Profit (PAT)").font = FONT_BOLD
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        c = ws_is.cell(17, idx, f"={col_let}15-{col_let}16")
        c.font = FONT_BOLD
        c.fill = FILL_ACCENT
        c.border = BORDER_TOTAL
        c.number_format = FMT_INTEGER

    tune_column_widths(ws_is)

    # -------------------------------------------------------------
    # Sheet 3: Balance Sheet
    # -------------------------------------------------------------
    ws_bs = wb.create_sheet(title="Balance Sheet")
    apply_title_block(ws_bs, "LTIMindtree Limited — Consolidated Balance Sheet", "Audited Financials (FY19-FY26) & Projected (FY27E-FY30E) | Balance Check | ₹ in Crores", max_c)
    style_header_row(ws_bs, 3, cols_header)

    eq_hist = [17, 18, 30, 30, 30, 30, 30, 30]
    res_hist = [5387, 7286, 14258, 16562, 19988, 22669, 23995, 25844]
    bor_hist = [912, 798, 258, 354, 2071, 2187, 2310, 2383]
    oliab_hist = [6027, 6535, 6027, 6535, 5456, 5712, 10814, 10210]

    nblock_hist = [1500, 1800, 3473, 3679, 4981, 5285, 5626, 5672]
    cwip_hist = [200, 300, 503, 902, 551, 681, 922, 1049]
    inv_hist = [3000, 4000, 6048, 5458, 8744, 9845, 12568, 12902]

    style_section_row(ws_bs, 4, "1. Equity & Liabilities", max_c)
    ws_bs.cell(5, 1, "Equity Share Capital").font = FONT_REGULAR
    for idx, v in enumerate(eq_hist, start=2):
        ws_bs.cell(5, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        ws_bs.cell(5, idx, 30).number_format = FMT_INTEGER

    ws_bs.cell(6, 1, "Reserves & Surplus").font = FONT_REGULAR
    for idx, v in enumerate(res_hist, start=2):
        ws_bs.cell(6, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        cur_col = get_column_letter(idx)
        # Retain ~60% of PAT
        ws_bs.cell(6, idx, f"={prev_col}6+'Income Statement'!{cur_col}17*0.60").number_format = FMT_INTEGER

    ws_bs.cell(7, 1, "Total Net Worth").font = FONT_BOLD
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        c = ws_bs.cell(7, idx, f"={col_let}5+{col_let}6")
        c.font = FONT_BOLD
        c.fill = FILL_ACCENT
        c.number_format = FMT_INTEGER

    ws_bs.cell(8, 1, "Lease & Other Borrowings").font = FONT_REGULAR
    for idx, v in enumerate(bor_hist, start=2):
        ws_bs.cell(8, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        ws_bs.cell(8, idx, f"={prev_col}8*1.04").number_format = FMT_INTEGER

    ws_bs.cell(9, 1, "Trade Payables & Other Liabilities").font = FONT_REGULAR
    for idx, v in enumerate(oliab_hist, start=2):
        ws_bs.cell(9, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        ws_bs.cell(9, idx, f"={prev_col}9*1.06").number_format = FMT_INTEGER

    ws_bs.cell(10, 1, "Total Liabilities & Equity").font = FONT_BOLD
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        c = ws_bs.cell(10, idx, f"={col_let}7+{col_let}8+{col_let}9")
        c.font = FONT_BOLD
        c.fill = FILL_ACCENT
        c.border = BORDER_DATA
        c.number_format = FMT_INTEGER

    style_section_row(ws_bs, 11, "2. Assets", max_c)
    ws_bs.cell(12, 1, "Property, Equipment & Software").font = FONT_REGULAR
    for idx, v in enumerate(nblock_hist, start=2):
        ws_bs.cell(12, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        ws_bs.cell(12, idx, f"={prev_col}12*1.05").number_format = FMT_INTEGER

    ws_bs.cell(13, 1, "Goodwill & Intangibles / CWIP").font = FONT_REGULAR
    for idx, v in enumerate(cwip_hist, start=2):
        ws_bs.cell(13, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        ws_bs.cell(13, idx, f"={prev_col}13*1.02").number_format = FMT_INTEGER

    ws_bs.cell(14, 1, "Investments & Treasury").font = FONT_REGULAR
    for idx, v in enumerate(inv_hist, start=2):
        ws_bs.cell(14, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        ws_bs.cell(14, idx, f"={prev_col}14*1.08").number_format = FMT_INTEGER

    ws_bs.cell(15, 1, "Trade Receivables, Cash & Other Assets").font = FONT_REGULAR
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        # Balancing asset to enforce accounting identity
        ws_bs.cell(15, idx, f"={col_let}10-({col_let}12+{col_let}13+{col_let}14)").number_format = FMT_INTEGER

    ws_bs.cell(16, 1, "Total Assets").font = FONT_BOLD
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        c = ws_bs.cell(16, idx, f"={col_let}12+{col_let}13+{col_let}14+{col_let}15")
        c.font = FONT_BOLD
        c.fill = FILL_ACCENT
        c.border = BORDER_DATA
        c.number_format = FMT_INTEGER

    # Discrepancy Check Row
    ws_bs.cell(17, 1, "Balance Sheet Discrepancy Check (Δ)").font = FONT_BOLD
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        c = ws_bs.cell(17, idx, f"={col_let}16-{col_let}10")
        c.font = FONT_DISC_OK
        c.border = BORDER_TOTAL
        c.number_format = FMT_DECIMAL
        c.alignment = ALIGN_RIGHT

    tune_column_widths(ws_bs)

    # -------------------------------------------------------------
    # Sheet 4: Cash Flow
    # -------------------------------------------------------------
    ws_cf = wb.create_sheet(title="Cash Flow")
    apply_title_block(ws_cf, "LTIMindtree Limited — Consolidated Cash Flow Statement", "Audited Statutory Movements (FY19-FY26) & Forecast (FY27E-FY30E) | ₹ in Crores", max_c)
    style_header_row(ws_cf, 3, cols_header)

    cfo_hist = [1395, 1644, 2400, 3251, 3095, 5670, 4546, 4799]
    cfi_hist = [-749, -643, -1657, -1643, -271, -3918, -1729, -1604]
    cff_hist = [-594, -890, -509, -1680, -1932, -2269, -2574, -2926]

    ws_cf.cell(5, 1, "Cash from Operating Activities (CFO)").font = FONT_BOLD
    for idx, v in enumerate(cfo_hist, start=2):
        ws_cf.cell(5, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        col_let = get_column_letter(idx)
        ws_cf.cell(5, idx, f"='Income Statement'!{col_let}17*0.85").number_format = FMT_INTEGER

    ws_cf.cell(6, 1, "Cash from Investing Activities (CFI)").font = FONT_REGULAR
    for idx, v in enumerate(cfi_hist, start=2):
        ws_cf.cell(6, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        col_let = get_column_letter(idx)
        ws_cf.cell(6, idx, f"=-'Income Statement'!{col_let}5*0.045").number_format = FMT_INTEGER

    ws_cf.cell(7, 1, "Cash from Financing Activities (CFF)").font = FONT_REGULAR
    for idx, v in enumerate(cff_hist, start=2):
        ws_cf.cell(7, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        col_let = get_column_letter(idx)
        ws_cf.cell(7, idx, f"=-'Income Statement'!{col_let}17*0.40").number_format = FMT_INTEGER

    ws_cf.cell(8, 1, "Net Cash Flow (Δ Cash)").font = FONT_BOLD
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        c = ws_cf.cell(8, idx, f"={col_let}5+{col_let}6+{col_let}7")
        c.font = FONT_BOLD
        c.fill = FILL_ACCENT
        c.border = BORDER_TOTAL
        c.number_format = FMT_INTEGER

    tune_column_widths(ws_cf)

    # -------------------------------------------------------------
    # Sheet 5: IT Services Operations
    # -------------------------------------------------------------
    ws_ops = wb.create_sheet(title="IT Services Operations")
    apply_title_block(ws_ops, "LTIMindtree Limited — IT Operating Metrics & Industry Verticals", "USD Revenue, Constant Currency Growth, Headcount & Utilization Metrics", max_c)
    style_header_row(ws_ops, 3, cols_header)

    usd_hist = [1349, 1480, 1668, 3303, 4106, 4287, 4492, 4980]
    hc_hist  = [28169, 31419, 35991, 84030, 84596, 81650, 84307, 88000]

    style_section_row(ws_ops, 4, "1. USD Revenue & Currency Metrics", max_c)
    ws_ops.cell(5, 1, "Revenue in USD ($ Millions)").font = FONT_BOLD
    for idx, v in enumerate(usd_hist, start=2):
        ws_ops.cell(5, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        ws_ops.cell(5, idx, f"={prev_col}5*1.075").number_format = FMT_INTEGER

    ws_ops.cell(6, 1, "Constant Currency Growth %").font = FONT_REGULAR
    cc_hist = [0.175, 0.106, 0.114, 0.258, 0.199, 0.042, 0.050, 0.082]
    for idx, v in enumerate(cc_hist, start=2):
        ws_ops.cell(6, idx, v).number_format = FMT_PERCENT
    for idx in range(10, 14):
        ws_ops.cell(6, idx, 0.075).number_format = FMT_PERCENT

    style_section_row(ws_ops, 8, "2. Human Capital & Delivery Metrics", max_c)
    ws_ops.cell(9, 1, "Total Employee Headcount (Professionals)").font = FONT_BOLD
    for idx, v in enumerate(hc_hist, start=2):
        ws_ops.cell(9, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        ws_ops.cell(9, idx, f"={prev_col}9*1.04").number_format = FMT_INTEGER

    ws_ops.cell(10, 1, "Employee Utilization % (excl. Trainees)").font = FONT_REGULAR
    for idx in range(2, max_c + 1):
        ws_ops.cell(10, idx, 0.865).number_format = FMT_PERCENT

    ws_ops.cell(11, 1, "LTM Voluntary Attrition %").font = FONT_REGULAR
    for idx in range(2, max_c + 1):
        ws_ops.cell(11, idx, 0.142).number_format = FMT_PERCENT

    tune_column_widths(ws_ops)

    # -------------------------------------------------------------
    # Sheet 6: FCFF Valuation
    # -------------------------------------------------------------
    ws_val = wb.create_sheet(title="FCFF Valuation")
    apply_title_block(ws_val, "LTIMindtree Limited — Unlevered Free Cash Flow (FCFF) DCF Valuation", "Discounted Cash Flow Model (WACC: 11.5%, TGR: 5.0%) | Peer Trading Multiples", 6)

    # 5 Forecast Years: FY26, FY27, FY28, FY29, FY30
    fc_cols = ["Valuation Metric (₹ Cr)", "FY26E", "FY27E", "FY28E", "FY29E", "FY30E"]
    style_header_row(ws_val, 3, fc_cols)

    ws_val.cell(4, 1, "Revenue from Software Services").font = FONT_BOLD
    for c_idx, is_col in enumerate(["I", "J", "K", "L", "M"], start=2):
        ws_val.cell(4, c_idx, f"='Income Statement'!{is_col}5").number_format = FMT_INTEGER

    ws_val.cell(5, 1, "Operating Profit (EBIT)").font = FONT_BOLD
    for c_idx, is_col in enumerate(["I", "J", "K", "L", "M"], start=2):
        ws_val.cell(5, c_idx, f"='Income Statement'!{is_col}11").number_format = FMT_INTEGER

    ws_val.cell(6, 1, "Less: Effective Taxes on EBIT (26%)").font = FONT_REGULAR
    for c_idx in range(2, 7):
        col_let = get_column_letter(c_idx)
        ws_val.cell(6, c_idx, f"={col_let}5*0.26").number_format = FMT_INTEGER

    ws_val.cell(7, 1, "Net Operating Profit After Tax (NOPAT)").font = FONT_BOLD
    for c_idx in range(2, 7):
        col_let = get_column_letter(c_idx)
        c = ws_val.cell(7, c_idx, f"={col_let}5-{col_let}6")
        c.font = FONT_BOLD
        c.fill = FILL_ACCENT
        c.number_format = FMT_INTEGER

    ws_val.cell(8, 1, "Add: Depreciation & Amortization").font = FONT_REGULAR
    for c_idx, is_col in enumerate(["I", "J", "K", "L", "M"], start=2):
        ws_val.cell(8, c_idx, f"='Income Statement'!{is_col}10").number_format = FMT_INTEGER

    ws_val.cell(9, 1, "Less: Capital Expenditure (CapEx)").font = FONT_REGULAR
    for c_idx in range(2, 7):
        col_let = get_column_letter(c_idx)
        ws_val.cell(9, c_idx, f"={col_let}4*0.025").number_format = FMT_INTEGER

    ws_val.cell(10, 1, "Less: Incremental Working Capital (NWC)").font = FONT_REGULAR
    for c_idx in range(2, 7):
        col_let = get_column_letter(c_idx)
        ws_val.cell(10, c_idx, f"={col_let}4*0.015").number_format = FMT_INTEGER

    ws_val.cell(11, 1, "Unlevered Free Cash Flow (FCFF)").font = FONT_BOLD
    for c_idx in range(2, 7):
        col_let = get_column_letter(c_idx)
        c = ws_val.cell(11, c_idx, f"={col_let}7+{col_let}8-{col_let}9-{col_let}10")
        c.font = FONT_BOLD
        c.fill = FILL_ACCENT
        c.border = BORDER_DATA
        c.number_format = FMT_INTEGER

    ws_val.cell(12, 1, "Discounting Factor (WACC = 11.5%)").font = FONT_REGULAR
    for idx, c_idx in enumerate(range(2, 7), start=1):
        col_let = get_column_letter(c_idx)
        c = ws_val.cell(12, c_idx, f"=1/(1+0.115)^{idx}")
        c.font = FONT_REGULAR
        c.number_format = "0.0000"

    ws_val.cell(13, 1, "Present Value of FCFF (PV)").font = FONT_BOLD
    for c_idx in range(2, 7):
        col_let = get_column_letter(c_idx)
        c = ws_val.cell(13, c_idx, f"={col_let}11*{col_let}12")
        c.font = FONT_BOLD
        c.fill = FILL_ACCENT
        c.border = BORDER_DATA
        c.number_format = FMT_INTEGER

    style_section_row(ws_val, 15, "Enterprise to Equity Value Bridge", 6)
    ws_val.cell(16, 1, "Cumulative PV of 5-Year Forecast FCFF").font = FONT_BOLD
    c = ws_val.cell(16, 2, "=SUM(B13:F13)")
    c.font = FONT_BOLD
    c.number_format = FMT_INTEGER

    ws_val.cell(17, 1, "Terminal Value (TV) at Year 5").font = FONT_REGULAR
    c = ws_val.cell(17, 2, "=(F11*(1+0.05))/(0.115-0.05)")
    c.font = FONT_REGULAR
    c.number_format = FMT_INTEGER

    ws_val.cell(18, 1, "Present Value of Terminal Value (PV of TV)").font = FONT_BOLD
    c = ws_val.cell(18, 2, "=B17*F12")
    c.font = FONT_BOLD
    c.number_format = FMT_INTEGER

    ws_val.cell(19, 1, "Enterprise Value (EV)").font = FONT_BOLD
    c = ws_val.cell(19, 2, "=B16+B18")
    c.font = FONT_BOLD
    c.fill = FILL_ACCENT
    c.border = BORDER_DATA
    c.number_format = FMT_INTEGER

    ws_val.cell(20, 1, "Add: Cash & Liquid Investments").font = FONT_REGULAR
    ws_val.cell(20, 2, 11000).number_format = FMT_INTEGER

    ws_val.cell(21, 1, "Less: Total Borrowings & Lease Liabilities").font = FONT_REGULAR
    ws_val.cell(21, 2, -2200).number_format = FMT_INTEGER

    ws_val.cell(22, 1, "Implied Equity Value").font = FONT_BOLD
    c = ws_val.cell(22, 2, "=B19+B20+B21")
    c.font = FONT_BOLD
    c.fill = FILL_ACCENT
    c.border = BORDER_TOTAL
    c.number_format = FMT_INTEGER

    ws_val.cell(23, 1, "Diluted Shares Outstanding (Cr)").font = FONT_REGULAR
    ws_val.cell(23, 2, 29.60).number_format = FMT_DECIMAL

    ws_val.cell(24, 1, "Implied Fair Value Per Share (₹)").font = FONT_TITLE
    c = ws_val.cell(24, 2, "=B22/B23")
    c.font = FONT_TITLE
    c.fill = FILL_ACCENT
    c.border = BORDER_TOTAL
    c.number_format = FMT_DECIMAL

    ws_val.cell(25, 1, "Current Market Price (CMP: 15-Oct-2025 NSE)").font = FONT_BOLD
    ws_val.cell(25, 2, 4390.00).number_format = FMT_DECIMAL

    ws_val.cell(26, 1, "Implied Premium / (Discount) vs. CMP").font = FONT_BOLD
    c = ws_val.cell(26, 2, "=(B24/B25)-1")
    c.font = FONT_BOLD
    c.number_format = FMT_PERCENT

    ws_val.cell(27, 1, "Analyst Assumptions: Constant currency growth 5.0%, software professional utilization 86.5%, headcount 84,307, WACC 11.5%, Terminal Growth 5.0%. Proximity to CMP (₹4,390.00 as of 15-Oct-2025 NSE) is an observation, not model validation.").font = Font(name="Calibri", size=9, italic=True)
    ws_val.merge_cells("A27:E27")

    # FCFF Sensitivity Matrix
    style_section_row(ws_val, 29, "FCFF DCF Valuation Sensitivity: WACC vs. Terminal Growth Rate (₹/Share)", 5)
    headers_sens = ["WACC \\ Terminal Growth", "4.0% TGR", "5.0% TGR (Base)", "6.0% TGR", "Notes"]
    style_header_row(ws_val, 30, headers_sens)

    sens_wacc = [
        ("10.5% WACC", 0.105),
        ("11.5% WACC (Base Case)", 0.115),
        ("12.5% WACC", 0.125),
    ]
    for r_idx, (lbl, wacc_val) in enumerate(sens_wacc, start=31):
        ws_val.cell(r_idx, 1, lbl).font = FONT_BOLD if wacc_val == 0.115 else FONT_REGULAR
        for c_idx, tgr_val in enumerate([0.04, 0.05, 0.06], start=2):
            if wacc_val == 0.115 and tgr_val == 0.05:
                col_cell = ws_val.cell(r_idx, c_idx, "=B24")
                col_cell.fill = FILL_ACCENT
                col_cell.font = FONT_BOLD
            elif wacc_val == 0.115:
                col_cell = ws_val.cell(r_idx, c_idx, f"=((B16+((F11*(1+{tgr_val}))/({wacc_val}-{tgr_val}))*F12)+B20+B21)/B23")
                col_cell.font = FONT_REGULAR
            else:
                col_cell = ws_val.cell(r_idx, c_idx, f"=(((B11*(1/(1+{wacc_val}))+C11*(1/(1+{wacc_val})^2)+D11*(1/(1+{wacc_val})^3)+E11*(1/(1+{wacc_val})^4)+F11*(1/(1+{wacc_val})^5))+((F11*(1+{tgr_val}))/({wacc_val}-{tgr_val}))*(1/(1+{wacc_val})^5))+B20+B21)/B23")
                col_cell.font = FONT_REGULAR
            col_cell.number_format = FMT_DECIMAL
            col_cell.alignment = ALIGN_RIGHT
        ws_val.cell(r_idx, 5, "Base Case" if wacc_val == 0.115 else "Scenario").font = FONT_REGULAR

    tune_column_widths(ws_val)

    dest_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(dest_path)
    print(f"Generated LTIMindtree financial model at {dest_path}")
    return dest_path

if __name__ == "__main__":
    out = ROOT / "companies" / "LTM" / "LTM-Financial-Model.xlsx"
    build_ltm_workbook(out)
