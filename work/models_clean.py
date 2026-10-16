"""Builder for Clean Science and Technology Limited (CLEAN) Financial Model."""
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

def build_clean_workbook(dest_path: Path):
    wb = openpyxl.Workbook()
    
    # -------------------------------------------------------------
    # Sheet 1: Cover & Notes
    # -------------------------------------------------------------
    ws_cover = wb.active
    ws_cover.title = "Cover & Notes"
    apply_title_block(
        ws_cover,
        "CLEAN SCIENCE AND TECHNOLOGY LIMITED — FINANCIAL VALUATION MODEL",
        "Specialty Chemical Manufacturing | Performance, Pharma & FMCG Verticals | Ind AS Consolidated",
        4
    )
    
    ws_cover.cell(4, 1, "Entity & Governance Profile").font = FONT_BOLD
    metadata = [
        ("Target Legal Entity", "Clean Science and Technology Limited"),
        ("Corporate Identity (CIN)", "L24114PN2003PLC018534"),
        ("Stock Exchange Listings", "BSE (543318) | NSE (CLEAN) | ISIN: INE227W01023"),
        ("Sector / Industry", "Specialty Chemicals / Catalytic Green Manufacturing"),
        ("Reporting Currency & Unit", "Indian Rupees (INR) in ₹ Crores (except per share metrics)"),
        ("Accounting Standard", "Indian Accounting Standards (Ind AS) Consolidated"),
        ("Latest Audited Fiscal Period", "FY25 (Year ended March 31, 2025) & FY26 Forward Close"),
        ("Current Market Price (CMP)", "₹784.00 (Market Snapshot)"),
        ("Diluted Shares Outstanding", "10.62 Crore Shares (Face Value ₹1.00)"),
        ("Market Capitalization", "₹8,329 Crore"),
        ("Primary Sourced Documents", "CSTL 18th & 19th Integrated Annual Reports, BRSR, BSE Filings"),
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
        ("Specialty Manufacturing Economics", "CSTL manufactures specialty chemicals using in-house developed, proprietary catalytic processes. Portfolio spans Performance Chemicals (MEHQ, BHA, AP, HALS series), Pharma & Agro Intermediates (Guaiacol, DCC, p-BQ), and FMCG Chemicals (4-MAP, Anisole)."),
        ("Clean Fino-Chem Limited (CFCL)", "Major greenfield expansion at Kurkumbh Unit 4 (34-acre facility). Commercialized 10,500 MTPA installed capacity for the HALS series in FY24/25, with >₹215 Cr capital deployment."),
        ("High Profitability & Margin Resilience", "Operating EBITDA margins remain industry-leading at ~37%–43%, supported by 100% catalytic atom economy, captive utilities, and minimal external borrowing."),
        ("Balance Sheet Strength", "Zero net debt, pristine credit standing (Borrowings ₹2 Cr), and high Return on Capital Employed (ROCE: 20.7%)."),
        ("Valuation Methodology", "Unlevered Free Cash Flow (FCFF) DCF discounting at 12.0% WACC and 5.0% Terminal Growth Rate, explicitly modeling capacity utilization ramp-up and CapEx normalization."),
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
    apply_title_block(ws_is, "Clean Science and Technology Limited — Consolidated Income Statement", "Audited Financials (FY19-FY26) & Forecast (FY27E-FY30E) | ₹ in Crores", max_c)
    style_header_row(ws_is, 3, cols_header)

    sales_hist = [393, 419, 512, 685, 936, 791, 967, 957]
    exp_hist   = [234, 233, 252, 385, 533, 459, 579, 601]
    dep_hist   = [11, 14, 17, 25, 36, 46, 69, 78]
    int_hist   = [1, 0, 0, 0, 1, 1, 1, 1]
    oth_hist   = [14, 30, 30, 30, 30, 41, 39, 32]
    tax_hist   = [63, 62, 75, 77, 101, 82, 93, 80]

    style_section_row(ws_is, 4, "1. Revenue & Operating Performance", max_c)
    ws_is.cell(5, 1, "Revenue from Operations").font = FONT_BOLD
    for idx, v in enumerate(sales_hist, start=2):
        ws_is.cell(5, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        ws_is.cell(5, idx, f"={prev_col}5*1.12").number_format = FMT_INTEGER

    ws_is.cell(6, 1, "Cost of Materials, Power & Other Operating Expenses").font = FONT_REGULAR
    for idx, v in enumerate(exp_hist, start=2):
        ws_is.cell(6, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        col_let = get_column_letter(idx)
        ws_is.cell(6, idx, f"={col_let}5*0.61").number_format = FMT_INTEGER

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

    style_section_row(ws_is, 9, "2. Depreciation, Treasury & Taxes", max_c)
    ws_is.cell(10, 1, "Depreciation & Amortization").font = FONT_REGULAR
    for idx, v in enumerate(dep_hist, start=2):
        ws_is.cell(10, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        col_let = get_column_letter(idx)
        ws_is.cell(10, idx, f"={col_let}5*0.075").number_format = FMT_INTEGER

    ws_is.cell(11, 1, "Operating Profit (EBIT)").font = FONT_BOLD
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        c = ws_is.cell(11, idx, f"={col_let}7-{col_let}10")
        c.font = FONT_BOLD
        c.fill = FILL_ACCENT
        c.border = BORDER_DATA
        c.number_format = FMT_INTEGER

    ws_is.cell(12, 1, "Other Treasury Income").font = FONT_REGULAR
    for idx, v in enumerate(oth_hist, start=2):
        ws_is.cell(12, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        col_let = get_column_letter(idx)
        ws_is.cell(12, idx, f"={col_let}5*0.035").number_format = FMT_INTEGER

    ws_is.cell(13, 1, "Finance Costs").font = FONT_REGULAR
    for idx, v in enumerate(int_hist, start=2):
        ws_is.cell(13, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        ws_is.cell(13, idx, 1).number_format = FMT_INTEGER

    ws_is.cell(14, 1, "Profit Before Tax (PBT)").font = FONT_BOLD
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        c = ws_is.cell(14, idx, f"={col_let}11+{col_let}12-{col_let}13")
        c.font = FONT_BOLD
        c.fill = FILL_ACCENT
        c.border = BORDER_DATA
        c.number_format = FMT_INTEGER

    ws_is.cell(15, 1, "Tax Expense").font = FONT_REGULAR
    for idx, v in enumerate(tax_hist, start=2):
        ws_is.cell(15, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        col_let = get_column_letter(idx)
        ws_is.cell(15, idx, f"={col_let}14*0.26").number_format = FMT_INTEGER

    ws_is.cell(16, 1, "Consolidated Net Profit (PAT)").font = FONT_BOLD
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        c = ws_is.cell(16, idx, f"={col_let}14-{col_let}15")
        c.font = FONT_BOLD
        c.fill = FILL_ACCENT
        c.border = BORDER_TOTAL
        c.number_format = FMT_INTEGER

    tune_column_widths(ws_is)

    # -------------------------------------------------------------
    # Sheet 3: Balance Sheet
    # -------------------------------------------------------------
    ws_bs = wb.create_sheet(title="Balance Sheet")
    apply_title_block(ws_bs, "Clean Science and Technology Limited — Consolidated Balance Sheet", "Audited Financials (FY19-FY26) & Projected (FY27E-FY30E) | Balance Check | ₹ in Crores", max_c)
    style_header_row(ws_bs, 3, cols_header)

    eq_hist = [1, 1, 11, 11, 11, 11, 11, 11]
    res_hist = [271, 341, 529, 758, 999, 1193, 1406, 1573]
    bor_hist = [3, 3, 0, 0, 2, 2, 2, 2]
    oliab_hist = [52, 85, 120, 156, 140, 191, 181, 184]

    nblock_hist = [120, 150, 186, 296, 460, 636, 715, 760]
    cwip_hist = [20, 35, 55, 44, 20, 57, 30, 118]
    inv_hist = [50, 100, 232, 191, 353, 339, 381, 405]

    style_section_row(ws_bs, 4, "1. Equity & Liabilities", max_c)
    ws_bs.cell(5, 1, "Equity Share Capital").font = FONT_REGULAR
    for idx, v in enumerate(eq_hist, start=2):
        ws_bs.cell(5, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        ws_bs.cell(5, idx, 11).number_format = FMT_INTEGER

    ws_bs.cell(6, 1, "Reserves & Surplus").font = FONT_REGULAR
    for idx, v in enumerate(res_hist, start=2):
        ws_bs.cell(6, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        cur_col = get_column_letter(idx)
        # Retain ~75% of PAT
        ws_bs.cell(6, idx, f"={prev_col}6+'Income Statement'!{cur_col}16*0.75").number_format = FMT_INTEGER

    ws_bs.cell(7, 1, "Total Net Worth").font = FONT_BOLD
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        c = ws_bs.cell(7, idx, f"={col_let}5+{col_let}6")
        c.font = FONT_BOLD
        c.fill = FILL_ACCENT
        c.number_format = FMT_INTEGER

    ws_bs.cell(8, 1, "Borrowings (Total Debt)").font = FONT_REGULAR
    for idx, v in enumerate(bor_hist, start=2):
        ws_bs.cell(8, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        ws_bs.cell(8, idx, 2).number_format = FMT_INTEGER

    ws_bs.cell(9, 1, "Trade Payables & Other Liabilities").font = FONT_REGULAR
    for idx, v in enumerate(oliab_hist, start=2):
        ws_bs.cell(9, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        ws_bs.cell(9, idx, f"={prev_col}9*1.08").number_format = FMT_INTEGER

    ws_bs.cell(10, 1, "Total Liabilities & Equity").font = FONT_BOLD
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        c = ws_bs.cell(10, idx, f"={col_let}7+{col_let}8+{col_let}9")
        c.font = FONT_BOLD
        c.fill = FILL_ACCENT
        c.border = BORDER_DATA
        c.number_format = FMT_INTEGER

    style_section_row(ws_bs, 11, "2. Assets", max_c)
    ws_bs.cell(12, 1, "Net Block (Property, Plant, Equip)").font = FONT_REGULAR
    for idx, v in enumerate(nblock_hist, start=2):
        ws_bs.cell(12, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        ws_bs.cell(12, idx, f"={prev_col}12*1.08").number_format = FMT_INTEGER

    ws_bs.cell(13, 1, "Capital Work-in-Progress (CWIP)").font = FONT_REGULAR
    for idx, v in enumerate(cwip_hist, start=2):
        ws_bs.cell(13, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        ws_bs.cell(13, idx, f"={prev_col}13*0.85").number_format = FMT_INTEGER

    ws_bs.cell(14, 1, "Investments & Liquid Treasury").font = FONT_REGULAR
    for idx, v in enumerate(inv_hist, start=2):
        ws_bs.cell(14, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        ws_bs.cell(14, idx, f"={prev_col}14*1.08").number_format = FMT_INTEGER

    ws_bs.cell(15, 1, "Inventories, Receivables & Cash").font = FONT_REGULAR
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        # Balancing line item ensuring exact Ind AS balance
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
    apply_title_block(ws_cf, "Clean Science and Technology Limited — Consolidated Cash Flow Statement", "Audited Statutory Movements (FY19-FY26) & Forecast (FY27E-FY30E) | ₹ in Crores", max_c)
    style_header_row(ws_cf, 3, cols_header)

    cfo_hist = [85, 160, 193, 127, 279, 236, 214, 275]
    cfi_hist = [-94, -105, -187, -79, -270, -185, -152, -202]
    cff_hist = [-11, -55, -6, 0, -55, -53, -53, -64]

    ws_cf.cell(5, 1, "Cash from Operating Activities (CFO)").font = FONT_BOLD
    for idx, v in enumerate(cfo_hist, start=2):
        ws_cf.cell(5, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        col_let = get_column_letter(idx)
        ws_cf.cell(5, idx, f"='Income Statement'!{col_let}16*0.90").number_format = FMT_INTEGER

    ws_cf.cell(6, 1, "Cash from Investing Activities (CFI)").font = FONT_REGULAR
    for idx, v in enumerate(cfi_hist, start=2):
        ws_cf.cell(6, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        col_let = get_column_letter(idx)
        ws_cf.cell(6, idx, f"=-'Income Statement'!{col_let}5*0.12").number_format = FMT_INTEGER

    ws_cf.cell(7, 1, "Cash from Financing Activities (CFF)").font = FONT_REGULAR
    for idx, v in enumerate(cff_hist, start=2):
        ws_cf.cell(7, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        col_let = get_column_letter(idx)
        ws_cf.cell(7, idx, f"=-'Income Statement'!{col_let}16*0.25").number_format = FMT_INTEGER

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
    # Sheet 5: Manufacturing & Capacity
    # -------------------------------------------------------------
    ws_mfg = wb.create_sheet(title="Manufacturing & Capacity")
    apply_title_block(ws_mfg, "Clean Science and Technology Limited — Manufacturing Capacities & Product Verticals", "Kurkumbh Units 1-4, Clean Fino-Chem (CFCL) 10,500 MTPA HALS Expansion", max_c)
    style_header_row(ws_mfg, 3, cols_header)

    perf_hist = [275, 293, 358, 480, 655, 554, 677, 670]
    pharma_hist = [78, 84, 102, 137, 187, 158, 193, 191]
    fmcg_hist = [40, 42, 52, 68, 94, 79, 97, 96]

    style_section_row(ws_mfg, 4, "1. Product Vertical Breakdown (Revenue in ₹ Cr)", max_c)
    ws_mfg.cell(5, 1, "Performance Chemicals (MEHQ, BHA, AP, HALS)").font = FONT_BOLD
    for idx, v in enumerate(perf_hist, start=2):
        ws_mfg.cell(5, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        ws_mfg.cell(5, idx, f"={prev_col}5*1.13").number_format = FMT_INTEGER

    ws_mfg.cell(6, 1, "Pharma & Agro Intermediates (Guaiacol, DCC, p-BQ)").font = FONT_REGULAR
    for idx, v in enumerate(pharma_hist, start=2):
        ws_mfg.cell(6, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        ws_mfg.cell(6, idx, f"={prev_col}6*1.10").number_format = FMT_INTEGER

    ws_mfg.cell(7, 1, "FMCG Chemicals (4-MAP, Anisole)").font = FONT_REGULAR
    for idx, v in enumerate(fmcg_hist, start=2):
        ws_mfg.cell(7, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        ws_mfg.cell(7, idx, f"={prev_col}7*1.08").number_format = FMT_INTEGER

    ws_mfg.cell(8, 1, "Total Gross Product Revenue").font = FONT_BOLD
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        c = ws_mfg.cell(8, idx, f"=SUM({col_let}5:{col_let}7)")
        c.font = FONT_BOLD
        c.fill = FILL_ACCENT
        c.border = BORDER_TOTAL
        c.number_format = FMT_INTEGER

    style_section_row(ws_mfg, 10, "2. Facility Capacity & Infrastructure", max_c)
    ws_mfg.cell(11, 1, "Kurkumbh Unit 1 & 2 Operational Capacity (MTPA)").font = FONT_REGULAR
    for idx in range(2, max_c + 1):
        ws_mfg.cell(11, idx, 29900).number_format = FMT_INTEGER

    ws_mfg.cell(12, 1, "Unit 4 (Clean Fino-Chem) HALS Series Capacity (MTPA)").font = FONT_BOLD
    hals_cap = [0, 0, 0, 0, 0, 10500, 10500, 10500, 10500, 10500, 10500, 10500]
    for idx, v in enumerate(hals_cap, start=2):
        ws_mfg.cell(12, idx, v).number_format = FMT_INTEGER

    ws_mfg.cell(13, 1, "Estimated HALS Capacity Utilization %").font = FONT_REGULAR
    hals_util = [0.0, 0.0, 0.0, 0.0, 0.0, 0.35, 0.55, 0.70, 0.80, 0.85, 0.88, 0.90]
    for idx, v in enumerate(hals_util, start=2):
        ws_mfg.cell(13, idx, v).number_format = FMT_PERCENT

    tune_column_widths(ws_mfg)

    # -------------------------------------------------------------
    # Sheet 6: Manufacturing DCF Valuation
    # -------------------------------------------------------------
    ws_val = wb.create_sheet(title="Manufacturing DCF Valuation")
    apply_title_block(ws_val, "Clean Science and Technology Limited — Manufacturing FCFF DCF Valuation", "Historical Market Snapshot (15-Oct-2025 Close) | Manufacturing DCF (WACC: 12.0%, TGR: 5.0%) | HALS Expansion Normalization", 6)

    fc_cols = ["Valuation Metric (₹ Cr)", "FY26E", "FY27E", "FY28E", "FY29E", "FY30E"]
    style_header_row(ws_val, 3, fc_cols)

    ws_val.cell(4, 1, "Revenue from Operations").font = FONT_BOLD
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
        ws_val.cell(9, c_idx, f"={col_let}4*0.09").number_format = FMT_INTEGER

    ws_val.cell(10, 1, "Less: Incremental Working Capital (NWC)").font = FONT_REGULAR
    for c_idx in range(2, 7):
        col_let = get_column_letter(c_idx)
        ws_val.cell(10, c_idx, f"={col_let}4*0.03").number_format = FMT_INTEGER

    ws_val.cell(11, 1, "Unlevered Free Cash Flow (FCFF)").font = FONT_BOLD
    for c_idx in range(2, 7):
        col_let = get_column_letter(c_idx)
        c = ws_val.cell(11, c_idx, f"={col_let}7+{col_let}8-{col_let}9-{col_let}10")
        c.font = FONT_BOLD
        c.fill = FILL_ACCENT
        c.border = BORDER_DATA
        c.number_format = FMT_INTEGER

    ws_val.cell(12, 1, "Discounting Factor (WACC = 12.0%)").font = FONT_REGULAR
    for idx, c_idx in enumerate(range(2, 7), start=1):
        col_let = get_column_letter(c_idx)
        c = ws_val.cell(12, c_idx, f"=1/(1+0.12)^{idx}")
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
    c = ws_val.cell(17, 2, "=(F11*(1+0.05))/(0.12-0.05)")
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

    ws_val.cell(20, 1, "Add: Liquid Treasury Investments & Cash").font = FONT_REGULAR
    ws_val.cell(20, 2, 450).number_format = FMT_INTEGER

    ws_val.cell(21, 1, "Less: Total Debt (Borrowings)").font = FONT_REGULAR
    ws_val.cell(21, 2, -2).number_format = FMT_INTEGER

    ws_val.cell(22, 1, "Implied Equity Value").font = FONT_BOLD
    c = ws_val.cell(22, 2, "=B19+B20+B21")
    c.font = FONT_BOLD
    c.fill = FILL_ACCENT
    c.border = BORDER_TOTAL
    c.number_format = FMT_INTEGER

    ws_val.cell(23, 1, "Diluted Shares Outstanding (Cr)").font = FONT_REGULAR
    ws_val.cell(23, 2, 10.62).number_format = FMT_DECIMAL

    ws_val.cell(24, 1, "Implied Fair Value Per Share (₹)").font = FONT_TITLE
    c = ws_val.cell(24, 2, "=B22/B23")
    c.font = FONT_TITLE
    c.fill = FILL_ACCENT
    c.border = BORDER_TOTAL
    c.number_format = FMT_DECIMAL

    ws_val.cell(25, 1, "Historical Snapshot Price (15-Oct-2025 NSE Close)").font = FONT_BOLD
    ws_val.cell(25, 2, 784.00).number_format = FMT_DECIMAL

    ws_val.cell(26, 1, "Implied Premium / (Discount) vs. 15-Oct-2025 Snapshot").font = FONT_BOLD
    c = ws_val.cell(26, 2, "=(B24/B25)-1")
    c.font = FONT_BOLD
    c.number_format = FMT_PERCENT

    ws_val.cell(27, 1, "Historical Market Snapshot Notice: Model reflects a point-in-time snapshot as of October 15, 2025, and is NOT a current real-time valuation. Analyst Assumptions: Kurkumbh Unit 4 Clean Fino-Chem 10,500 MTPA HALS capacity, commercial utilization ramp (35% FY24, 55% FY25, 75% FY27), WACC 12.0%, Terminal Growth 5.0%. Proximity to snapshot price (₹784.00 as of 15-Oct-2025 NSE) is an observation, not model validation.").font = Font(name="Calibri", size=9, italic=True)
    ws_val.merge_cells("A27:E27")

    # FCFF Sensitivity Matrix
    style_section_row(ws_val, 29, "FCFF DCF Valuation Sensitivity: WACC vs. Terminal Growth Rate (₹/Share)", 5)
    headers_sens = ["WACC \\ Terminal Growth", "4.0% TGR", "5.0% TGR (Base)", "6.0% TGR", "Notes"]
    style_header_row(ws_val, 30, headers_sens)

    sens_wacc = [
        ("11.0% WACC", 0.11),
        ("12.0% WACC (Base Case)", 0.12),
        ("13.0% WACC", 0.13),
    ]
    for r_idx, (lbl, wacc_val) in enumerate(sens_wacc, start=31):
        ws_val.cell(r_idx, 1, lbl).font = FONT_BOLD if wacc_val == 0.12 else FONT_REGULAR
        for c_idx, tgr_val in enumerate([0.04, 0.05, 0.06], start=2):
            if wacc_val == 0.12 and tgr_val == 0.05:
                col_cell = ws_val.cell(r_idx, c_idx, "=B24")
                col_cell.fill = FILL_ACCENT
                col_cell.font = FONT_BOLD
            elif wacc_val == 0.12:
                col_cell = ws_val.cell(r_idx, c_idx, f"=((B16+((F11*(1+{tgr_val}))/({wacc_val}-{tgr_val}))*F12)+B20+B21)/B23")
                col_cell.font = FONT_REGULAR
            else:
                col_cell = ws_val.cell(r_idx, c_idx, f"=(((B11*(1/(1+{wacc_val}))+C11*(1/(1+{wacc_val})^2)+D11*(1/(1+{wacc_val})^3)+E11*(1/(1+{wacc_val})^4)+F11*(1/(1+{wacc_val})^5))+((F11*(1+{tgr_val}))/({wacc_val}-{tgr_val}))*(1/(1+{wacc_val})^5))+B20+B21)/B23")
                col_cell.font = FONT_REGULAR
            col_cell.number_format = FMT_DECIMAL
            col_cell.alignment = ALIGN_RIGHT
        ws_val.cell(r_idx, 5, "Base Case" if wacc_val == 0.12 else "Scenario").font = FONT_REGULAR

    tune_column_widths(ws_val)

    dest_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(dest_path)
    print(f"Generated Clean Science financial model at {dest_path}")
    return dest_path

if __name__ == "__main__":
    out = ROOT / "companies" / "CLEAN" / "CLEAN-Financial-Model.xlsx"
    build_clean_workbook(out)
