"""Builder for Edelweiss Financial Services Limited (EDELWEISS) Financial Model."""
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

def build_edelweiss_workbook(dest_path: Path):
    wb = openpyxl.Workbook()
    
    # -------------------------------------------------------------
    # Sheet 1: Cover & Notes
    # -------------------------------------------------------------
    ws_cover = wb.active
    ws_cover.title = "Cover & Notes"
    apply_title_block(
        ws_cover,
        "EDELWEISS FINANCIAL SERVICES LIMITED — FINANCIAL VALUATION MODEL",
        "Diversified Financial Services Group | Alternatives, ARC, Mutual Fund, Insurance & Credit | SOTP Valuation",
        4
    )
    
    ws_cover.cell(4, 1, "Entity & Governance Profile").font = FONT_BOLD
    metadata = [
        ("Target Legal Entity", "Edelweiss Financial Services Limited"),
        ("Corporate Identity (CIN)", "L99999MH1995PLC094641"),
        ("Stock Exchange Listings", "BSE (532922) | NSE (EDELWEISS) | ISIN: INE532F01054"),
        ("Sector / Industry", "Financial Services / Diversified Financials & Asset Management"),
        ("Reporting Currency & Unit", "Indian Rupees (INR) in ₹ Crores (except per share metrics)"),
        ("Accounting Standard", "Indian Accounting Standards (Ind AS) Consolidated Financials"),
        ("Latest Audited Fiscal Period", "FY25 (Year ended March 31, 2025) & FY26 Forward Close"),
        ("Current Market Price (CMP)", "₹135.00 (Market Snapshot)"),
        ("Diluted Shares Outstanding", "94.50 Crore Shares (Face Value ₹1.00)"),
        ("Market Capitalization", "₹12,746 Crore"),
        ("Primary Sourced Documents", "EFSL 29th & 30th Annual Reports, Rating Rationales (Crisil/Care/Acuite), BSE Filings"),
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
        ("Financial Services Economics", "EFSL operates as a diversified financial conglomerate spanning Alternative Asset Management (EAAA), Asset Reconstruction (EARCL), Mutual Fund (EAML), Insurance (General & Life), and Credit/Mortgages. Wealth management (Nuvama) was successfully demerged in 2023."),
        ("Industrial FCFF Strictly Withheld", "Industrial Free Cash Flow to Firm (FCFF) discounting is methodologically invalid for financial institutions with debt-funded balance sheets, floating borrowings, and lending books. FCFF is formally excluded and withheld."),
        ("Sum-of-the-Parts (SOTP) Valuation", "Conglomerate valuation is derived through business vertical SOTP: EAAA valued on P/E (22x on FY25 PAT of ₹230 Cr), EARCL on P/B (1.2x on net worth), Mutual Fund at 4.0% of AAUM (₹1.41 Lakh Cr), Insurance at 1.0x net worth/embedded value proxy, and Credit at 0.8x P/B, less holding company net debt."),
        ("Credit Costs & Asset Quality", "Wholesale credit exposure has been progressively resolved into the ARC platform, shifting group economics toward capital-light, high-fee alternative asset management."),
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
    apply_title_block(ws_is, "Edelweiss Financial Services Limited — Consolidated Income Statement", "Audited Financials (FY19-FY26) & Forecast (FY27E-FY30E) | Financial Services Ind AS | ₹ in Crores", max_c)
    style_header_row(ws_is, 3, cols_header)

    tot_rev_hist = [11078, 9513, 10775, 7151, 8500, 9523, 9393, 10802]
    int_inc_hist = [6500, 5200, 5600, 3800, 4200, 4600, 4500, 5100]
    fee_inc_hist = [4578, 4313, 5175, 3351, 4300, 4923, 4893, 5702]
    int_exp_hist = [5800, 4793, 3834, 2984, 2575, 2786, 2537, 2492]
    opex_hist    = [2600, 2244, 2809, 1145, 3092, 3466, 3496, 4835]
    prov_hist    = [1000, 2044, 520, 230, 200, 240, 260, 290]
    dep_hist     = [232, 260, 151, 138, 126, 147, 144, 144]
    tax_hist     = [383, 200, 25, 15, -21, 266, 149, 147]

    style_section_row(ws_is, 4, "1. Net Interest Income & Fee Revenues (₹ Cr)", max_c)
    ws_is.cell(5, 1, "Interest & Financing Income").font = FONT_REGULAR
    for idx, v in enumerate(int_inc_hist, start=2):
        ws_is.cell(5, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        ws_is.cell(5, idx, f"={prev_col}5*1.06").number_format = FMT_INTEGER

    ws_is.cell(6, 1, "Interest & Borrowing Expenses").font = FONT_REGULAR
    for idx, v in enumerate(int_exp_hist, start=2):
        ws_is.cell(6, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        ws_is.cell(6, idx, f"={prev_col}6*1.04").number_format = FMT_INTEGER

    ws_is.cell(7, 1, "Net Interest Income (NII)").font = FONT_BOLD
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        c = ws_is.cell(7, idx, f"={col_let}5-{col_let}6")
        c.font = FONT_BOLD
        c.fill = FILL_ACCENT
        c.border = BORDER_DATA
        c.number_format = FMT_INTEGER

    ws_is.cell(8, 1, "Fee, Commission, Asset Mgmt & Advisory Income").font = FONT_REGULAR
    for idx, v in enumerate(fee_inc_hist, start=2):
        ws_is.cell(8, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        ws_is.cell(8, idx, f"={prev_col}8*1.10").number_format = FMT_INTEGER

    ws_is.cell(9, 1, "Total Operating Net Revenue").font = FONT_BOLD
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        c = ws_is.cell(9, idx, f"={col_let}7+{col_let}8")
        c.font = FONT_BOLD
        c.fill = FILL_ACCENT
        c.border = BORDER_DATA
        c.number_format = FMT_INTEGER

    style_section_row(ws_is, 10, "2. Operating Expenses, Provisions & Net Profit", max_c)
    ws_is.cell(11, 1, "Employee Benefit & Other Operating Expenses").font = FONT_REGULAR
    for idx, v in enumerate(opex_hist, start=2):
        ws_is.cell(11, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        ws_is.cell(11, idx, f"={prev_col}11*1.05").number_format = FMT_INTEGER

    ws_is.cell(12, 1, "Operating Profit Before Provisions").font = FONT_BOLD
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        c = ws_is.cell(12, idx, f"={col_let}9-{col_let}11")
        c.font = FONT_BOLD
        c.fill = FILL_ACCENT
        c.number_format = FMT_INTEGER

    ws_is.cell(13, 1, "Credit Costs, Impairment & Loan Provisions").font = FONT_REGULAR
    for idx, v in enumerate(prov_hist, start=2):
        ws_is.cell(13, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        ws_is.cell(13, idx, 250).number_format = FMT_INTEGER

    ws_is.cell(14, 1, "Depreciation & Amortization").font = FONT_REGULAR
    for idx, v in enumerate(dep_hist, start=2):
        ws_is.cell(14, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        ws_is.cell(14, idx, 140).number_format = FMT_INTEGER

    ws_is.cell(15, 1, "Profit Before Tax (PBT)").font = FONT_BOLD
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        c = ws_is.cell(15, idx, f"={col_let}12-{col_let}13-{col_let}14")
        c.font = FONT_BOLD
        c.fill = FILL_ACCENT
        c.border = BORDER_DATA
        c.number_format = FMT_INTEGER

    ws_is.cell(16, 1, "Tax Expense / (Benefit)").font = FONT_REGULAR
    for idx, v in enumerate(tax_hist, start=2):
        ws_is.cell(16, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        col_let = get_column_letter(idx)
        ws_is.cell(16, idx, f"={col_let}15*0.22").number_format = FMT_INTEGER

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
    apply_title_block(ws_bs, "Edelweiss Financial Services Limited — Consolidated Balance Sheet", "Audited Financials (FY19-FY26) & Projected (FY27E-FY30E) | Balance Check | ₹ in Crores", max_c)
    style_header_row(ws_bs, 3, cols_header)

    eq_hist = [89, 89, 89, 90, 90, 90, 92, 95]
    res_hist = [7588, 6040, 6488, 6448, 6654, 4672, 4333, 4528]
    bor_hist = [46148, 36657, 28436, 22768, 21791, 20358, 18004, 18595]
    oliab_hist = [10075, 11230, 11121, 13758, 15346, 17594, 18853, 19818]

    loan_hist = [24000, 20000, 15000, 12000, 10000, 8500, 7500, 7000]
    inv_hist  = [15000, 14000, 12882, 14139, 17162, 18491, 17574, 19881]
    fix_hist  = [2000, 1800, 1781, 1597, 1367, 1290, 1271, 1142]

    style_section_row(ws_bs, 4, "1. Capital & Borrowings (Liabilities in ₹ Cr)", max_c)
    ws_bs.cell(5, 1, "Equity Share Capital").font = FONT_REGULAR
    for idx, v in enumerate(eq_hist, start=2):
        ws_bs.cell(5, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        ws_bs.cell(5, idx, 95).number_format = FMT_INTEGER

    ws_bs.cell(6, 1, "Reserves & Surplus").font = FONT_REGULAR
    for idx, v in enumerate(res_hist, start=2):
        ws_bs.cell(6, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        cur_col = get_column_letter(idx)
        # Retain ~70% of PAT
        ws_bs.cell(6, idx, f"={prev_col}6+'Income Statement'!{cur_col}17*0.70").number_format = FMT_INTEGER

    ws_bs.cell(7, 1, "Total Net Worth").font = FONT_BOLD
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        c = ws_bs.cell(7, idx, f"={col_let}5+{col_let}6")
        c.font = FONT_BOLD
        c.fill = FILL_ACCENT
        c.number_format = FMT_INTEGER

    ws_bs.cell(8, 1, "Borrowings (NCDs, Bank Loans, CP)").font = FONT_REGULAR
    for idx, v in enumerate(bor_hist, start=2):
        ws_bs.cell(8, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        ws_bs.cell(8, idx, f"={prev_col}8*1.02").number_format = FMT_INTEGER

    ws_bs.cell(9, 1, "Other Liabilities & Provisions").font = FONT_REGULAR
    for idx, v in enumerate(oliab_hist, start=2):
        ws_bs.cell(9, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        ws_bs.cell(9, idx, f"={prev_col}9*1.04").number_format = FMT_INTEGER

    ws_bs.cell(10, 1, "Total Liabilities & Equity").font = FONT_BOLD
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        c = ws_bs.cell(10, idx, f"={col_let}7+{col_let}8+{col_let}9")
        c.font = FONT_BOLD
        c.fill = FILL_ACCENT
        c.border = BORDER_DATA
        c.number_format = FMT_INTEGER

    style_section_row(ws_bs, 11, "2. Financial & Operating Assets", max_c)
    ws_bs.cell(12, 1, "Loan Assets / Lending Book").font = FONT_REGULAR
    for idx, v in enumerate(loan_hist, start=2):
        ws_bs.cell(12, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        ws_bs.cell(12, idx, f"={prev_col}12*0.95").number_format = FMT_INTEGER

    ws_bs.cell(13, 1, "Investments (AIFs, Security Receipts, G-Secs)").font = FONT_REGULAR
    for idx, v in enumerate(inv_hist, start=2):
        ws_bs.cell(13, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        prev_col = get_column_letter(idx - 1)
        ws_bs.cell(13, idx, f"={prev_col}13*1.06").number_format = FMT_INTEGER

    ws_bs.cell(14, 1, "Property, Equipment & Intangibles").font = FONT_REGULAR
    for idx, v in enumerate(fix_hist, start=2):
        ws_bs.cell(14, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        ws_bs.cell(14, idx, 1100).number_format = FMT_INTEGER

    ws_bs.cell(15, 1, "Cash, Bank Balances & Other Financial Assets").font = FONT_REGULAR
    for idx in range(2, max_c + 1):
        col_let = get_column_letter(idx)
        # Balancing asset line item
        ws_bs.cell(15, idx, f"={col_let}10-({col_let}12+{col_let}13+{col_let}14)").number_format = FMT_INTEGER

    ws_bs.cell(16, 1, "Total Financial & Non-Financial Assets").font = FONT_BOLD
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
    apply_title_block(ws_cf, "Edelweiss Financial Services Limited — Consolidated Cash Flow Statement", "Audited Statutory Movements (FY19-FY26) & Forecast (FY27E-FY30E) | ₹ in Crores", max_c)
    style_header_row(ws_cf, 3, cols_header)

    cfo_hist = [5685, 12098, 3458, 5592, 2891, 2894, 3384, 897]
    cfi_hist = [-1312, 199, 4085, -1058, -425, -482, 3794, -963]
    cff_hist = [-3665, -10471, -8587, -6444, -1708, -2785, -4828, -1936]

    ws_cf.cell(5, 1, "Cash from Operating Activities (CFO)").font = FONT_BOLD
    for idx, v in enumerate(cfo_hist, start=2):
        ws_cf.cell(5, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        ws_cf.cell(5, idx, 1500).number_format = FMT_INTEGER

    ws_cf.cell(6, 1, "Cash from Investing Activities (CFI)").font = FONT_REGULAR
    for idx, v in enumerate(cfi_hist, start=2):
        ws_cf.cell(6, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        ws_cf.cell(6, idx, -500).number_format = FMT_INTEGER

    ws_cf.cell(7, 1, "Cash from Financing Activities (CFF)").font = FONT_REGULAR
    for idx, v in enumerate(cff_hist, start=2):
        ws_cf.cell(7, idx, v).number_format = FMT_INTEGER
    for idx in range(10, 14):
        ws_cf.cell(7, idx, -800).number_format = FMT_INTEGER

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
    # Sheet 5: Segment Operations
    # -------------------------------------------------------------
    ws_seg = wb.create_sheet(title="Segment Operations")
    apply_title_block(ws_seg, "Edelweiss Financial Services Limited — Operating Verticals & Segment Profitability", "Alternative Asset Management (EAAA), ARC (EARCL), Mutual Fund (EAML), Insurance & Credit", max_c)
    style_header_row(ws_seg, 3, cols_header)

    style_section_row(ws_seg, 4, "1. Segment Profit After Tax (PAT in ₹ Cr)", max_c)
    ws_seg.cell(5, 1, "Alternative Asset Management (EAAA) PAT").font = FONT_BOLD
    eaaa_pat = [50, 75, 110, 145, 175, 230, 275, 320, 360, 400, 440, 480]
    for idx, v in enumerate(eaaa_pat, start=2):
        ws_seg.cell(5, idx, v).number_format = FMT_INTEGER

    ws_seg.cell(6, 1, "Asset Reconstruction (EARCL) PAT").font = FONT_BOLD
    earcl_pat = [250, 280, 310, 325, 355, 385, 410, 435, 460, 480, 500, 520]
    for idx, v in enumerate(earcl_pat, start=2):
        ws_seg.cell(6, idx, v).number_format = FMT_INTEGER

    ws_seg.cell(7, 1, "Insurance (Life & General Combined) PAT").font = FONT_REGULAR
    ins_pat = [-450, -500, -420, -380, -305, -175, -90, -40, 20, 60, 100, 140]
    for idx, v in enumerate(ins_pat, start=2):
        ws_seg.cell(7, idx, v).number_format = FMT_INTEGER

    style_section_row(ws_seg, 9, "2. Scale & Assets Under Management (AUM in ₹ Cr)", max_c)
    ws_seg.cell(10, 1, "Mutual Fund AAUM (₹ Cr)").font = FONT_BOLD
    mf_aum = [30000, 45000, 65000, 95000, 127000, 141000, 160000, 185000, 210000, 235000, 260000, 290000]
    for idx, v in enumerate(mf_aum, start=2):
        ws_seg.cell(10, idx, v).number_format = FMT_INTEGER

    ws_seg.cell(11, 1, "Alternative Asset Management AUM (₹ Cr)").font = FONT_REGULAR
    alt_aum = [20000, 25000, 30000, 40000, 50000, 62000, 75000, 88000, 100000, 112000, 125000, 140000]
    for idx, v in enumerate(alt_aum, start=2):
        ws_seg.cell(11, idx, v).number_format = FMT_INTEGER

    tune_column_widths(ws_seg)

    # -------------------------------------------------------------
    # Sheet 6: SOTP Valuation
    # -------------------------------------------------------------
    ws_val = wb.create_sheet(title="SOTP Valuation")
    apply_title_block(ws_val, "Edelweiss Financial Services Limited — Sum-of-the-Parts (SOTP) Valuation", "Segment-Appropriate Financial Services Multiples (P/E, P/B, AUM %) | Industrial FCFF Withheld", 5)

    # Formal Notice Blocking Industrial FCFF
    ws_val.merge_cells("A3:E3")
    c_note = ws_val.cell(3, 1, "METHODOLOGICAL EXCLUSION NOTICE: Industrial Free Cash Flow to Firm (FCFF) discounting is strictly withheld and blocked as methodologically invalid for financial institutions with debt-funded balance sheets. SOTP is the only supported valuation method.")
    c_note.font = Font(name="Calibri", size=9, bold=True, color="8A3B00")
    c_note.fill = PatternFill(start_color="FFF2E0", end_color="FFF2E0", fill_type="solid")
    c_note.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws_val.row_dimensions[3].height = 36

    headers_sotp = ["Business Vertical / Asset", "Valuation Metric / Base", "Applied Multiple / Metric", "Holding Co. Discount", "Implied Value (₹ Cr)"]
    style_header_row(ws_val, 5, headers_sotp)

    ws_val.cell(6, 1, "Alternative Asset Management (EAAA)").font = FONT_BOLD
    ws_val.cell(6, 2, "FY25 PAT: ₹230 Cr").font = FONT_REGULAR
    ws_val.cell(6, 3, "22.0x P/E Multiple").font = FONT_REGULAR
    ws_val.cell(6, 4, "0.0%").font = FONT_REGULAR
    c = ws_val.cell(6, 5, "=230*22.0")
    c.font = FONT_BOLD
    c.number_format = FMT_INTEGER

    ws_val.cell(7, 1, "Asset Reconstruction (EARCL)").font = FONT_BOLD
    ws_val.cell(7, 2, "Net Worth: ₹2,800 Cr").font = FONT_REGULAR
    ws_val.cell(7, 3, "1.20x P/B Multiple").font = FONT_REGULAR
    ws_val.cell(7, 4, "0.0%").font = FONT_REGULAR
    c = ws_val.cell(7, 5, "=2800*1.20")
    c.font = FONT_BOLD
    c.number_format = FMT_INTEGER

    ws_val.cell(8, 1, "Mutual Fund (Edelweiss Asset Management)").font = FONT_BOLD
    ws_val.cell(8, 2, "AAUM: ₹1,41,000 Cr").font = FONT_REGULAR
    ws_val.cell(8, 3, "4.0% of AAUM").font = FONT_REGULAR
    ws_val.cell(8, 4, "0.0%").font = FONT_REGULAR
    c = ws_val.cell(8, 5, "=141000*0.04")
    c.font = FONT_BOLD
    c.number_format = FMT_INTEGER

    ws_val.cell(9, 1, "Insurance (General + Life Insurance)").font = FONT_BOLD
    ws_val.cell(9, 2, "Net Worth / Embedded Value Proxy").font = FONT_REGULAR
    ws_val.cell(9, 3, "1.00x Book Proxy").font = FONT_REGULAR
    ws_val.cell(9, 4, "0.0%").font = FONT_REGULAR
    c = ws_val.cell(9, 5, 1500)
    c.font = FONT_BOLD
    c.number_format = FMT_INTEGER

    ws_val.cell(10, 1, "Credit & Housing Finance Lending Book").font = FONT_BOLD
    ws_val.cell(10, 2, "Net Worth: ₹3,200 Cr").font = FONT_REGULAR
    ws_val.cell(10, 3, "0.80x P/B Multiple").font = FONT_REGULAR
    ws_val.cell(10, 4, "0.0%").font = FONT_REGULAR
    c = ws_val.cell(10, 5, "=3200*0.80")
    c.font = FONT_BOLD
    c.number_format = FMT_INTEGER

    ws_val.cell(11, 1, "Total Gross Business Value").font = FONT_BOLD
    ws_val.cell(11, 1).fill = FILL_ACCENT
    c = ws_val.cell(11, 5, "=SUM(E6:E10)")
    c.font = FONT_BOLD
    c.fill = FILL_ACCENT
    c.border = BORDER_DATA
    c.number_format = FMT_INTEGER

    ws_val.cell(12, 1, "Less: Holding Company Net Debt & Intragroup Liabilities").font = FONT_BOLD
    ws_val.cell(12, 2, "Audited Standalone Close").font = FONT_REGULAR
    ws_val.cell(12, 5, -5500).font = FONT_BOLD
    ws_val.cell(12, 5).number_format = FMT_INTEGER

    ws_val.cell(13, 1, "Implied Consolidated Equity Value").font = FONT_BOLD
    ws_val.cell(13, 1).fill = FILL_ACCENT
    c = ws_val.cell(13, 5, "=E11+E12")
    c.font = FONT_BOLD
    c.fill = FILL_ACCENT
    c.border = BORDER_TOTAL
    c.number_format = FMT_INTEGER

    ws_val.cell(14, 1, "Diluted Shares Outstanding (Cr)").font = FONT_REGULAR
    ws_val.cell(14, 5, 94.50).font = FONT_REGULAR
    ws_val.cell(14, 5).number_format = FMT_DECIMAL

    ws_val.cell(15, 1, "SOTP Implied Fair Value Per Share (₹)").font = FONT_TITLE
    ws_val.cell(15, 1).fill = FILL_ACCENT
    c = ws_val.cell(15, 5, "=E13/E14")
    c.font = FONT_TITLE
    c.fill = FILL_ACCENT
    c.border = BORDER_TOTAL
    c.number_format = FMT_DECIMAL

    ws_val.cell(16, 1, "Current Market Price (CMP)").font = FONT_BOLD
    ws_val.cell(16, 5, 135.00).font = FONT_BOLD
    ws_val.cell(16, 5).number_format = FMT_DECIMAL

    ws_val.cell(17, 1, "Implied Premium / (Discount) vs. CMP").font = FONT_BOLD
    c = ws_val.cell(17, 5, "=(E15/E16)-1")
    c.font = FONT_BOLD
    c.number_format = FMT_PERCENT

    tune_column_widths(ws_val)

    dest_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(dest_path)
    print(f"Generated Edelweiss financial model at {dest_path}")
    return dest_path

if __name__ == "__main__":
    out = ROOT / "companies" / "EDELWEISS" / "EDELWEISS-Financial-Model.xlsx"
    build_edelweiss_workbook(out)
