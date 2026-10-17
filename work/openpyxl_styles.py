"""Shared Institutional OpenPyXL Styling & Formatting Helpers for Excel Model Maker."""
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Institutional Typography
FONT_TITLE = Font(name="Calibri", size=14, bold=True, color="1F4E79")
FONT_SUBTITLE = Font(name="Calibri", size=10, italic=True, color="333333")
FONT_SEC_HDR = Font(name="Calibri", size=11, bold=True, color="1F4E79")
FONT_TBL_HDR = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
FONT_REGULAR = Font(name="Calibri", size=10, bold=False, color="000000")
FONT_BOLD = Font(name="Calibri", size=10, bold=True, color="000000")
FONT_FORMULA = Font(name="Calibri", size=10, bold=False, color="000000")
FONT_LINK = Font(name="Calibri", size=10, bold=False, color="006100")
FONT_SMALL_ITALIC = Font(name="Calibri", size=9, italic=True, color="444444")

# Institutional Fills
FILL_NAVY = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
FILL_SEC_HDR = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
FILL_SUB_HDR = PatternFill(start_color="EDEDED", end_color="EDEDED", fill_type="solid")
FILL_ZEBRA = PatternFill(start_color="F9FAFC", end_color="F9FAFC", fill_type="solid")
FILL_TOTAL = PatternFill(start_color="EAEEF7", end_color="EAEEF7", fill_type="solid")
FILL_ACCENT = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

# Borders
THIN_SIDE = Side(border_style="thin", color="D9D9D9")
MED_NAVY = Side(border_style="medium", color="1F4E79")
DOUBLE_BOTTOM = Side(border_style="double", color="1F4E79")
TOP_THIN_NAVY = Side(border_style="thin", color="1F4E79")

BORDER_DATA = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
BORDER_TOTAL = Border(top=TOP_THIN_NAVY, bottom=DOUBLE_BOTTOM, left=THIN_SIDE, right=THIN_SIDE)
BORDER_HEADER = Border(left=THIN_SIDE, right=THIN_SIDE, top=MED_NAVY, bottom=MED_NAVY)

# Alignments
ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")
ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)

# Number Formats
FMT_INR_CR = "#,##0"
FMT_DEC2 = "#,##0.00"
FMT_PCT = "0.0%"
FMT_INT = "#,##0"
FMT_MULT = "0.0\"x\""


def style_range(ws, min_row: int, min_col: int, max_row: int, max_col: int, font=None, fill=None, border=None, alignment=None):
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            cell = ws.cell(r, c)
            if font is not None:
                cell.font = font
            if fill is not None:
                cell.fill = fill
            if border is not None:
                cell.border = border
            if alignment is not None:
                cell.alignment = alignment


def set_col_widths(ws, widths: dict[str, float]):
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width


def setup_page(ws, header_title: str = ""):
    ws.views.sheetView[0].showGridLines = True
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    if header_title:
        ws.oddHeader.center.text = header_title
