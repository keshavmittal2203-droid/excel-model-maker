"""Build authentic, fully articulated Excel models for M&M, LTM, CLEAN, and EDELWEISS."""
import json
import os
import sys
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent.parent

# Corporate Theme Colors
NAVY_HEADER = "1B365D"      # Dark Navy Blue
SLATE_SECTION = "2C4D75"    # Medium Slate Blue
ACCENT_FILL = "E8EEF5"      # Very Light Blue Accent
ZEBRA_FILL = "F9FAFC"       # Off-white Zebra
WHITE = "FFFFFF"
BORDER_GRAY = "D3D3D3"
DARK_GRAY = "333333"

# Fonts
FONT_TITLE = Font(name="Calibri", size=14, bold=True, color="1B365D")
FONT_SUBTITLE = Font(name="Calibri", size=10, italic=True, color="555555")
FONT_HEADER = Font(name="Calibri", size=11, bold=True, color=WHITE)
FONT_SECTION = Font(name="Calibri", size=11, bold=True, color=WHITE)
FONT_BOLD = Font(name="Calibri", size=10, bold=True, color=DARK_GRAY)
FONT_REGULAR = Font(name="Calibri", size=10, bold=False, color=DARK_GRAY)
FONT_DISC_OK = Font(name="Calibri", size=10, bold=True, color="008000")
FONT_DISC_ERR = Font(name="Calibri", size=10, bold=True, color="FF0000")

# Fills
FILL_HEADER = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
FILL_SECTION = PatternFill(start_color=SLATE_SECTION, end_color=SLATE_SECTION, fill_type="solid")
FILL_ACCENT = PatternFill(start_color=ACCENT_FILL, end_color=ACCENT_FILL, fill_type="solid")
FILL_ZEBRA = PatternFill(start_color=ZEBRA_FILL, end_color=ZEBRA_FILL, fill_type="solid")

# Borders
THIN_GRAY = Side(style="thin", color=BORDER_GRAY)
MEDIUM_GRAY = Side(style="medium", color=NAVY_HEADER)
DOUBLE_BOTTOM = Side(style="double", color=NAVY_HEADER)

BORDER_DATA = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
BORDER_TOTAL = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=DOUBLE_BOTTOM)
BORDER_SECTION = Border(left=THIN_GRAY, right=THIN_GRAY, top=MEDIUM_GRAY, bottom=MEDIUM_GRAY)

# Alignments
ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")
ALIGN_CENTER = Alignment(horizontal="center", vertical="center")

# Number Formats
FMT_INTEGER = "#,##0"
FMT_DECIMAL = "#,##0.00"
FMT_PERCENT = "0.0%"
FMT_DATE = "YYYY-MM-DD"


def apply_title_block(ws, title: str, subtitle: str, max_col: int):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    c1 = ws.cell(1, 1, title)
    c1.font = FONT_TITLE
    c1.alignment = ALIGN_LEFT
    ws.row_dimensions[1].height = 24

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
    c2 = ws.cell(2, 1, subtitle)
    c2.font = FONT_SUBTITLE
    c2.alignment = ALIGN_LEFT
    ws.row_dimensions[2].height = 18


def style_header_row(ws, row_idx: int, cols: list[str]):
    ws.row_dimensions[row_idx].height = 24
    for col_idx, col_name in enumerate(cols, start=1):
        cell = ws.cell(row_idx, col_idx, col_name)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = ALIGN_CENTER if col_idx > 1 else ALIGN_LEFT
        cell.border = BORDER_DATA


def style_section_row(ws, row_idx: int, title: str, max_col: int):
    ws.row_dimensions[row_idx].height = 20
    ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=max_col)
    cell = ws.cell(row_idx, 1, title)
    cell.font = FONT_SECTION
    cell.fill = FILL_SECTION
    cell.alignment = ALIGN_LEFT
    for c in range(1, max_col + 1):
        ws.cell(row_idx, c).border = BORDER_SECTION


def tune_column_widths(ws, min_width: int = 14, max_width: int = 42):
    ws.views.sheetView[0].showGridLines = True
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            # Skip title merged row
            if cell.row in [1, 2]:
                continue
            v = str(cell.value or "")
            if len(v) > max_len:
                max_len = len(v)
        col_width = max(min_width, min(max_len + 4, max_width))
        if col[0].column == 1:
            col_width = max(col_width, 28)
        ws.column_dimensions[col_letter].width = col_width


print("Loaded styling infrastructure for four targets.")
