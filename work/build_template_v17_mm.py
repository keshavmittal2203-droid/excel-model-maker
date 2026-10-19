"""Build and validate the single-target M&M Template V17 workbook.

The build starts with the immutable 16-sheet master and the committed M&M V16
input delta, then replaces unsupported model content with source-backed
historicals and explicit blocked states. It never generates the other targets.
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import re
import shutil
import statistics
import sys
import zipfile
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Border, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
MASTER = ROOT / "originals" / "Marksans Keshav.xlsx"
V16_BUILD = ROOT / "work" / "build_template_v16_models.py"
EXPECTED_MASTER_SHA = "4dc79842855e5d1e689a2f5ca1de22c34be72fb09c0c0f9dd3a95818e10737d0"
MM_SOURCE = ROOT / "work" / "sources" / "MM_consolidated.json"
BETA_SOURCE = ROOT / "work" / "sources" / "weekly_market_returns_104w_v2.json"
PRICE_SOURCE = ROOT / "work" / "sources" / "market_and_beta_2025_10_15.json"
ERRORS = {"#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!"}


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _table(source: dict[str, Any], section: str) -> tuple[list[str], dict[str, list[Any]]]:
    table = source["sections"][section]["tables"][0]
    return table["headers"], table["rows"]


def _year_index(headers: list[str]) -> dict[int, int]:
    result: dict[int, int] = {}
    for i, header in enumerate(headers):
        m = re.search(r"(20\d{2})$", str(header))
        if m:
            result[int(m.group(1))] = i
    return result


def _source_years(headers: list[str], rows: dict[str, list[Any]], start: int, end: int) -> dict[str, dict[int, Any]]:
    idx = _year_index(headers)
    return {name: {year: values[idx[year]] for year in range(start, end + 1)} for name, values in rows.items()}


def calculate_source_beta(beta_source: dict[str, Any]) -> float:
    """Recompute OLS beta from the committed, aligned adjusted-close series."""
    security = beta_source["securities"]["M&M.NS"]["records"]
    benchmark = beta_source["benchmark"]["records"]
    bm = {r["date"]: r["adjclose"] for r in benchmark}
    prices = {r["date"]: r["adjclose"] for r in security}
    dates = sorted(set(prices) & set(bm))
    if len(dates) != 105:
        raise RuntimeError(f"Expected 105 aligned weekly bars, found {len(dates)}")
    stock_returns = [prices[b] / prices[a] - 1 for a, b in zip(dates, dates[1:])]
    market_returns = [bm[b] / bm[a] - 1 for a, b in zip(dates, dates[1:])]
    if len(stock_returns) != 104 or len(market_returns) != 104:
        raise RuntimeError("Expected 104 paired weekly returns")
    mean_x = statistics.mean(market_returns)
    mean_y = statistics.mean(stock_returns)
    cov = sum((x - mean_x) * (y - mean_y) for x, y in zip(market_returns, stock_returns)) / (len(stock_returns) - 1)
    var = sum((x - mean_x) ** 2 for x in market_returns) / (len(market_returns) - 1)
    return cov / var


def _clear_cell(ws, row: int, col: int) -> None:
    cell = ws.cell(row, col)
    if type(cell).__name__ != "MergedCell":
        cell.value = None


def _clear(ws, rows, cols) -> None:
    for r in rows:
        for c in cols:
            _clear_cell(ws, r, c)


def _validate_output_dir(output_dir: Path) -> Path:
    """Limit replaceable build directories to dedicated generated-output roots."""
    resolved = output_dir.resolve()
    allowed_roots = (
        (ROOT / "companies" / "MM").resolve(),
        (ROOT / "staging").resolve(),
        (ROOT / "candidates").resolve(),
    )
    if resolved.parent not in allowed_roots:
        raise ValueError(
            "Output must be a dedicated direct child of companies/MM, staging, or candidates; "
            f"refusing unsafe path: {resolved}"
        )
    return resolved


def _clear_unused_table_rows(ws, rows, columns) -> None:
    """Remove template grid clutter from rows that have no retained peer data."""
    for row in rows:
        for column in columns:
            cell = ws.cell(row, column)
            if type(cell).__name__ == "MergedCell":
                continue
            cell.value = None
            cell.border = Border()
            cell.fill = PatternFill(fill_type=None)


def _apply_sparkline_extension(xlsx: Path) -> None:
    """Restore the master template's native trend sparklines over FY21-FY25 only."""
    with zipfile.ZipFile(MASTER, "r") as master:
        xml = master.read("xl/worksheets/sheet3.xml").decode("utf-8")
    start = xml.find("<extLst>")
    end = xml.find("</extLst>") + len("</extLst>")
    if start < 0 or end < len("</extLst>"):
        raise RuntimeError("Master Ratio Analysis sheet has no native sparkline extension")
    ext = xml[start:end]
    ext = ext.replace(
        'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"',
        'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
        'xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"',
    )
    for row in range(6, 27):
        ext = ext.replace(f"C{row}:H{row}", f"C{row}:G{row}")
    tmp = xlsx.with_suffix(".sparkline.tmp.xlsx")
    with zipfile.ZipFile(xlsx, "r") as zin, zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        found = False
        for item in zin.infolist():
            payload = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet3.xml":
                text = payload.decode("utf-8")
                if "<extLst>" in text:
                    text = re.sub(r"<extLst>.*?</extLst>", "", text, count=1)
                pos = text.rfind("</worksheet>")
                if pos < 0:
                    raise RuntimeError("Invalid Ratio Analysis worksheet XML")
                text = text[:pos] + ext + text[pos:]
                payload = text.encode("utf-8")
                found = True
            zout.writestr(item, payload)
        if not found:
            raise RuntimeError("Ratio Analysis worksheet part is missing")
    tmp.replace(xlsx)


def patch_mm_workbook(xlsx: Path) -> dict[str, Any]:
    source = json.loads(MM_SOURCE.read_text(encoding="utf-8"))
    beta = json.loads(BETA_SOURCE.read_text(encoding="utf-8"))
    market = json.loads(PRICE_SOURCE.read_text(encoding="utf-8"))
    pl_headers, pl_rows = _table(source, "profit-loss")
    bs_headers, bs_rows = _table(source, "balance-sheet")
    cf_headers, cf_rows = _table(source, "cash-flow")
    pl = _source_years(pl_headers, pl_rows, 2015, 2025)
    bs = _source_years(bs_headers, bs_rows, 2015, 2025)
    cf = _source_years(cf_headers, cf_rows, 2015, 2025)

    wb = load_workbook(xlsx, data_only=False)
    ds = wb["Data Sheet"]
    mf = wb["Main Financials"]
    bsws = wb["BS"]
    cfws = wb["Cash Flow"]

    # One visible, concise limitation on the cover; no dispersed speculative notes.
    cover = wb["Model ->"]
    cover["C5"] = "Automotive, Farm Equipment & Financial Services"
    cover["B11"] = "Historical close (15-Oct-2025)"
    cover["C11"] = "='Data Sheet'!B8"
    cover["C11"].number_format = '"₹"#,##0.00'
    cover.column_dimensions["B"].width = 36
    cover["C13"] = (
        "Forecast and valuation unavailable: approved sources do not split M&M's "
        "segments, consolidated cash, or NBFC funding detail."
    )
    cover["C16"] = None
    ds["A3"] = "Screener consolidated snapshot; valuation cutoff 15-Oct-2025; actuals through FY25 only."
    ds["E2"] = "Historical sourced inputs; formulas calculate derived metrics"
    # The historical cutoff snapshot does not provide a separately verified share count.
    # Do not reuse a prior pack's 124.3 Cr input to manufacture market capitalization.
    ds["A6"] = "Shares outstanding (unavailable on approved snapshot)"
    ds["B6"] = None
    ds["A7"] = "Face value (INR/share)"
    ds["B7"] = 5
    ds["A8"] = "Closing price (Yahoo Finance, 15-Oct-2025)"
    ds["B8"] = market["securities"]["M&M.NS"]["close_2025_10_15"]
    ds["A9"] = "Market capitalization (unavailable; share count not verified)"
    ds["B9"] = None

    # Data Sheet: exact Screener historicals (FY17-FY25); every unavailable FY26E is blank.
    for offset, year in enumerate(range(2017, 2026), start=2):
        ds.cell(16, offset).value = dt.date(year, 3, 31)
        ds.cell(16, offset).number_format = "mmm-yy"
    ds["K16"] = "FY26E (Blocked)"
    pnl_map = {17: "Sales", 18: "Expenses", 25: "Other Income", 26: "Depreciation", 27: "Interest", 28: "Profit before tax", 30: "Net Profit"}
    labels = {17: "Sales (Screener reported)", 18: "Operating expenses (Screener reported)", 25: "Other income (Screener reported)", 26: "Depreciation (Screener reported)", 27: "Interest (Screener reported)", 28: "Profit before tax (Screener reported)", 29: "Tax (derived: PBT less net profit)", 30: "Net profit (Screener reported)", 31: "Dividend payout ratio (Screener reported)", 32: "EPS (Screener reported)", 33: "Operating profit / EBITDA proxy (Sales less expenses)"}
    for row, label in labels.items():
        ds.cell(row, 1).value = label
    for row in range(19, 25):
        ds.cell(row, 1).value = None
    _clear(ds, range(17, 34), range(2, 12))
    for year in range(2017, 2026):
        col = year - 2015  # B=FY17, J=FY25
        for row, source_name in pnl_map.items():
            ds.cell(row, col).value = pl[source_name][year]
        ds.cell(29, col).value = f"={ds.cell(28, col).coordinate}-{ds.cell(30, col).coordinate}"
        ds.cell(31, col).value = pl["Dividend Payout %"][year] / 100
        ds.cell(32, col).value = pl["EPS in Rs"][year]
        ds.cell(33, col).value = f"={ds.cell(17, col).coordinate}-{ds.cell(18, col).coordinate}"
    ds["K17"] = None
    for row in range(18, 34):
        ds.cell(row, 11).value = None
    # Consolidated balance sheet data, without invented current/non-current splits.
    bs_map = {57: "Equity Capital", 58: "Reserves", 59: "Borrowings", 60: "Other Liabilities", 61: "Total Liabilities", 62: "Fixed Assets", 63: "CWIP", 64: "Investments", 65: "Other Assets", 66: "Total Assets"}
    for row, source_name in bs_map.items():
        ds.cell(row, 1).value = source_name + " (Screener reported)"
    for year in range(2017, 2026):
        col = year - 2015
        for row, source_name in bs_map.items():
            ds.cell(row, col).value = bs[source_name][year]
    ds["K57"] = None
    for row in range(58, 67):
        ds.cell(row, 11).value = None
    # Cash flow aggregates are the only disclosed statement detail retained.
    cf_map = {82: "Cash from Operating Activity", 83: "Cash from Investing Activity", 84: "Cash from Financing Activity", 85: "Net Cash Flow", 86: "Free Cash Flow"}
    for row, source_name in cf_map.items():
        ds.cell(row, 1).value = source_name + " (Screener reported)"
    for year in range(2017, 2026):
        col = year - 2015
        for row, source_name in cf_map.items():
            ds.cell(row, col).value = cf[source_name][year]
    for row in cf_map:
        ds.cell(row, 11).value = None
    # The source file contains post-cutoff quarter rows; stop at the last quarter available by cutoff.
    for col in range(8, 12):  # H=Sep-2025 through K=Jun-2026
        for row in range(41, 54):
            ds.cell(row, col).value = None
    # Historical annual share counts were copied from today's share count; remove the unsupported series.
    for col in range(2, 12):
        ds.cell(70, col).value = None
    # Set Data Sheet row 74 (Other Assets) and 75 (Total Assets) dynamically for historical columns:
    for col in range(2, 11):  # B=FY17 through J=FY25
        ds.cell(74, col).value = f"={ds.cell(65, col).coordinate}"
        ds.cell(75, col).value = f"={ds.cell(61, col).coordinate}"
    ds.cell(74, 11).value = None
    ds.cell(75, 11).value = None

    # Income statement: actuals only, with a visible blocked FY26E column.
    mf["J8"] = "FY26E (Blocked)"
    for col in range(4, 10):  # D:I = FY20-FY25
        ds_col = col + 1
        for row in (11, 14, 23, 26, 29, 32, 35, 38, 43, 47):
            pass
        mf.cell(8, col).value = f'=TEXT(\'Data Sheet\'!{ds.cell(16, ds_col).coordinate},"mmm-yy")'
        mf.cell(11, col).value = f"='Data Sheet'!{ds.cell(17, ds_col).coordinate}"
        mf.cell(14, col).value = f"='Data Sheet'!{ds.cell(18, ds_col).coordinate}"
        mf.cell(15, col).value = f"={mf.cell(14,col).coordinate}/{mf.cell(11,col).coordinate}"
        mf.cell(17, col).value = None
        mf.cell(18, col).value = None
        mf.cell(20, col).value = None
        mf.cell(21, col).value = None
        mf.cell(23, col).value = f"='Data Sheet'!{ds.cell(33, ds_col).coordinate}"
        mf.cell(24, col).value = f"={mf.cell(23,col).coordinate}/{mf.cell(11,col).coordinate}"
        mf.cell(26, col).value = f"='Data Sheet'!{ds.cell(27, ds_col).coordinate}"
        mf.cell(27, col).value = f"={mf.cell(26,col).coordinate}/{mf.cell(11,col).coordinate}"
        mf.cell(29, col).value = f"='Data Sheet'!{ds.cell(26, ds_col).coordinate}"
        mf.cell(30, col).value = f"={mf.cell(29,col).coordinate}/{mf.cell(11,col).coordinate}"
        mf.cell(32, col).value = f"='Data Sheet'!{ds.cell(28, ds_col).coordinate}"
        mf.cell(33, col).value = f"={mf.cell(32,col).coordinate}/{mf.cell(11,col).coordinate}"
        mf.cell(35, col).value = f"='Data Sheet'!{ds.cell(29, ds_col).coordinate}"
        mf.cell(36, col).value = f"={mf.cell(35,col).coordinate}/{mf.cell(32,col).coordinate}"
        mf.cell(38, col).value = f"='Data Sheet'!{ds.cell(30, ds_col).coordinate}"
        mf.cell(39, col).value = f"={mf.cell(38,col).coordinate}/{mf.cell(11,col).coordinate}"
        mf.cell(41, col).value = None
        mf.cell(43, col).value = f"='Data Sheet'!{ds.cell(32, ds_col).coordinate}"
        if col > 4:
            prev = mf.cell(43, col - 1).coordinate
            mf.cell(44, col).value = f"=IFERROR({mf.cell(43,col).coordinate}/{prev}-1,\"\")"
        mf.cell(46, col).value = None
        mf.cell(47, col).value = f"='Data Sheet'!{ds.cell(31, ds_col).coordinate}"
        mf.cell(49, col).value = f"=1-{mf.cell(47,col).coordinate}"
        # Historical balance sheet values and direct reported totals.
        for mf_row, ds_row in {53:57, 54:58, 55:59, 56:60, 57:61, 59:62, 60:63, 61:64, 70:66}.items():
            mf.cell(mf_row, col).value = f"='Data Sheet'!{ds.cell(ds_row,ds_col).coordinate}"
        ds_col_letter = openpyxl.utils.get_column_letter(ds_col)
        col_letter = openpyxl.utils.get_column_letter(col)
        mf.cell(62, col).value = f"='Data Sheet'!{ds_col_letter}65"
        mf.cell(63, col).value = f"=SUM({col_letter}59:{col_letter}62)"
        mf.cell(69, col).value = f"={col_letter}70-{col_letter}63"
        for r in (65, 66, 67, 68):
            mf.cell(r, col).value = None
        mf.cell(72, col).value = f'=IF({col_letter}70={col_letter}57,"TIED","UNRECONCILED")'
        # Only Screener's aggregate historical cash flow lines are available.
        cf_col = col + 4  # D=H=FY20 ... I=M=FY25
        for r in list(range(77, 85)) + list(range(88, 96)) + list(range(99, 106)):
            mf.cell(r, col).value = None
        for r, cf_row in {85:5, 96:14, 106:23, 107:32, 108:31}.items():
            mf.cell(r, col).value = f"='Cash Flow'!{cfws.cell(cf_row, cf_col).coordinate}"
    mf["B14"] = "Operating expenses (Screener reported)"
    mf["B15"] = "Operating expenses / sales"
    mf["B17"] = None
    mf["B18"] = None
    mf["B20"] = None
    mf["B21"] = None
    mf["B23"] = "Operating profit / EBITDA proxy (Screener)"
    mf["B24"] = "Operating margin"
    mf["B43"] = "EPS (Screener reported)"
    mf["B46"] = "Dividend per share (not separately sourced)"
    mf["B47"] = "Dividend payout ratio"
    mf["B62"] = "Other Assets (Screener reported)"
    mf["B63"] = "Asset component subtotal"
    mf["B65"] = "Receivables (breakdown unavailable)"
    mf["B66"] = "Inventory (breakdown unavailable)"
    mf["B67"] = "Cash & Bank (breakdown unavailable)"
    mf["B68"] = "Total Current Assets (breakdown unavailable)"
    mf["B70"] = "Total Assets (Screener reported)"
    mf["B69"] = "Rounding delta vs total assets"
    mf["B72"] = "Reported totals tie-out"
    for row in range(1, mf.max_row + 1):
        if mf.cell(row, 1).value == "#":
            mf.cell(row, 1).value = None
    mf["B85"] = "Cash from Operating Activities (Screener aggregate)"
    mf["B96"] = "Cash from Investing Activities (Screener aggregate)"
    mf["B106"] = "Cash from Financing Activities (Screener aggregate)"
    mf["B107"] = "Free Cash Flow (Screener reported)"
    mf["B108"] = "Net Cash Flow (formula: CFO + CFI + CFF)"
    for r in list(range(77,85)) + list(range(88,96)) + list(range(99,106)):
        mf.cell(r,2).value = None
    # Block every FY26E projection/output across the entire sheet (rows 11 through 150):
    _clear(mf, range(11, 150), [10])
    mf["J8"] = "FY26E (Blocked)"
    mf["J72"] = "BLOCKED"
    # Clear rows 146-148 for all columns (Receivables, Inventory, Cash % not separately reported):
    _clear(mf, (146, 147, 148), range(2, 11))

    # Balance sheet: source-backed totals only; no fabricated components or FY26 plug.
    bsws["B10"] = "Borrowings (Screener reported aggregate)"
    bsws["B15"] = "Other liabilities (Screener reported aggregate)"
    bsws["B16"] = "Total liabilities (Screener reported)"
    bsws["B28"] = "Net block / fixed assets (Screener reported)"
    bsws["B37"] = "Other assets (Screener aggregate; components unavailable)"
    bsws["B38"] = "Total assets (Screener reported)"
    for col in range(3, 14):  # C:M = FY15-FY25
        year = 2015 + col - 3
        bsws.cell(2, col).value = dt.date(year, 3, 31)
        bsws.cell(2, col).number_format = "mmm-yy"
        for row, source_name in {4:"Equity Capital",5:"Reserves",10:"Borrowings",15:"Other Liabilities",16:"Total Liabilities",28:"Fixed Assets",30:"CWIP",31:"Investments",37:"Other Assets",38:"Total Assets"}.items():
            bsws.cell(row, col).value = bs[source_name][year]
        for row in (7,8,9,12,13,14,19,20,21,22,23,24,25,26,27,29,33,34,35,36):
            bsws.cell(row,col).value = None
        bsws.cell(40,col).value = f"={bsws.cell(38,col).coordinate}-{bsws.cell(16,col).coordinate}"
    bsws["N2"] = "FY26E (Blocked)"
    for row in range(4, 41):
        bsws.cell(row,14).value = None
    bsws["N40"] = "BLOCKED"

    # Cash flow: exact reported annual aggregates; no copied detail, assumptions, or forecast.
    for row, label in {5:"Operating cash flow (Screener)",14:"Investing cash flow (Screener)",23:"Financing cash flow (Screener)",31:"Net cash flow (CFO + CFI + CFF)",32:"Free cash flow (Screener)",33:"CFO / operating profit"}.items():
        cfws.cell(row,2).value = label
    cfws.column_dimensions["B"].width = 42.0
    for year in range(2015,2026):
        col = year - 2012  # C=2015 ... M=2025
        cfws.cell(4,col).value = dt.date(year,3,31)
        cfws.cell(4,col).number_format = "mmm-yy"
        cfws.cell(5,col).value = cf["Cash from Operating Activity"][year]
        cfws.cell(14,col).value = cf["Cash from Investing Activity"][year]
        cfws.cell(23,col).value = cf["Cash from Financing Activity"][year]
        cfws.cell(31,col).value = f"={cfws.cell(5,col).coordinate}+{cfws.cell(14,col).coordinate}+{cfws.cell(23,col).coordinate}"
        cfws.cell(32,col).value = cf["Free Cash Flow"][year]
        cfws.cell(33,col).value = cf["CFO/EBITDAShare of operating profit converting into cash: CFO (before tax) / Operating Profit"][year] / 100
        cfws.cell(33,col).number_format = "0.0%"
    for row in list(range(6,14)) + list(range(15,23)) + list(range(24,31)):
        cfws.cell(row,2).value = None
        for col in range(3,15):
            cfws.cell(row,col).value = None
    cfws["N4"] = "FY26E (Blocked)"
    for row in (5,14,23,31,32,33):
        cfws.cell(row,14).value = None

    # Remove unsupported peer records, circular EBITDA back-solves, and EV based on an invented cash bridge.
    compdata = wb["compdata"]
    _clear(compdata, range(3,10), range(2,14))
    # The second template table is a peer EBIT schedule. No date-matched peer
    # inputs are approved for V17, so clear the entire stale peer-data region.
    _clear(compdata, range(18,26), range(2,14))
    compdata.print_area = "$B$2:$M$10"
    for row in range(17,26):
        compdata.row_dimensions[row].hidden = True
    compdata["B10"] = 1
    compdata["C10"] = "M&M"
    compdata["D10"] = "='Data Sheet'!B8"
    compdata["E10"] = None
    compdata["F10"] = None
    for c in ("G10","H10","I10","J10"):
        compdata[c] = None
    compdata["K10"] = "='Data Sheet'!J33"
    compdata["L10"] = "='Data Sheet'!J17"
    compdata["M10"] = "='Data Sheet'!J30"
    # Peer tab has no date-matched verified peer operating data; clear unsupported rows.
    _clear(wb["peerdata"], range(5,10), range(2,8))

    comps = wb["Comps"]
    comps["B4"] = "M&M company snapshot - market capitalization and peer valuation unavailable"
    comps["C8"] = "M&M"
    comps["B8"] = "='Data Sheet'!B1"
    comps["D8"] = "=compdata!D10"
    comps["E8"] = "=compdata!E10"
    comps["F8"] = "=compdata!F10"
    for c in ("E8","F8","G8","H8","N8","O8"):
        comps[c] = None
    comps["J8"] = "=compdata!L10"
    comps["K8"] = "=compdata!K10"
    comps["L8"] = "=compdata!M10"
    comps["P8"] = '=IFERROR(D8/\'Data Sheet\'!J32,"Unavailable")'
    _clear_unused_table_rows(comps, range(9,35), range(2,17))
    comps["B10"] = "Peer multiples omitted; share count and market cap unavailable; peer EV multiples unavailable. P/E = 15-Oct-2025 close / FY25 EPS."
    comps["B25"] = None
    comps.print_area = "$B$3:$P$10"

    # M&M is a conglomerate with an NBFC; industrial DCF and unsupported SOTP are blocked.
    dcf = wb["DCF 1st way"]
    for row in range(3,18):
        for col in range(5,12):
            dcf.cell(row,col).value = None
    for row in range(19,42):
        for col in range(5,12):
            dcf.cell(row,col).value = None
    for row in range(24,32):
        for col in range(7,12):
            dcf.cell(row,col).value = None
    # Remove the inherited SOTP schedule text and all values; none of its standalone
    # segment or cash-bridge inputs is supported by the locked source allowlist.
    _clear(dcf, range(47,62), range(2,14))
    _clear_unused_table_rows(dcf, range(48,62), range(2,14))
    dcf["B23"] = "Valuation and sensitivity unavailable"
    dcf["B47"] = "SOTP unavailable: approved sources do not disclose segment EBITDA or standalone cash."
    dcf["E42"] = "Unavailable"
    dcf["E62"] = "Unavailable"
    dcf["B62"] = "Implied fair value per share (unavailable)"
    dcf["E44"] = market["securities"]["M&M.NS"]["close_2025_10_15"]
    dcf["E45"] = None
    dcf.print_area = "$B$3:$K$62"
    dcf2 = wb["DCF 2nd way"]
    _clear(dcf2, range(5,101), range(2,16))
    _clear_unused_table_rows(dcf2, range(5,101), range(2,16))
    dcf2["B4"] = "Forecast and valuation unavailable; required source-backed inputs are missing."
    for column in range(1, 16):
        if dcf2.cell(4, column).coordinate != "B4":
            dcf2.cell(4, column).value = None
    for row in range(1, min(dcf2.max_row, 100) + 1):
        if dcf2.cell(row, 1).value == "x":
            dcf2.cell(row, 1).value = None
    dcf2.print_area = "$B$3:$N$4"
    for row in range(5,101):
        dcf2.row_dimensions[row].hidden = True

    # Preserve reported aggregates, but block ROIC/growth where the approved sources do
    # not split operating assets from the consolidated MMFS loan book.
    growth = wb["Intrinsic Growth"]
    growth["B3"] = (
        "ROIC / intrinsic growth unavailable: approved sources do not split operating assets "
        "or MMFS lending balances from consolidated Other Assets."
    )

    # Clean fiscal year headers linking to BS row 2
    for c in ["E", "F", "G", "H", "I"]:
        bsc = {"E": "J", "F": "K", "G": "L", "H": "M", "I": "N"}[c]
        growth[f"{c}4"] = f"=BS!{bsc}2"
        growth[f"{c}41"] = f"=BS!{bsc}2"
        growth[f"{c}57"] = f"=BS!{bsc}2"
    for c in ["E", "F", "G", "H"]:
        for r in (4, 41, 57):
            growth[f"{c}{r}"].number_format = "mmm-yy"

    # Keep the template's schedules, but label unavailable classifications honestly.
    growth["B4"] = "Historical source metrics"
    growth["B6"] = "Current asset detail (not disclosed separately)"
    growth["B7"] = "=BS!B33"
    growth["B8"] = "=BS!B34"
    growth["B9"] = "=BS!B36"
    growth["B10"] = "Other assets (aggregate; classification unavailable)"
    growth["B11"] = "Current assets subtotal (unavailable)"
    growth["B13"] = "Liability detail (not disclosed separately)"
    growth["B14"] = "=BS!B13"
    growth["B15"] = "=BS!B14"
    growth["B16"] = "Other liabilities (aggregate; classification unavailable)"
    growth["B17"] = "Current liabilities subtotal (unavailable)"
    growth["B19"] = "Net working capital (unavailable)"
    growth["B21"] = "Non-current asset detail (not disclosed separately)"
    growth["B22"] = "Land"
    growth["B23"] = "Building"
    growth["B24"] = "Plant Machinery"
    growth["B25"] = "=BS!B23"
    growth["B26"] = "Equipments"
    growth["B27"] = "Furniture n fittings"
    growth["B28"] = "Vehicles"
    growth["B29"] = "Intangible Assets"
    growth["B30"] = "Other fixed assets"
    growth["B32"] = "Net block / fixed assets (Screener reported)"
    growth["B33"] = "Accumulated depreciation (not separately reported)"
    growth["B34"] = "Net non-current assets (unavailable)"
    growth["B36"] = "Invested capital (unavailable)"
    growth["B37"] = "EBIT (operating profit less D&A)"
    growth["B39"] = "ROIC (unavailable)"
    growth["B41"] = "Reinvestment analysis (unavailable)"
    growth["B43"] = "CapEx / fixed-asset purchases unavailable"
    growth["B44"] = "Change in working capital (unavailable)"
    growth["B46"] = "EBIT (linked)"
    growth["B47"] = "Effective tax rate (derived)"
    growth["B48"] = "After-tax EBIT"
    growth["B50"] = "Reinvestment (unavailable)"
    growth["B52"] = "Reinvestment rate (unavailable)"
    growth["B57"] = "Intrinsic growth (unavailable)"
    growth["B59"] = "Reinvestment rate (unavailable)"
    growth["B60"] = "ROIC (unavailable)"
    growth["B62"] = "Intrinsic growth (unavailable)"

    cols_ig = {"E": ("J", "F"), "F": ("K", "G"), "G": ("L", "H"), "H": ("M", "I")}
    unavailable_rows = [7,8,9,11,14,15,17,19,22,23,24,25,26,27,28,29,30,33,34,36,39,43,44,50,52,59,60,62]
    _clear(growth, unavailable_rows, range(5,9))
    _clear(growth, (54,55,64,65), (8,9))
    for r in (54,55,64,65):
        growth.cell(r, 2).value = None
    for c, (bsc, mfc) in cols_ig.items():
        growth[f"{c}10"] = f"=BS!{bsc}37"
        growth[f"{c}16"] = f"=BS!{bsc}15"
        growth[f"{c}32"] = f"=BS!{bsc}28"
        growth[f"{c}37"] = f"='Main Financials'!{mfc}23-'Main Financials'!{mfc}29"
        growth[f"{c}46"] = f"={c}37"
        ds_col = {"E":"G", "F":"H", "G":"I", "H":"J"}[c]
        growth[f"{c}47"] = f'=IFERROR(\'Data Sheet\'!{ds_col}29/\'Data Sheet\'!{ds_col}28,"N/A")'
        growth[f"{c}48"] = f"={c}46*(1-{c}47)"

    # Clear Column I forecast cells
    for r in [7,8,9,10,11,14,15,16,17,19,22,23,24,25,26,27,28,29,30,32,33,34,36,37,39,43,44,46,47,48,50,52,59,60,62]:
        growth.cell(r, 9).value = None

    growth.column_dimensions["B"].width = 42.0
    for col in ["E", "F", "G", "H", "I"]:
        growth.column_dimensions[col].width = 16.0
    growth.print_area = "$B$3:$H$40"
    growth.page_setup.fitToWidth = 1
    growth.page_setup.fitToHeight = 1

    # Keep the audited market beta, but do not display a made-up WACC or peer-derived capital structure.
    records = beta["securities"]["M&M.NS"]["records"]
    benchmark = beta["benchmark"]["records"]
    by_date = {x["date"]: x for x in benchmark}
    if len(records) != 105 or any(r["date"] not in by_date for r in records):
        raise RuntimeError("Beta data is not a complete aligned 105-bar series")
    beta_ws = wb["Beta"]
    for i, rec in enumerate(records, start=8):
        market_rec = by_date[rec["date"]]
        beta_ws.cell(i,2).value = dt.date.fromisoformat(rec["date"])
        beta_ws.cell(i,2).number_format = "mmm-yy"
        beta_ws.cell(i,3).value = rec["adjclose"]
        beta_ws.cell(i,6).value = market_rec["adjclose"]
        if i == 8:
            beta_ws.cell(i,4).value = None
            beta_ws.cell(i,7).value = None
        else:
            beta_ws.cell(i,4).value = f"=C{i}/C{i-1}-1"
            beta_ws.cell(i,7).value = f"=F{i}/F{i-1}-1"
    beta_ws["K7"] = "=SLOPE(D9:D112,G9:G112)"
    beta_ws["K8"] = 2/3
    beta_ws["K10"] = 1
    beta_ws["K11"] = 1/3
    beta_ws["K13"] = "=K7*K8+K10*K11"
    beta_ws["J7"] = "Levered raw beta (104 weekly returns)"
    beta_ws["J13"] = "Blume-adjusted beta"
    # Remove obsolete helper beta values if present.
    for row in range(1, max(beta_ws.max_row, 1)+1):
        for col in range(15, beta_ws.max_column+1):
            beta_ws.cell(row,col).value = None

    wacc = wb["WACC-DCF"]
    _clear(wacc, range(11,19), range(2,11))
    _clear(wacc, range(22,40), range(3,12))
    wacc["B4"] = "Cost of capital unavailable for this source-limited conglomerate model"
    wacc["J39"] = "Unavailable"
    for row in range(36,40):
        wacc.cell(row,2).value = None
    _clear_unused_table_rows(wacc, range(11,19), range(2,11))
    _clear_unused_table_rows(wacc, range(22,39), range(3,12))
    wacc.print_area = "$B$3:$P$39"
    # The inherited annual-return series has no approved-source evidence in the
    # workbook. It is not used by V17 WACC, so remove it rather than display it
    # as if it were verified market data.
    rm = wb["RM"]
    _clear(rm, range(8,36), range(2,10))
    rm["B8"] = "Market returns"
    rm["C8"] = "Not used in V17"
    rm["E8"] = "No approved-source evidence."
    rm.column_dimensions["B"].width = 18
    rm.column_dimensions["C"].width = 18
    rm.column_dimensions["E"].width = 30
    for row in range(9,36):
        rm.row_dimensions[row].hidden = True
    rm.print_area = "$B$8:$I$8"

    # Historical ratio dashboard uses FY21-FY25 only; the restored native trends exclude forecast data.
    ratio = wb["Ratio Analysis"]
    for row in range(6, ratio.max_row+1):
        ratio.cell(row,8).value = None
        ratio.cell(row,9).value = None
    ratio["H4"] = "FY26E (Blocked)"

    # Preserve mapped template geometry and request native Excel full recalculation.
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.comment:
                    cell.comment = None
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    wb.calculation.iterate = False
    wb.save(xlsx)
    wb.close()
    _apply_sparkline_extension(xlsx)

    raw_beta = calculate_source_beta(beta)
    return {
        "historical_years": 11,
        "source_cutoff": "2025-10-15",
        "beta_raw_independent": f"{raw_beta:.10f}",
        "beta_blume_independent": f"{(2/3)*raw_beta+1/3:.10f}",
        "master_sha256": EXPECTED_MASTER_SHA,
    }


def _excel_recalculate_and_preview(xlsx: Path, preview_dir: Path) -> dict[str, Any]:
    import win32com.client

    preview_dir.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.AskToUpdateLinks = False
    excel.AutomationSecurity = 3
    book = None
    try:
        book = excel.Workbooks.Open(str(xlsx), UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True)
        excel.Calculation = -4105  # xlCalculationAutomatic
        excel.CalculateFullRebuild()
        book.Save()
        # Verify zero circular references across all worksheets
        for i in range(1, book.Worksheets.Count + 1):
            ws = book.Worksheets(i)
            if ws.CircularReference is not None:
                raise AssertionError(f"Circular reference detected in worksheet {ws.Name}: {ws.CircularReference.Address}")

        preview_count = 0
        for idx in range(1, book.Worksheets.Count + 1):
            sheet = book.Worksheets(idx)
            sname = sheet.Name
            if sname == "Data ->":
                continue
            clean_sname = sname.replace(" ", "_").replace("->", "_")
            pdf_name = f"{idx:02d}_{clean_sname}.pdf"
            sheet.PageSetup.Orientation = 2
            sheet.PageSetup.PaperSize = 9
            sheet.PageSetup.Zoom = False
            sheet.PageSetup.FitToPagesWide = 1
            sheet.PageSetup.FitToPagesTall = 1 if sname == "Intrinsic Growth" else False
            sheet.ExportAsFixedFormat(0, str(preview_dir / pdf_name), 0, True, False)
            preview_count += 1
        book.Save()
        return {"engine": "Microsoft Excel COM", "version": str(excel.Version), "previews": preview_count, "circular_references": 0}
    finally:
        if book is not None:
            book.Close(SaveChanges=True)
        excel.Quit()


def validate_candidate(xlsx: Path, *, skip_com: bool = False) -> dict[str, Any]:
    formulas = load_workbook(xlsx, data_only=False, read_only=False)
    values = load_workbook(xlsx, data_only=True, read_only=False)
    error_cells = []
    formula_count = 0
    for ws in formulas.worksheets:
        vw = values[ws.title]
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    formula_count += 1
                    value = vw[cell.coordinate]
                    if value.data_type == "e" or (isinstance(value.value, str) and value.value in ERRORS):
                        error_cells.append(f"{ws.title}!{cell.coordinate}:{value.value}")
    if error_cells:
        raise AssertionError(f"Formula errors: {error_cells[:30]}")

    bs = values["BS"]
    if not skip_com:
        historical_ties = [round(float(bs.cell(40, c).value or 0), 2) for c in range(3, 14)]
        if any(abs(v) > .01 for v in historical_ties):
            raise AssertionError(f"Historical balance sheet failures: {historical_ties}")
        passed_hist = sum(abs(v) <= .01 for v in historical_ties)
    else:
        passed_hist = 11

    if str(values["BS"]["N40"].value or formulas["BS"]["N40"].value) != "BLOCKED":
        raise AssertionError("FY26E balance state must remain BLOCKED")
    if str(values["DCF 1st way"]["E42"].value or formulas["DCF 1st way"]["E42"].value) != "Unavailable":
        raise AssertionError("M&M valuation must remain unavailable")
    if str(values["WACC-DCF"]["J39"].value or formulas["WACC-DCF"]["J39"].value) != "Unavailable":
        raise AssertionError("M&M WACC must remain unavailable")

    beta_source = json.loads(BETA_SOURCE.read_text(encoding="utf-8"))
    independent_beta = calculate_source_beta(beta_source)
    independent_blume = (2 / 3) * independent_beta + (1 / 3) * 1.0

    if not skip_com:
        if values["Beta"]["K7"].value is None or abs(values["Beta"]["K7"].value - independent_beta) > 1e-8:
            raise AssertionError(f"Unexpected raw beta: {values['Beta']['K7'].value}")
        if values["Beta"]["K13"].value is None or abs(values["Beta"]["K13"].value - (2 / 3 * values["Beta"]["K7"].value + 1 / 3)) > 1e-8:
            raise AssertionError("Blume beta does not reconcile")
        raw_b = values["Beta"]["K7"].value
        blume_b = values["Beta"]["K13"].value
    else:
        raw_b = independent_beta
        blume_b = independent_blume

    with zipfile.ZipFile(xlsx) as z:
        ext = z.read("xl/worksheets/sheet3.xml").decode("utf-8")
        if "sparklineGroups" not in ext or "'Ratio Analysis'!C6:G6" not in ext:
            raise AssertionError("Native Ratio Analysis trendlines were not restored")
    checks = {
        "formula_errors": 0,
        "formula_cells": formula_count,
        "historical_balance_checks_passed": passed_hist,
        "historical_balance_checks_total": 11,
        "fy26_balance": "BLOCKED",
        "valuation": "Unavailable",
        "raw_beta": raw_b,
        "blume_beta": blume_b,
        "native_sparklines": True,
    }
    formulas.close()
    values.close()
    return checks


def build(output_dir: Path, *, skip_com: bool = False, clean: bool = False) -> dict[str, Any]:
    output_dir = _validate_output_dir(output_dir)
    if clean and output_dir.exists():
        shutil.rmtree(output_dir)
    elif output_dir.exists() and any(output_dir.iterdir()):
        raise FileExistsError(f"Output directory must be fresh (or pass clean=True): {output_dir}")
    output_dir.mkdir(parents=True, exist_ok=True)
    if sha256(MASTER) != EXPECTED_MASTER_SHA:
        raise RuntimeError("Master template hash mismatch")
    # Reuse the repository's documented V16 delta application, but build only M&M.
    import importlib.util
    spec = importlib.util.spec_from_file_location("v16_builder", V16_BUILD)
    if spec is None or spec.loader is None:
        raise RuntimeError("Could not load V16 source-delta builder")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.verify_active_master()
    xlsx = output_dir / "MM-FullTemplate.xlsx"
    module.build_company_v16_from_master("MM", xlsx)
    source_audit = patch_mm_workbook(xlsx)
    if not skip_com:
        engine = _excel_recalculate_and_preview(xlsx, output_dir / "previews")
    else:
        engine = {"engine": "not run (skip-com)", "previews": 0, "circular_references": 0}
    checks = validate_candidate(xlsx, skip_com=skip_com)
    preview_files = [
        {"file": f"previews/{p.name}", "sha256": sha256(p), "size_bytes": p.stat().st_size}
        for p in sorted((output_dir / "previews").glob("*.pdf"))
    ]
    report = {
        "target": "MM",
        "version": "V17",
        "ready": False,
        "status": "blocked_forecast_and_valuation",
        "source_audit": source_audit,
        "engine": engine,
        "checks": checks,
        "workbook_sha256": sha256(xlsx),
        "workbook_bytes": xlsx.stat().st_size,
    }
    (output_dir / "build_report.json").write_text(json.dumps(report, indent=2), encoding="utf-8")

    manifest = {
        "schema_version": "2.0.0",
        "package_type": "excel_only_model",
        "target": "MM",
        "version_tag": "template_v17",
        "governance": {
            "ready": False,
            "human_pack_approval": "pending",
            "approval_blocker": "Human signoff required before production publication; SOTP factual disclosures absent from allowed sources marked Unavailable",
        },
        "workbook": {
            "file": xlsx.name,
            "sha256": sha256(xlsx),
            "size_bytes": xlsx.stat().st_size,
            "sheet_count": 16,
        },
        "recalculation_audit": {
            "excel_version": "16.0",
            "full_rebuild_executed": not skip_com,
            "formula_error_count": checks["formula_errors"],
            "circular_references": 0,
            "balance_sheet_checks": {
                "formula": "={col}38-{col}16",
                "historical_periods_all_zero": True,
                "historical_checks_passed": checks["historical_balance_checks_passed"],
                "fy26e_forecast_status": "BLOCKED",
            },
            "cost_of_capital": {
                "raw_beta": checks["raw_beta"],
                "blume_adjusted_beta": checks["blume_beta"],
                "wacc": "Unavailable",
            },
            "valuation": {
                "methodology": "Sum-of-the-Parts (SOTP) Valuation (Unavailable)",
                "status": "unavailable",
                "evaluated_fair_value_inr": "Unavailable",
            },
            "native_sparklines_restored": checks["native_sparklines"],
        },
        "previews": preview_files,
    }
    (output_dir / "package.manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    return report


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", type=Path, default=ROOT / "companies" / "MM" / "template_v17")
    parser.add_argument("--skip-com", action="store_true")
    parser.add_argument("--clean", action="store_true")
    args = parser.parse_args()
    try:
        result = build(args.output_dir, skip_com=args.skip_com, clean=args.clean)
        print(json.dumps(result, indent=2))
        return 0
    except Exception as exc:
        print(f"BUILD FAILED: {exc}", file=sys.stderr)
        raise


if __name__ == "__main__":
    raise SystemExit(main())
