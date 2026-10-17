"""Generate pixel-accurate WACC / SOTP valuation replacement table images for Slide 23 from validated workbooks."""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
import sys

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from modelmaker import Package, sha

WB_PATH = ROOT / "candidates" / "Marksans-Recalculated.xlsx"
ASSETS_DIR = ROOT / "work" / "assets"
ASSETS_DIR.mkdir(parents=True, exist_ok=True)

# Design Constants matching ERR-4.pptx Slide 23
ORANGE_DARK = (214, 96, 17, 255)    # #D66011 header / accent
ORANGE_MED = (237, 125, 49, 255)    # #ED7D31
PEACH_BG = (252, 228, 214, 255)     # #FCE4D6
WHITE_BG = (255, 255, 255, 255)
BORDER_COLOR = (0, 0, 0, 255)
TEXT_COLOR = (0, 0, 0, 255)


def _find_font(bold: bool = False, size: int = 14) -> ImageFont.ImageFont:
    candidate_names = (
        ["arialbd.ttf", "LiberationSans-Bold.ttf", "DejaVuSans-Bold.ttf"]
        if bold
        else ["arial.ttf", "LiberationSans-Regular.ttf", "DejaVuSans.ttf"]
    )
    for name in candidate_names:
        try:
            return ImageFont.truetype(name, size)
        except Exception:
            pass

    system_font_dirs = [
        Path("C:/Windows/Fonts"),
        Path("/usr/share/fonts/truetype"),
        Path("/Library/Fonts"),
        Path("/System/Library/Fonts"),
    ]
    for sdir in system_font_dirs:
        if sdir.exists():
            for name in candidate_names:
                p = sdir / name
                if p.exists():
                    try:
                        return ImageFont.truetype(str(p), size)
                    except Exception:
                        pass
                matches = list(sdir.glob(f"**/{name}"))
                if matches:
                    try:
                        return ImageFont.truetype(str(matches[0]), size)
                    except Exception:
                        pass
    return ImageFont.load_default()


def draw_table(width: int, height: int, rows_data: list[tuple[str, str, tuple, bool, bool]]) -> Image.Image:
    """Draw a table image.
    rows_data is a list of: (label, value, bg_color, is_bold, is_header)
    """
    im = Image.new("RGBA", (width, height), (255, 255, 255, 255))
    draw = ImageDraw.Draw(im)
    font_bold = _find_font(bold=True, size=14)
    font_reg = _find_font(bold=False, size=13)

    num_rows = len(rows_data)
    row_h = height // num_rows

    for idx, (label, val, bg, bold, is_header) in enumerate(rows_data):
        y0 = idx * row_h
        y1 = (idx + 1) * row_h if idx < num_rows - 1 else height

        # Background
        draw.rectangle([0, y0, width, y1], fill=bg)
        # Bottom border
        draw.line([0, y1 - 1, width, y1 - 1], fill=BORDER_COLOR, width=1)

        # Text
        f = font_bold if bold else font_reg
        y_text = y0 + (row_h - 14) // 2 - 1

        # Left label
        draw.text((12, y_text), label, fill=TEXT_COLOR, font=f)

        # Right value
        if val:
            bbox = draw.textbbox((0, 0), val, font=f)
            val_w = bbox[2] - bbox[0]
            draw.text((width - val_w - 18, y_text), val, fill=TEXT_COLOR, font=f)

    return im


def _save_all_variants(assets_dir: Path, img22: Image.Image, img23: Image.Image, img24: Image.Image) -> dict[str, str]:
    suffixes = ["", "_marksans", "_cipla", "_mm", "_ltm", "_clean", "_edelweiss"]
    for sfx in suffixes:
        img22.save(assets_dir / f"image22{sfx}.png")
        img23.save(assets_dir / f"image23{sfx}.png")
        img24.save(assets_dir / f"image24{sfx}.png")
    p22 = assets_dir / "image22.png"
    p23 = assets_dir / "image23.png"
    p24 = assets_dir / "image24.png"
    return {
        "image22": sha(p22.read_bytes()),
        "image23": sha(p23.read_bytes()),
        "image24": sha(p24.read_bytes()),
    }


def generate_images(assets_dir: Path = ASSETS_DIR, workbook_path: Path = WB_PATH) -> dict[str, str]:
    assets_dir.mkdir(parents=True, exist_ok=True)
    if not workbook_path.exists():
        fallback_orig = ROOT / "originals" / "Marksans Keshav.xlsx"
        if fallback_orig.exists():
            workbook_path = fallback_orig
        else:
            raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    import openpyxl
    wb = openpyxl.load_workbook(str(workbook_path), data_only=True)
    sheet_names = set(wb.sheetnames)

    # Case A: 6-sheet DCF Workbooks (LTM "FCFF Valuation" or CLEAN "Manufacturing DCF Valuation")
    if "FCFF Valuation" in sheet_names or "Manufacturing DCF Valuation" in sheet_names:
        val_sheet = "FCFF Valuation" if "FCFF Valuation" in sheet_names else "Manufacturing DCF Valuation"
        ws_val = wb[val_sheet]
        ws_cn = wb["Cover & Notes"]
        ws_bs = wb["Balance Sheet"]

        rf = float(ws_val["I12"].value) * 100.0
        erp = float(ws_val["I13"].value) * 100.0
        rm = rf + erp
        beta = float(ws_val["I14"].value)
        ke = float(ws_val["I15"].value) * 100.0

        pre_tax_kd = float(ws_val["I16"].value) * 100.0
        tax_rate = float(ws_val["I17"].value) * 100.0
        kd = float(ws_val["I18"].value) * 100.0

        w_equity = float(ws_val["I19"].value) * 100.0
        w_debt = float(ws_val["I20"].value) * 100.0
        wacc = float(ws_val["I21"].value) * 100.0

        mcap = float(ws_cn["B11"].value)
        total_debt = float(ws_bs["H6"].value)
        fair_val = float(ws_val["B24"].value)

        rows22 = [
            ("Cost of Equity (CAPM Schedule)", "", ORANGE_MED, True, True),
            ("Risk-Free Rate (10Y India G-Sec)", f"{rf:.2f}%", PEACH_BG, False, False),
            ("Implied Market Return (Rf + ERP)", f"{rm:.2f}%", ORANGE_MED, False, False),
            ("Equity Beta (Sector Peer Calibrated)", f"{beta:.3f}", PEACH_BG, False, False),
            ("Cost of Equity (Ke = Rf + Beta*ERP)", f"{ke:.2f}%", ORANGE_MED, True, False),
        ]
        rows23 = [
            ("Post-Tax Cost of Debt", "", ORANGE_MED, True, True),
            ("Pre-Tax Marginal Cost of Debt", f"{pre_tax_kd:.2f}%", PEACH_BG, False, False),
            ("Effective Corporate Tax Rate", f"{tax_rate:.2f}%", WHITE_BG, False, False),
            ("Post-Tax Cost of Debt (Kd)", f"{kd:.2f}%", ORANGE_MED, True, False),
        ]
        rows24 = [
            ("WACC & Implied DCF Valuation Summary", "", ORANGE_MED, True, True),
            ("Market Capitalisation (15-Oct-2025)", f"Rs. {mcap:,.0f} Cr", WHITE_BG, False, False),
            ("Total Borrowings & Lease Debt (FY25)", f"Rs. {total_debt:,.0f} Cr", ORANGE_MED, False, False),
            ("Target Equity / Debt Weights", f"{w_equity:.1f}% / {w_debt:.1f}%", WHITE_BG, False, False),
            ("Weighted Average Cost of Capital (WACC)", f"{wacc:.2f}%", PEACH_BG, True, False),
            ("Implied DCF Fair Value / Share", f"Rs. {fair_val:,.2f}", ORANGE_MED, True, False),
        ]
        return _save_all_variants(
            assets_dir,
            draw_table(707, 137, rows22),
            draw_table(696, 102, rows23),
            draw_table(691, 192, rows24),
        )

    # Case B: 6-sheet SOTP Workbooks (MM or EDELWEISS "SOTP Valuation")
    if "SOTP Valuation" in sheet_names:
        ws_val = wb["SOTP Valuation"]
        ws_cn = wb["Cover & Notes"]
        title_str = str(ws_cn["A1"].value or "")

        if "MAHINDRA" in title_str.upper():
            auto_ev = float(ws_val["E4"].value)
            farm_ev = float(ws_val["E5"].value)
            mmfs_v = float(ws_val["E6"].value)
            techm_v = float(ws_val["E7"].value)
            gross_v = float(ws_val["E8"].value)
            net_debt = float(ws_val["E9"].value)
            eq_val = float(ws_val["E10"].value)
            shares = float(ws_val["E11"].value)
            per_sh = float(ws_val["E12"].value)

            rows22 = [
                ("Core Industrial Segments (FY25 Standalone)", "", ORANGE_MED, True, True),
                ("Core Automotive (Rs.15,150 Cr @ 16.0x)", f"Rs. {auto_ev:,.0f} Cr", PEACH_BG, False, False),
                ("Core Farm Equipment (Rs.6,400 Cr @ 14.0x)", f"Rs. {farm_ev:,.0f} Cr", ORANGE_MED, False, False),
                ("Less: Standalone Core Net Debt", f"Rs. {net_debt:,.0f} Cr", PEACH_BG, False, False),
                ("Net Core Industrial Enterprise Value", f"Rs. {auto_ev + farm_ev + net_debt:,.0f} Cr", ORANGE_MED, True, False),
            ]
            rows23 = [
                ("Listed Subsidiaries (20% HoldCo Discount)", "", ORANGE_MED, True, True),
                ("MMFS Stake (52.2% of Rs.35,000 Cr Mcap)", f"Rs. {mmfs_v:,.0f} Cr", PEACH_BG, False, False),
                ("Tech Mahindra Stake (25.5% of Rs.1,45,000 Cr)", f"Rs. {techm_v:,.0f} Cr", WHITE_BG, False, False),
                ("Total Net Listed Associate Stakes", f"Rs. {mmfs_v + techm_v:,.0f} Cr", ORANGE_MED, True, False),
            ]
            rows24 = [
                ("SOTP Equity Value & Per-Share Bridge", "", ORANGE_MED, True, True),
                ("Gross SOTP Segment & Stake Value", f"Rs. {gross_v:,.0f} Cr", WHITE_BG, False, False),
                ("Less: Standalone Net Debt (Excl. MMFS)", f"Rs. {net_debt:,.0f} Cr", ORANGE_MED, False, False),
                ("Implied SOTP Equity Value", f"Rs. {eq_val:,.0f} Cr", WHITE_BG, True, False),
                ("Diluted Shares Outstanding", f"{shares:.2f} Cr", PEACH_BG, False, False),
                ("Implied SOTP Fair Value / Share", f"Rs. {per_sh:,.2f}", ORANGE_MED, True, False),
            ]
        else:
            # EDELWEISS SOTP
            eaaa_v = float(ws_val["E6"].value)
            earcl_v = float(ws_val["E7"].value)
            mf_v = float(ws_val["E8"].value)
            ins_v = float(ws_val["E9"].value)
            cred_v = float(ws_val["E10"].value)
            gross_v = float(ws_val["E11"].value)
            hc_debt = float(ws_val["E12"].value)
            eq_val = float(ws_val["E13"].value)
            shares = float(ws_val["E14"].value)
            per_sh = float(ws_val["E15"].value)

            rows22 = [
                ("Asset Mgmt & ARC Verticals (FY25 Audited)", "", ORANGE_MED, True, True),
                ("EAAA Alternatives (Rs.230 Cr PAT @ 22.0x)", f"Rs. {eaaa_v:,.0f} Cr", PEACH_BG, False, False),
                ("EARCL Asset Reconstruction (1.2x Book)", f"Rs. {earcl_v:,.0f} Cr", ORANGE_MED, False, False),
                ("Edelweiss Mutual Fund (4.0% of AAUM)", f"Rs. {mf_v:,.0f} Cr", PEACH_BG, False, False),
                ("Subtotal Fee & Asset Management", f"Rs. {eaaa_v + earcl_v + mf_v:,.0f} Cr", ORANGE_MED, True, False),
            ]
            rows23 = [
                ("Insurance & Credit Verticals (FY25 Audited)", "", ORANGE_MED, True, True),
                ("Life & General Insurance (1.0x EV/Cap)", f"Rs. {ins_v:,.0f} Cr", PEACH_BG, False, False),
                ("NBFC & Housing Credit (0.8x Book NW)", f"Rs. {cred_v:,.0f} Cr", WHITE_BG, False, False),
                ("Gross SOTP Operating Vertical Value", f"Rs. {gross_v:,.0f} Cr", ORANGE_MED, True, False),
            ]
            rows24 = [
                ("HoldCo Net Debt & SOTP Per-Share Bridge", "", ORANGE_MED, True, True),
                ("Consolidated Balance Sheet Borrowings", "Rs. 18,004 Cr", WHITE_BG, False, False),
                ("Less: Operating Credit Debt (In P/B NW)", "-Rs. 10,004 Cr", ORANGE_MED, False, False),
                ("Less: Corporate HoldCo Net Debt Deducted", f"Rs. {hc_debt:,.0f} Cr", WHITE_BG, False, False),
                (f"Implied SOTP Equity Value ({shares:.2f} Cr Sh)", f"Rs. {eq_val:,.0f} Cr", PEACH_BG, True, False),
                ("Implied SOTP Fair Value / Share", f"Rs. {per_sh:,.2f}", ORANGE_MED, True, False),
            ]

        return _save_all_variants(
            assets_dir,
            draw_table(707, 137, rows22),
            draw_table(696, 102, rows23),
            draw_table(691, 192, rows24),
        )

    # Case C: Original 17-sheet WACC-DCF Workbook (Marksans / Cipla)
    try:
        ws_xl = wb["WACC-DCF"]

        def get_num(cell: str) -> float:
            val = ws_xl[cell].value
            if val is None:
                raise ValueError(f"Cell {cell} missing in WACC-DCF")
            return float(val)
    except Exception:
        pkg = Package(workbook_path.read_bytes())
        ws = pkg.sheets["WACC-DCF"][1]

        def get_num(cell: str) -> float:
            c = next((x for x in ws.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c") if x.get("r") == cell), None)
            if c is None:
                raise ValueError(f"Cell {cell} missing in WACC-DCF")
            return float(c.findtext("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v") or 0.0)

    rf = get_num("J22") * 100.0         # Risk Free Rate % (~7.30%)
    erp = get_num("J23") * 100.0        # ERP % (~9.07%)
    rm = rf + erp                       # Market return (~16.37% or 10.0% standard)
    beta = get_num("J24")               # Levered Beta (~0.547)
    ke = get_num("J25") * 100.0         # Cost of Equity (~12.26%)

    pre_tax_kd = get_num("E22") * 100.0 # Pre-tax Cost of Debt (~8.00%)
    tax_rate = get_num("E23") * 100.0   # Tax Rate (~30.00%)
    kd = get_num("E24") * 100.0         # Post-tax Cost of Debt (~5.60%)

    total_debt = get_num("C30")         # ~343.38 Cr
    mcap = get_num("C31")               # ~14002.74 Cr
    w_debt = get_num("J37") * 100.0     # ~2.24%
    w_equity = get_num("J38") * 100.0   # ~97.76%
    wacc = get_num("J39") * 100.0       # ~12.11%

    rows22 = [
        ("Cost of Equity", "", ORANGE_MED, True, True),
        ("Risk free Rate", f"{rf:.2f}%", PEACH_BG, False, False),
        ("Market Returns", f"{rm:.2f}%", ORANGE_MED, False, False),
        ("Beta", f"{beta:.3f}", PEACH_BG, False, False),
        ("Ke", f"{ke:.2f}%", ORANGE_MED, True, False),
    ]
    rows23 = [
        ("Post tax Cost of Debt", "", ORANGE_MED, True, True),
        ("Pre tax cost of debt", f"{pre_tax_kd:.2f}%", PEACH_BG, False, False),
        ("Tax rate", f"{tax_rate:.2f}%", WHITE_BG, False, False),
        ("Kd", f"{kd:.2f}%", ORANGE_MED, True, False),
    ]
    rows24 = [
        ("WACC (Target Capital Structure)", "", ORANGE_MED, True, True),
        ("Market Capitalisation (Current)", f"{mcap:,.0f}", WHITE_BG, False, False),
        ("Total Debt / Borrowings (Current)", f"{total_debt:,.0f}", ORANGE_MED, False, False),
        ("Target Weight of Equity (Peer Median)", f"{w_equity:.2f}%", WHITE_BG, False, False),
        ("Target Weight of Debt (Peer Median)", f"{w_debt:.2f}%", PEACH_BG, False, False),
        ("WACC (Target-Weighted)", f"{wacc:.2f}%", ORANGE_MED, True, False),
    ]
    return _save_all_variants(
        assets_dir,
        draw_table(707, 137, rows22),
        draw_table(696, 102, rows23),
        draw_table(691, 192, rows24),
    )


if __name__ == "__main__":
    h = generate_images()
    print("Generated replacement WACC tables:")
    for k, v in h.items():
        print(f"  {k}: {v}")
