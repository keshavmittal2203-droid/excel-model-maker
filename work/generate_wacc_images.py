"""Generate pixel-accurate WACC replacement table images for Slide 23 from validated Marksans workbook."""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
import xml.etree.ElementTree as ET
import sys
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from modelmaker import Package, sha
WB_PATH = ROOT / "candidates" / "Marksans-Recalculated.xlsx"
ASSETS_DIR = ROOT / "work" / "assets"
ASSETS_DIR.mkdir(parents=True, exist_ok=True)

# Read validated values from recalculated workbook
pkg = Package(WB_PATH.read_bytes())
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

# Design Constants
ORANGE_DARK = (214, 96, 17, 255)    # #D66011 header / accent
ORANGE_MED = (237, 125, 49, 255)    # #ED7D31
PEACH_BG = (252, 228, 214, 255)     # #FCE4D6
WHITE_BG = (255, 255, 255, 255)
BORDER_COLOR = (0, 0, 0, 255)
TEXT_COLOR = (0, 0, 0, 255)

try:
    font_bold = ImageFont.truetype("C:/Windows/Fonts/arialbd.ttf", 15)
    font_reg = ImageFont.truetype("C:/Windows/Fonts/arial.ttf", 14)
except Exception:
    font_bold = ImageFont.load_default()
    font_reg = ImageFont.load_default()


def draw_table(width: int, height: int, rows_data: list[tuple[str, str, tuple, bool, bool]]) -> Image.Image:
    """Draw a table image.
    rows_data is a list of: (label, value, bg_color, is_bold, is_header)
    """
    im = Image.new("RGBA", (width, height), (255, 255, 255, 255))
    draw = ImageDraw.Draw(im)
    
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


# 1. Cost of Equity (image22.png, 707x137)
rows22 = [
    ("Cost of Equity", "", ORANGE_MED, True, True),
    ("Risk free Rate", f"{rf:.2f}%", PEACH_BG, False, False),
    ("Market Returns", f"{rm:.1f}%", ORANGE_MED, False, False),
    ("Beta", f"{beta:.3f}", PEACH_BG, False, False),
    ("Ke", f"{ke:.2f}%", ORANGE_MED, True, False),
]
img22 = draw_table(707, 137, rows22)
p22 = ASSETS_DIR / "image22_marksans.png"
img22.save(p22)

# 2. Post tax Cost of Debt (image23.png, 696x102)
rows23 = [
    ("Post tax Cost of Debt", "", ORANGE_MED, True, True),
    ("Pre tax cost of debt", f"{pre_tax_kd:.2f}%", PEACH_BG, False, False),
    ("Tax rate", f"{tax_rate:.0f}%", WHITE_BG, False, False),
    ("Kd", f"{kd:.2f}%", ORANGE_MED, True, False),
]
img23 = draw_table(696, 102, rows23)
p23 = ASSETS_DIR / "image23_marksans.png"
img23.save(p23)

# 3. WACC Table (image24.png, 691x192)
rows24 = [
    ("WACC", "", ORANGE_MED, True, True),
    ("Market Capitalisation", f"{mcap:,.0f}", WHITE_BG, False, False),
    ("Long term Debt", f"{total_debt:,.0f}", ORANGE_MED, False, False),
    ("Weight of Equity", f"{w_equity:.2f}%", WHITE_BG, False, False),
    ("Weight of Debt", f"{w_debt:.2f}%", PEACH_BG, False, False),
    ("WACC", f"{wacc:.2f}%", ORANGE_MED, True, False),
]
img24 = draw_table(691, 192, rows24)
p24 = ASSETS_DIR / "image24_marksans.png"
img24.save(p24)

print("Generated replacement WACC tables:")
print(f"  image22: {p22} (SHA256: {sha(p22.read_bytes())})")
print(f"  image23: {p23} (SHA256: {sha(p23.read_bytes())})")
print(f"  image24: {p24} (SHA256: {sha(p24.read_bytes())})")
