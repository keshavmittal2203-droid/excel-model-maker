from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
PREVIEWS_DIR = ROOT / "candidates" / "run_native_baseline_v6" / "previews"

slides = [PREVIEWS_DIR / f"slide_{i}.png" for i in range(1, 6)]
images = [Image.open(s) for s in slides if s.exists()]

if len(images) == 5:
    # 5 slides arranged side by side or 3 + 2 grid
    # Thumbnails at width 540, height 720
    thumb_w, thumb_h = 540, 720
    thumbs = [img.resize((thumb_w, thumb_h), Image.Resampling.LANCZOS) for img in images]
    
    # 5 columns or 3x2 grid
    # Let's do 5 columns: 5 * 540 = 2700 wide, 720 + 80 banner high
    sheet_w = 5 * thumb_w + 6 * 20
    sheet_h = thumb_h + 100
    sheet = Image.new("RGB", (sheet_w, sheet_h), color=(240, 243, 246))
    draw = ImageDraw.Draw(sheet)
    
    for idx, thumb in enumerate(thumbs):
        x = 20 + idx * (thumb_w + 20)
        y = 60
        sheet.paste(thumb, (x, y))
        draw.rectangle([x, y, x + thumb_w, y + thumb_h], outline=(180, 190, 200), width=2)
        draw.text((x + 10, 20), f"Slide {idx + 1}", fill=(30, 40, 60))
        
    out_path = PREVIEWS_DIR / "contact_sheet.png"
    sheet.save(out_path, format="PNG")
    print(f"Generated contact sheet: {out_path} ({out_path.stat().st_size} bytes)")
    
    # Also copy previews to candidates/previews for convenience
    dest_previews = ROOT / "candidates" / "previews"
    dest_previews.mkdir(parents=True, exist_ok=True)
    import shutil
    for s in PREVIEWS_DIR.glob("*.*"):
        shutil.copy2(s, dest_previews / s.name)
    print(f"Synced previews to {dest_previews}")
