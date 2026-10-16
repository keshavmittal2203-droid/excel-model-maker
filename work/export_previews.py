"""Export candidate presentation slides to PNG and candidate workbook sheets to PDF."""
from pathlib import Path
import pythoncom
import win32com.client

ROOT = Path(__file__).resolve().parent.parent
PPT_PATH = ROOT / "candidates" / "ERR-4-Marksans-Candidate.pptx"
WB_PATH = ROOT / "candidates" / "Marksans-Recalculated.xlsx"
PREVIEWS_DIR = ROOT / "candidates" / "previews"
PREVIEWS_DIR.mkdir(parents=True, exist_ok=True)

def export_ppt():
    print(f"Exporting slides from {PPT_PATH}...")
    pythoncom.CoInitialize()
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    # Open presentation in read-only, without window if possible
    pres = ppt_app.Presentations.Open(str(PPT_PATH), ReadOnly=True, Untitled=False, WithWindow=False)
    try:
        slide_count = pres.Slides.Count
        print(f"Total slides in presentation: {slide_count}")
        for i in range(1, slide_count + 1):
            slide = pres.Slides(i)
            out_png = PREVIEWS_DIR / f"slide_{i}.png"
            # Slide size is 4:3 portrait (cx=6858000, cy=9144000)
            # Export at high resolution: 1080x1440
            slide.Export(str(out_png), "PNG", 1080, 1440)
            print(f"  Exported slide {i} -> {out_png.name} ({out_png.stat().st_size} bytes)")
    finally:
        pres.Close()
        ppt_app.Quit()
        pythoncom.CoUninitialize()

def export_excel():
    print(f"Exporting sheets from {WB_PATH}...")
    pythoncom.CoInitialize()
    xl_app = win32com.client.Dispatch("Excel.Application")
    xl_app.Visible = False
    xl_app.DisplayAlerts = False
    wb = xl_app.Workbooks.Open(str(WB_PATH), ReadOnly=True)
    try:
        sheets_to_export = [
            "Data Sheet",
            "Ratio Analysis",
            "Intrinsic Growth",
            "WACC-DCF",
            "DCF 1st way",
            "DCF 2nd way",
        ]
        for sname in sheets_to_export:
            try:
                ws = wb.Sheets(sname)
                clean_name = sname.replace(" ", "_")
                out_pdf = PREVIEWS_DIR / f"{clean_name}.pdf"
                # ExportAsFixedFormat 0 = xlTypePDF
                ws.ExportAsFixedFormat(0, str(out_pdf))
                print(f"  Exported sheet {sname!r} -> {out_pdf.name} ({out_pdf.stat().st_size} bytes)")
            except Exception as e:
                print(f"  Failed to export sheet {sname}: {e}")
    finally:
        wb.Close(SaveChanges=False)
        xl_app.Quit()
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    export_ppt()
    export_excel()
    print("Export previews complete.")
