import zipfile
import xml.etree.ElementTree as ET

C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

with zipfile.ZipFile("originals/ERR-4.pptx") as z:
    for cname in sorted([n for n in z.namelist() if n.startswith("ppt/charts/chart") and n.endswith(".xml")]):
        root = ET.fromstring(z.read(cname))
        title_elem = root.find(f".//{{{C_NS}}}title")
        t_text = "".join(title_elem.itertext()) if title_elem is not None else "No title"
        print(f"=== {cname} ===")
        print(f"  Title: {t_text.strip().encode('ascii', 'backslashreplace').decode('ascii')}")
        for ser in root.findall(f".//{{{C_NS}}}ser"):
            idx = ser.find(f"{{{C_NS}}}idx").get("val")
            tx_elem = ser.find(f"{{{C_NS}}}tx")
            tx = "".join(tx_elem.itertext()) if tx_elem is not None else "No tx"
            cat = [p.find(f"{{{C_NS}}}v").text for p in ser.findall(f".//{{{C_NS}}}cat//{{{C_NS}}}pt")]
            val = [p.find(f"{{{C_NS}}}v").text for p in ser.findall(f".//{{{C_NS}}}val//{{{C_NS}}}pt")]
            print(f"    Series {idx} ({tx.strip().encode('ascii', 'backslashreplace').decode('ascii')}):")
            print(f"      Categories ({len(cat)}): {cat[:6]}")
            print(f"      Values ({len(val)}): {val[:6]}")
