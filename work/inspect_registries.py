import json
from pathlib import Path

for target in ["MM", "LTM", "CLEAN", "EDELWEISS"]:
    reg_path = Path(f"companies/{target}/source_registry.json")
    if not reg_path.exists():
        continue
    items = json.loads(reg_path.read_text(encoding="utf-8"))
    print(f"\n=================== {target} ({len(items)} items) ===================")
    for i, it in enumerate(items, 1):
        metric = it.get("metric", "")
        cell = it.get("workbook_cell", "")
        status = it.get("status", "")
        url = it.get("official_url", "")
        sec_url = it.get("secondary_url", "")
        print(f"{i:2d}. {metric[:50]:50s} | Cell: {cell:25s} | Status: {status:25s} | URL: {str(url)[:50]}")
