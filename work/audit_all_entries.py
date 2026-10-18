import json
from pathlib import Path

for target in ["MM", "LTM", "CLEAN", "EDELWEISS"]:
    reg_path = Path(f"companies/{target}/source_registry.json")
    items = json.loads(reg_path.read_text(encoding="utf-8"))
    print(f"\n=================== TARGET: {target} ({len(items)} items) ===================")
    for i, it in enumerate(items, 1):
        print(f"[{i:2d}] {it['metric']}")
        print(f"     Cell: {it.get('workbook_cell')} | Status: {it.get('status')}")
        print(f"     Official URL: {it.get('official_url')}")
        print(f"     Source Org:   {it.get('source_org')}")
