import json

for t in ['MM', 'LTM', 'CLEAN', 'EDELWEISS']:
    with open(f'companies/{t}/source_registry.json', encoding='utf-8') as f:
        data = json.load(f)
    print(f"\n=== {t}: {len(data)} items ===")
    for d in data:
        print(f"  {d['metric']} -> {d['workbook_cell']}")
