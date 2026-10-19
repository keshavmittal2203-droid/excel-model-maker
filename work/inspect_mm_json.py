import json

with open('work/sources/MM_consolidated.json') as f:
    d = json.load(f)

print("Sections:", d['sections'].keys())
for sname in ['balance-sheet', 'cash-flow']:
    if sname in d['sections']:
        sec = d['sections'][sname]
        print(f"\n=== Section: {sname} ===")
        print("Columns:", sec.get('columns', []))
        for row in sec.get('rows', []):
            line = row.get('line', '')
            vals = row.get('values', {})
            mar24 = vals.get('Mar 2024')
            mar25 = vals.get('Mar 2025')
            print(f"  {line:35s} | Mar 2024: {mar24} | Mar 2025: {mar25}")
