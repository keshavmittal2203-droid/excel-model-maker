import json

with open('companies/LTM/source_registry.json') as f:
    reg = json.load(f)

print("Number of records in LTM source_registry:", len(reg))
for r in reg[:15]:
    print(f"  {r.get('sheet')}: {r.get('cell')} -> {r.get('value')} ({r.get('line_item')})")
