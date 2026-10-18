import json
with open('companies/LTM/source_registry.json', encoding='utf-8') as f:
    reg = json.load(f)
print("Keys:", list(reg[0].keys()))
for k in ['item', 'category', 'sheet', 'cell_address', 'reported_value', 'statement_period']:
    print(f"  {k}: {reg[0].get(k)}")
