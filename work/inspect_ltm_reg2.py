import json
with open('companies/LTM/source_registry.json') as f:
    reg = json.load(f)
print(reg[0])
