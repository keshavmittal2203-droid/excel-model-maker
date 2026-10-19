import json

with open('work/sources/MM_consolidated.json') as f:
    d = json.load(f)

cf = d['sections']['cash-flow']
print("cf tables:", cf['tables'])
