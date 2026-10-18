import json

with open('work/sources/MM_consolidated.json') as f:
    d = json.load(f)

bs = d['sections']['balance-sheet']
print("bs keys:", bs.keys())
print("bs type:", type(bs))
if isinstance(bs, dict):
    for k in list(bs.keys())[:10]:
        print(f"  {k}: {bs[k]}")
