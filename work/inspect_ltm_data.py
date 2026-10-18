import json

with open('work/sources/LTM_consolidated.json') as f:
    d = json.load(f)

print("LTM Balance Sheet rows:")
bs = d['sections']['balance-sheet']
for row_name, vals in bs['tables'][0]['rows'].items():
    print(f"  {row_name:25s} | FY24: {vals[-2]} | FY25: {vals[-1]}")

print("\nLTM Schedules:")
with open('work/sources/LTM_schedules.json') as f:
    sch = json.load(f)
for k, v in sch.items():
    print(f"  Schedule: {k}")
    if isinstance(v, dict):
        for sub_k, sub_v in v.items():
            if isinstance(sub_v, list):
                print(f"    {sub_k:30s}: {sub_v[-2:]}")
