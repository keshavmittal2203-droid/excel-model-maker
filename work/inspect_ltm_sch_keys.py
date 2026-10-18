import json

with open('work/sources/LTM_schedules.json') as f:
    sch = json.load(f)

print(sch.keys())
if 'consolidated_schedules' in sch:
    print(sch['consolidated_schedules'].keys())
    for k, v in sch['consolidated_schedules'].items():
        print(f"  {k}: {list(v.keys())}")
        if 'Other Assets' in k or 'Other Assets' in str(v):
            print(f"    detail: {v}")
