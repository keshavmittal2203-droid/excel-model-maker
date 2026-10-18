import json

with open('work/sources/LTM_consolidated.json') as f:
    d = json.load(f)

bs = d['sections']['balance-sheet']['tables'][0]
headers = bs['headers']
m_idx = headers.index('Mar 2025')
print("LTM Mar 2025 Main Table:")
for r_name, vals in bs['rows'].items():
    print(f"  {r_name:25s}: {vals[m_idx]}")

with open('work/sources/LTM_schedules.json') as f:
    sch = json.load(f)

print("\nLTM Mar 2025 Schedules:")
for section in ['balance-sheet:Borrowings', 'balance-sheet:Other Liabilities', 'balance-sheet:Fixed Assets', 'balance-sheet:Other Assets']:
    if section in sch:
        print(f"--- {section} ---")
        for line, vals in sch[section].items():
            if isinstance(vals, dict):
                print(f"  {line:25s}: {vals.get('Mar 2025')}")
