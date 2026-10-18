import json

with open('work/sources/CLEAN_consolidated.json') as f:
    d = json.load(f)

bs = d['sections']['balance-sheet']['tables'][0]
headers = bs['headers']
m_idx = headers.index('Mar 2025')
print("CLEAN Mar 2025 Main Table:")
for r_name, vals in bs['rows'].items():
    print(f"  {r_name:25s}: {vals[m_idx]}")
