"""Inspect the fetched consolidated source data for the 4 companies."""
import json
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

sources = Path("work/sources")
for p in sources.glob("*_consolidated.json"):
    data = json.loads(p.read_text(encoding="utf-8"))
    print("=" * 60)
    print(f"Company: {data.get('company_name')} ({data.get('key')})")
    print(f"BSE: {data.get('bse')}, NSE: {data.get('nse')}")
    print(f"Stats: {data.get('stats')}")
    sec = data.get("sections", {})
    pl = sec.get("profit-loss", {}).get("tables", [{}])[0]
    bs = sec.get("balance-sheet", {}).get("tables", [{}])[0]
    cf = sec.get("cash-flow", {}).get("tables", [{}])[0]
    print(f"P&L periods: {pl.get('headers')}")
    print(f"P&L lines: {list(pl.get('rows', {}).keys())}")
    print(f"B/S lines: {list(bs.get('rows', {}).keys())}")
    print(f"C/F lines: {list(cf.get('rows', {}).keys())}")
    # Print last 3 years of Sales, Operating Profit, PAT
    headers = pl.get("headers", [])
    if headers:
        sales = pl.get("rows", {}).get("Sales", [])
        pat = pl.get("rows", {}).get("Net Profit", [])
        print("Recent P&L:")
        for h, s, p in zip(headers[-4:], sales[-4:], pat[-4:]):
            print(f"  {h}: Sales={s}, PAT={p}")
