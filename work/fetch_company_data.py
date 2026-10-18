"""Fetch and structure full financial data for M&M, LTM, CLEAN, EDELWEISS."""
import urllib.request
import json
import sys
from pathlib import Path
from bs4 import BeautifulSoup

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent.parent
SOURCES_DIR = ROOT / "work" / "sources"
SOURCES_DIR.mkdir(parents=True, exist_ok=True)

TARGETS = {
    "MM": {"screener_ticker": "M&M", "name": "Mahindra & Mahindra Limited", "bse": "500520", "nse": "M&M"},
    "LTM": {"screener_ticker": "LTM", "name": "LTIMindtree Limited", "bse": "540005", "nse": "LTIM"},
    "CLEAN": {"screener_ticker": "CLEAN", "name": "Clean Science and Technology Limited", "bse": "543318", "nse": "CLEAN"},
    "EDELWEISS": {"screener_ticker": "EDELWEISS", "name": "Edelweiss Financial Services Limited", "bse": "532922", "nse": "EDELWEISS"},
}

def clean_val(v: str):
    v = v.replace(",", "").replace("%", "").strip()
    if not v or v == "--" or v == "":
        return None
    try:
        if "." in v:
            return float(v)
        return int(v)
    except ValueError:
        return v

def fetch_company(key: str, info: dict) -> dict:
    ticker = info["screener_ticker"]
    url = f"https://www.screener.in/company/{ticker}/consolidated/"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
    with urllib.request.urlopen(req) as resp:
        html = resp.read().decode("utf-8")
    
    soup = BeautifulSoup(html, "html.parser")
    h1 = soup.find("h1")
    company_name = h1.get_text(strip=True) if h1 else info["name"]
    
    # Top stats
    ratios = {}
    for li in soup.find_all("li", class_="flex flex-space-between"):
        name_elem = li.find("span", class_="name")
        val_elem = li.find("span", class_="nowrap value")
        if name_elem and val_elem:
            k = name_elem.get_text(strip=True)
            val_text = val_elem.get_text(strip=True).replace("₹", "").replace(",", "").strip()
            ratios[k] = clean_val(val_text)
            
    # Links
    links = {}
    for a in soup.find_all("a"):
        href = a.get("href", "")
        text = a.get_text(strip=True)
        if "bseindia.com" in href:
            links["bse"] = href
        elif "nseindia.com" in href:
            links["nse"] = href
        elif "annual" in href.lower() or "annual-reports" in href.lower():
            links["annual_reports"] = href

    # Tables
    sections = {}
    for sec in soup.find_all("section"):
        sec_id = sec.get("id")
        if not sec_id:
            continue
        h2 = sec.find("h2")
        title = h2.get_text(strip=True) if h2 else sec_id
        
        tables = sec.find_all("table")
        if not tables:
            continue
            
        sec_table_list = []
        for t in tables:
            headers = [th.get_text(strip=True) for th in t.find_all("th")]
            rows = {}
            for tr in t.find_all("tr"):
                cells = [td.get_text(strip=True) for td in tr.find_all("td")]
                if cells:
                    row_name = cells[0].replace("+", "").strip()
                    row_values = [clean_val(c) for c in cells[1:]]
                    rows[row_name] = row_values
            sec_table_list.append({"headers": headers[1:], "rows": rows})
            
        sections[sec_id] = {"title": title, "tables": sec_table_list}
        
    doc = {
        "key": key,
        "company_name": company_name,
        "screener_ticker": ticker,
        "bse": info["bse"],
        "nse": info["nse"],
        "url": url,
        "stats": ratios,
        "links": links,
        "sections": sections,
    }
    
    out_file = SOURCES_DIR / f"{key}_consolidated.json"
    out_file.write_text(json.dumps(doc, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Saved {key} ({company_name}) to {out_file.name}")
    return doc

if __name__ == "__main__":
    import time
    for k, info in TARGETS.items():
        out_file = SOURCES_DIR / f"{k}_consolidated.json"
        if out_file.exists():
            print(f"Skipping {k}, already exists at {out_file.name}")
            continue
        time.sleep(6)
        doc = fetch_company(k, info)
        pl_table = doc["sections"].get("profit-loss", {}).get("tables", [{}])[0]
        bs_table = doc["sections"].get("balance-sheet", {}).get("tables", [{}])[0]
        print(f"  P&L periods: {pl_table.get('headers')}")
        print(f"  Sales: {pl_table.get('rows', {}).get('Sales')}")
        print(f"  Net Profit: {pl_table.get('rows', {}).get('Net Profit')}")
        print(f"  Total Assets: {bs_table.get('rows', {}).get('Total Assets')}")
        print(f"  Market Cap: {doc['stats'].get('Market Cap')}, Price: {doc['stats'].get('Current Price')}")
