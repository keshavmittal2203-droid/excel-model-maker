import os
import json
import re
from pathlib import Path
from urllib.parse import urlparse

ALLOWED_HOSTS = {
    'screener.in', 'www.screener.in',
    'finance.yahoo.com', 'yahoo.com', 'query1.finance.yahoo.com',
    'nseindia.com', 'www.nseindia.com',
    'tradingeconomics.com', 'www.tradingeconomics.com',
    'topstockresearch.com', 'www.topstockresearch.com',
    'investing.com', 'www.investing.com', 'in.investing.com',
    'primeinvestor.in', 'www.primeinvestor.in',
    'youtube.com', 'www.youtube.com', 'youtu.be',
    'github.com'
}

url_pattern = re.compile(r'https?://[^\s"\'<>]+')

results = {}
for root, dirs, files in os.walk('.'):
    # skip .git, .pytest_cache, __pycache__, originals
    norm_root = root.replace('\\', '/')
    if any(p in norm_root for p in ['.git', '.pytest_cache', '__pycache__', 'originals', 'archive']):
        continue
    for f in files:
        if f.endswith(('.json', '.md', '.py')):
            path = Path(root) / f
            try:
                text = path.read_text(encoding='utf-8', errors='ignore')
            except Exception:
                continue
            urls = url_pattern.findall(text)
            for u in urls:
                u_clean = u.rstrip('.,);]')
                parsed = urlparse(u_clean)
                host = parsed.netloc.lower()
                if host and host not in ALLOWED_HOSTS:
                    results.setdefault(host, []).append((str(path).replace('\\', '/'), u_clean))

print(f"Total disallowed host types found: {len(results)}")
for host, items in sorted(results.items()):
    print(f"\n=== {host} ({len(items)} occurrences) ===")
    unique_files = sorted(set(p for p, _ in items))
    print(f"Files ({len(unique_files)}): {', '.join(unique_files[:10])}")
    if len(unique_files) > 10:
        print(f"  ... and {len(unique_files)-10} more files")
    for p, u in items[:3]:
        print(f"  Sample from {p}: {u}")
