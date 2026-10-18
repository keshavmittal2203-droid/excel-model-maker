"""Fetch and archive 104-week market price observations and regression diagnostics.

Covers:
- Benchmark: Nifty 50 (^NSEI)
- Targets: M&M.NS, LTM.NS, CLEAN.NS, EDELWEISS.NS
- Listed Stakes: M&MFIN.NS, TECHM.NS
- LTM Peers: TCS.NS, INFY.NS, HCLTECH.NS, WIPRO.NS
- CLEAN Peers: VINATIORGA.NS, FINEORG.NS, NAVINFLUOR.NS, PIIND.NS, SRF.NS
- MM Peers: MARUTI.NS, BAJAJ-AUTO.NS, EICHERMOT.NS, ESCORTS.NS

Valuation Cutoff Date: 2025-10-15 (Strictly zero observations post-cutoff).
Observation Window: 2023-10-16 to 2025-10-13 (105 weekly price bars -> 104 paired returns).
"""

import datetime
import hashlib
import json
import math
from pathlib import Path
import urllib.request

ROOT = Path(__file__).resolve().parent.parent
SOURCES_DIR = ROOT / "work" / "sources"
OUT_JSON = SOURCES_DIR / "weekly_market_returns_104w.json"

TICKERS = [
    "^NSEI",
    "M&M.NS",
    "LTM.NS",
    "CLEAN.NS",
    "EDELWEISS.NS",
    "M&MFIN.NS",
    "TECHM.NS",
    "TCS.NS",
    "INFY.NS",
    "HCLTECH.NS",
    "WIPRO.NS",
    "VINATIORGA.NS",
    "FINEORG.NS",
    "NAVINFLUOR.NS",
    "PIIND.NS",
    "SRF.NS",
    "MARUTI.NS",
    "BAJAJ-AUTO.NS",
    "EICHERMOT.NS",
    "ESCORTS.NS",
]

# Unix timestamps for 2023-10-15 to 2025-10-16
PERIOD1 = 1697328000  # 2023-10-15
PERIOD2 = 1760572800  # 2025-10-16


def fetch_weekly_data(ticker: str) -> dict:
    url_ticker = ticker.replace("^", "%5E")
    url = f"https://query1.finance.yahoo.com/v8/finance/chart/{url_ticker}?period1={PERIOD1}&period2={PERIOD2}&interval=1wk"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
    with urllib.request.urlopen(req) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    
    result = data["chart"]["result"][0]
    timestamps = result["timestamp"]
    quote = result["indicators"]["quote"][0]
    adjclose = result["indicators"]["adjclose"][0]["adjclose"] if "adjclose" in result["indicators"] else quote["close"]
    
    records = []
    for i, ts in enumerate(timestamps):
        dt = datetime.datetime.fromtimestamp(ts, tz=datetime.timezone.utc)
        dt_str = dt.strftime("%Y-%m-%d")
        # Ensure cutoff on or before 2025-10-15
        if dt_str > "2025-10-15":
            continue
        c = quote["close"][i]
        ac = adjclose[i] if adjclose and i < len(adjclose) else c
        if c is not None and not math.isnan(c):
            records.append({
                "date": dt_str,
                "timestamp": ts,
                "close": round(c, 2),
                "adjclose": round(ac, 2) if ac is not None else round(c, 2),
            })
    return {"ticker": ticker, "records": records}


def compute_weekly_returns(records: list[dict]) -> list[dict]:
    ret_list = []
    for i in range(1, len(records)):
        prev_p = records[i - 1]["close"]
        curr_p = records[i]["close"]
        ret = (curr_p / prev_p) - 1.0 if prev_p else 0.0
        ret_list.append({
            "date": records[i]["date"],
            "prev_close": prev_p,
            "close": curr_p,
            "return": round(ret, 6),
        })
    return ret_list


def compute_ols_beta(stock_returns: list[float], mkt_returns: list[float]) -> dict:
    n = len(stock_returns)
    if n != len(mkt_returns) or n < 10:
        raise ValueError(f"Mismatched return counts: stock={len(stock_returns)}, mkt={len(mkt_returns)}")
    
    mean_s = sum(stock_returns) / n
    mean_m = sum(mkt_returns) / n
    
    cov = sum((stock_returns[i] - mean_s) * (mkt_returns[i] - mean_m) for i in range(n)) / (n - 1)
    var_m = sum((mkt_returns[i] - mean_m) ** 2 for i in range(n)) / (n - 1)
    var_s = sum((stock_returns[i] - mean_s) ** 2 for i in range(n)) / (n - 1)
    
    raw_beta = cov / var_m if var_m != 0 else 1.0
    blume_beta = (2.0 / 3.0) * raw_beta + (1.0 / 3.0) * 1.0
    
    std_s = math.sqrt(var_s)
    std_m = math.sqrt(var_m)
    corr = cov / (std_s * std_m) if (std_s * std_m) != 0 else 0.0
    r_squared = corr ** 2
    
    # Regression Residual Standard Error: sqrt(SSE / (n - 2))
    # y = alpha + beta * x
    alpha = mean_s - raw_beta * mean_m
    residuals = [stock_returns[i] - (alpha + raw_beta * mkt_returns[i]) for i in range(n)]
    sse = sum(r ** 2 for r in residuals)
    se_regression = math.sqrt(sse / (n - 2)) if n > 2 else 0.0
    se_beta = se_regression / math.sqrt((n - 1) * var_m) if var_m > 0 else 0.0
    t_stat = raw_beta / se_beta if se_beta > 0 else 0.0
    
    return {
        "n_observations": n,
        "mean_stock_weekly_return": round(mean_s, 6),
        "mean_mkt_weekly_return": round(mean_m, 6),
        "covariance": round(cov, 8),
        "market_variance": round(var_m, 8),
        "stock_variance": round(var_s, 8),
        "stock_weekly_volatility": round(std_s, 6),
        "market_weekly_volatility": round(std_m, 6),
        "raw_ols_beta": round(raw_beta, 4),
        "blume_adjusted_beta": round(blume_beta, 4),
        "correlation_r": round(corr, 4),
        "r_squared": round(r_squared, 4),
        "alpha_weekly": round(alpha, 6),
        "se_regression": round(se_regression, 6),
        "se_beta": round(se_beta, 4),
        "t_statistic": round(t_stat, 2),
    }


def main():
    print(f"Fetching weekly data for {len(TICKERS)} tickers (Cutoff: 2025-10-15)...")
    raw_data = {}
    for t in TICKERS:
        d = fetch_weekly_data(t)
        raw_data[t] = d
        print(f"  {t}: {len(d['records'])} bars, start={d['records'][0]['date']}, end={d['records'][-1]['date']}")
    
    # Align dates against benchmark ^NSEI
    benchmark_records = raw_data["^NSEI"]["records"]
    benchmark_returns = compute_weekly_returns(benchmark_records)
    bench_ret_map = {r["date"]: r["return"] for r in benchmark_returns}
    
    analysis_results = {
        "metadata": {
            "title": "Empirical Beta & Cost of Capital Derivation Dataset",
            "valuation_cutoff_date": "2025-10-15",
            "benchmark_index": "Nifty 50 (^NSEI)",
            "observation_window": f"{benchmark_records[0]['date']} to {benchmark_records[-1]['date']}",
            "weekly_sampling_convention": "Weekly bars ending Monday (incorporating trading week ending Wednesday 2025-10-15)",
            "data_provider": "National Stock Exchange of India (NSE) via Yahoo Finance API",
            "observation_count": len(benchmark_returns),
            "generated_at": datetime.datetime.now(datetime.timezone.utc).isoformat(),
        },
        "benchmark": {
            "ticker": "^NSEI",
            "total_bars": len(benchmark_records),
            "total_returns": len(benchmark_returns),
            "records": benchmark_records,
            "returns": benchmark_returns,
        },
        "securities": {},
    }
    
    for t in TICKERS:
        if t == "^NSEI":
            continue
        sec_records = raw_data[t]["records"]
        sec_returns = compute_weekly_returns(sec_records)
        sec_ret_map = {r["date"]: r["return"] for r in sec_returns}
        
        # Paired returns on common dates
        common_dates = [d for d in bench_ret_map if d in sec_ret_map]
        stock_aligned_rets = [sec_ret_map[d] for d in common_dates]
        mkt_aligned_rets = [bench_ret_map[d] for d in common_dates]
        
        diagnostics = compute_ols_beta(stock_aligned_rets, mkt_aligned_rets)
        
        analysis_results["securities"][t] = {
            "ticker": t,
            "latest_close_2025_10_15": sec_records[-1]["close"],
            "latest_adjclose_2025_10_15": sec_records[-1]["adjclose"],
            "common_return_count": len(common_dates),
            "diagnostics": diagnostics,
            "records": sec_records,
            "returns": sec_returns,
        }
        print(f"  -> {t}: Beta={diagnostics['raw_ols_beta']}, Blume={diagnostics['blume_adjusted_beta']}, R2={diagnostics['r_squared']}, n={diagnostics['n_observations']}")
    
    # Save structured json
    SOURCES_DIR.mkdir(parents=True, exist_ok=True)
    json_bytes = json.dumps(analysis_results, indent=2).encode("utf-8")
    OUT_JSON.write_bytes(json_bytes)
    sha_hash = hashlib.sha256(json_bytes).hexdigest()
    print(f"\nSaved structured weekly returns archive to: {OUT_JSON}")
    print(f"File size: {len(json_bytes)} bytes | SHA-256: {sha_hash}")


if __name__ == "__main__":
    main()
