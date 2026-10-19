"""Build immutable Yahoo Finance market snapshots ending on 2025-10-15.

The older 104-week file was labeled as ending before the cutoff, but its final
price was the 2025-10-16 close. This builder samples aligned Wednesday closes
through the valuation cutoff and refuses any observation after that date.
"""

from __future__ import annotations

import datetime as dt
import json
import math
import urllib.parse
import urllib.request
from pathlib import Path
from statistics import mean

ROOT = Path(__file__).resolve().parent.parent
SOURCE_DIR = ROOT / "work" / "sources"
CUTOFF = dt.date(2025, 10, 15)
START = dt.date(2023, 10, 18)
TICKERS = {
    "M&M.NS": "M%26M.NS",
    "LTM.NS": "LTM.NS",
    "CLEAN.NS": "CLEAN.NS",
    "EDELWEISS.NS": "EDELWEISS.NS",
    "M&MFIN.NS": "M%26MFIN.NS",
    "TECHM.NS": "TECHM.NS",
}
BENCHMARK = "^NSEI"


def fetch_daily(ticker: str) -> dict[dt.date, dict[str, float]]:
    encoded = urllib.parse.quote(ticker, safe="%^")
    start_ts = int(dt.datetime(2023, 10, 15, tzinfo=dt.timezone.utc).timestamp())
    end_ts = int(dt.datetime(2025, 10, 16, tzinfo=dt.timezone.utc).timestamp())
    url = (
        f"https://query1.finance.yahoo.com/v8/finance/chart/{encoded}"
        f"?period1={start_ts}&period2={end_ts}&interval=1d"
    )
    request = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(request, timeout=30) as response:
        payload = json.load(response)
    results = payload.get("chart", {}).get("result") or []
    if not results:
        raise ValueError(f"Yahoo Finance returned no daily history for {ticker}")
    result = results[0]
    quote = result["indicators"]["quote"][0]
    adj = (result.get("indicators", {}).get("adjclose") or [{}])[0].get("adjclose", [])
    closes = quote.get("close", [])
    data: dict[dt.date, dict[str, float]] = {}
    for i, timestamp in enumerate(result.get("timestamp", [])):
        day = dt.datetime.fromtimestamp(timestamp, dt.timezone.utc).date()
        close = closes[i] if i < len(closes) else None
        if close is None or day > CUTOFF:
            continue
        row = {"close": float(close)}
        if i < len(adj) and adj[i] is not None:
            row["adjclose"] = float(adj[i])
        data[day] = row
    return data


def weekly_dates() -> list[dt.date]:
    dates = []
    day = START
    while day <= CUTOFF:
        dates.append(day)
        day += dt.timedelta(days=7)
    if len(dates) != 105 or dates[-1] != CUTOFF:
        raise AssertionError(f"Unexpected weekly sampling window: {len(dates)} points, {dates[-1]}")
    return dates


def sample_aligned(daily: dict[dt.date, dict[str, float]], slots: list[dt.date]) -> list[dict]:
    available = sorted(day for day in daily if START - dt.timedelta(days=7) <= day <= CUTOFF)
    records = []
    previous_actual_date = None
    for slot in slots:
        candidates = [day for day in available if day <= slot and (previous_actual_date is None or day > previous_actual_date)]
        if not candidates:
            raise ValueError(f"No Yahoo close available on or before weekly slot {slot}")
        actual_date = candidates[-1]
        records.append({
            "date": actual_date.isoformat(),
            "week_ending": slot.isoformat(),
            **daily[actual_date],
        })
        previous_actual_date = actual_date
    if records[-1]["date"] != CUTOFF.isoformat():
        raise AssertionError("Last sampled observation is not the valuation cutoff")
    return records


def slope_and_r2(stock_returns: list[float], market_returns: list[float]) -> tuple[float, float]:
    if len(stock_returns) != 104 or len(market_returns) != 104:
        raise ValueError("OLS requires exactly 104 paired weekly returns")
    xbar, ybar = mean(market_returns), mean(stock_returns)
    cov = sum((x - xbar) * (y - ybar) for x, y in zip(market_returns, stock_returns))
    var_x = sum((x - xbar) ** 2 for x in market_returns)
    var_y = sum((y - ybar) ** 2 for y in stock_returns)
    beta = cov / var_x
    r2 = (cov * cov) / (var_x * var_y)
    if not math.isfinite(beta) or not math.isfinite(r2):
        raise ValueError("Non-finite beta result")
    return beta, r2


def build() -> tuple[dict, dict]:
    slots = weekly_dates()
    raw = {ticker: fetch_daily(ticker) for ticker in [BENCHMARK, *TICKERS.values()]}
    sampled = {ticker: sample_aligned(raw[ticker], slots) for ticker in raw}
    benchmark_records = sampled[BENCHMARK]
    price = lambda row: row.get("adjclose", row["close"])
    benchmark_returns = [
        price(benchmark_records[i]) / price(benchmark_records[i - 1]) - 1
        for i in range(1, len(benchmark_records))
    ]
    weekly_securities = {}
    snapshot_securities = {}
    for display_ticker, yahoo_ticker in TICKERS.items():
        records = sampled[yahoo_ticker]
        stock_returns = [price(records[i]) / price(records[i - 1]) - 1 for i in range(1, len(records))]
        beta, r2 = slope_and_r2(stock_returns, benchmark_returns)
        weekly_securities[display_ticker] = {
            "provider_symbol": yahoo_ticker,
            "records": records,
            "return_observations": len(stock_returns),
            "raw_beta": beta,
            "blume_adjusted_beta": (2 / 3) * beta + (1 / 3),
            "r_squared": r2,
        }
        daily = raw[yahoo_ticker]
        daily_cutoff = daily.get(CUTOFF)
        if daily_cutoff is None:
            raise AssertionError(f"No exact 2025-10-15 daily close for {display_ticker}")
        snapshot_securities[display_ticker] = {
            "provider_symbol": yahoo_ticker,
            "close_2025_10_15": daily_cutoff["close"],
            "adjclose_2025_10_15": daily_cutoff.get("adjclose"),
            "price_field": "Yahoo Finance chart API quote.close (unadjusted close)",
            "source_url": f"https://finance.yahoo.com/quote/{urllib.parse.quote(yahoo_ticker, safe='%^')}/history/",
        }

    weekly = {
        "metadata": {
            "schema_version": "2.0.0",
            "title": "Cutoff-safe empirical beta dataset",
            "valuation_cutoff_date": CUTOFF.isoformat(),
            "benchmark_index": "NIFTY 50 (^NSEI)",
            "data_provider": "Yahoo Finance historical chart data for NSE-listed securities",
            "sampling_convention": "Aligned Wednesday close; if Wednesday is not a trading day, use the last trading close on or before Wednesday",
            "return_basis": "Yahoo Finance adjusted-close simple return (corporate-action adjusted where supplied): P_t/P_(t-1)-1",
            "observation_count": 104,
            "price_bar_count": 105,
            "window_start_sample_date": slots[0].isoformat(),
            "window_end_sample_date": slots[-1].isoformat(),
            "lookahead_check": "passed: no price date after 2025-10-15",
        },
        "benchmark": {
            "ticker": BENCHMARK,
            "records": benchmark_records,
        },
        "securities": weekly_securities,
    }
    market = {
        "metadata": {
            "schema_version": "2.0.0",
            "valuation_cutoff_date": CUTOFF.isoformat(),
            "data_provider": "Yahoo Finance historical chart data for NSE-listed securities",
            "price_field": "quote.close (unadjusted closing price)",
            "all_observations_at_or_before_cutoff": True,
        },
        "securities": snapshot_securities,
    }
    return weekly, market


if __name__ == "__main__":
    weekly, market = build()
    (SOURCE_DIR / "weekly_market_returns_104w_v2.json").write_text(
        json.dumps(weekly, indent=2), encoding="utf-8"
    )
    (SOURCE_DIR / "market_and_beta_2025_10_15_v2.json").write_text(
        json.dumps(market, indent=2), encoding="utf-8"
    )
    print("Wrote cutoff-safe v2 market datasets")
    for ticker, data in weekly["securities"].items():
        print(
            f"{ticker}: last={data['records'][-1]['date']} "
            f"adjusted_close={data['records'][-1].get('adjclose', data['records'][-1]['close']):.2f} "
            f"raw_beta={data['raw_beta']:.4f} "
            f"blume={data['blume_adjusted_beta']:.4f} "
            f"R2={data['r_squared']:.4f}"
        )
