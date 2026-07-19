#!/usr/bin/env python3
"""Fetch A-share / ETF quotes for portfolio tracking."""

from __future__ import annotations

import time
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, Optional

import akshare as ak


@dataclass
class Quote:
    code: str
    price: float
    change_pct: Optional[float] = None
    trade_date: Optional[str] = None
    source: str = ""


ETF_CODES = {"588810", "159819", "515880", "510880", "510300"}
ALL_CODES = [
    "588810", "159819", "002384", "002475", "603228", "600183",
    "515880", "600487", "001359", "002156", "002222", "510880", "510300",
]


def _stock_symbol(code: str) -> str:
    code = str(code).zfill(6)
    if code.startswith(("6", "5", "9")):
        return f"sh{code}"
    return f"sz{code}"


def _retry(fn, retries: int = 3, delay: float = 2.0):
    last_err = None
    for i in range(retries):
        try:
            return fn()
        except Exception as exc:  # noqa: BLE001
            last_err = exc
            if i < retries - 1:
                time.sleep(delay * (i + 1))
    raise last_err


def fetch_etf_quotes(codes: Optional[list[str]] = None) -> Dict[str, Quote]:
    codes = set(codes or ETF_CODES)
    result: Dict[str, Quote] = {}

    def _load():
        return ak.fund_etf_spot_em()

    df = _retry(_load, retries=4, delay=3.0)
    for code in codes:
        rows = df[df["代码"] == code]
        if rows.empty:
            continue
        row = rows.iloc[0]
        result[code] = Quote(
            code=code,
            price=float(row["最新价"]),
            change_pct=float(row["涨跌幅"]),
            trade_date=str(datetime.now().date()),
            source="fund_etf_spot_em",
        )
    return result


def fetch_stock_quote(code: str) -> Quote:
    symbol = _stock_symbol(code)

    def _load():
        return ak.stock_zh_a_daily(symbol=symbol, adjust="qfq")

    df = _retry(_load, retries=3, delay=2.0)
    if df.empty:
        raise ValueError(f"No daily data for {code}")

    last = df.iloc[-1]
    price = float(last["close"])
    trade_date = str(last["date"])[:10]
    change_pct = None
    if len(df) >= 2:
        prev = float(df.iloc[-2]["close"])
        if prev:
            change_pct = round((price - prev) / prev * 100, 2)

    return Quote(
        code=code,
        price=price,
        change_pct=change_pct,
        trade_date=trade_date,
        source="stock_zh_a_daily",
    )


def fetch_stock_quotes(codes: list[str], pause: float = 0.4) -> Dict[str, Quote]:
    result: Dict[str, Quote] = {}
    for code in codes:
        if code in ETF_CODES:
            continue
        try:
            result[code] = fetch_stock_quote(code)
        except Exception as exc:  # noqa: BLE001
            print(f"  [warn] stock {code}: {exc}")
        time.sleep(pause)
    return result


def fetch_index_shanghai() -> Quote:
    def _load():
        return ak.stock_zh_index_daily(symbol="sh000001")

    df = _retry(_load, retries=3, delay=2.0)
    last = df.iloc[-1]
    prev = df.iloc[-2]
    price = float(last["close"])
    change_pct = round((price - float(prev["close"])) / float(prev["close"]) * 100, 2)
    return Quote(
        code="000001",
        price=price,
        change_pct=change_pct,
        trade_date=str(last["date"])[:10],
        source="stock_zh_index_daily",
    )


def fetch_limit_down_count() -> Optional[int]:
    def _load():
        df = ak.stock_zh_a_spot_em()
        return int((df["涨跌幅"] <= -9.9).sum())

    try:
        return _retry(_load, retries=3, delay=5.0)
    except Exception as exc:  # noqa: BLE001
        print(f"  [warn] limit-down count unavailable: {exc}")
        return None


def fetch_all_quotes(codes: Optional[list[str]] = None) -> Dict[str, Quote]:
    codes = codes or ALL_CODES
    etf_codes = [c for c in codes if c in ETF_CODES]
    stock_codes = [c for c in codes if c not in ETF_CODES]

    print("Fetching ETF quotes...")
    quotes = fetch_etf_quotes(etf_codes)

    print("Fetching stock quotes...")
    quotes.update(fetch_stock_quotes(stock_codes))

    return quotes


if __name__ == "__main__":
    qs = fetch_all_quotes()
    for code in ALL_CODES:
        q = qs.get(code)
        if q:
            print(f"{code}: {q.price} ({q.change_pct}%) [{q.trade_date}] via {q.source}")
        else:
            print(f"{code}: MISSING")

    print("\nFetching Shanghai index...")
    idx = fetch_index_shanghai()
    print(f"上证指数: {idx.price} ({idx.change_pct}%) [{idx.trade_date}]")

    print("\nFetching limit-down count...")
    n = fetch_limit_down_count()
    print(f"跌停家数: {n if n is not None else 'N/A'}")
