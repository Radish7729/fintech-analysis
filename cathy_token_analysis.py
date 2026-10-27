"""
Cathy's group: tokenized Money Market Funds, ETFs, Private Credit
Daily price + volume (5y) from yfinance for tokens and underlying/proxy assets,
then token-vs-underlying comparison.

Run locally (needs internet):  pip install yfinance pandas openpyxl && python cathy_token_analysis.py
Output: cathy_token_analysis.xlsx  (one sheet per series, a Coverage sheet, a Comparison sheet)

IMPORTANT: Yahoo only lists a subset of tokens. Every ticker below is a CANDIDATE.
The Coverage sheet records which ones actually returned data; anything that fails
needs a different source (CoinGecko, rwa.xyz, DefiLlama, issuer NAV pages).
"""
import datetime as dt
import numpy as np
import pandas as pd
import yfinance as yf

END = dt.date.today()
START = END - dt.timedelta(days=5 * 365 + 2)

# (group, token label, token ticker, underlying label, underlying/proxy ticker, note)
PAIRS = [
    # --- Money market funds ---
    ("MMF", "Franklin BENJI (FOBXX token)", None, "Franklin OnChain US Govt MMF", "FOBXX",
     "Token is a NAV-$1 fund share; mostly not on Yahoo. Underlying fund is."),
    ("MMF", "BlackRock BUIDL", None, "T-bill proxy: SGOV", "SGOV",
     "BUIDL has no Yahoo listing; SGOV (0-3m T-bills) is a proxy, not the true underlying."),
    ("MMF", "BlackRock BUIDL", None, "T-bill proxy: BIL", "BIL", "Proxy only."),
    # --- ETFs (tokenized equity/ETF wrappers) ---
    ("ETF", "SPYx (xStocks)", "SPYX-USD", "SPY", "SPY", "Candidate Yahoo crypto ticker; verify."),
    ("ETF", "QQQx (xStocks)", "QQQX-USD", "QQQ", "QQQ", "Candidate Yahoo crypto ticker; verify."),
    ("ETF", "SPYon (Ondo Global Markets)", "SPYON-USD", "SPY", "SPY", "Candidate; may not exist."),
    ("ETF", "QQQon (Ondo Global Markets)", "QQQON-USD", "QQQ", "QQQ", "Candidate; may not exist."),
    # --- Private credit ---
    ("PrivCredit", "Maple SYRUP governance token", "SYRUP-USD", "Listed BDC proxy: BIZD", "BIZD",
     "SYRUP is a governance token, NOT syrupUSDC (the yield-bearing credit token)."),
    ("PrivCredit", "Maple syrupUSDC", None, "Listed BDC proxy: ARCC", "ARCC",
     "syrupUSDC has no Yahoo listing; use DefiLlama/rwa.xyz/CoinGecko for its price/NAV."),
    ("PrivCredit", "Centrifuge JTRSY / JAAA", None, "Private credit ETF proxy: PRIV", "PRIV",
     "Not on Yahoo; proxy only."),
]


def fetch(ticker):
    """Daily OHLCV; returns empty frame on any failure."""
    try:
        df = yf.Ticker(ticker).history(start=START, end=END, interval="1d", auto_adjust=False)
        if df is None or df.empty:
            return pd.DataFrame()
        df.index = pd.to_datetime(df.index).tz_localize(None).normalize()
        return df[["Open", "High", "Low", "Close", "Adj Close", "Volume"]]
    except Exception as e:  # network block, delisted, bad ticker
        print(f"  ! {ticker}: {type(e).__name__}: {e}")
        return pd.DataFrame()


def compare(tok, und):
    """Token vs underlying: premium/discount, tracking error, correlation, volume."""
    j = pd.concat([tok["Close"].rename("tok"), und["Close"].rename("und")], axis=1).dropna()
    if len(j) < 30:
        return None
    r = j.pct_change().dropna()
    ratio = j["tok"] / j["und"]
    base = ratio.iloc[:20].median()          # normalise: tokens often have a different unit price
    prem = ratio / base - 1                  # premium/discount vs. its own starting relationship
    diff = r["tok"] - r["und"]
    return {
        "overlap_days": len(j),
        "first_date": j.index[0].date(),
        "last_date": j.index[-1].date(),
        "return_corr": r["tok"].corr(r["und"]),
        "tracking_error_ann": diff.std() * np.sqrt(252),
        "mean_prem_disc": prem.mean(),
        "max_premium": prem.max(),
        "max_discount": prem.min(),
        "token_total_return": j["tok"].iloc[-1] / j["tok"].iloc[0] - 1,
        "underlying_total_return": j["und"].iloc[-1] / j["und"].iloc[0] - 1,
    }


def main():
    series, coverage, comps = {}, [], []
    cache = {}

    def get(t):
        if t not in cache:
            print(f"fetching {t}")
            cache[t] = fetch(t)
        return cache[t]

    for grp, tl, tt, ul, ut, note in PAIRS:
        tdf = get(tt) if tt else pd.DataFrame()
        udf = get(ut)
        for label, tick, df in ((tl, tt, tdf), (ul, ut, udf)):
            if tick:
                coverage.append({
                    "group": grp, "asset": label, "ticker": tick,
                    "rows": len(df),
                    "first": df.index.min().date() if len(df) else None,
                    "last": df.index.max().date() if len(df) else None,
                    "has_volume": bool(len(df) and df["Volume"].sum() > 0),
                    "status": "OK" if len(df) else "NO DATA ON YAHOO / FETCH FAILED",
                })
            else:
                coverage.append({"group": grp, "asset": label, "ticker": None, "rows": 0,
                                 "first": None, "last": None, "has_volume": False,
                                 "status": "NO YAHOO TICKER - use alternative source"})
            if tick and len(df):
                series[tick] = df
        if tt and len(tdf) and len(udf):
            c = compare(tdf, udf)
            if c:
                comps.append({"group": grp, "token": tl, "token_ticker": tt,
                              "underlying": ul, "underlying_ticker": ut, **c, "note": note})

    with pd.ExcelWriter("cathy_token_analysis.xlsx") as xw:
        pd.DataFrame(coverage).drop_duplicates(["asset", "ticker"]).to_excel(xw, "Coverage", index=False)
        pd.DataFrame(comps).to_excel(xw, "Comparison", index=False)
        for t, df in series.items():
            df.to_excel(xw, t.replace("-", "_")[:31])
    print("wrote cathy_token_analysis.xlsx")


if __name__ == "__main__":
    main()
