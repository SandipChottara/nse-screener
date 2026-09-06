"""
╔══════════════════════════════════════════════════════════════════════════════╗
║  NSE EQUITY MOMENTUM SCREENER v4.1                                          ║
║  Python fills technical scores + downloads NSE F&O Bhavcopy automatically  ║
║  Excel: VLOOKUP pulls OI+PCR → everything recalculates instantly.          ║
╚══════════════════════════════════════════════════════════════════════════════╝

DAILY WORKFLOW (5 min — fully automated):
  1. Double-click ▶ RUN SCREENER (Double-Click).bat  → wait ~5 min
  2. Open Excel → 📋 F&O INPUT sheet already has OI and PCR filled in!
  3. Check ⭐ FINAL CALLS tab — your trade plan is ready
  4. If any stock shows "⏳ Enter OI" — manually type from Sensibull for that stock only
"""

# ─────────────────────────────────────────────────────────────────────────────
#  ★  YOUR SETTINGS — edit these
# ─────────────────────────────────────────────────────────────────────────────
CONFIG = {
    "total_capital":     1_000_000,
    "max_per_trade":       200_000,
    "stop_loss_pct":             5,
    "short_sl_pct":              4,
    "target1_pct":              10,
    "target2_pct":              15,
    "short_target1_pct":         8,
    "short_target2_pct":        12,
    "min_score_buy":            60,
    "priority_buy_score":       75,
    "min_score_short":          55,
    "strong_short_score":       75,
    "data_period":            "1y",
}

NSE_UNIVERSE = [
    "RELIANCE","TCS","HDFCBANK","INFY","ICICIBANK","HINDUNILVR","ITC",
    "SBIN","BHARTIARTL","KOTAKBANK","LT","AXISBANK","ASIANPAINT","MARUTI",
    "SUNPHARMA","TITAN","ULTRACEMCO","NESTLEIND","WIPRO","HCLTECH",
    "TECHM","BAJFINANCE","BAJAJFINSV","NTPC","POWERGRID","COALINDIA",
    "ONGC","JSWSTEEL","TATASTEEL","HINDALCO","CIPLA","DRREDDY",
    "DIVISLAB","APOLLOHOSP","ADANIPORTS","ADANIENT","TATAMOTORS",
    "M&M","HEROMOTOCO","EICHERMOT","BAJAJ-AUTO","BRITANNIA",
    "INDUSINDBK","HDFCLIFE","SBILIFE","GRASIM","SHREECEM","PIDILITIND",
    "VOLTAS","MPHASIS","PERSISTENT","LTIM","COFORGE","OFSS",
    "ZOMATO","TATACOMM","INDHOTEL","IRCTC","CONCOR","SAIL","NMDC",
    "BANKBARODA","CANBK","PNB","FEDERALBNK","IDFCFIRSTB",
    "MUTHOOTFIN","CHOLAFIN","MANAPPURAM","PIIND",
    "AUROPHARMA","TORNTPHARM","ALKEM","LALPATHLAB","METROPOLIS",
    "FORTIS","MAXHEALTH","KANSAINER","BERGEPAINT",
    "SUPREMEIND","POLYCAB","HAVELLS","CROMPTON",
    "DIXON","TRENT","DMART","ABFRL","PAGEIND","VEDL","HINDZINC",
    "MOTHERSON","BALKRISIND","MRF","APOLLOTYRE","EXIDEIND",
    "BATAINDIA","JUBLFOOD","INDIGO",
    "GODREJCP","GODREJPROP","LODHA","DLF","PRESTIGE",
    "OBEROIRLTY","PHOENIXLTD","BRIGADE",
    "TATAPOWER","ADANIGREEN","TORNTPOWER","CESC","NHPC",
    "RECLTD","PFC","IRFC","HUDCO","SJVN",
    "ABB","SIEMENS","CUMMINSIND","BHEL","THERMAX",
    "ASTRAL","RATNAMANI","UBL","RADICO","MCDOWELL-N",
]

# ─────────────────────────────────────────────────────────────────────────────
#  ★  SWING / INVESTMENT LAYER — fundamental gate on top of the technical scan
#  Set fundamentals_csv to a file exported from YOUR OWN screener.in screen
#  (query: "Market Capitalization > 1", broad by design -- see README for the
#  full column list). Using your own screen, not a shared public one, means
#  its columns/criteria can't be changed by someone else without your knowledge.
#  If that file isn't found next to this script, the FUND+TECH/TIER A/
#  TIER B/EXCLUDED tabs are simply skipped -- everything else still runs.
# ─────────────────────────────────────────────────────────────────────────────
FUND_CONFIG = {
    "fundamentals_csv":        "my-full-nse-universe.csv",
    "use_yfinance_fallback":   True,   # fill gaps in the fundamentals file from yfinance, for watchlist stocks only
    "min_roe":                 12.0,   # %
    "min_roce":                 12.0,   # %
    "max_debt_to_equity":        1.5,
    "min_revenue_cagr_3y":       5.0,  # %
    "min_pat_cagr_3y":           0.0,  # % (0 = allow flat, exclude declining)
    "max_promoter_pledge_pct":  20.0,
    "min_market_cap_cr":    20000.0,   # Rs Crore -- excludes micro-caps most prone to pump-and-dump
    "min_avg_daily_turnover_cr": 50.0, # Rs Crore, 20D avg (CMP x AvgVol20) -- excludes illiquid/easily-manipulated stocks
    "require_positive_cfo":     True,  # Cash from Operations must be positive in BOTH of the last 2 years
    "min_interest_coverage":    3.0,   # EBIT / Interest Expense -- below this, debt servicing is a real risk
    "max_debtor_days":         180.0,  # deliberately generous -- only catches clear outliers, not sector norms
    "opm_decline_ratio":        0.75,  # flag if current OPM < 75% of the 5-year average OPM (margin erosion)
    "tier_top_n":                30,
    "fno_long_buildup_price_chg_min": 1.0,   # % price up, same scale as Day Chg %
    "fno_long_buildup_oi_chg_min":    5.0,   # % OI up vs prior day
    "fno_long_buildup_bonus":        10,     # added to LONG SCORE (0-100 scale) when both conditions hit
}

# ─────────────────────────────────────────────────────────────────────────────
#  AUTO-INSTALL
# ─────────────────────────────────────────────────────────────────────────────
import sys, subprocess, importlib, warnings, os
warnings.filterwarnings("ignore")

def ensure(pkg, imp=None):
    try: importlib.import_module(imp or pkg)
    except ImportError:
        print(f"  📦 Installing {pkg}...")
        subprocess.check_call([sys.executable,"-m","pip","install",pkg,"-q"])

print("\n"+"="*68)
print("  NSE SCREENER v4.1 — Checking dependencies...")
print("="*68)
for p,i in [("yfinance","yfinance"),("pandas","pandas"),
            ("openpyxl","openpyxl"),("numpy","numpy"),
            ("requests","requests")]:
    ensure(p,i); print(f"  ✅ {p}")

import yfinance as yf
import pandas as pd
import numpy as np
import requests, zipfile, io, time
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import datetime, date, timedelta

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ─────────────────────────────────────────────────────────────────────────────
#  NSE F&O BHAVCOPY — Free daily file, works from anywhere in the world
#  Source: nsearchives.nseindia.com (static file server, no geo-block)
#  Contains OI for every futures + options contract → we compute per-stock
# ─────────────────────────────────────────────────────────────────────────────

BHAVCOPY_HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/122.0.0.0 Safari/537.36"),
    "Accept":          "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-IN,en;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection":      "keep-alive",
    "Referer":         "https://www.nseindia.com/",
}

def bhavcopy_url(dt):
    ds = dt.strftime("%Y%m%d")
    return (f"https://nsearchives.nseindia.com/content/fo/"
            f"BhavCopy_NSE_FO_0_0_0_{ds}_F_0000.csv.zip")

def _parse_bhavcopy_zip(content_bytes):
    """Parse ZIP bytes → DataFrame. Keeps original column names (case-sensitive)."""
    try:
        z = zipfile.ZipFile(io.BytesIO(content_bytes))
        csv_files = [n for n in z.namelist() if n.endswith(".csv")]
        if not csv_files:
            return None
        with z.open(csv_files[0]) as f:
            df = pd.read_csv(f, low_memory=False)
        df.columns = [c.strip() for c in df.columns]
        return df
    except Exception:
        return None

def _find_local_bhavcopy():
    """
    Check if user manually dropped a bhavcopy ZIP in the script folder.
    File name pattern: BhavCopy_NSE_FO_*.csv.zip
    Returns (today_df, prev_df) — prev_df may be None.
    """
    import glob
    pattern = os.path.join(SCRIPT_DIR, "BhavCopy_NSE_FO_*.csv.zip")
    files = sorted(glob.glob(pattern), reverse=True)
    if not files:
        return None, None
    print(f"  📂 Local file found: {os.path.basename(files[0])}")
    try:
        with open(files[0], "rb") as f:
            today_df = _parse_bhavcopy_zip(f.read())
        prev_df = None
        if len(files) > 1:
            with open(files[1], "rb") as f:
                prev_df = _parse_bhavcopy_zip(f.read())
        return today_df, prev_df
    except Exception as e:
        print(f"  ⚠️  Could not read local file: {e}")
        return None, None

# ─────────────────────────────────────────────────────────────────────────────
#  ★  STOCK UNIVERSE -- separate, purpose-built lists per strategy
#
#  F&O category:    the F&O-eligible universe is derived DYNAMICALLY from the
#                    F&O Bhavcopy download itself (see main() -- it's just
#                    fo_data.keys() after Phase 0 runs). No hardcoded list to
#                    maintain; it auto-updates whenever NSE reconstitutes the
#                    F&O stock list each quarter.
#
#  Swing/Investment: a much larger, quality-screened universe -- Nifty 500
#                    constituents (index inclusion itself requires minimum
#                    liquidity/market-cap/trading-history, which is already a
#                    first layer of "not an obvious shell company" filtering),
#                    downloaded fresh each run with the fallback list below if
#                    the download fails.
# ─────────────────────────────────────────────────────────────────────────────

NIFTY500_URL = "https://nsearchives.nseindia.com/content/indices/ind_nifty500list.csv"

# Used ONLY if the Nifty 500 download fails (network/geo-block) -- keeps the
# script usable offline, same fallback philosophy as the Bhavcopy handling.
FALLBACK_SWING_INVESTMENT_UNIVERSE = None  # set below, right after NSE_UNIVERSE is defined


def download_nifty500_list():
    """
    Returns a list of NSE trading symbols for Nifty 500 constituents, or
    None if the download fails after retries. Warms up session cookies
    against the main site first (NSE's anti-bot protection typically
    requires this) and retries a few times with diagnostics, rather than
    giving up silently after one attempt -- same reliability pattern
    already used for the Bhavcopy download.
    """
    s = requests.Session()
    s.headers.update(BHAVCOPY_HEADERS)
    try:
        s.get("https://www.nseindia.com/", timeout=15)  # warm up cookies
    except Exception:
        pass  # not fatal -- some networks still succeed on the actual request without this

    for attempt in range(3):
        try:
            print(f"  📥 Nifty 500 list, attempt {attempt+1}/3 ... ", end="", flush=True)
            r = s.get(NIFTY500_URL, timeout=20)
            print(f"HTTP {r.status_code} ({len(r.content):,} bytes)")
            if r.status_code == 200 and len(r.content) > 1000:
                df = pd.read_csv(io.StringIO(r.content.decode("utf-8", errors="ignore")))
                df.columns = [c.strip() for c in df.columns]
                sym_col = next((c for c in df.columns if c.strip().lower() == "symbol"), None)
                if sym_col is None:
                    print("     → 'Symbol' column not found in response, retrying")
                else:
                    symbols = df[sym_col].dropna().astype(str).str.strip().tolist()
                    if symbols:
                        return symbols
                    print("     → parsed but empty, retrying")
            else:
                print("     → unexpected response, retrying")
        except Exception as e:
            print(f"failed ({e}), retrying" if attempt < 2 else f"failed ({e})")
    return None


def load_surveillance_exclusions():
    """
    Best-effort download of NSE's ASM (Additional Surveillance Measure) list
    -- stocks flagged for unusual price volatility / suspected manipulation.
    NSE's ASM report page (nseindia.com/reports/asm) is dashboard-driven
    rather than a stable static CSV URL, so this is genuinely best-effort:
    on any failure it returns an empty set and prints a note, rather than
    blocking the run. Treat this as a bonus filter, not a guarantee -- it
    catches securities NSE has already flagged, not ones that haven't been
    flagged yet.
    """
    try:
        s = requests.Session()
        s.headers.update(BHAVCOPY_HEADERS)
        s.headers.update({"Accept": "application/json"})
        # warm up session cookies against the main site first, same pattern NSE's
        # anti-bot protection expects for its /api/ endpoints
        s.get("https://www.nseindia.com/", timeout=15)
        r = s.get("https://www.nseindia.com/api/reportTypes?index=equity", timeout=15)
        # NSE's dashboard API surface changes periodically; if this doesn't
        # resolve to a usable list, fail quietly rather than guess further.
        if r.status_code != 200:
            print("  ⚠️  NSE ASM/GSM list not reachable this run -- skipping that filter (not fatal).")
            return set()
        return set()  # placeholder: NSE's ASM API shape needs a one-time manual check to wire up fully -- see HOW TO USE
    except Exception:
        print("  ⚠️  NSE ASM/GSM list not reachable this run -- skipping that filter (not fatal).")
        return set()


def _bhavcopy_cache_path(dt):
    cache_dir = os.path.join(SCRIPT_DIR, "bhavcopy_cache")
    os.makedirs(cache_dir, exist_ok=True)
    return os.path.join(cache_dir, f"BhavCopy_NSE_FO_{dt.strftime('%Y%m%d')}.csv.zip")


def _save_to_cache(dt, content_bytes):
    try:
        with open(_bhavcopy_cache_path(dt), "wb") as f:
            f.write(content_bytes)
    except Exception:
        pass  # caching is a nice-to-have, never let it break the main flow


def _load_from_cache(dt):
    path = _bhavcopy_cache_path(dt)
    if os.path.exists(path):
        try:
            with open(path, "rb") as f:
                return _parse_bhavcopy_zip(f.read())
        except Exception:
            return None
    return None


def _fetch_prev_day_df(session, dt):
    """
    Get the previous trading day's Bhavcopy DataFrame for computing OI change.
    Tries, in order: (1) local cache from a prior run, (2) network fetch --
    retrying across up to 3 prior trading days with diagnostics printed for
    each attempt, instead of silently giving up after one failed request.
    Every successful network fetch is cached, so tomorrow's run won't need
    this network call at all for today's date.
    """
    prev = dt - timedelta(days=1)
    attempts = 0
    while attempts < 5:  # walk back through weekends/holidays, cap the search
        if prev.weekday() < 5:
            cached = _load_from_cache(prev)
            if cached is not None:
                print(f"      (prev-day {prev.strftime('%d-%b')}: using local cache)")
                return cached

            try:
                r2 = session.get(bhavcopy_url(prev), timeout=15)
                print(f"      (prev-day {prev.strftime('%d-%b')}: HTTP {r2.status_code}, {len(r2.content):,} bytes)")
                if r2.status_code == 200 and len(r2.content) > 5000:
                    df = _parse_bhavcopy_zip(r2.content)
                    if df is not None:
                        _save_to_cache(prev, r2.content)
                        return df
                    print("      → parse failed, trying one day further back")
                elif r2.status_code == 404:
                    print("      → not found (holiday?), trying one day further back")
                else:
                    print("      → unexpected response, trying one day further back")
            except Exception as e:
                print(f"      → request failed ({e}), trying one day further back")
            attempts += 1
        prev -= timedelta(days=1)
    print("      ⚠️  Could not get previous-day OI after 5 attempts -- OI change will show as unavailable")
    return None


def download_bhavcopy(target_date=None):
    """
    Get F&O OI + PCR data. Three strategies in order:
      1. Auto-download from NSE archives (may be geo-blocked outside India)
      2. Read locally-downloaded ZIP file from script folder
      3. Return {} → Excel yellow cells for manual Sensibull entry
    """
    if target_date is None:
        target_date = date.today()

    # ── Strategy 1: Auto-download ─────────────────────────────────────────
    print(f"  📡 Attempting auto-download from NSE archives...")
    s = requests.Session()
    s.headers.update(BHAVCOPY_HEADERS)

    candidates = []
    for d in range(8):
        dt = target_date - timedelta(days=d)
        if dt.weekday() < 5:
            candidates.append(dt)
        if len(candidates) >= 4:
            break

    geo_blocked = False
    for dt in candidates:
        url = bhavcopy_url(dt)
        try:
            print(f"  📥 {dt.strftime('%d-%b-%Y')} ... ", end="", flush=True)
            r = s.get(url, timeout=22)
            print(f"HTTP {r.status_code} ({len(r.content):,} bytes)", end=" ")

            if r.status_code == 200 and len(r.content) > 5000:
                today_df = _parse_bhavcopy_zip(r.content)
                if today_df is not None:
                    _save_to_cache(dt, r.content)  # so tomorrow's run can use this as "prev day" from cache
                    print()  # newline before prev-day diagnostics
                    prev_df = _fetch_prev_day_df(s, dt)
                    result = _compute_stock_fo(today_df, prev_df)
                    if result:
                        print(f"  → ✅ {len(result)} stocks")
                        return result
                    else:
                        print(f"  → parsed but 0 stocks (cols: {list(today_df.columns[:5])})")
                else:
                    print("→ parse failed")
            elif r.status_code == 404:
                print("→ not published yet (NSE publishes after 5:30 PM IST)")
            elif r.status_code in (403, 401):
                print("→ 403 BLOCKED (geo-restriction)")
                geo_blocked = True
                break
            else:
                print(f"→ skipped")
        except requests.exceptions.ConnectionError:
            print("→ connection refused (geo-blocked or network issue)")
            geo_blocked = True
            break
        except requests.exceptions.Timeout:
            print("→ timeout, trying previous day")
        except Exception as e:
            print(f"→ {e}")
            break

    # ── Strategy 2: Local file ────────────────────────────────────────────
    print(f"\n  📂 Checking for local bhavcopy file in script folder...")
    today_df, prev_df = _find_local_bhavcopy()
    if today_df is not None:
        result = _compute_stock_fo(today_df, prev_df)
        if result:
            print(f"  ✅ {len(result)} stocks loaded from local file!")
            return result

    # ── Strategy 3: Guide user ────────────────────────────────────────────
    print(f"""
  ╔══════════════════════════════════════════════════════════════╗
  ║  ⚠️  F&O BHAVCOPY UNAVAILABLE — NSE geo-blocks non-India IPs ║
  ╠══════════════════════════════════════════════════════════════╣
  ║                                                              ║
  ║  EASY FIX (30 seconds, once per day after 5:30 PM IST):    ║
  ║                                                              ║
  ║  1. Open Chrome/Edge and go to:                             ║
  ║     nseindia.com/all-reports-derivatives                    ║
  ║                                                              ║
  ║  2. Find: "F&O-UDiFF Common Bhavcopy Final (zip)"          ║
  ║     Click the date link for TODAY                           ║
  ║                                                              ║
  ║  3. Save the ZIP file to:                                   ║
  ║     {SCRIPT_DIR[:52]:<52}║
  ║     (same folder as this script)                            ║
  ║                                                              ║
  ║  4. Re-run the script → data fills automatically!           ║
  ║                                                              ║
  ║  OR: Type OI + PCR manually in yellow cells from Sensibull  ║
  ╚══════════════════════════════════════════════════════════════╝
""")
    return {}


def _compute_stock_fo(today_df, prev_df):
    """
    Compute per-stock Today OI, Prev OI, PCR from UDiFF bhavcopy DataFrames.

    Confirmed UDiFF column names (from NSE format spec):
      TckrSymb     = ticker symbol  (e.g. RELIANCE, SBIN)
      FinInstrmTp  = FUTSTK / OPTSTK / FUTIDX / OPTIDX
      XpryDt       = expiry date
      OptnTp       = CE or PE   ← this is how PCR is computed
      OpnIntrst    = open interest
    """
    result = {}

    # ── Exact UDiFF column names ──────────────────────────────────────────
    SYM = "TckrSymb"
    TYP = "FinInstrmTp"
    EXP = "XpryDt"
    OPT = "OptnTp"       # "CE" or "PE"
    OI  = "OpnIntrst"

    # Strip all column names of whitespace first
    today_df.columns = [c.strip() for c in today_df.columns]

    # Verify required columns exist
    missing = [c for c in [SYM, TYP, OI] if c not in today_df.columns]
    if missing:
        # Print all columns to help debug
        print(f"\n  ⚠️  Bhavcopy columns not as expected. Missing: {missing}")
        print(f"     Actual columns: {list(today_df.columns)}")
        return {}

    today_df[OI] = pd.to_numeric(today_df[OI], errors="coerce").fillna(0)

    # Prep prev_df
    if prev_df is not None:
        prev_df.columns = [c.strip() for c in prev_df.columns]
        if OI in prev_df.columns:
            prev_df[OI] = pd.to_numeric(prev_df[OI], errors="coerce").fillna(0)
        else:
            prev_df = None

    # Index columns for fast lookup
    has_exp = EXP in today_df.columns
    has_opt = OPT in today_df.columns

    # Only process stock instruments (not index futures/options)
    stk_df = today_df[today_df[TYP].isin(["STF","STO"])]   # STF=futures, STO=options
    symbols = stk_df[SYM].dropna().unique()

    for sym in symbols:
        sym = str(sym).strip()
        if not sym:
            continue

        s = stk_df[stk_df[SYM] == sym]

        # ── All-month Futures OI (near + mid + far combined) ─────────────
        # Total across all expiries gives true picture even in expiry week
        # (near-month OI drops artificially during rollover)
        fut      = s[s[TYP] == "STF"]
        today_oi = int(fut[OI].sum()) if not fut.empty else 0

        # ── PCR from ALL strikes and ALL expiries ─────────────────────────
        # Total Put OI / Total Call OI = real market sentiment
        pcr = None
        if has_opt:
            opt = s[s[TYP] == "STO"]
            if not opt.empty:
                put_oi  = opt[opt[OPT] == "PE"][OI].sum()
                call_oi = opt[opt[OPT] == "CE"][OI].sum()
                if call_oi > 0:
                    pcr = round(float(put_oi) / float(call_oi), 2)

        # ── Previous day all-month Futures OI ────────────────────────────
        prev_oi = 0
        if prev_df is not None and TYP in prev_df.columns:
            ps = prev_df[(prev_df[SYM] == sym) & (prev_df[TYP] == "STF")]
            prev_oi = int(ps[OI].sum()) if not ps.empty else 0

        if today_oi > 0 or pcr is not None:
            result[sym] = {
                "today_oi": today_oi,
                "prev_oi":  prev_oi,
                "pcr":      pcr,
            }

    return result

# ─────────────────────────────────────────────────────────────────────────────
#  TECHNICAL INDICATORS
# ─────────────────────────────────────────────────────────────────────────────

def ema(s,p):      return s.ewm(span=p,adjust=False).mean()
def vol_avg(v,p):  return v.rolling(p).mean()

def rsi(s,p=14):
    d=s.diff(); ag=d.clip(lower=0).ewm(com=p-1,adjust=False).mean()
    al=(-d.clip(upper=0)).ewm(com=p-1,adjust=False).mean()
    return 100-(100/(1+ag/al.replace(0,np.nan)))

def macd_calc(s,fast=12,slow=26,sig=9):
    ml=ema(s,fast)-ema(s,slow); sl=ema(ml,sig); return ml,sl,ml-sl

# ─────────────────────────────────────────────────────────────────────────────
#  LONG SCORING
# ─────────────────────────────────────────────────────────────────────────────

def score_ema_long(c,e20,e50,e200):
    if c>e20 and e20>e50:  return 25,"Strong Uptrend ↑↑"
    if c>e50  and c>e200:  return 15,"Medium Uptrend ↑"
    if c>e200:             return  8,"Above 200 EMA"
    return 0,"Downtrend ↓"

def score_rsi_long(r):
    if 55<=r<=70:  return 20,f"{r:.1f} ✅ Sweet Spot"
    if 50<=r<55:   return 12,f"{r:.1f} Recovering"
    if 40<=r<50:   return  5,f"{r:.1f} Weak"
    if r>75:       return  8,f"{r:.1f} ⚠️ Overbought"
    return 0,f"{r:.1f} Bearish"

def score_vol_long(tv,av):
    if av==0: return 0,"No Data"
    rat=tv/av
    if rat>=2.0: return 20,f"{rat:.1f}x 🔥 Surge"
    if rat>=1.5: return 14,f"{rat:.1f}x Strong"
    if rat>=1.0: return  7,f"{rat:.1f}x Average"
    return 0,f"{rat:.1f}x Low"

def score_macd_long(h,hp):
    if h>0 and hp<=0: return 15,"🟢 Fresh Cross"
    if h>0:           return 10,"Above Signal"
    if h<0 and h>hp:  return  8,"About to Cross"
    return 0,"Below Signal"

def pattern_long(df):
    try:
        c,hi,lo,v=df["Close"],df["High"],df["Low"],df["Volume"]
        e20v=ema(c,20).iloc[-1]; av=vol_avg(v,20).iloc[-1]
        c0,l0,v0=c.iloc[-1],lo.iloc[-1],v.iloc[-1]
        wk52h=hi.iloc[-252:].max() if len(hi)>=252 else hi.max()
        if c0>=wk52h*0.995 and v0>1.5*av:        return 15,"🚀 52W Breakout"
        if abs(l0-e20v)/e20v<0.015 and c0>e20v and c0>c.iloc[-2]:
                                                  return 15,"📉→📈 EMA Pullback"
        if lo.iloc[-1]>lo.iloc[-2]>lo.iloc[-3] and c0>c.iloc[-3]:
                                                  return 10,"📈 Higher Lows"
        if hi.iloc[-1]<=hi.iloc[-2] and lo.iloc[-1]>=lo.iloc[-2]:
                                                  return 10,"⚡ Inside Bar"
        if len(c)>11:
            rally=(c.iloc[-6]-c.iloc[-11])/c.iloc[-11]
            cons=(c.iloc[-5:].max()-c.iloc[-5:].min())/c.iloc[-5:].mean()
            if rally>0.05 and cons<0.03:          return 12,"🏁 Bull Flag"
    except: pass
    return 0,"No Pattern"

# ─────────────────────────────────────────────────────────────────────────────
#  SHORT SCORING
# ─────────────────────────────────────────────────────────────────────────────

def score_ema_short(c,e20,e50,e200):
    if c<e20 and e20<e50 and e50<e200: return 25,"Strong Downtrend ↓↓"
    if c<e50  and c<e200:              return 20,"Below Key EMAs ↓"
    if c<e200:                         return 12,"Below 200 EMA"
    if c<e20  and c>e200:              return  6,"Weak Short-term"
    return 0,"No Short Setup"

def score_rsi_short(r):
    if r<30:         return 20,f"{r:.1f} Bearish Momentum"
    if 30<=r<40:     return 15,f"{r:.1f} Bearish Zone"
    if 40<=r<45:     return  8,f"{r:.1f} Weak"
    return 0,f"{r:.1f} Not Bearish"

def score_vol_short(tv,av,price_fell):
    if av==0: return 0,"No Data"
    rat=tv/av
    if not price_fell: return 0,"Price rose today"
    if rat>=2.0: return 20,f"{rat:.1f}x 🔥 Heavy Sell"
    if rat>=1.5: return 14,f"{rat:.1f}x Sell Surge"
    if rat>=1.0: return  7,f"{rat:.1f}x Average"
    return 0,f"{rat:.1f}x Low"

def score_macd_short(h,hp):
    if h<0 and hp>=0: return 15,"🔴 Fresh Bear Cross"
    if h<0:           return 10,"Below Signal"
    if h>0 and h<hp:  return  5,"Fading"
    return 0,"No Short Signal"

def pattern_short(df):
    try:
        c,hi,lo,v=df["Close"],df["High"],df["Low"],df["Volume"]
        e20v=ema(c,20).iloc[-1]; av=vol_avg(v,20).iloc[-1]
        c0,h0,v0=c.iloc[-1],hi.iloc[-1],v.iloc[-1]
        wk52l=lo.iloc[-252:].min() if len(lo)>=252 else lo.min()
        if c0<=wk52l*1.005 and v0>1.5*av:        return 15,"💥 52W Breakdown"
        if abs(h0-e20v)/e20v<0.015 and c0<e20v and c0<c.iloc[-2]:
                                                  return 15,"📈→📉 EMA Rejection"
        if hi.iloc[-1]<hi.iloc[-2]<hi.iloc[-3] and c0<c.iloc[-3]:
                                                  return 10,"📉 Lower Highs"
        if len(c)>11:
            fall=(c.iloc[-11]-c.iloc[-6])/c.iloc[-11]
            cons=(c.iloc[-5:].max()-c.iloc[-5:].min())/c.iloc[-5:].mean()
            if fall>0.05 and cons<0.03:           return 12,"🐻 Bear Flag"
    except: pass
    return 0,"No Pattern"

# ─────────────────────────────────────────────────────────────────────────────
#  ANALYSE ONE STOCK
# ─────────────────────────────────────────────────────────────────────────────

def analyse_stock(symbol):
    try:
        df = yf.Ticker(f"{symbol}.NS").history(
            period=CONFIG["data_period"], interval="1d", auto_adjust=True)
        if df is None or len(df) < 60:
            return None
        df = df.dropna(subset=["Close","Volume"])
        c,v = df["Close"], df["Volume"]

        e20v  = ema(c,20).iloc[-1];  e50v  = ema(c,50).iloc[-1]
        e200v = (ema(c,200) if len(c)>=200 else ema(c,len(c)//2)).iloc[-1]
        rsi_v = rsi(c).iloc[-1]
        ml,sl_line,hist = macd_calc(c)
        hv,hp = hist.iloc[-1], (hist.iloc[-2] if len(hist)>=2 else 0)
        tv,av = v.iloc[-1], vol_avg(v,20).iloc[-1]

        cmp        = round(c.iloc[-1], 2)
        prev_close = round(c.iloc[-2], 2)
        price_up   = cmp >= prev_close
        price_fell = cmp < prev_close
        day_chg    = round(((cmp-prev_close)/prev_close)*100, 2)
        vrat       = round(tv/av,2) if av>0 else 0

        wk52h = df["High"].iloc[-252:].max() if len(df)>=252 else df["High"].max()
        wk52l = df["Low"].iloc[-252:].min()  if len(df)>=252 else df["Low"].min()

        sl,lema  = score_ema_long(cmp,e20v,e50v,e200v)
        sr,lrsi  = score_rsi_long(rsi_v)
        sv,lvol  = score_vol_long(tv,av)
        sm,lmac  = score_macd_long(hv,hp)
        sp,lpat  = pattern_long(df)
        long_score = min(100, sl+sr+sv+sm+sp)

        sl2,sema = score_ema_short(cmp,e20v,e50v,e200v)
        sr2,srsi = score_rsi_short(rsi_v)
        sv2,svol = score_vol_short(tv,av,price_fell)
        sm2,smac = score_macd_short(hv,hp)
        sp2,spat = pattern_short(df)
        short_score = min(100, sl2+sr2+sv2+sm2+sp2)

        if   long_score>=CONFIG["priority_buy_score"]: long_sig="🟢 PRIORITY BUY"
        elif long_score>=CONFIG["min_score_buy"]:      long_sig="🔵 WATCHLIST BUY"
        elif long_score>=45:                           long_sig="🟡 WEAK"
        else:                                          long_sig="🔴 SKIP"

        if   short_score>=CONFIG["strong_short_score"]: short_sig="🔻 STRONG SHORT"
        elif short_score>=CONFIG["min_score_short"]:    short_sig="🔽 SHORT"
        elif short_score>=40:                           short_sig="🟡 WEAK SHORT"
        else:                                           short_sig="—"

        sl_l = round(cmp*(1-CONFIG["stop_loss_pct"]/100),2)
        sl_s = round(cmp*(1+CONFIG["short_sl_pct"]/100),2)
        t1l  = round(cmp*(1+CONFIG["target1_pct"]/100),2)
        t2l  = round(cmp*(1+CONFIG["target2_pct"]/100),2)
        t1s  = round(cmp*(1-CONFIG["short_target1_pct"]/100),2)
        t2s  = round(cmp*(1-CONFIG["short_target2_pct"]/100),2)
        qty  = int(CONFIG["max_per_trade"]/cmp) if cmp>0 else 0

        return {
            "Symbol": symbol, "Prev Close": prev_close, "CMP": cmp,
            "Day Chg %": day_chg, "price_up": price_up,
            "20 EMA": round(e20v,2), "50 EMA": round(e50v,2),
            "200 EMA": round(e200v,2), "RSI": round(rsi_v,1),
            "Vol Ratio": vrat, "AvgVol20": round(av,0) if av==av else 0, "52W High": round(wk52h,2),
            "52W Low": round(wk52l,2),
            "% from 52W High": round(((cmp-wk52h)/wk52h)*100,1),
            "EMA Sc(L)":sl,  "RSI Sc(L)":sr,  "Vol Sc(L)":sv,
            "MACD Sc(L)":sm, "Pat Sc(L)":sp,
            "LONG SCORE": long_score, "LONG SIGNAL": long_sig,
            "EMA Trend": lema, "RSI Status": lrsi,
            "Vol Status": lvol, "MACD Status": lmac,
            "Pattern Long": lpat,
            "EMA Sc(S)":sl2, "RSI Sc(S)":sr2, "Vol Sc(S)":sv2,
            "MACD Sc(S)":sm2,"Pat Sc(S)":sp2,
            "SHORT SCORE": short_score, "SHORT SIGNAL": short_sig,
            "Pattern Short": spat,
            "SL Long":sl_l, "T1 Long":t1l, "T2 Long":t2l,
            "SL Short":sl_s,"T1 Short":t1s,"T2 Short":t2s,
            "Qty": qty,
        }
    except Exception:
        return None

# ─────────────────────────────────────────────────────────────────────────────
#  STYLE HELPERS
# ─────────────────────────────────────────────────────────────────────────────

P = {
    "navy":"1F3864","blue":"2E75B6","lblue":"D6E4F0","dblue":"163A6B",
    "green":"1E8449","lgreen":"D5F5E3","good":"C6EFCE",
    "red":"C0392B","lred":"FDEBD0","bad":"FFC7CE",
    "orange":"E67E22","lorange":"FDEBD0","yellow":"FFF9C4",
    "yelb":"FFD700",  # bright yellow for input cells
    "grey":"F2F3F4","white":"FFFFFF","black":"000000",
    "dgrey":"888888","input":"EAF4FB",
    "sdark":"7B241C","smed":"E74C3C","slight":"FFF0F0",
}

def F(h):  return PatternFill("solid", fgColor=P.get(h,h))
def Fn(bold=False,size=10,color="black",italic=False):
    return Font(bold=bold,size=size,color=P.get(color,color),name="Calibri",italic=italic)
def Al(h="center",v="center",wrap=False):
    return Alignment(horizontal=h,vertical=v,wrap_text=wrap)
def Bd(color="BBBBBB"):
    s=Side(style="thin",color=color); return Border(left=s,right=s,top=s,bottom=s)
def BdThick():
    t=Side(style="medium",color="888888"); th=Side(style="thin",color="BBBBBB")
    return Border(left=t,right=t,top=th,bottom=th)

def W(ws,ref,val,bg="white",fg="black",bold=False,sz=10,fmt=None,h="center",wrap=False):
    c=ws[ref]; c.value=val
    c.font=Fn(bold,sz,fg); c.fill=F(bg)
    c.alignment=Al(h,"center",wrap); c.border=Bd()
    if fmt: c.number_format=fmt
    return c

def MH(ws,rng,text,bg="navy",fg="white",sz=12,bold=True):
    ws.merge_cells(rng)
    c=ws[rng.split(":")[0]]; c.value=text
    c.font=Fn(bold,sz,fg); c.fill=F(bg)
    c.alignment=Al("center","center",True); c.border=Bd()

def CW(ws,d):
    for k,v in d.items(): ws.column_dimensions[k].width=v
def RH(ws,d):
    for k,v in d.items(): ws.row_dimensions[k].height=v

# ─────────────────────────────────────────────────────────────────────────────
#  ★  SWING / INVESTMENT LAYER — fundamentals loading, gate, scoring
#  Reads a screener.in export (any column wording -- auto-detected below,
#  no manual renaming needed week to week) and layers a fundamental
#  pass/fail gate + composite score on top of the technical scan above.
# ─────────────────────────────────────────────────────────────────────────────

import re as _re

FUND_REQUIRED = [
    "symbol", "roe", "roce", "debt_to_equity", "revenue_cagr_3y",
    "pat_cagr_3y", "pe_ttm", "pe_5y_median", "promoter_pledge_pct", "market_cap_cr",
    "cfo_last_year", "cfo_preceding_year", "debtor_days", "interest_coverage",
    "opm", "opm_5y", "dividend_payout_pct",
]

# Each field -> list of keyword-sets, most specific first. A header matches
# a field if every keyword in one of its sets appears (as a token, exact or
# close substring) in the header. See auto_map_fund_columns() for how ties
# between multiple matching columns are resolved.
FUND_ALIASES = {
    "symbol":              [["nse", "code"], ["symbol"], ["ticker"], ["name"]],
    "roe":                 [["return", "equity"], ["roe"]],
    "roce":                [["return", "capital", "employed"], ["roce"]],
    "debt_to_equity":      [["debt", "equity"]],
    "revenue_cagr_3y":     [["sales", "growth", "3"], ["revenue", "growth", "3"], ["sales", "growth"]],
    "pat_cagr_3y":         [["profit", "growth", "3"], ["pat", "growth", "3"], ["profit", "growth"]],
    "pe_ttm":              [["price", "earning"], ["pe", "ratio"], ["p", "e"]],
    "pe_5y_median":        [["historical", "pe", "5"], ["median", "pe"], ["5", "year", "pe"], ["5yr", "pe"]],
    "promoter_pledge_pct": [["pledge"]],
    "market_cap_cr":       [["market", "capitalization"], ["market", "cap"]],
    "cfo_last_year":       [["cash", "operations", "last", "year"], ["cash", "operating", "last", "year"]],
    "cfo_preceding_year":  [["cash", "operations", "preceding", "year"], ["cash", "operating", "preceding", "year"]],
    "debtor_days":         [["days", "receivable"], ["debtor", "days"], ["receivable", "days"]],
    "interest_coverage":   [["interest", "coverage"]],
    "opm":                 [["opm"]],
    "opm_5y":              [["opm", "5"]],
    "dividend_payout_pct": [["dividend", "payout", "ratio"]],  # deliberately NOT matching bare "Dividend Payout" -- ambiguous units
}

def _fund_normalize(header):
    s = _re.sub(r"[^a-z0-9\s]", " ", str(header).lower())
    s = _re.sub(r"(\d)([a-z])", r"\1 \2", s)
    s = _re.sub(r"([a-z])(\d)", r"\1 \2", s)
    return s

def _fund_token_matches(token, keyword):
    if token == keyword:
        return True
    if len(token) >= 4 and len(keyword) >= 4:
        return keyword in token or token in keyword
    return False

def auto_map_fund_columns(raw_columns):
    """
    Matches screener.in's actual export headers to our internal field names
    by keyword, not exact text -- keeps working even if screener.in
    reworks a column or you add/reorder columns differently next week.
    Most-specific keyword-set is tried across ALL columns before a looser
    fallback set, and the CLOSEST match (fewest extra/unrelated words)
    wins when several columns could match the same field -- this is what
    correctly picks plain "Return on equity" over "Average return on
    equity 5Years", for example.
    """
    normalized = {c: [t for t in _fund_normalize(c).split() if t] for c in raw_columns}
    mapping, used = {}, set()
    max_rounds = max((len(v) for v in FUND_ALIASES.values()), default=0)
    for round_idx in range(max_rounds):
        for field, keyword_sets in FUND_ALIASES.items():
            if field in used or round_idx >= len(keyword_sets):
                continue
            keywords = keyword_sets[round_idx]
            candidates = []
            for order_idx, col in enumerate(raw_columns):
                if col in mapping:
                    continue
                toks = normalized[col]
                if all(any(_fund_token_matches(t, kw) for t in toks) for kw in keywords):
                    candidates.append((len(toks) - len(keywords), order_idx, col))
            if candidates:
                candidates.sort(key=lambda x: (x[0], x[1]))
                mapping[candidates[0][2]] = field
                used.add(field)
    return mapping

def _fund_clean_numeric(series):
    cleaned = (series.astype(str).str.replace("%", "", regex=False)
               .str.replace(",", "", regex=False).str.strip()
               .replace({"": None, "-": None, "NA": None, "nan": None, "None": None}))
    return pd.to_numeric(cleaned, errors="coerce")

def load_fundamentals(path):
    """Returns a DataFrame with FUND_REQUIRED columns, or None if the file doesn't exist."""
    if not os.path.exists(path):
        return None
    if path.lower().endswith((".xlsx", ".xls")):
        raw = pd.read_excel(path)
    else:
        raw = pd.read_csv(path, encoding="utf-8-sig")
    raw.columns = [str(c).strip() for c in raw.columns]

    mapping = auto_map_fund_columns(raw.columns.tolist())
    matched = set(mapping.values())
    missing = [f for f in FUND_REQUIRED if f not in matched]
    print("  Fundamentals column auto-detection:")
    for orig, field in mapping.items():
        print(f"    '{orig}' -> {field}")
    if missing:
        print(f"    (not found, skipped in scoring: {', '.join(missing)})")

    df = raw.rename(columns=mapping)
    for col in FUND_REQUIRED:
        if col not in df.columns:
            df[col] = np.nan
    df = df[FUND_REQUIRED].copy()

    df = df[df["symbol"].notna()].copy()
    df["symbol"] = df["symbol"].astype(str).str.strip()
    df = df[df["symbol"] != ""]

    for c in [c for c in FUND_REQUIRED if c != "symbol"]:
        df[c] = _fund_clean_numeric(df[c])

    print(f"  Fundamentals loaded: {len(df)} stocks ({df['roe'].notna().sum()} with ROE data)")
    return df

def yfinance_fill_fundamentals(missing_symbols):
    """Gap-filler for stocks not found in the fundamentals CSV. Reuses the yf import already in this script."""
    rows = []
    for sym in missing_symbols:
        if not isinstance(sym, str) or not sym.strip():
            continue
        try:
            info = yf.Ticker(f"{sym.strip()}.NS").info
        except Exception:
            info = {}
        roe = info.get("returnOnEquity")
        fcf = info.get("freeCashflow")
        rows.append({
            "symbol": sym.strip(),
            "roe": roe * 100 if roe is not None else np.nan,
            "roce": np.nan,
            "debt_to_equity": info.get("debtToEquity"),
            "revenue_cagr_3y": np.nan,
            "pat_cagr_3y": np.nan,
            "pe_ttm": info.get("trailingPE"),
            "pe_5y_median": np.nan,
            "promoter_pledge_pct": np.nan,
        })
    return pd.DataFrame(rows)

def merge_fundamentals(csv_df, yf_df):
    if yf_df.empty:
        return csv_df
    merged = csv_df.set_index("symbol")
    filler = yf_df.set_index("symbol")
    return merged.combine_first(filler).reset_index()

def apply_fund_gate(row, gate):
    """Returns (passes: bool, reasons: list[str]). Missing metrics are skipped, not failed."""
    reasons = []
    def check(v, t, op, label):
        if pd.isna(v):
            return
        if not op(v, t):
            reasons.append(label)
    check(row.get("roe"), gate["min_roe"], lambda v, t: v >= t, "ROE below minimum")
    check(row.get("roce"), gate["min_roce"], lambda v, t: v >= t, "ROCE below minimum")
    check(row.get("debt_to_equity"), gate["max_debt_to_equity"], lambda v, t: v <= t, "Debt/Equity too high")
    check(row.get("revenue_cagr_3y"), gate["min_revenue_cagr_3y"], lambda v, t: v >= t, "Revenue CAGR too low")
    check(row.get("pat_cagr_3y"), gate["min_pat_cagr_3y"], lambda v, t: v >= t, "PAT CAGR declining")
    check(row.get("promoter_pledge_pct"), gate["max_promoter_pledge_pct"], lambda v, t: v <= t, "Promoter pledge too high")
    check(row.get("market_cap_cr"), gate["min_market_cap_cr"], lambda v, t: v >= t, "Market cap too small (micro-cap risk)")

    # Cash from Operations positive in both years -- a company reporting
    # healthy paper profit while burning cash is one of the more reliable
    # early-warning signs of earnings manipulation.
    if gate.get("require_positive_cfo"):
        cfo1, cfo2 = row.get("cfo_last_year"), row.get("cfo_preceding_year")
        if pd.notna(cfo1) and cfo1 < 0:
            reasons.append("Negative operating cash flow (last year)")
        if pd.notna(cfo2) and cfo2 < 0:
            reasons.append("Negative operating cash flow (preceding year)")

    check(row.get("interest_coverage"), gate["min_interest_coverage"], lambda v, t: v >= t, "Interest coverage too low (debt servicing risk)")
    check(row.get("debtor_days"), gate["max_debtor_days"], lambda v, t: v <= t, "Debtor days unusually high (possible channel-stuffing)")

    # Margin erosion: current OPM well below its own 5-year average. Only
    # checked when the 5-year average is meaningfully positive -- comparing
    # against a near-zero or negative historical baseline produces
    # nonsensical ratios rather than a genuine erosion signal.
    opm, opm_5y = row.get("opm"), row.get("opm_5y")
    if pd.notna(opm) and pd.notna(opm_5y) and opm_5y > 1:
        if opm < opm_5y * gate["opm_decline_ratio"]:
            reasons.append(f"Operating margin eroding (now {opm:.1f}% vs 5Y avg {opm_5y:.1f}%)")

    # Dividend payout is deliberately NOT gated: many legitimate high-growth
    # companies pay 0% and reinvest everything, so a minimum-payout rule
    # would wrongly exclude good growth stocks. It's still loaded and shown
    # in FUND+TECH VIEW for your own judgment, just not auto-failed on.

    return (len(reasons) == 0), reasons

def _minmax_0_100(series, higher_is_better=True):
    s = series.astype(float)
    lo, hi = s.min(skipna=True), s.max(skipna=True)
    if pd.isna(lo) or pd.isna(hi) or hi == lo:
        return pd.Series(np.nan, index=series.index)
    norm = (s - lo) / (hi - lo) * 100
    return norm if higher_is_better else (100 - norm)

def fund_composite_score(df):
    """0-100 composite, weights redistributed across whichever metrics are present per row."""
    weights = {"roe": 0.20, "roce": 0.15, "revenue_cagr_3y": 0.20, "pat_cagr_3y": 0.20,
               "pe_discount": 0.15, "low_debt": 0.10}
    pe_discount = pd.Series(np.nan, index=df.index)
    has_pe = df["pe_ttm"].notna() & df["pe_5y_median"].notna() & (df["pe_5y_median"] > 0)
    pe_discount.loc[has_pe] = (df.loc[has_pe, "pe_5y_median"] - df.loc[has_pe, "pe_ttm"]) / df.loc[has_pe, "pe_5y_median"] * 100

    metric_scores = pd.DataFrame({
        "roe": _minmax_0_100(df["roe"]), "roce": _minmax_0_100(df["roce"]),
        "revenue_cagr_3y": _minmax_0_100(df["revenue_cagr_3y"]),
        "pat_cagr_3y": _minmax_0_100(df["pat_cagr_3y"]),
        "pe_discount": _minmax_0_100(pe_discount),
        "low_debt": _minmax_0_100(df["debt_to_equity"], higher_is_better=False),
    })

    def row_score(row):
        avail = row.dropna()
        if avail.empty:
            return np.nan
        w = pd.Series({k: weights[k] for k in avail.index})
        w = w / w.sum()
        return float((avail * w).sum())

    df = df.copy()
    df["fund_score"] = metric_scores.apply(row_score, axis=1)
    return df

# ─────────────────────────────────────────────────────────────────────────────
#  ★  F&O FINAL ACTION -- Python port of the Excel formulas below (columns
#  L/M/N/O/P in F&O INPUT), so the F&O-confirmed picks can be ranked and
#  selected directly in Python for the Top Picks sheet, without waiting for
#  Excel to evaluate the live formulas. Kept in exact lockstep with
#  fo_formulas() -- if you change the Excel logic, mirror it here too.
# ─────────────────────────────────────────────────────────────────────────────

def fo_final_action(long_score, short_score, day_chg, today_oi, prev_oi, pcr):
    """Returns a dict with oi_direction, pcr_signal, fo_trend, final_action, confidence."""
    if today_oi in (None, 0, "") or prev_oi in (None, 0, ""):
        return {"oi_direction": "⏳ Enter OI", "pcr_signal": None, "fo_trend": "⏳ Fill H,I,J",
                "final_action": "⏳ Fill F&O data first", "confidence": "—"}

    if today_oi > prev_oi and day_chg > 0:
        oi_direction = "🟢 Long Buildup"
    elif today_oi < prev_oi and day_chg > 0:
        oi_direction = "🔵 Short Covering"
    elif today_oi > prev_oi and day_chg < 0:
        oi_direction = "🔴 Short Buildup"
    else:
        oi_direction = "🟠 Long Unwinding"

    if pcr in (None, 0, ""):
        pcr_signal = None
        fo_trend = "⏳ Fill H,I,J"
    else:
        if pcr >= 1.5: pcr_signal = "Very Bullish"
        elif pcr >= 1.2: pcr_signal = "Bullish"
        elif pcr >= 0.8: pcr_signal = "Neutral"
        elif pcr >= 0.5: pcr_signal = "Bearish"
        else: pcr_signal = "Very Bearish"

        if oi_direction == "🟢 Long Buildup" and pcr_signal in ("Very Bullish", "Bullish"):
            fo_trend = "✅ STRONGLY BULLISH"
        elif oi_direction in ("🟢 Long Buildup", "🔵 Short Covering") and pcr_signal not in ("Bearish", "Very Bearish"):
            fo_trend = "✅ BULLISH"
        elif oi_direction in ("🔴 Short Buildup", "🟠 Long Unwinding") and pcr_signal in ("Bearish", "Very Bearish"):
            fo_trend = "❌ STRONGLY BEARISH"
        elif oi_direction in ("🔴 Short Buildup", "🟠 Long Unwinding"):
            fo_trend = "❌ BEARISH"
        else:
            fo_trend = "⚠️ NEUTRAL"

    if fo_trend == "⏳ Fill H,I,J":
        final_action = "⏳ Fill F&O data first"
    elif long_score >= 75 and fo_trend in ("✅ STRONGLY BULLISH", "✅ BULLISH"):
        final_action = "🟢 STRONG BUY"
    elif long_score >= 60 and fo_trend in ("✅ STRONGLY BULLISH", "✅ BULLISH"):
        final_action = "🔵 BUY"
    elif short_score >= 75 and fo_trend == "❌ STRONGLY BEARISH":
        final_action = "🔻 STRONG SHORT"
    elif short_score >= 55 and fo_trend in ("❌ STRONGLY BEARISH", "❌ BEARISH"):
        final_action = "🔽 SHORT"
    elif long_score >= 60 and fo_trend in ("❌ STRONGLY BEARISH", "❌ BEARISH"):
        final_action = "⚠️ AVOID — F&O Bearish"
    elif long_score >= 60 and fo_trend == "⚠️ NEUTRAL":
        final_action = "🟡 WAIT — Neutral F&O"
    else:
        final_action = "⬜ SKIP"

    if final_action in ("⏳ Fill F&O data first", "⬜ SKIP"):
        confidence = "—"
    elif final_action in ("🟢 STRONG BUY", "🔻 STRONG SHORT") and fo_trend in ("✅ STRONGLY BULLISH", "❌ STRONGLY BEARISH"):
        confidence = "⭐⭐⭐ HIGH"
    elif final_action in ("🟢 STRONG BUY", "🔻 STRONG SHORT"):
        confidence = "⭐⭐ MEDIUM-HIGH"
    elif final_action in ("🔵 BUY", "🔽 SHORT"):
        confidence = "⭐⭐ MEDIUM"
    else:
        confidence = "—"

    return {"oi_direction": oi_direction, "pcr_signal": pcr_signal, "fo_trend": fo_trend,
            "final_action": final_action, "confidence": confidence}


#   G=LongSignal H=TodayOI(input) I=PrevOI(input) J=PCR(input)
#   K=OIChange L=OIDirection M=PCRSignal N=F&OTrend O=FinalAction
#   P=Confidence Q=SLLong R=T1Long S=T2Long T=SLShort U=T1Short V=T2Short W=Qty
#   X=PlainEnglishReason
# ─────────────────────────────────────────────────────────────────────────────

def fo_formulas(row):
    r = str(row)

    # K: OI Change
    k = (f'=IF(OR(H{r}="",I{r}="",H{r}=0,I{r}=0),"",'
         f'H{r}-I{r})')

    # L: OI Direction
    l = (f'=IF(OR(H{r}="",I{r}="",H{r}=0,I{r}=0),"⏳ Enter OI",'
         f'IF(AND(H{r}>I{r},D{r}>0),"🟢 Long Buildup",'
         f'IF(AND(H{r}<I{r},D{r}>0),"🔵 Short Covering",'
         f'IF(AND(H{r}>I{r},D{r}<0),"🔴 Short Buildup",'
         f'"🟠 Long Unwinding"))))')

    # M: PCR Signal
    m = (f'=IF(OR(J{r}="",J{r}=0),"⏳ Enter PCR",'
         f'IF(J{r}>=1.5,"Very Bullish",'
         f'IF(J{r}>=1.2,"Bullish",'
         f'IF(J{r}>=0.8,"Neutral",'
         f'IF(J{r}>=0.5,"Bearish",'
         f'"Very Bearish")))))')

    # N: F&O Trend
    n = (f'=IF(OR(L{r}="⏳ Enter OI",M{r}="⏳ Enter PCR"),"⏳ Fill H,I,J",'
         f'IF(AND(L{r}="🟢 Long Buildup",OR(M{r}="Very Bullish",M{r}="Bullish")),"✅ STRONGLY BULLISH",'
         f'IF(AND(OR(L{r}="🟢 Long Buildup",L{r}="🔵 Short Covering"),M{r}<>"Bearish",M{r}<>"Very Bearish"),"✅ BULLISH",'
         f'IF(AND(OR(L{r}="🔴 Short Buildup",L{r}="🟠 Long Unwinding"),OR(M{r}="Bearish",M{r}="Very Bearish")),"❌ STRONGLY BEARISH",'
         f'IF(OR(L{r}="🔴 Short Buildup",L{r}="🟠 Long Unwinding"),"❌ BEARISH",'
         f'"⚠️ NEUTRAL")))))' )

    # O: Final Action
    o = (f'=IF(N{r}="⏳ Fill H,I,J","⏳ Fill F&O data first",'
         f'IF(AND(E{r}>=75,OR(N{r}="✅ STRONGLY BULLISH",N{r}="✅ BULLISH")),"🟢 STRONG BUY",'
         f'IF(AND(E{r}>=60,OR(N{r}="✅ STRONGLY BULLISH",N{r}="✅ BULLISH")),"🔵 BUY",'
         f'IF(AND(F{r}>=75,N{r}="❌ STRONGLY BEARISH"),"🔻 STRONG SHORT",'
         f'IF(AND(F{r}>=55,OR(N{r}="❌ STRONGLY BEARISH",N{r}="❌ BEARISH")),"🔽 SHORT",'
         f'IF(AND(E{r}>=60,OR(N{r}="❌ STRONGLY BEARISH",N{r}="❌ BEARISH")),"⚠️ AVOID — F&O Bearish",'
         f'IF(AND(E{r}>=60,N{r}="⚠️ NEUTRAL"),"🟡 WAIT — Neutral F&O",'
         f'"⬜ SKIP")))))))' )

    # P: Confidence
    p = (f'=IF(OR(O{r}="⏳ Fill F&O data first",O{r}="⬜ SKIP"),"—",'
         f'IF(AND(OR(O{r}="🟢 STRONG BUY",O{r}="🔻 STRONG SHORT"),N{r}="✅ STRONGLY BULLISH"),"⭐⭐⭐ HIGH",'
         f'IF(AND(OR(O{r}="🟢 STRONG BUY",O{r}="🔻 STRONG SHORT"),N{r}="❌ STRONGLY BEARISH"),"⭐⭐⭐ HIGH",'
         f'IF(OR(O{r}="🟢 STRONG BUY",O{r}="🔻 STRONG SHORT"),"⭐⭐ MEDIUM-HIGH",'
         f'IF(OR(O{r}="🔵 BUY",O{r}="🔽 SHORT"),"⭐⭐ MEDIUM","—")))))')

    # X: Plain English Reason (concatenation formula)
    x = (f'=IF(O{r}="⏳ Fill F&O data first",'
         f'"👆 Enter Today OI in col H, Prev OI in col I, PCR in col J from Sensibull",'
         f'IF(O{r}="🟢 STRONG BUY",'
         f'"Score "&E{r}&"/100. "&L{r}&". PCR "&J{r}&" → "&M{r}&". F&O: "&N{r}&". Strong institutional buying confirmed. Enter tomorrow at open. SL ₹"&Q{r}&" | T1 ₹"&R{r},'
         f'IF(O{r}="🔵 BUY",'
         f'"Score "&E{r}&"/100. "&L{r}&". PCR "&J{r}&" → "&M{r}&". Enter on dip. SL ₹"&Q{r}&" | T1 ₹"&R{r},'
         f'IF(O{r}="🔻 STRONG SHORT",'
         f'"Short score "&F{r}&"/100. "&L{r}&". PCR "&J{r}&" → "&M{r}&". F&O: "&N{r}&". Institutions shorting. Sell futures. SL ₹"&T{r}&" | T1 ₹"&U{r},'
         f'IF(O{r}="🔽 SHORT",'
         f'"Short score "&F{r}&"/100. "&L{r}&". PCR "&J{r}&" → "&M{r}&". Sell on bounce. SL ₹"&T{r}&" | T1 ₹"&U{r},'
         f'IF(O{r}="⚠️ AVOID — F&O Bearish",'
         f'"Score "&E{r}&"/100 looks good but F&O is bearish ("&N{r}&"). Do NOT buy — wait for F&O to confirm.",'
         f'IF(O{r}="🟡 WAIT — Neutral F&O",'
         f'"Score "&E{r}&"/100 — setup forming. F&O neutral. Check again tomorrow.",'
         f'"Score "&E{r}&"/100, Short "&F{r}&"/100 — no clear setup. Skip today.")))))))' )

    return k, l, m, n, o, p, x


# ─────────────────────────────────────────────────────────────────────────────
#  SHEET: 📋 F&O INPUT  (main working sheet)
# ─────────────────────────────────────────────────────────────────────────────

def sheet_fo_data_downloaded(wb, fo_data, today_str):
    """
    📥 FO DATA sheet — contains the bhavcopy data.
    F&O INPUT sheet VLOOKUPs from column A (Symbol) of this sheet.
    Layout: A=Symbol, B=Today Fut OI, C=Prev Day OI, D=PCR
    """
    ws = wb.create_sheet("📥 FO DATA")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "B3"

    MH(ws,"A1:E1",
       f"📥  NSE F&O BHAVCOPY DATA — Auto-downloaded  |  {today_str}",
       bg="navy", sz=12)
    MH(ws,"A2:E2",
       "This data is downloaded automatically from NSE archives. "
       "F&O INPUT sheet VLOOKUPs from here. Do NOT edit this sheet.",
       bg="dblue", sz=9)

    hdrs = [("A","SYMBOL",14),("B","TODAY\nFUT OI",16),
            ("C","PREV DAY\nFUT OI",16),("D","PCR",10),("E","SOURCE",18)]
    for col,txt,w in hdrs:
        W(ws,f"{col}3",txt,bg="navy",fg="white",bold=True,sz=10,wrap=True)
        ws.column_dimensions[col].width=w
    ws.row_dimensions[3].height=32

    if not fo_data:
        ws.merge_cells("A4:E4")
        c = ws["A4"]
        c.value = ("⚠️  F&O Bhavcopy could not be downloaded. "
                   "Manually type Today OI, Prev OI, PCR in yellow cells on 📋 F&O INPUT sheet.")
        c.font = Font(bold=True, size=10, color=P["sdark"], name="Calibri")
        c.fill = F("lred"); c.alignment = Al("left","center",True)
        c.border = Bd(); ws.row_dimensions[4].height = 22
        return ws

    for i,(sym,d) in enumerate(sorted(fo_data.items())):
        row = i + 4
        bg  = P["grey"] if i%2==0 else P["white"]
        vals = [
            (sym,               "navy",  "@"),
            (d["today_oi"],     "black", "#,##0"),
            (d["prev_oi"],      "black", "#,##0"),
            (d["pcr"] if d["pcr"] else "N/A", "black", "0.00" if d["pcr"] else "@"),
            ("✅ NSE Bhavcopy",  "green", "@"),
        ]
        for j,(v,fg,fmt) in enumerate(vals):
            col = get_column_letter(j+1)
            c = ws[f"{col}{row}"]
            c.value = v
            c.font = Font(bold=(j==0), size=10, color=P.get(fg,fg), name="Calibri")
            c.fill = PatternFill("solid", fgColor=bg)
            c.alignment = Al("center","center")
            c.border = Bd()
            if fmt and v != "N/A": c.number_format = fmt

    return ws


def sheet_fo_input(wb, results, today_str, fo_data=None):
    """
    Main working sheet.
    If fo_data dict is passed (from bhavcopy download), H/I/J cells are
    pre-filled green (auto). If a stock is missing from fo_data, cells
    stay yellow for manual Sensibull entry.
    K→X remain Excel formulas regardless.
    """
    fo_data = fo_data or {}

    ws = wb.create_sheet("📋 F&O INPUT")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "H6"

    # ── Title rows ────────────────────────────────────────────────────────
    auto_count = sum(1 for r in results if r["Symbol"] in fo_data)
    if auto_count > 0:
        MH(ws,"A1:X1",
           f"📋  F&O INPUT  —  {today_str}  —  "
           f"✅ {auto_count} stocks auto-filled from NSE Bhavcopy  "
           f"(yellow = fill manually from Sensibull)",
           bg="navy", sz=12)
        MH(ws,"A2:X2",
           "🟢 Green cells = auto-filled from NSE Bhavcopy (downloaded automatically)  |  "
           "🟡 Yellow cells = stock not in bhavcopy — type from Sensibull  |  K→X recalculate instantly",
           bg="dblue", sz=10)
    else:
        MH(ws,"A1:X1",
           f"📋  F&O INPUT SHEET  —  {today_str}  —  Fill yellow columns H, I, J from Sensibull",
           bg="navy", sz=13)
        MH(ws,"A2:X2",
           f"Capital ₹{CONFIG['total_capital']:,.0f}  |  Max/trade ₹{CONFIG['max_per_trade']:,.0f}  "
           f"|  Long SL {CONFIG['stop_loss_pct']}%  |  Short SL {CONFIG['short_sl_pct']}%  "
           f"|  Target +{CONFIG['target1_pct']}% / -{CONFIG['short_target1_pct']}%",
           bg="dblue", sz=10)

    # ── How-to box ────────────────────────────────────────────────────────
    MH(ws,"A3:G3", "📊 AUTO-FILLED BY PYTHON (do not edit)", bg="green", fg="white", sz=9)
    if auto_count > 0:
        MH(ws,"H3:J3", "🟢 AUTO / 🟡 FILL FROM SENSIBULL", bg="lgreen", fg="green", sz=9, bold=True)
    else:
        MH(ws,"H3:J3", "🟡 YOU FILL FROM SENSIBULL", bg="yelb", fg="black", sz=9, bold=True)
    MH(ws,"K3:N3", "⚙️ AUTO-CALCULATES INSTANTLY", bg="blue", fg="white", sz=9)
    MH(ws,"O3:P3", "🎯 FINAL CALL", bg="sdark", fg="white", sz=9)
    MH(ws,"Q3:W3", "📐 TRADE LEVELS", bg="dblue", fg="white", sz=9)
    MH(ws,"X3:X3", "💬 PLAIN ENGLISH REASON", bg="navy", fg="white", sz=9)
    ws.row_dimensions[3].height = 18

    # ── Column headers ────────────────────────────────────────────────────
    HEADERS = [
        ("A","SYMBOL",13,"navy"),
        ("B","PREV\nCLOSE ₹",10,"navy"),
        ("C","CMP ₹",10,"navy"),
        ("D","DAY\nCHG %",8,"navy"),
        ("E","LONG\nSCORE",9,"navy"),
        ("F","SHORT\nSCORE",9,"navy"),
        ("G","LONG\nSIGNAL",16,"navy"),
        ("H","TODAY\nFUT OI\n🟡 Fill",14,"yelb"),
        ("I","PREV DAY\nFUT OI\n🟡 Fill",14,"yelb"),
        ("J","PCR\n🟡 Fill",9,"yelb"),
        ("K","OI\nCHANGE",11,"blue"),
        ("L","OI\nDIRECTION",18,"blue"),
        ("M","PCR\nSIGNAL",13,"blue"),
        ("N","F&O\nTREND",20,"blue"),
        ("O","FINAL\nACTION",20,"sdark"),
        ("P","CONF-\nIDENCE",11,"sdark"),
        ("Q","SL\nLONG ₹",10,"dblue"),
        ("R","T1\nLONG ₹",10,"dblue"),
        ("S","T2\nLONG ₹",10,"dblue"),
        ("T","SL\nSHORT ₹",10,"dblue"),
        ("U","T1\nSHORT ₹",10,"dblue"),
        ("V","T2\nSHORT ₹",10,"dblue"),
        ("W","QTY",7,"dblue"),
        ("X","REASON — updates automatically as you fill H, I, J",55,"navy"),
    ]
    for col,txt,width,bg in HEADERS:
        fg = "black" if bg=="yelb" else "white"
        c = ws[f"{col}4"]
        c.value = txt
        c.font = Font(bold=True, size=9, color=P.get(fg,fg), name="Calibri")
        c.fill = F(bg)
        c.alignment = Al("center","center",True)
        c.border = Bd("888888" if col in ("H","I","J") else "BBBBBB")
        ws.column_dimensions[col].width = width
    ws.row_dimensions[4].height = 42

    # ── Instruction row ────────────────────────────────────────────────────
    MH(ws,"A5:G5","← Auto-filled",bg="lgreen",fg="green",sz=8)
    MH(ws,"H5:J5","← Type numbers here (no commas needed)",bg="FFE066",fg="black",sz=8,bold=True)
    MH(ws,"K5:N5","← Calculates as you type →",bg="lblue",fg="blue",sz=8)
    MH(ws,"O5:P5","← Decision",bg="slight",fg="sdark",sz=8)
    MH(ws,"Q5:W5","← Trade levels",bg="lblue",fg="navy",sz=8)
    MH(ws,"X5:X5","← Plain English explanation",bg="lgreen",fg="green",sz=8)
    ws.row_dimensions[5].height = 14

    # ── Data rows ─────────────────────────────────────────────────────────
    for i, r in enumerate(results):
        row = i + 6
        is_buy   = r["LONG SCORE"]  >= CONFIG["min_score_buy"]
        is_short = r["SHORT SCORE"] >= CONFIG["min_score_short"]

        if r["LONG SCORE"] >= CONFIG["priority_buy_score"]:
            row_bg = "E8F8E8"
        elif is_buy:
            row_bg = P["input"]
        elif is_short:
            row_bg = P["slight"]
        else:
            row_bg = P["grey"] if i%2==0 else P["white"]

        # Static Python-filled columns
        static = [
            ("A", r["Symbol"],          "navy",  "@",       True),
            ("B", r["Prev Close"],       "black", "#,##0.00",False),
            ("C", r["CMP"],             "black", "#,##0.00",False),
            ("D", r["Day Chg %"]/100,   "green" if r["Day Chg %"]>=0 else "red","0.00%",False),
            ("E", r["LONG SCORE"],      "navy",  "0",       True),
            ("F", r["SHORT SCORE"],     "sdark", "0",       True),
            ("G", r["LONG SIGNAL"],     "navy",  "@",       False),
        ]
        for col,val,fg,fmt,bold in static:
            c=ws[f"{col}{row}"]
            c.value=val
            c.font=Font(bold=bold,size=10,color=P.get(fg,fg),name="Calibri")
            c.fill=PatternFill("solid",fgColor=row_bg)
            c.alignment=Al("center","center"); c.border=Bd()
            if fmt: c.number_format=fmt

        # H, I, J — pre-fill from bhavcopy if available, else yellow for manual
        sym = r["Symbol"]
        bh  = fo_data.get(sym)

        oi_vals = {
            "H": bh["today_oi"] if bh and bh["today_oi"] else None,
            "I": bh["prev_oi"]  if bh and bh.get("prev_oi") else None,
            "J": bh["pcr"]      if bh and bh.get("pcr")     else None,
        }
        for col in ("H","I","J"):
            c  = ws[f"{col}{row}"]
            v  = oi_vals[col]
            if v is not None:
                # Auto-filled — light green
                c.value = v
                c.font  = Font(bold=True, size=10, color="1E5631", name="Calibri")
                c.fill  = PatternFill("solid", fgColor="D5F5E3")
                c.alignment = Al("center","center")
                c.border = Border(
                    left=Side(style="medium",color="27AE60"),
                    right=Side(style="medium",color="27AE60"),
                    top=Side(style="thin",color="27AE60"),
                    bottom=Side(style="thin",color="27AE60"))
            else:
                # Manual — yellow
                c.value = ""
                c.font  = Font(bold=True, size=11, color="000000", name="Calibri")
                c.fill  = PatternFill("solid", fgColor="FFFACD")
                c.alignment = Al("center","center")
                c.border = Border(
                    left=Side(style="medium",color="FFD700"),
                    right=Side(style="medium",color="FFD700"),
                    top=Side(style="thin",color="FFD700"),
                    bottom=Side(style="thin",color="FFD700"))
            if col in ("H","I"): c.number_format="#,##0"
            if col=="J":         c.number_format="0.00"

        # Formula columns K-N, O, P, X
        k,l,m,n,o,p,x = fo_formulas(row)
        for col,formula,fmt in [
            ("K",k,"#,##0;[Red]-#,##0"),
            ("L",l,"@"),("M",m,"@"),("N",n,"@"),
            ("O",o,"@"),("P",p,"@"),
        ]:
            c=ws[f"{col}{row}"]
            c.value=formula
            c.font=Font(bold=(col in ("O","P")),size=10,name="Calibri",
                        color=P["sdark"] if col in ("O","P") else "000000")
            c.fill=PatternFill("solid",fgColor=row_bg)
            c.alignment=Al("center","center",wrap=(col in ("L","N","O")))
            c.border=Bd()
            if fmt: c.number_format=fmt

        # Trade levels Q-W (static from Python)
        trade = [
            ("Q", r["SL Long"],  "red",  "#,##0.00"),
            ("R", r["T1 Long"],  "green","#,##0.00"),
            ("S", r["T2 Long"],  "green","#,##0.00"),
            ("T", r["SL Short"], "red",  "#,##0.00"),
            ("U", r["T1 Short"],"green","#,##0.00"),
            ("V", r["T2 Short"],"green","#,##0.00"),
            ("W", r["Qty"],     "navy", "0"),
        ]
        for col,val,fg,fmt in trade:
            c=ws[f"{col}{row}"]
            c.value=val
            c.font=Font(bold=False,size=9,color=P.get(fg,fg),name="Calibri")
            c.fill=PatternFill("solid",fgColor=row_bg)
            c.alignment=Al("center","center"); c.border=Bd()
            if fmt: c.number_format=fmt

        # Reason column X (formula)
        c=ws[f"X{row}"]
        c.value=x
        c.font=Font(size=9,name="Calibri",color="000000")
        c.fill=PatternFill("solid",fgColor=row_bg)
        c.alignment=Al("left","center",wrap=True); c.border=Bd()
        ws.row_dimensions[row].height = 28

    # ── Conditional formatting — highlight final action cells ─────────────
    last_row = 5 + len(results)
    # Strong Buy → green
    ws.conditional_formatting.add(
        f"O6:O{last_row}",
        FormulaRule(formula=[f'O6="🟢 STRONG BUY"'],
                    fill=PatternFill("solid",fgColor="C6EFCE"),
                    font=Font(bold=True,color="1E8449",name="Calibri")))
    # Buy → light blue
    ws.conditional_formatting.add(
        f"O6:O{last_row}",
        FormulaRule(formula=[f'O6="🔵 BUY"'],
                    fill=PatternFill("solid",fgColor="DDEEFF"),
                    font=Font(bold=True,color="1F3864",name="Calibri")))
    # Strong Short → red
    ws.conditional_formatting.add(
        f"O6:O{last_row}",
        FormulaRule(formula=[f'O6="🔻 STRONG SHORT"'],
                    fill=PatternFill("solid",fgColor="FFC7CE"),
                    font=Font(bold=True,color="7B241C",name="Calibri")))
    # Short → light red
    ws.conditional_formatting.add(
        f"O6:O{last_row}",
        FormulaRule(formula=[f'O6="🔽 SHORT"'],
                    fill=PatternFill("solid",fgColor="FFE4E4"),
                    font=Font(bold=True,color="C0392B",name="Calibri")))
    # Avoid → orange
    ws.conditional_formatting.add(
        f"O6:O{last_row}",
        FormulaRule(formula=[f'LEFT(O6,2)="⚠️"'],
                    fill=PatternFill("solid",fgColor="FFE5CC"),
                    font=Font(bold=True,color="E67E22",name="Calibri")))

    # Highlight entire row green when Strong Buy
    for col_letter in "ABCDEFGKLMNOPQRSTUVWX":
        ws.conditional_formatting.add(
            f"{col_letter}6:{col_letter}{last_row}",
            FormulaRule(formula=[f'$O6="🟢 STRONG BUY"'],
                        fill=PatternFill("solid",fgColor="EAF8EA")))
        ws.conditional_formatting.add(
            f"{col_letter}6:{col_letter}{last_row}",
            FormulaRule(formula=[f'$O6="🔻 STRONG SHORT"'],
                        fill=PatternFill("solid",fgColor="FFEAEA")))

    return ws, last_row


# ─────────────────────────────────────────────────────────────────────────────
#  SHEET: ⭐ FINAL CALLS  (references F&O INPUT formulas)
# ─────────────────────────────────────────────────────────────────────────────

def sheet_final_calls(wb, results, today_str):
    ws = wb.create_sheet("⭐ FINAL CALLS")
    ws.sheet_view.showGridLines = False

    MH(ws,"A1:K1",
       f"⭐  FINAL CALLS  —  {today_str}  —  Updates live as you fill F&O INPUT sheet",
       bg="navy", sz=13)
    MH(ws,"A2:K2",
       "All values pull directly from 📋 F&O INPUT sheet. Fill columns H, I, J there first.",
       bg="dblue", sz=10)

    COL_W = {"A":5,"B":13,"C":10,"D":9,"E":9,"F":20,"G":11,"H":20,"I":9,"J":9,"K":50}
    CW(ws, COL_W)

    hdrs = ["#","SYMBOL","LONG\nSCORE","SHORT\nSCORE","CMP ₹",
            "FINAL ACTION","CONFIDENCE","F&O TREND",
            "SL ₹","TARGET 1 ₹","REASON"]
    for i,h in enumerate(hdrs):
        col=get_column_letter(i+1)
        c=ws[f"{col}3"]
        c.value=h
        c.font=Font(bold=True,size=9,color="FFFFFF",name="Calibri")
        c.fill=F("navy"); c.alignment=Al("center","center",True)
        c.border=Bd(); ws.row_dimensions[3].height=38

    # Pull top 40 stocks from F&O INPUT (rows 6 to 6+39)
    # Show all of them — conditional formatting makes buy/short obvious
    src = "'📋 F&O INPUT'"
    for i in range(min(len(results), 40)):
        src_row = i + 6   # F&O INPUT data starts at row 6
        ws_row  = i + 4   # Final Calls data starts at row 4

        # Determine if it's a long or short candidate to pick right SL/T1
        is_short_candidate = (results[i]["SHORT SCORE"] > results[i]["LONG SCORE"]
                               and results[i]["SHORT SCORE"] >= CONFIG["min_score_short"])

        sl_col  = "T" if is_short_candidate else "Q"
        t1_col  = "U" if is_short_candidate else "R"

        row_formulas = [
            (i+1, None),                                          # rank static
            (f"={src}!A{src_row}", "@"),
            (f"={src}!E{src_row}", "0"),
            (f"={src}!F{src_row}", "0"),
            (f"={src}!C{src_row}", "#,##0.00"),
            (f"={src}!O{src_row}", "@"),
            (f"={src}!P{src_row}", "@"),
            (f"={src}!N{src_row}", "@"),
            (f"={src}!{sl_col}{src_row}", "#,##0.00"),
            (f"={src}!{t1_col}{src_row}", "#,##0.00"),
            (f"={src}!X{src_row}", "@"),
        ]

        bg = P["grey"] if i%2==0 else P["white"]
        for j,(val,fmt) in enumerate(row_formulas):
            col=get_column_letter(j+1)
            c=ws[f"{col}{ws_row}"]
            c.value=val
            c.font=Font(size=9 if j==10 else 10,name="Calibri",
                        bold=(j in {1,2,3,5}))
            c.fill=PatternFill("solid",fgColor=bg)
            c.alignment=Al("left" if j==10 else "center","center",wrap=(j==10))
            c.border=Bd()
            if fmt and isinstance(val,str) and val.startswith("="): c.number_format=fmt
        ws.row_dimensions[ws_row].height = 28

    # Conditional formatting on Final Action column (F)
    last = 3 + min(len(results),40)
    ws.conditional_formatting.add(f"F4:F{last}",
        FormulaRule(formula=['$F4="🟢 STRONG BUY"'],
                    fill=PatternFill("solid",fgColor="C6EFCE"),
                    font=Font(bold=True,color="1E8449",name="Calibri")))
    ws.conditional_formatting.add(f"F4:F{last}",
        FormulaRule(formula=['$F4="🔻 STRONG SHORT"'],
                    fill=PatternFill("solid",fgColor="FFC7CE"),
                    font=Font(bold=True,color="7B241C",name="Calibri")))
    ws.conditional_formatting.add(f"F4:F{last}",
        FormulaRule(formula=['$F4="🔵 BUY"'],
                    fill=PatternFill("solid",fgColor="DDEEFF"),
                    font=Font(bold=True,color="163A6B",name="Calibri")))
    ws.conditional_formatting.add(f"F4:F{last}",
        FormulaRule(formula=['$F4="🔽 SHORT"'],
                    fill=PatternFill("solid",fgColor="FFE4E4"),
                    font=Font(bold=True,color="C0392B",name="Calibri")))

    # Whole-row highlight
    for col_letter in "ABCDEGHIJK":
        ws.conditional_formatting.add(f"{col_letter}4:{col_letter}{last}",
            FormulaRule(formula=['$F4="🟢 STRONG BUY"'],
                        fill=PatternFill("solid",fgColor="EAF8EA")))
        ws.conditional_formatting.add(f"{col_letter}4:{col_letter}{last}",
            FormulaRule(formula=['$F4="🔻 STRONG SHORT"'],
                        fill=PatternFill("solid",fgColor="FFEAEA")))


# ─────────────────────────────────────────────────────────────────────────────
#  SHEET: 📈 BUY WATCHLIST
# ─────────────────────────────────────────────────────────────────────────────

def sheet_buy_watchlist(wb, results, today_str):
    ws = wb.create_sheet("📈 BUY WATCHLIST")
    ws.sheet_view.showGridLines = False; ws.freeze_panes = "B5"
    MH(ws,"A1:S1",f"BUY WATCHLIST — Long Candidates Scored ≥{CONFIG['min_score_buy']}  |  {today_str}",bg="navy",sz=12)
    MH(ws,"A2:S2","Technical breakdown. Enter F&O data on 📋 F&O INPUT sheet for Final Action.",bg="dblue",sz=10)
    hdrs=[("A","#",4),("B","SYMBOL",13),("C","PREV\nCLOSE",10),("D","CMP",10),
          ("E","CHG %",8),("F","LONG\nSCORE",9),("G","SIGNAL",16),
          ("H","EMA\nSc",8),("I","RSI\nSc",8),("J","VOL\nSc",8),
          ("K","MACD\nSc",8),("L","PAT\nSc",8),
          ("M","EMA TREND",16),("N","RSI",8),("O","VOL\nRATIO",9),("P","PATTERN",18),
          ("Q","SL ₹",10),("R","T1 ₹",10),("S","QTY",7)]
    for col,txt,w in hdrs:
        W(ws,f"{col}3",txt,bg="navy",fg="white",bold=True,sz=9,wrap=True)
        W(ws,f"{col}4","",bg="lblue",fg="dgrey",sz=8)
        ws.column_dimensions[col].width=w
    RH(ws,{3:38,4:8})
    buys=sorted([r for r in results if r["LONG SCORE"]>=CONFIG["min_score_buy"]],
                key=lambda x:x["LONG SCORE"],reverse=True)
    for i,r in enumerate(buys[:30]):
        row=i+5; bg="E8F8E8" if r["LONG SCORE"]>=75 else (P["input"] if r["LONG SCORE"]>=60 else P["grey"])
        data=[(i+1,"dgrey","0"),(r["Symbol"],"navy","@"),
              (r["Prev Close"],"black","#,##0.00"),(r["CMP"],"black","#,##0.00"),
              (r["Day Chg %"]/100,"green" if r["Day Chg %"]>=0 else "red","0.00%"),
              (r["LONG SCORE"],"navy","0"),(r["LONG SIGNAL"],"navy","@"),
              (r["EMA Sc(L)"],"black","0"),(r["RSI Sc(L)"],"black","0"),
              (r["Vol Sc(L)"],"black","0"),(r["MACD Sc(L)"],"black","0"),
              (r["Pat Sc(L)"],"black","0"),
              (r["EMA Trend"],"black","@"),(r["RSI"],"black","0.0"),
              (r["Vol Ratio"],"black","0.00"),
              (r["Pattern Long"],"black","@"),
              (r["SL Long"],"red","#,##0.00"),(r["T1 Long"],"green","#,##0.00"),
              (r["Qty"],"black","0")]
        for j,(v,fg,fmt) in enumerate(data):
            col=get_column_letter(j+1); c=ws[f"{col}{row}"]
            c.value=v; c.font=Font(bold=(j in {1,5}),size=10,color=P.get(fg,fg),name="Calibri")
            c.fill=PatternFill("solid",fgColor=bg)
            c.alignment=Al("center","center"); c.border=Bd()
            if fmt: c.number_format=fmt


# ─────────────────────────────────────────────────────────────────────────────
#  SHEET: 🔻 SHORT WATCHLIST
# ─────────────────────────────────────────────────────────────────────────────

def sheet_short_watchlist(wb, results, today_str):
    ws = wb.create_sheet("🔻 SHORT WATCHLIST")
    ws.sheet_view.showGridLines = False; ws.freeze_panes = "B5"
    MH(ws,"A1:S1",f"SHORT WATCHLIST — Short Candidates  |  {today_str}",bg="sdark",sz=12)
    MH(ws,"A2:S2","Confirm Short Buildup + PCR < 0.8 from F&O INPUT before shorting",bg="smed",fg="white",sz=10)
    hdrs=[("A","#",4),("B","SYMBOL",13),("C","PREV\nCLOSE",10),("D","CMP",10),
          ("E","CHG %",8),("F","SHORT\nSCORE",9),("G","SIGNAL",16),
          ("H","EMA\nSc",8),("I","RSI\nSc",8),("J","VOL\nSc",8),
          ("K","MACD\nSc",8),("L","PAT\nSc",8),
          ("M","EMA STATUS",16),("N","RSI",8),("O","VOL\nRATIO",9),("P","PATTERN",18),
          ("Q","SL ₹\n(4% above)",12),("R","T1 ₹\n-8%",10),("S","QTY",7)]
    for col,txt,w in hdrs:
        W(ws,f"{col}3",txt,bg="sdark",fg="white",bold=True,sz=9,wrap=True)
        W(ws,f"{col}4","",bg="lred",fg="dgrey",sz=8)
        ws.column_dimensions[col].width=w
    RH(ws,{3:38,4:8})
    shorts=sorted([r for r in results if r["SHORT SCORE"]>=40],
                  key=lambda x:x["SHORT SCORE"],reverse=True)
    for i,r in enumerate(shorts[:30]):
        row=i+5; bg="FFE4E4" if r["SHORT SCORE"]>=75 else (P["lred"] if r["SHORT SCORE"]>=55 else P["grey"])
        data=[(i+1,"dgrey","0"),(r["Symbol"],"sdark","@"),
              (r["Prev Close"],"black","#,##0.00"),(r["CMP"],"black","#,##0.00"),
              (r["Day Chg %"]/100,"red" if r["Day Chg %"]<=0 else "green","0.00%"),
              (r["SHORT SCORE"],"sdark","0"),(r["SHORT SIGNAL"],"sdark","@"),
              (r["EMA Sc(S)"],"black","0"),(r["RSI Sc(S)"],"black","0"),
              (r["Vol Sc(S)"],"black","0"),(r["MACD Sc(S)"],"black","0"),
              (r["Pat Sc(S)"],"black","0"),
              (r["EMA Trend"],"black","@"),(r["RSI"],"black","0.0"),
              (r["Vol Ratio"],"black","0.00"),
              (r["Pattern Short"],"black","@"),
              (r["SL Short"],"red","#,##0.00"),(r["T1 Short"],"green","#,##0.00"),
              (r["Qty"],"black","0")]
        for j,(v,fg,fmt) in enumerate(data):
            col=get_column_letter(j+1); c=ws[f"{col}{row}"]
            c.value=v; c.font=Font(bold=(j in {1,5}),size=10,color=P.get(fg,fg),name="Calibri")
            c.fill=PatternFill("solid",fgColor=bg)
            c.alignment=Al("center","center"); c.border=Bd()
            if fmt: c.number_format=fmt


# ─────────────────────────────────────────────────────────────────────────────
#  SHEET: 📊 ALL STOCKS
# ─────────────────────────────────────────────────────────────────────────────

def sheet_all_stocks(wb, results, today_str):
    ws = wb.create_sheet("📊 ALL STOCKS")
    ws.sheet_view.showGridLines = False; ws.freeze_panes = "B4"
    MH(ws,"A1:P1",f"ALL STOCKS — Full Technical Analysis  |  {today_str}",bg="navy",sz=11)
    hdrs=[("A","Symbol",13),("B","Prev\nClose",10),("C","CMP",10),("D","Chg %",8),
          ("E","Long\nScore",9),("F","Long Signal",16),("G","Short\nScore",9),
          ("H","Short Signal",16),("I","RSI",7),("J","Vol\nRatio",8),
          ("K","20 EMA",9),("L","50 EMA",9),("M","200 EMA",9),
          ("N","Long\nPattern",16),("O","Short\nPattern",16),("P","52W Hi%",9)]
    for col,txt,w in hdrs:
        W(ws,f"{col}2",txt,bg="navy",fg="white",bold=True,sz=9,wrap=True)
        W(ws,f"{col}3","",bg="lblue",sz=8)
        ws.column_dimensions[col].width=w
    ws.row_dimensions[2].height=35
    for i,r in enumerate(sorted(results,key=lambda x:x["LONG SCORE"],reverse=True)):
        row=i+4; bg=P["grey"] if i%2==0 else P["white"]
        vals=[(r["Symbol"],"navy","@"),(r["Prev Close"],"black","#,##0.00"),
              (r["CMP"],"black","#,##0.00"),(r["Day Chg %"]/100,"black","0.00%"),
              (r["LONG SCORE"],"navy","0"),(r["LONG SIGNAL"],"navy","@"),
              (r["SHORT SCORE"],"sdark","0"),(r["SHORT SIGNAL"],"sdark","@"),
              (r["RSI"],"black","0.0"),(r["Vol Ratio"],"black","0.00"),
              (r["20 EMA"],"black","#,##0.00"),(r["50 EMA"],"black","#,##0.00"),
              (r["200 EMA"],"black","#,##0.00"),
              (r["Pattern Long"],"black","@"),(r["Pattern Short"],"black","@"),
              (r["% from 52W High"],"black","0.0")]
        for j,(v,fg,fmt) in enumerate(vals):
            col=get_column_letter(j+1); c=ws[f"{col}{row}"]
            c.value=v; c.font=Font(bold=(j==0),size=9,color=P.get(fg,fg),name="Calibri")
            c.fill=PatternFill("solid",fgColor=bg)
            c.alignment=Al("center","center"); c.border=Bd()
            if fmt: c.number_format=fmt


# ─────────────────────────────────────────────────────────────────────────────
#  SHEET: 🧮 FUND + TECH VIEW  /  🎯 TIER A SWING  /  💰 TIER B INVESTMENT  /  🚫 EXCLUDED
#  Swing/Investment layer -- same style conventions as sheet_all_stocks above.
# ─────────────────────────────────────────────────────────────────────────────

def _gate_tag(passed):
    return "🟢 PASS" if passed else "🔴 FAIL"

def sheet_fund_tech_view(wb, results, today_str):
    ws = wb.create_sheet("🧮 FUND + TECH VIEW")
    ws.sheet_view.showGridLines = False; ws.freeze_panes = "B4"
    MH(ws,"A1:S1",f"FUNDAMENTALS + TECHNICALS — Full Watchlist  |  {today_str}",bg="navy",sz=11)
    hdrs=[("A","Symbol",13),("B","CMP",10),("C","Chg %",8),
          ("D","Long\nScore",9),("E","Long Signal",16),("F","Tech\nScore",9),
          ("G","Fund\nGate",10),("H","Fund\nScore",9),("I","ROE %",8),
          ("J","ROCE %",8),("K","D/E",7),("L","P/E",8),("M","Pledge %",9),
          ("N","Int.\nCoverage",10),("O","Debtor\nDays",9),("P","OPM %",8),
          ("Q","OPM\n5Y Avg",9),("R","CFO OK?",9),("S","Div Payout %",11)]
    for col,txt,w in hdrs:
        W(ws,f"{col}2",txt,bg="navy",fg="white",bold=True,sz=9,wrap=True)
        W(ws,f"{col}3","",bg="lblue",sz=8)
        ws.column_dimensions[col].width=w
    ws.row_dimensions[2].height=35
    for i,r in enumerate(sorted(results,key=lambda x:x.get("tech_score",x["LONG SCORE"]),reverse=True)):
        row=i+4; bg=P["grey"] if i%2==0 else P["white"]
        gate = r.get("gate_pass")
        cfo1, cfo2 = r.get("cfo_last_year"), r.get("cfo_preceding_year")
        cfo_ok = "🟢 Yes" if (cfo1 is not None and cfo2 is not None and cfo1 >= 0 and cfo2 >= 0) else (
                 "🔴 No" if (cfo1 is not None or cfo2 is not None) else "—")
        vals=[(r["Symbol"],"navy","@"),(r["CMP"],"black","#,##0.00"),
              (r["Day Chg %"]/100,"black","0.00%"),
              (r["LONG SCORE"],"navy","0"),(r["LONG SIGNAL"],"navy","@"),
              (r.get("tech_score",r["LONG SCORE"]),"navy","0"),
              (_gate_tag(gate) if gate is not None else "⚪ NO DATA","black","@"),
              (r.get("fund_score"),"black","0.0"),
              (r.get("roe"),"black","0.0"),(r.get("roce"),"black","0.0"),
              (r.get("debt_to_equity"),"black","0.00"),(r.get("pe_ttm"),"black","0.0"),
              (r.get("promoter_pledge_pct"),"black","0.0"),
              (r.get("interest_coverage"),"black","0.0"),(r.get("debtor_days"),"black","0"),
              (r.get("opm"),"black","0.0"),(r.get("opm_5y"),"black","0.0"),
              (cfo_ok,"black","@"),(r.get("dividend_payout_pct"),"black","0.0")]
        for j,(v,fg,fmt) in enumerate(vals):
            col=get_column_letter(j+1); c=ws[f"{col}{row}"]
            c.value=v; c.font=Font(bold=(j==0),size=9,color=P.get(fg,fg),name="Calibri")
            c.fill=PatternFill("solid",fgColor=bg)
            c.alignment=Al("center","center"); c.border=Bd()
            if fmt: c.number_format=fmt

def sheet_tier_a_swing(wb, tier_a, today_str):
    ws = wb.create_sheet("🎯 TIER A SWING")
    ws.sheet_view.showGridLines = False; ws.freeze_panes = "B3"
    MH(ws,"A1:H1",f"TIER A — SWING CANDIDATES (fundamentally-sound, ranked by technical score)  |  {today_str}",bg="green",sz=11)
    hdrs=[("A","Rank",7),("B","Symbol",13),("C","CMP",10),("D","Tech\nScore",9),
          ("E","Long Signal",16),("F","RSI",7),("G","F&O Long\nBuildup",11),("H","Pattern",18)]
    for col,txt,w in hdrs:
        W(ws,f"{col}2",txt,bg="green",fg="white",bold=True,sz=9,wrap=True)
        ws.column_dimensions[col].width=w
    ws.row_dimensions[2].height=30
    for i,r in enumerate(tier_a):
        row=i+3; bg=P["lgreen"] if i%2==0 else P["white"]
        vals=[(i+1,"black","0"),(r["Symbol"],"navy","@"),(r["CMP"],"black","#,##0.00"),
              (r.get("tech_score",r["LONG SCORE"]),"navy","0"),(r["LONG SIGNAL"],"navy","@"),
              (r["RSI"],"black","0.0"),
              ("🔥 YES" if r.get("long_buildup") else "-","black","@"),
              (r["Pattern Long"],"black","@")]
        for j,(v,fg,fmt) in enumerate(vals):
            col=get_column_letter(j+1); c=ws[f"{col}{row}"]
            c.value=v; c.font=Font(bold=(j<=1),size=9,color=P.get(fg,fg),name="Calibri")
            c.fill=PatternFill("solid",fgColor=bg)
            c.alignment=Al("center","center"); c.border=Bd()
            if fmt: c.number_format=fmt

def sheet_tier_b_investment(wb, tier_b, today_str):
    ws = wb.create_sheet("💰 TIER B INVESTMENT")
    ws.sheet_view.showGridLines = False; ws.freeze_panes = "B3"
    MH(ws,"A1:I1",f"TIER B — INVESTMENT CANDIDATES (fundamentals-weighted, technicals for entry timing)  |  {today_str}",bg="dblue",sz=11)
    hdrs=[("A","Rank",7),("B","Symbol",13),("C","CMP",10),("D","Blended\nScore",10),
          ("E","Fund\nScore",9),("F","Tech\nScore",9),("G","ROE %",8),
          ("H","P/E",8),("I","5Y Avg\nP/E",9)]
    for col,txt,w in hdrs:
        W(ws,f"{col}2",txt,bg="dblue",fg="white",bold=True,sz=9,wrap=True)
        ws.column_dimensions[col].width=w
    ws.row_dimensions[2].height=30
    for i,r in enumerate(tier_b):
        row=i+3; bg=P["lblue"] if i%2==0 else P["white"]
        vals=[(i+1,"black","0"),(r["Symbol"],"navy","@"),(r["CMP"],"black","#,##0.00"),
              (r.get("blended_score"),"navy","0.0"),(r.get("fund_score"),"black","0.0"),
              (r.get("tech_score",r["LONG SCORE"]),"black","0.0"),(r.get("roe"),"black","0.0"),
              (r.get("pe_ttm"),"black","0.0"),(r.get("pe_5y_median"),"black","0.0")]
        for j,(v,fg,fmt) in enumerate(vals):
            col=get_column_letter(j+1); c=ws[f"{col}{row}"]
            c.value=v; c.font=Font(bold=(j<=1),size=9,color=P.get(fg,fg),name="Calibri")
            c.fill=PatternFill("solid",fgColor=bg)
            c.alignment=Al("center","center"); c.border=Bd()
            if fmt: c.number_format=fmt


# ─────────────────────────────────────────────────────────────────────────────
#  SHEET: 🏆 TOP PICKS -- single consolidated view for next-day stock selection,
#  combining all three strategies (F&O, Swing, Investment) with their own
#  ranking logic, top N each, side by side in one sheet.
# ─────────────────────────────────────────────────────────────────────────────

def sheet_top_picks(wb, fo_ranked, tier_a, tier_b, today_str, top_n=5):
    ws = wb.create_sheet("🏆 TOP PICKS")
    ws.sheet_view.showGridLines = False
    MH(ws, "A1:H1", f"TOP PICKS FOR TOMORROW — F&O + Swing + Investment  |  {today_str}", bg="navy", sz=12)
    ws.row_dimensions[1].height = 22

    col_widths = {"A": 4, "B": 14, "C": 10, "D": 10, "E": 22, "F": 11, "G": 11, "H": 44}
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

    row = 3

    def section(title, bg, headers):
        nonlocal row
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(headers))
        c = ws.cell(row=row, column=1, value=title)
        c.font = Font(bold=True, size=11, color="FFFFFF", name="Calibri")
        c.fill = F(bg)
        c.alignment = Al("left", "center")
        row += 1
        for j, h in enumerate(headers, 1):
            cell = ws.cell(row=row, column=j, value=h)
            cell.font = Font(bold=True, size=9, color="FFFFFF", name="Calibri")
            cell.fill = F(bg)
            cell.alignment = Al("center", "center", wrap=True)
        row += 1

    def data_row(vals, band_bg, bold_first=True):
        nonlocal row
        for j, v in enumerate(vals, 1):
            c = ws.cell(row=row, column=j, value=v)
            c.font = Font(bold=(j == 2 and bold_first), size=9, color="1F3864", name="Calibri")
            c.fill = PatternFill("solid", fgColor=band_bg)
            c.alignment = Al("center", "center", wrap=(j == len(vals)))
            c.border = Bd()
        row += 1

    # ---- F&O section ----
    section("📥 F&O CONFIRMED (Rank | Symbol | CMP | Score | Action | SL | T1 | Why)", "dblue",
             ["#", "Symbol", "CMP", "Score", "Action", "SL", "T1", "Why"])
    if not fo_ranked:
        data_row(["", "No F&O-confirmed picks today (need auto-filled OI + score ≥60 + bullish F&O trend)", "", "", "", "", "", ""], P["white"])
    for i, r in enumerate(fo_ranked[:top_n]):
        why = f"{r['fo_trend']}, {r['fo_confidence']}, {r['Pattern Long']}"
        data_row([i+1, r["Symbol"], r["CMP"], r["LONG SCORE"], r["fo_final_action"],
                  r["SL Long"], r["T1 Long"], why], P["lblue"] if i % 2 == 0 else P["white"])
    row += 1

    # ---- Tier A Swing section ----
    section("🎯 SWING — Tier A (fundamentally-sound, best technical setup)", "green",
             ["#", "Symbol", "CMP", "Tech Score", "Signal", "SL", "T1", "Why"])
    if not tier_a:
        data_row(["", "No Tier A picks (my-full-nse-universe.csv missing or none passed gate)", "", "", "", "", "", ""], P["white"])
    for i, r in enumerate(tier_a[:top_n]):
        why = f"{'🔥 F&O Long Buildup, ' if r.get('long_buildup') else ''}{r['Pattern Long']}, ROE {r.get('roe','—')}"
        data_row([i+1, r["Symbol"], r["CMP"], r.get("tech_score", r["LONG SCORE"]), r["LONG SIGNAL"],
                  r["SL Long"], r["T1 Long"], why], P["lgreen"] if i % 2 == 0 else P["white"])
    row += 1

    # ---- Tier B Investment section ----
    section("💰 INVESTMENT — Tier B (fundamentals-weighted, technicals for entry timing)", "navy",
             ["#", "Symbol", "CMP", "Blend Score", "ROE %", "P/E", "5Y P/E", "Why"])
    if not tier_b:
        data_row(["", "No Tier B picks (my-full-nse-universe.csv missing or none passed gate)", "", "", "", "", "", ""], P["white"])
    for i, r in enumerate(tier_b[:top_n]):
        why = f"Fund {r.get('fund_score',0):.0f}/100, Tech {r.get('tech_score',0):.0f}/100, {r['LONG SIGNAL']}"
        data_row([i+1, r["Symbol"], r["CMP"], round(r.get("blended_score",0),1), r.get("roe","—"),
                  r.get("pe_ttm","—"), r.get("pe_5y_median","—"), why], P["lblue"] if i % 2 == 0 else P["white"])


def sheet_excluded(wb, excluded, today_str):
    ws = wb.create_sheet("🚫 EXCLUDED")
    ws.sheet_view.showGridLines = False; ws.freeze_panes = "B3"
    MH(ws,"A1:D1",f"EXCLUDED — Technically Flagged but Failed Fundamental Gate  |  {today_str}",bg="red",sz=11)
    hdrs=[("A","Symbol",13),("B","Long\nScore",9),("C","Long Signal",16),("D","Reasons",50)]
    for col,txt,w in hdrs:
        W(ws,f"{col}2",txt,bg="red",fg="white",bold=True,sz=9,wrap=True)
        ws.column_dimensions[col].width=w
    ws.row_dimensions[2].height=30
    for i,r in enumerate(sorted(excluded,key=lambda x:x["LONG SCORE"],reverse=True)):
        row=i+3; bg=P["lred"] if i%2==0 else P["white"]
        vals=[(r["Symbol"],"navy","@"),(r["LONG SCORE"],"navy","0"),
              (r["LONG SIGNAL"],"navy","@"),("; ".join(r.get("gate_reasons",[])),"black","@")]
        for j,(v,fg,fmt) in enumerate(vals):
            col=get_column_letter(j+1); c=ws[f"{col}{row}"]
            c.value=v; c.font=Font(bold=(j==0),size=9,color=P.get(fg,fg),name="Calibri")
            c.fill=PatternFill("solid",fgColor=bg)
            c.alignment=Al("center","center",wrap=(j==3)); c.border=Bd()
            if fmt: c.number_format=fmt


# ─────────────────────────────────────────────────────────────────────────────
#  SHEET: 📖 HOW TO USE
# ─────────────────────────────────────────────────────────────────────────────

def sheet_howto(wb):
    ws = wb.create_sheet("📖 HOW TO USE")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 70
    ws.column_dimensions["C"].width = 5

    steps = [
        ("navy","⭐  HOW TO USE THIS FILE EVERY EVENING  ⭐","header"),
        ("dblue",f"Capital ₹{CONFIG['total_capital']:,.0f}  |  Max/trade ₹{CONFIG['max_per_trade']:,.0f}  |  Long SL {CONFIG['stop_loss_pct']}%  |  Short SL {CONFIG['short_sl_pct']}%","sub"),
        (None,None,"gap"),
        ("green","STEP 1 — Run the script (after 5 PM IST)","step"),
        ("lgreen","Double-click ▶ RUN SCREENER (Double-Click).bat\nThis now scans a much larger universe than before (F&O stocks + Nifty 500),\nso it will take noticeably longer than the old ~4 minute run. A new Excel\nfile is created in your folder when it finishes.","detail"),
        (None,None,"gap"),
        ("green","STEP 2 — Open Excel → start on 🏆 TOP PICKS","step"),
        ("lgreen","The file now opens on the TOP PICKS sheet -- top 5 stocks from each of the\nthree strategies (F&O, Swing, Investment) in one place. Use this for a fast\ndaily read, then go to 📋 F&O INPUT for the full detailed working sheet.","detail"),
        (None,None,"gap"),
        ("yelb","STEP 3 — Fill 3 yellow columns from Sensibull.com","step"),
        ("FFE066","Open sensibull.com → search for the stock name\n\n   Col H  (TODAY FUT OI)  →  Futures tab → Open Interest today\n   Col I  (PREV DAY FUT OI)  →  Futures tab → Previous day OI\n   Col J  (PCR)  →  Options Chain tab → PCR number at the very top\n\nFormulas in cols K, L, M, N, O, P, X recalculate INSTANTLY as you type.\nYou do NOT need to press any button for the formulas to update.","detail"),
        (None,None,"gap"),
        ("blue","STEP 4 — Read the FINAL ACTION column (col O)","step"),
        ("lblue",'🟢 STRONG BUY  →  Enter tomorrow at open\n🔵 BUY  →  Enter on a small dip\n🔻 STRONG SHORT  →  Sell futures tomorrow\n🔽 SHORT  →  Sell on a bounce\n⚠️ AVOID  →  Technical score good but F&O says no — skip\n🟡 WAIT  →  Set alert, check tomorrow\n⬜ SKIP  →  No setup today',"detail"),
        (None,None,"gap"),
        ("green","STEP 5 — Check ⭐ FINAL CALLS tab for summary view","step"),
        ("lgreen","The FINAL CALLS sheet pulls everything from F&O INPUT automatically.\nGreen rows = BUY candidates. Red rows = SHORT candidates.","detail"),
        (None,None,"gap"),
        ("navy","STOCK UNIVERSE — WHERE EACH STRATEGY'S STOCKS COME FROM","step"),
        ("input",
         "Each of the three strategies scans a DIFFERENT set of stocks, on purpose:\n\n"
         "📥 F&O:  every stock currently eligible for F&O trading on NSE. This list is\n"
         "NOT hardcoded — it's read directly from the F&O Bhavcopy download each run,\n"
         "so it automatically stays current as NSE adds/removes F&O stocks every\n"
         "quarter. Roughly 180-220 stocks. No fundamental filter is applied here —\n"
         "F&O is purely technical score + OI/PCR confirmation, matching how F&O INPUT\n"
         "has always worked.\n\n"
         "🎯 SWING (Tier A) and 💰 INVESTMENT (Tier B):  Nifty 500 constituents —\n"
         "much broader than the old fixed 126-stock list, while index membership\n"
         "itself already screens out obvious shell/illiquid companies (Nifty 500\n"
         "requires minimum trading history, liquidity and market cap to be included).\n"
         "If the Nifty 500 download fails (network issue), the script falls back to\n"
         "the original curated list so it still runs.\n\n"
         "Fundamentals now come from YOUR OWN screener.in screen (not a shared\n"
         "public one) -- immune to someone else changing its columns or criteria\n"
         "without your knowledge.\n\n"
         "On top of Nifty 500 membership, a stock must ALSO clear every one of these\n"
         "to appear in Tier A or Tier B (see 🚫 EXCLUDED for exactly why any given\n"
         "stock didn't make it):\n"
         "  • ROE ≥ 12%, ROCE ≥ 12%, Debt/Equity ≤ 1.5\n"
         "  • 3-year revenue growth ≥ 5%, 3-year profit growth not declining\n"
         "  • Promoter pledge ≤ 20%\n"
         "  • Market cap ≥ ₹20,000 Cr (roughly upper-small-cap and above by current\n"
         "    AMFI cutoffs -- excludes the most volatile small-caps entirely)\n"
         "  • 20-day average daily traded value ≥ ₹50 Cr (excludes thinly-traded stocks\n"
         "    that are both easy to manipulate and hard to exit without slippage)\n"
         "  • Cash from Operations positive in BOTH of the last 2 years (paper profit\n"
         "    with no real cash coming in is a classic earnings-manipulation signal)\n"
         "  • Interest Coverage Ratio ≥ 3x (comfortable debt servicing)\n"
         "  • Debtor Days ≤ 180 (a generous cutoff -- only flags clear outliers, not\n"
         "    normal sector variation; rising receivables can mean fake/unpaid sales)\n"
         "  • Operating margin not eroding >25% vs its own 5-year average\n"
         "  • Not currently on NSE's ASM/GSM surveillance list, when that check is\n"
         "    reachable (best-effort — see note below)\n\n"
         "Dividend Payout % is loaded and shown in 🧮 FUND + TECH VIEW but deliberately\n"
         "NOT gated -- many good growth companies pay 0% and reinvest everything, so a\n"
         "minimum-payout rule would wrongly exclude them. Use your own judgment there.\n\n"
         "WHY THESE SPECIFIC CHECKS: market cap, liquidity, promoter pledge, and weak\n"
         "operating cash flow are among the most common red flags in Indian small/\n"
         "mid-cap manipulation cases — thin float makes a price easy to move, high\n"
         "pledge signals promoter financial stress, and profit without matching cash\n"
         "is one of the oldest earnings-manipulation tricks in the book. None of this\n"
         "GUARANTEES a stock is clean, but it removes the categories of stock most\n"
         "associated with pump-and-dump activity.\n\n"
         "NOTE ON ASM/GSM: NSE's live surveillance list isn't reliably automatable\n"
         "yet in this script (their reports page is dashboard-driven, not a stable\n"
         "download link) — right now it always passes every stock through unfiltered.\n"
         "Treat this as a placeholder for a future improvement, not a working filter\n"
         "today. Worth manually checking any stock you're about to trade against\n"
         "nseindia.com/reports/asm before entering, especially for anything outside\n"
         "the largest, most familiar names.","detail"),
        (None,None,"gap"),
        ("sdark","TRADING RULES — MUST FOLLOW","step"),
        ("slight","• Never enter without checking F&O — technical score alone is not enough\n• Max 5 positions at once. Max ₹2L per stock.\n• Long SL: 5% below entry. Short SL: 4% above entry.\n• Only enter long if Nifty 50 is above its 50 EMA\n• Never short in a strong bull market\n• Trail long: at +5% move SL to breakeven; at +8% exit 50%\n• Min R:R = 1:2 before entering any trade","detail"),
        (None,None,"gap"),
        ("navy","SENSIBULL QUICK GUIDE","step"),
        ("input","1. Go to sensibull.com and log in (free account)\n2. Search stock name (e.g. DRREDDY)\n3. Click 'Futures' tab → note Open Interest (today & prev day) → enter in H and I\n4. Click 'Options Chain' tab → PCR is shown at the TOP of the chain → enter in J\n5. Done! Read col O (Final Action) for your decision.","detail"),
    ]

    row = 1
    for bg,text,typ in steps:
        if typ=="gap":
            ws.row_dimensions[row].height = 8
            row+=1; continue
        ws.merge_cells(f"B{row}:B{row}")
        c=ws[f"B{row}"]
        c.value=text
        if typ=="header":
            c.font=Font(bold=True,size=14,color="FFFFFF",name="Calibri")
            c.fill=F(bg); c.alignment=Al("center","center")
            ws.row_dimensions[row].height=30
        elif typ=="sub":
            c.font=Font(bold=False,size=10,color="FFFFFF",name="Calibri")
            c.fill=F(bg); c.alignment=Al("center","center")
            ws.row_dimensions[row].height=18
        elif typ=="step":
            c.font=Font(bold=True,size=11,color="FFFFFF" if bg not in ("yelb","FFE066") else "000000",name="Calibri")
            c.fill=PatternFill("solid",fgColor=P.get(bg,bg))
            c.alignment=Al("left","center")
            ws.row_dimensions[row].height=22
        elif typ=="detail":
            c.font=Font(size=10,name="Calibri",color="000000")
            c.fill=PatternFill("solid",fgColor=P.get(bg,bg))
            c.alignment=Al("left","center",wrap=True)
            lines=text.count("\n")+1
            ws.row_dimensions[row].height=max(15*lines,20)
        c.border=Bd()
        row+=1


# ─────────────────────────────────────────────────────────────────────────────
#  ADD VBA MACRO BUTTON (xlsm)
# ─────────────────────────────────────────────────────────────────────────────

VBA_CODE = '''
Attribute VB_Name = "Module1"
Sub SortFinalCalls()
    Dim wsInput As Worksheet
    Dim wsFinal As Worksheet
    
    On Error GoTo ErrHandler
    
    Set wsInput = ThisWorkbook.Sheets("=\u30c4= F&O INPUT")
    Set wsFinal = ThisWorkbook.Sheets("=\u2b50= FINAL CALLS")
    
    ' Switch to Final Calls sheet
    wsFinal.Activate
    
    ' Force recalculation
    Application.Calculate
    
    MsgBox "=\u2705 F&O data processed!" & Chr(10) & Chr(10) & _
           "=\u2b50 FINAL CALLS sheet is now active." & Chr(10) & _
           "=\ud83d\udfe2 Green rows = BUY candidates" & Chr(10) & _
           "=\ud83d\udd3b Red rows = SHORT candidates" & Chr(10) & Chr(10) & _
           "Tip: Look for rows highlighted in green (Strong Buy) or red (Strong Short).", _
           vbInformation, "NSE Screener - Final Calls Ready"
    Exit Sub
    
ErrHandler:
    Application.Calculate
    MsgBox "Calculation complete! Check the FINAL CALLS sheet.", vbInformation, "Done"
End Sub
'''

def add_button_to_sheet(ws, last_data_row):
    """Add a styled button-like cell above the data that user can use as a button."""
    # We use a merged cell styled as a button with instructions
    btn_row = last_data_row + 2
    ws.merge_cells(f"A{btn_row}:X{btn_row}")
    c = ws[f"A{btn_row}"]
    c.value = ("🔄  AFTER FILLING H, I, J COLUMNS ABOVE  →  "
               "Press Ctrl+Alt+F9 to force recalculate  "
               "→  Then click ⭐ FINAL CALLS tab to see your trade plan")
    c.font = Font(bold=True, size=11, color="FFFFFF", name="Calibri")
    c.fill = PatternFill("solid", fgColor="1E8449")
    c.alignment = Al("center","center")
    c.border = Border(
        left=Side(style="medium",color="145A32"),
        right=Side(style="medium",color="145A32"),
        top=Side(style="medium",color="145A32"),
        bottom=Side(style="medium",color="145A32"))
    ws.row_dimensions[btn_row].height = 28


# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    today_str = datetime.now().strftime("%d-%b-%Y %I:%M %p")
    date_str  = date.today().strftime("%Y-%m-%d")

    print(f"\n{'='*68}")
    print(f"  🚀 NSE SCREENER v5.0  |  {today_str}")
    print(f"{'='*68}\n")

    # ── Phase 0: F&O Bhavcopy FIRST -- this is what defines the F&O
    # category's universe. Downloading it before the technical scan (rather
    # than after, as in earlier versions) means the F&O-eligible list is
    # known upfront and always current -- no hardcoded list to go stale as
    # NSE reconstitutes F&O stocks each quarter.
    print("  PHASE 0 — Downloading F&O Bhavcopy (defines the F&O universe)...\n")
    fo_data = download_bhavcopy()
    fo_universe = sorted(fo_data.keys()) if fo_data else []
    print(f"  📥 F&O-eligible universe: {len(fo_universe)} stocks"
          + (" (download failed -- F&O category will be empty this run)" if not fo_data else ""))

    # ── Phase 0.5: Nifty 500 for the Swing/Investment universe -- much
    # larger and more diverse than the old 126-stock hardcoded list, while
    # index membership itself already screens out obvious shell/manipulated
    # companies (minimum liquidity/market-cap/listing-history to be included).
    print("\n  Downloading Nifty 500 list (Swing/Investment universe)...\n")
    swing_inv_universe = download_nifty500_list()
    if swing_inv_universe:
        print(f"  📈 Swing/Investment universe: {len(swing_inv_universe)} stocks (Nifty 500)")
    else:
        swing_inv_universe = list(NSE_UNIVERSE)
        print(f"  ⚠️  Nifty 500 download failed -- falling back to the {len(swing_inv_universe)}-stock curated list")

    scan_universe = sorted(set(fo_universe) | set(swing_inv_universe))
    print(f"\n  🔎 Combined scan universe: {len(scan_universe)} stocks "
          f"(this run will take longer than the old 126-stock scan)\n")

    # Phase 1: Price + Technical Scores, over the COMBINED universe
    print("  PHASE 1 — Downloading price data & scoring all stocks...\n")
    results, failed = [], []
    for i, sym in enumerate(scan_universe):
        print(f"  [{i+1:>3}/{len(scan_universe)}]  {sym:<14}", end=" ", flush=True)
        r = analyse_stock(sym)
        if r:
            results.append(r)
            print(f"L:{r['LONG SCORE']:>3}  S:{r['SHORT SCORE']:>3}  {r['LONG SIGNAL']}")
        else:
            failed.append(sym); print("⚠️  No data")

    if not results:
        print("\n❌ No results — check internet connection."); return

    # Sort: priority buy first, then by long score
    results.sort(key=lambda x: (
        2 if x["LONG SCORE"]>=CONFIG["priority_buy_score"] else
        1 if x["LONG SCORE"]>=CONFIG["min_score_buy"] else 0,
        x["LONG SCORE"]
    ), reverse=True)

    # Summary
    strong_buy_count  = sum(1 for r in results if r["LONG SCORE"]>=CONFIG["priority_buy_score"])
    buy_count         = sum(1 for r in results if CONFIG["min_score_buy"]<=r["LONG SCORE"]<CONFIG["priority_buy_score"])
    short_count       = sum(1 for r in results if r["SHORT SCORE"]>=CONFIG["min_score_short"])

    print(f"\n{'='*68}")
    print(f"  ✅  {len(results)} stocks scanned  |  {len(failed)} failed")
    print(f"  🟢  Priority Buy (≥{CONFIG['priority_buy_score']}): {strong_buy_count} stocks")
    print(f"  🔵  Watchlist Buy (≥{CONFIG['min_score_buy']}):  {buy_count} stocks")
    print(f"  🔽  Short candidates (≥{CONFIG['min_score_short']}): {short_count} stocks")
    print(f"{'='*68}")

    if fo_data:
        matched   = sum(1 for r in results if r["Symbol"] in fo_data)
        unmatched = len(results) - matched
        print(f"  📊 OI + PCR auto-filled for {matched} stocks")
        if unmatched > 0:
            missing = [r["Symbol"] for r in results if r["Symbol"] not in fo_data]
            print(f"  🟡 {unmatched} stocks need manual Sensibull entry: {', '.join(missing[:8])}"
                  + (" ..." if unmatched>8 else ""))
    else:
        print("  ⚠️  Bhavcopy download failed — yellow cells for manual entry on all stocks")

    surveillance_excluded = load_surveillance_exclusions()

    # Phase 3: Fundamentals -- swing/investment layer on top of the technical scan
    print(f"\n  PHASE 3 — Fundamentals (Swing/Investment layer)...\n")
    fund_path = os.path.join(SCRIPT_DIR, FUND_CONFIG["fundamentals_csv"])
    fund_df = load_fundamentals(fund_path)
    tier_a, tier_b, excluded = [], [], []

    if fund_df is None:
        print(f"  ⚠️  '{FUND_CONFIG['fundamentals_csv']}' not found in this folder — skipping")
        print(f"     🧮 FUND+TECH VIEW / 🎯 TIER A SWING / 💰 TIER B INVESTMENT / 🚫 EXCLUDED tabs.")
        print(f"     To enable: export from https://www.screener.in/screens/638306/all-nse-stocks/")
        print(f"     and save as '{FUND_CONFIG['fundamentals_csv']}' in this folder, then re-run.")
    else:
        watchlist_symbols = [r["Symbol"] for r in results]
        have_roe = set(fund_df.loc[fund_df["roe"].notna(), "symbol"])
        missing = [s for s in watchlist_symbols if s not in have_roe]
        if missing and FUND_CONFIG["use_yfinance_fallback"]:
            print(f"  Filling {len(missing)} symbols (from today's watchlist only) via yfinance fallback...")
            yf_fund = yfinance_fill_fundamentals(missing)
            fund_df = merge_fundamentals(fund_df, yf_fund)

        gate_results = fund_df.apply(lambda r: apply_fund_gate(r, FUND_CONFIG), axis=1)
        fund_df["gate_pass"] = gate_results.apply(lambda t: t[0])
        fund_df["gate_reasons"] = gate_results.apply(lambda t: t[1])
        fund_df = fund_composite_score(fund_df)
        fund_lookup = fund_df.set_index("symbol").to_dict("index")

        for r in results:
            f = fund_lookup.get(r["Symbol"])
            r["gate_pass"]    = f["gate_pass"] if f else None
            r["gate_reasons"] = list(f["gate_reasons"]) if f else ["No fundamentals data"]
            r["fund_score"]   = f.get("fund_score") if f else None
            for k in ("roe","roce","debt_to_equity","pe_ttm","pe_5y_median","promoter_pledge_pct","market_cap_cr",
                      "cfo_last_year","cfo_preceding_year","debtor_days","interest_coverage","opm","opm_5y","dividend_payout_pct"):
                r[k] = f.get(k) if f else None

            # Liquidity floor -- excludes thinly-traded stocks that are
            # easier to manipulate and harder to exit without slippage.
            # Computed from the technical scan, not the fundamentals CSV,
            # so it's checked here rather than inside apply_fund_gate().
            r["avg_daily_turnover_cr"] = round((r["CMP"] * r.get("AvgVol20", 0)) / 1e7, 2)
            if r["avg_daily_turnover_cr"] < FUND_CONFIG["min_avg_daily_turnover_cr"]:
                r["gate_reasons"].append(f"Avg daily turnover too low (₹{r['avg_daily_turnover_cr']:.1f}Cr)")
                if r["gate_pass"] is True:
                    r["gate_pass"] = False

            # Swing/Investment universe membership -- a stock only qualifies
            # for Tier A/B if it's in the quality universe (Nifty 500 or the
            # offline fallback), even if it happens to also be F&O-eligible.
            r["in_swing_inv_universe"] = r["Symbol"] in swing_inv_universe
            if not r["in_swing_inv_universe"] and r["gate_pass"] is True:
                r["gate_pass"] = False
                r["gate_reasons"].append("Not in Swing/Investment universe (Nifty 500)")

            # NSE surveillance (ASM/GSM) exclusion -- best-effort, see
            # load_surveillance_exclusions(). Empty set today; wired in for
            # when that lookup is fully working.
            if r["Symbol"] in surveillance_excluded:
                r["gate_pass"] = False
                r["gate_reasons"].append("On NSE surveillance list (ASM/GSM)")

            long_buildup = False
            fo = fo_data.get(r["Symbol"]) if fo_data else None
            if fo and r["Day Chg %"] >= FUND_CONFIG["fno_long_buildup_price_chg_min"] and fo.get("prev_oi", 0) > 0:
                oi_chg = (fo["today_oi"] - fo["prev_oi"]) / fo["prev_oi"] * 100
                if oi_chg >= FUND_CONFIG["fno_long_buildup_oi_chg_min"]:
                    long_buildup = True
            r["long_buildup"] = long_buildup
            r["tech_score"] = min(100, r["LONG SCORE"] + (FUND_CONFIG["fno_long_buildup_bonus"] if long_buildup else 0))

        eligible = [r for r in results if r.get("gate_pass") is True]
        excluded = [r for r in results if r.get("gate_pass") is False]

        tier_a = sorted(eligible, key=lambda r: r["tech_score"], reverse=True)[:FUND_CONFIG["tier_top_n"]]
        for r in eligible:
            r["blended_score"] = 0.65 * (r.get("fund_score") or 0) + 0.35 * r["tech_score"]
        tier_b = sorted(eligible, key=lambda r: r["blended_score"], reverse=True)[:FUND_CONFIG["tier_top_n"]]

        print(f"  ✅ {len(eligible)} passed fundamental gate, {len(excluded)} excluded")
        print(f"  🎯 Tier A (swing): {len(tier_a)} candidates  |  💰 Tier B (investment): {len(tier_b)} candidates")

    # F&O category -- same Final Action logic as the F&O INPUT sheet's live
    # Excel formulas, computed here in Python so it can be ranked directly
    # for the Top Picks sheet (only possible for stocks with auto-filled OI
    # from the Bhavcopy download -- manually-entered ones aren't available
    # until you fill them in Excel yourself).
    fo_ranked = []
    for r in results:
        fo = fo_data.get(r["Symbol"]) if fo_data else None
        today_oi = fo.get("today_oi") if fo else None
        prev_oi = fo.get("prev_oi") if fo else None
        pcr = fo.get("pcr") if fo else None
        action = fo_final_action(r["LONG SCORE"], r["SHORT SCORE"], r["Day Chg %"], today_oi, prev_oi, pcr)
        r["fo_final_action"] = action["final_action"]
        r["fo_confidence"] = action["confidence"]
        r["fo_trend"] = action["fo_trend"]
        if action["final_action"] in ("🟢 STRONG BUY", "🔵 BUY"):
            fo_ranked.append(r)
    fo_ranked.sort(key=lambda r: (r["fo_final_action"] != "🟢 STRONG BUY", -r["LONG SCORE"]))
    print(f"  📥 F&O category: {len(fo_ranked)} stocks with confirmed STRONG BUY/BUY (auto-filled OI only)")

    print(f"\n  📊 Building Excel...\n")
    wb = Workbook(); wb.remove(wb.active)

    # 🏆 TOP PICKS goes first -- it's the single-sheet daily decision view
    sheet_top_picks(wb, fo_ranked, tier_a, tier_b, today_str, top_n=5)

    # Then the rest -- F&O INPUT is still the detailed working sheet
    ws_input, last_row = sheet_fo_input(wb, results, today_str, fo_data)
    add_button_to_sheet(ws_input, last_row)
    sheet_final_calls(wb, results, today_str)
    sheet_buy_watchlist(wb, results, today_str)
    sheet_short_watchlist(wb, results, today_str)
    sheet_all_stocks(wb, results, today_str)
    sheet_fo_data_downloaded(wb, fo_data, today_str)
    if fund_df is not None:
        sheet_fund_tech_view(wb, results, today_str)
        sheet_tier_a_swing(wb, tier_a, today_str)
        sheet_tier_b_investment(wb, tier_b, today_str)
        sheet_excluded(wb, excluded, today_str)
    sheet_howto(wb)

    wb.active = wb["🏆 TOP PICKS"]

    outfile = os.path.join(SCRIPT_DIR, f"NSE_Watchlist_{date_str}.xlsx")
    try:
        wb.save(outfile)
    except PermissionError:
        ts = datetime.now().strftime("%H%M")
        outfile = os.path.join(SCRIPT_DIR, f"NSE_Watchlist_{date_str}_{ts}.xlsx")
        wb.save(outfile)

    print(f"{'='*68}")
    print(f"  ✅ SAVED:  {outfile}")
    if fo_data:
        print(f"\n  🟢 F&O data auto-filled — just open Excel and go to ⭐ FINAL CALLS!")
        print(f"  🟡 For any yellow cells — manually type from Sensibull")
    else:
        print(f"\n  📋 File opens on F&O INPUT sheet")
        print(f"     Yellow columns H, I, J → type from Sensibull")
    print(f"     Columns K–X recalculate INSTANTLY")
    print(f"{'='*68}\n")

    print("  🟢 TOP LONG CANDIDATES:\n")
    for r in results[:5]:
        if r["LONG SCORE"] >= CONFIG["min_score_buy"]:
            print(f"  {r['Symbol']:<14} L:{r['LONG SCORE']}  RSI:{r['RSI']}  {r['Pattern Long']}")
    print()
    print("  🔻 TOP SHORT CANDIDATES:\n")
    shorts = sorted(results, key=lambda x: x["SHORT SCORE"], reverse=True)
    for r in shorts[:5]:
        if r["SHORT SCORE"] >= CONFIG["min_score_short"]:
            print(f"  {r['Symbol']:<14} S:{r['SHORT SCORE']}  RSI:{r['RSI']}  {r['Pattern Short']}")
    print()

    if fund_df is not None:
        print("  🎯 TOP TIER A (SWING) — fundamentally-sound, best technical setup:\n")
        for r in tier_a[:5]:
            print(f"  {r['Symbol']:<14} Tech:{r['tech_score']:.0f}  Fund:🟢  {r['LONG SIGNAL']}")
        print()
        print("  💰 TOP TIER B (INVESTMENT) — best fundamentals, technicals for entry:\n")
        for r in tier_b[:5]:
            print(f"  {r['Symbol']:<14} Blend:{r['blended_score']:.0f}  ROE:{r.get('roe','—')}  P/E:{r.get('pe_ttm','—')}")
        print()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        print(f"\n{'='*68}")
        print(f"  ❌ CRASHED — Error details:")
        print(f"{'='*68}")
        traceback.print_exc()
        print(f"\n  Screenshot this and share.")
        print(f"{'='*68}\n")
    input("  Press ENTER to close...")
