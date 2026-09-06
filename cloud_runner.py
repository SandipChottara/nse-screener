"""
CLOUD RUNNER — runs the screener on GitHub Actions and writes docs/data.json
for the phone web app to read.

Reuses nse_momentum_screener.py's OWN functions (analyse_stock, the
fundamental gate, fo_final_action, etc) rather than reimplementing
anything -- same logic as your PC run, just JSON output instead of Excel.

HONEST LIMITATION -- NSE GEO-BLOCKING:
GitHub Actions runners live in Microsoft Azure datacentres (mostly US).
NSE's archive server frequently blocks non-Indian datacentre IPs. So:
  - yfinance price data (Yahoo): reliably works from anywhere
  - NSE Bhavcopy (OI/PCR for F&O): MAY be blocked from Actions runners
If Bhavcopy fails, this still produces a complete Swing + Investment
result and marks F&O as unavailable, rather than failing the whole run.
The app shows that state honestly instead of pretending F&O is empty.
"""

import json, os, sys, traceback
from datetime import datetime, timedelta
import pandas as pd

sys.argv = ["cloud_runner"]  # keep v5's arg handling quiet
import nse_momentum_screener as v5

OUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "docs", "data.json")
HIST_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "docs", "history.json")
HIST_WINDOW_DAYS = 45   # keep a bit more than 30 so the 30-day view is always full


# ─────────────────────────────────────────────────────────────────────────────
#  TRACK RECORD — rolling history of every pick, with realized / unrealized P&L
#
#  HOW IT WORKS: each run appends today's picks (one entry per symbol per
#  category, no duplicates while a position is still open), then re-checks
#  every open position against actual daily High/Low bars since entry to see
#  if Stop Loss or Target was hit. Exits are marked realized; anything still
#  running is marked unrealized and marked-to-market at the latest close.
#
#  HONEST LIMITATION: history only accumulates from the first run onward --
#  it cannot be backfilled, because knowing what WOULD have been picked on a
#  past date requires re-scoring that date (that's what backtest_screener.py
#  does separately). So the first few days will look sparse; a full 30-day
#  window takes about six trading weeks to fill.
#
#  Exit rules mirror the live SL/Target: -5% stop, +10% Target 1 (books half,
#  stop moves to breakeven), +15% Target 2. Checked against daily High/Low,
#  not just closes, so an intraday stop-out is caught.
# ─────────────────────────────────────────────────────────────────────────────

#  Which category wins when the same stock qualifies in more than one on the
#  same day. F&O first (it has the tightest, fastest-resolving exits and needs
#  the OI confirmation to still be valid), then Swing, then Investment. Change
#  this order if you'd rather a dual-qualifying stock be logged as a long-term
#  investment than a short-term trade.
CATEGORY_PRIORITY = ("F&O", "Swing", "Investment")
PICKS_PER_CATEGORY = 3


def CATEGORY_PRIORITY_LISTS(fno, tier_a, tier_b):
    # fno here is the COMBINED long+short list, already ranked together
    by_name = {"F&O": fno, "Swing": tier_a, "Investment": tier_b}
    return [(c, by_name[c]) for c in CATEGORY_PRIORITY]


def load_history():
    if os.path.exists(HIST_PATH):
        try:
            with open(HIST_PATH, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"picks": []}


def update_track_record(fno, tier_a, tier_b, results_by_sym, fno_short_or_none=None):
    """Append today's picks, then mark every open position to market."""
    import yfinance as yf
    hist = load_history()
    picks = hist.get("picks", [])
    today = datetime.utcnow().strftime("%Y-%m-%d")

    # A stock is held ONCE, not once per category -- you'd only buy it one time.
    # So dedup is by symbol across every category, and against anything already
    # open. When the same stock qualifies in more than one category on the same
    # day, the first category in CATEGORY_PRIORITY below wins and the others
    # skip it. That means the recorded list is often shorter than 9 -- which is
    # correct: it mirrors what you'd actually buy.
    # One-time cleanup: files written before the cross-category dedup fix can
    # contain the same symbol under two categories on the same date. Keep the
    # highest-priority one and drop the rest.
    _prio = {c: i for i, c in enumerate(CATEGORY_PRIORITY)}
    _seen, _clean, _removed = set(), [], 0
    for p in sorted(picks, key=lambda x: (x["date"], _prio.get(x["category"], 9))):
        key = (p["symbol"], p["date"])
        if key in _seen:
            _removed += 1
            continue
        _seen.add(key)
        _clean.append(p)
    if _removed:
        print(f"  Cleaned {_removed} duplicate entries from earlier runs")

    # Trim dates that exceeded the cap from repeat runs before this was fixed.
    _bycat, _trimmed, _kept = {}, 0, []
    for p in _clean:
        k = (p["date"], p["category"])
        _bycat[k] = _bycat.get(k, 0) + 1
        if _bycat[k] > PICKS_PER_CATEGORY:
            _trimmed += 1
            continue
        _kept.append(p)
    if _trimmed:
        print(f"  Trimmed {_trimmed} entries over the {PICKS_PER_CATEGORY}/day cap")
    picks = _kept

    held = {p["symbol"] for p in picks if p["status"] == "open"}
    taken_today = set()
    # Count what today already logged per category. Without this, a second run
    # on the same day adds 3 more (the old code only checked "already open?",
    # not "how many today?"), pushing a single date past the cap.
    already_today = {}
    for p in picks:
        if p["date"] == today:
            already_today[p["category"]] = already_today.get(p["category"], 0) + 1

    # Defensive: merge F&O shorts into one list so long+short share ONE cap of
    # 3, not 3 each -- guards against any caller passing them separately.
    if fno_short_or_none:
        for r in fno:
            r.setdefault("fo_side", "LONG")
            r.setdefault("fo_rank_score", r.get("LONG SCORE", 0))
        for r in fno_short_or_none:
            r["fo_side"] = "SHORT"
            r["fo_rank_score"] = r.get("SHORT SCORE", 0)
        fno = sorted(list(fno) + list(fno_short_or_none),
                     key=lambda r: -r.get("fo_rank_score", 0))

    for cat, lst in CATEGORY_PRIORITY_LISTS(fno, tier_a, tier_b):
        room = PICKS_PER_CATEGORY - already_today.get(cat, 0)
        if room <= 0:
            continue
        for r in lst:
            if room <= 0:
                break
            sym = r["Symbol"]
            if sym in held or sym in taken_today:
                continue
            taken_today.add(sym); room -= 1
            entry = r["CMP"]
            short = r.get("fo_side") == "SHORT"
            if short:   # stop ABOVE entry, targets BELOW
                sl = entry * (1 + v5.CONFIG["short_sl_pct"] / 100)
                t1 = entry * (1 - v5.CONFIG["short_target1_pct"] / 100)
                t2 = entry * (1 - v5.CONFIG["short_target2_pct"] / 100)
                sc = r["SHORT SCORE"]
            else:
                sl = entry * (1 - v5.CONFIG["stop_loss_pct"] / 100)
                t1 = entry * (1 + v5.CONFIG["target1_pct"] / 100)
                t2 = entry * (1 + v5.CONFIG["target2_pct"] / 100)
                sc = r.get("tech_score", r["LONG SCORE"])
            picks.append({
                "date": today, "category": cat, "symbol": sym,
                "side": "SHORT" if short else "LONG",
                "entry": safe(entry), "sl": safe(sl), "t1": safe(t1), "t2": safe(t2),
                "score": safe(sc),
                "status": "open", "exit_date": None, "exit_price": None,
                "exit_reason": None, "pnl_pct": None, "half_booked": False,
            })

    # drop anything older than the window so the file doesn't grow forever
    cutoff = (datetime.utcnow() - timedelta(days=HIST_WINDOW_DAYS * 2)).strftime("%Y-%m-%d")
    picks = [p for p in picks if p["date"] >= cutoff]

    # mark open positions to market against real daily bars
    open_syms = sorted({p["symbol"] for p in picks if p["status"] == "open"})
    bars = {}
    for sym in open_syms:
        try:
            df = yf.Ticker(f"{sym}.NS").history(period="3mo", interval="1d", auto_adjust=True)
            if df is not None and len(df):
                df.index = pd.to_datetime(df.index).tz_localize(None)
                bars[sym] = df
        except Exception:
            pass

    for p in picks:
        if p["status"] != "open":
            continue
        df = bars.get(p["symbol"])
        if df is None or not len(df):
            continue
        entry_dt = pd.to_datetime(p["date"])
        after = df[df.index >= entry_dt]
        if not len(after):
            continue

        entry, sl, t1, t2 = p["entry"], p["sl"], p["t1"], p["t2"]
        is_short = p.get("side") == "SHORT"
        half = p.get("half_booked", False)
        closed = False

        # Shorts profit when price FALLS: stop is breached by a HIGH above it,
        # targets by a LOW below them, and P&L is (entry - exit) not (exit - entry).
        def pnl(x):
            return safe(((entry - x) if is_short else (x - entry)) / entry * 100)
        def hit_stop(row, lvl):
            return row["High"] >= lvl if is_short else row["Low"] <= lvl
        def hit_tgt(row, lvl):
            return row["Low"] <= lvl if is_short else row["High"] >= lvl

        for dt, row in after.iterrows():
            if not half:
                if hit_stop(row, sl):
                    p.update(status="closed", exit_date=dt.strftime("%Y-%m-%d"),
                             exit_price=safe(sl), exit_reason="Stop Loss",
                             pnl_pct=pnl(sl))
                    closed = True; break
                if hit_tgt(row, t1):
                    half = True; p["half_booked"] = True; sl = entry  # stop to breakeven
            else:
                if hit_stop(row, sl):
                    p.update(status="closed", exit_date=dt.strftime("%Y-%m-%d"),
                             exit_price=safe(sl),
                             exit_reason="Target 1 hit, rest exited at breakeven",
                             pnl_pct=safe((pnl(t1) + pnl(sl)) / 2))
                    closed = True; break
                if hit_tgt(row, t2):
                    p.update(status="closed", exit_date=dt.strftime("%Y-%m-%d"),
                             exit_price=safe(t2), exit_reason="Target 2",
                             pnl_pct=safe((pnl(t1) + pnl(t2)) / 2))
                    closed = True; break
        if not closed:
            last = float(after["Close"].iloc[-1])
            p["current"] = safe(last)
            p["pnl_pct"] = pnl(last)
            p["days_held"] = (datetime.utcnow() - pd.to_datetime(p["date"]).to_pydatetime()).days

    hist["picks"] = picks
    hist["updated_at"] = datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")
    with open(HIST_PATH, "w", encoding="utf-8") as f:
        json.dump(hist, f, indent=1, ensure_ascii=False)

    # build a 30-trading-day summary for the app
    recent_cut = (datetime.utcnow() - timedelta(days=44)).strftime("%Y-%m-%d")
    recent = [p for p in picks if p["date"] >= recent_cut]
    summary = {}
    for cat in ("F&O", "Swing", "Investment"):
        c = [p for p in recent if p["category"] == cat]
        closed = [p for p in c if p["status"] == "closed" and p["pnl_pct"] is not None]
        openp = [p for p in c if p["status"] == "open" and p["pnl_pct"] is not None]
        wins = [p for p in closed if p["pnl_pct"] > 0]
        summary[cat] = {
            "total": len(c), "closed": len(closed), "open": len(openp),
            "realized_avg": safe(sum(p["pnl_pct"] for p in closed) / len(closed)) if closed else None,
            "unrealized_avg": safe(sum(p["pnl_pct"] for p in openp) / len(openp)) if openp else None,
            "win_rate": safe(len(wins) / len(closed) * 100) if closed else None,
        }
    return summary, sorted(recent, key=lambda p: (p["date"], p["symbol"]), reverse=True)


TOP_N = 10  # how many per category to publish to the app


def safe(v):
    """JSON can't hold NaN/inf; convert those to None."""
    if v is None:
        return None
    try:
        f = float(v)
        if f != f or f in (float("inf"), float("-inf")):
            return None
        return round(f, 2)
    except (TypeError, ValueError):
        return v


def main():
    print("=" * 60)
    print("  CLOUD RUNNER — NSE Screener")
    print("=" * 60)

    notes = []

    # ---- F&O universe (defines F&O category; may be geo-blocked) ----
    print("\n[1/4] F&O Bhavcopy...")
    try:
        fo_data = v5.download_bhavcopy()
    except Exception as e:
        print(f"  Bhavcopy failed: {e}")
        fo_data = {}
    if not fo_data:
        notes.append("NSE Bhavcopy unavailable from the cloud runner (likely geo-blocked) — F&O category could not be computed this run. Swing and Investment are unaffected.")
    fo_universe = sorted(fo_data.keys()) if fo_data else []
    print(f"  F&O universe: {len(fo_universe)} stocks")

    # ---- Swing/Investment universe ----
    print("\n[2/4] Nifty 500 list...")
    swing_inv = v5.download_nifty500_list()
    if not swing_inv:
        swing_inv = list(v5.NSE_UNIVERSE)
        notes.append(f"Nifty 500 list unavailable — fell back to the {len(swing_inv)}-stock curated list.")
    print(f"  Swing/Investment universe: {len(swing_inv)} stocks")

    scan = sorted(set(fo_universe) | set(swing_inv))
    print(f"  Combined scan universe: {len(scan)}")

    # ---- Technical scan ----
    print(f"\n[3/4] Scoring {len(scan)} stocks (this is the slow part)...")
    results = []
    for i, sym in enumerate(scan):
        r = v5.analyse_stock(sym)
        if r:
            results.append(r)
        if (i + 1) % 50 == 0:
            print(f"  {i+1}/{len(scan)} done, {len(results)} scored")
    print(f"  Scored {len(results)} stocks")

    if not results:
        raise RuntimeError("No stocks scored — yfinance may be failing entirely.")

    # ---- Fundamentals + tiers ----
    print("\n[4/4] Fundamentals + tiers...")
    fund_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              v5.FUND_CONFIG["fundamentals_csv"])
    fund_df = v5.load_fundamentals(fund_path)

    tier_a, tier_b = [], []
    if fund_df is None:
        notes.append(f"'{v5.FUND_CONFIG['fundamentals_csv']}' not found in the repo — Swing/Investment tiers need it. Commit your screener.in export to the repo root.")
    else:
        gate = fund_df.apply(lambda r: v5.apply_fund_gate(r, v5.FUND_CONFIG), axis=1)
        fund_df["gate_pass"] = gate.apply(lambda t: t[0])
        fund_df["gate_reasons"] = gate.apply(lambda t: t[1])
        fund_df = v5.fund_composite_score(fund_df)
        lookup = fund_df.set_index("symbol").to_dict("index")

        for r in results:
            f = lookup.get(r["Symbol"])
            r["gate_pass"] = f["gate_pass"] if f else None
            r["gate_reasons"] = list(f["gate_reasons"]) if f else ["No fundamentals data"]
            r["fund_score"] = f.get("fund_score") if f else None
            for k in ("roe", "roce", "pe_ttm", "pe_5y_median", "market_cap_cr"):
                r[k] = f.get(k) if f else None

            r["avg_daily_turnover_cr"] = round((r["CMP"] * r.get("AvgVol20", 0)) / 1e7, 2)
            if r["avg_daily_turnover_cr"] < v5.FUND_CONFIG["min_avg_daily_turnover_cr"]:
                r["gate_reasons"].append(f"Avg daily turnover too low (Rs {r['avg_daily_turnover_cr']:.1f}Cr)")
                if r["gate_pass"] is True:
                    r["gate_pass"] = False
            if r["Symbol"] not in swing_inv and r["gate_pass"] is True:
                r["gate_pass"] = False
                r["gate_reasons"].append("Not in Swing/Investment universe")

            lb = False
            fo = fo_data.get(r["Symbol"]) if fo_data else None
            if fo and r["Day Chg %"] >= v5.FUND_CONFIG["fno_long_buildup_price_chg_min"] and fo.get("prev_oi", 0) > 0:
                if (fo["today_oi"] - fo["prev_oi"]) / fo["prev_oi"] * 100 >= v5.FUND_CONFIG["fno_long_buildup_oi_chg_min"]:
                    lb = True
            r["long_buildup"] = lb
            r["tech_score"] = min(100, r["LONG SCORE"] + (v5.FUND_CONFIG["fno_long_buildup_bonus"] if lb else 0))

        eligible = [r for r in results if r.get("gate_pass") is True]
        tier_a = sorted(eligible, key=lambda r: r["tech_score"], reverse=True)[:TOP_N]
        for r in eligible:
            r["blended_score"] = 0.65 * (r.get("fund_score") or 0) + 0.35 * r["tech_score"]
        tier_b = sorted(eligible, key=lambda r: r["blended_score"], reverse=True)[:TOP_N]
        print(f"  {len(eligible)} passed gate")

    # ---- F&O picks: BOTH sides. fo_final_action() already produces short
    # signals (STRONG SHORT / SHORT) when the technical short score and the
    # OI/PCR flow are both bearish -- they just weren't being surfaced before.
    fno, fno_short = [], []
    for r in results:
        fo = fo_data.get(r["Symbol"]) if fo_data else None
        if not fo:
            continue
        a = v5.fo_final_action(r["LONG SCORE"], r["SHORT SCORE"], r["Day Chg %"],
                                fo.get("today_oi"), fo.get("prev_oi"), fo.get("pcr"))
        if a["final_action"] in ("🟢 STRONG BUY", "🔵 BUY", "🔻 STRONG SHORT", "🔽 SHORT"):
            r["fo_action"] = a["final_action"]
            r["fo_confidence"] = a["confidence"]
            r["fo_trend"] = a["fo_trend"]
            (fno if "BUY" in a["final_action"] else fno_short).append(r)
    # ONE combined ranking: longs and shorts compete on their own side's score,
    # so the top slots go to whichever setups are genuinely strongest today.
    for r in fno: r["fo_side"], r["fo_rank_score"] = "LONG", r["LONG SCORE"]
    for r in fno_short: r["fo_side"], r["fo_rank_score"] = "SHORT", r["SHORT SCORE"]
    fno_all = sorted(fno + fno_short, key=lambda r: -r["fo_rank_score"])[:TOP_N]
    fno = [r for r in fno_all if r["fo_side"] == "LONG"]
    fno_short = [r for r in fno_all if r["fo_side"] == "SHORT"]

    def build_why(r, kind):
        """Plain-English reasons, built from the actual data that drove the pick."""
        w = []
        if kind == "fno":
            if r.get("fo_trend"):
                w.append(f"F&O flow: {r['fo_trend'].replace('✅','').replace('❌','').strip().title()}")
            if r.get("long_buildup"):
                w.append("Long buildup — price up AND open interest rising together")
            if r.get("fo_confidence") and r["fo_confidence"] != "—":
                w.append(f"Confidence: {r['fo_confidence'].replace('⭐','').strip()}")
        if r.get("Pattern Long") and r["Pattern Long"] not in ("No Pattern", "—", None):
            w.append(f"Chart pattern: {r['Pattern Long']}")
        if r.get("EMA Trend"):
            w.append(f"Trend: {r['EMA Trend']}")
        rsi = r.get("RSI")
        if rsi is not None:
            if 40 <= rsi <= 65:
                w.append(f"RSI {rsi:.0f} — in the healthy entry zone")
            elif rsi > 70:
                w.append(f"RSI {rsi:.0f} — overbought, may pull back first")
            elif rsi < 40:
                w.append(f"RSI {rsi:.0f} — weak momentum")
        if r.get("Vol Status") in ("Strong", "Very Strong"):
            w.append(f"Volume: {r['Vol Status'].lower()} vs its 20-day average")
        if kind in ("a", "b"):
            if r.get("roe") is not None:
                w.append(f"ROE {r['roe']:.0f}% — passed the quality gate")
            if r.get("pe_ttm") is not None and r.get("pe_5y_median") is not None:
                if r["pe_ttm"] < r["pe_5y_median"]:
                    w.append(f"P/E {r['pe_ttm']:.0f} vs its own 5-year average of {r['pe_5y_median']:.0f} — cheaper than usual")
                else:
                    w.append(f"P/E {r['pe_ttm']:.0f} vs 5-year average {r['pe_5y_median']:.0f} — richer than usual")
            if r.get("market_cap_cr"):
                w.append(f"Market cap Rs {r['market_cap_cr']:,.0f} Cr")
            if r.get("avg_daily_turnover_cr"):
                w.append(f"Trades ~Rs {r['avg_daily_turnover_cr']:,.0f} Cr/day — liquid enough to exit")
        if kind == "b" and r.get("fund_score") is not None and r.get("tech_score") is not None:
            w.append(f"Ranked on fundamentals ({r['fund_score']:.0f}/100) weighted 65%, technicals ({r['tech_score']:.0f}/100) 35%")
        return w

    def pack(r, kind):
        d = {
            "symbol": r["Symbol"], "cmp": safe(r["CMP"]), "chg": safe(r["Day Chg %"]),
            "score": safe(r.get("tech_score", r["LONG SCORE"])),
            "signal": r["LONG SIGNAL"], "rsi": safe(r["RSI"]),
            "pattern": r.get("Pattern Long"),
            "sl": safe(r.get("SL Long")), "t1": safe(r.get("T1 Long")), "t2": safe(r.get("T2 Long")),
            "why": build_why(r, kind),
        }
        if kind == "fno":
            d.update({"action": r.get("fo_action"), "confidence": r.get("fo_confidence"),
                      "trend": r.get("fo_trend"), "buildup": bool(r.get("long_buildup"))})
        elif kind == "fno_short":
            # Short side: SL sits ABOVE entry, targets BELOW.
            d.update({"action": r.get("fo_action"), "confidence": r.get("fo_confidence"),
                      "trend": r.get("fo_trend"), "side": "short",
                      "score": safe(r["SHORT SCORE"]), "pattern": r.get("Pattern Short"),
                      "sl": safe(r.get("SL Short")), "t1": safe(r.get("T1 Short")),
                      "t2": safe(r.get("T2 Short"))})
        elif kind == "a":
            d.update({"roe": safe(r.get("roe")), "buildup": bool(r.get("long_buildup"))})
        else:
            d.update({"blended": safe(r.get("blended_score")), "fund": safe(r.get("fund_score")),
                      "roe": safe(r.get("roe")), "pe": safe(r.get("pe_ttm")),
                      "pe5y": safe(r.get("pe_5y_median"))})
        return d

    print("\n[5/5] Updating 30-day track record...")
    try:
        results_by_sym = {r["Symbol"]: r for r in results}
        track_summary, track_picks = update_track_record(fno_all, tier_a, tier_b, results_by_sym)
        closed_n = sum(s["closed"] for s in track_summary.values())
        open_n = sum(s["open"] for s in track_summary.values())
        print(f"  Track record: {closed_n} closed, {open_n} open positions")
    except Exception as e:
        print(f"  Track record update failed (non-fatal): {e}")
        track_summary, track_picks = {}, []
        notes.append("Track record could not be updated this run.")

    payload = {
        "generated_at": datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
        "universe_scanned": len(results),
        "fno_universe": len(fo_universe),
        "notes": notes,
        "fno": [pack(r, "fno") for r in fno],
        "fno_short": [pack(r, "fno_short") for r in fno_short],
        "tier_a": [pack(r, "a") for r in tier_a],
        "tier_b": [pack(r, "b") for r in tier_b],
        "track_summary": track_summary,
        "track_picks": track_picks[:120],
    }

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=1, ensure_ascii=False)
    print(f"\n✅ Wrote {OUT_PATH}")
    print(f"   F&O: {len(payload['fno'])} | Tier A: {len(payload['tier_a'])} | Tier B: {len(payload['tier_b'])}")
    for n in notes:
        print(f"   NOTE: {n}")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        sys.exit(1)
