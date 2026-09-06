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
from datetime import datetime

sys.argv = ["cloud_runner"]  # keep v5's arg handling quiet
import nse_momentum_screener as v5

OUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "docs", "data.json")
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

    # ---- F&O picks ----
    fno = []
    for r in results:
        fo = fo_data.get(r["Symbol"]) if fo_data else None
        if not fo:
            continue
        a = v5.fo_final_action(r["LONG SCORE"], r["SHORT SCORE"], r["Day Chg %"],
                                fo.get("today_oi"), fo.get("prev_oi"), fo.get("pcr"))
        if a["final_action"] in ("🟢 STRONG BUY", "🔵 BUY"):
            r["fo_action"] = a["final_action"]
            r["fo_confidence"] = a["confidence"]
            r["fo_trend"] = a["fo_trend"]
            fno.append(r)
    fno.sort(key=lambda r: (r["fo_action"] != "🟢 STRONG BUY", -r["LONG SCORE"]))
    fno = fno[:TOP_N]

    def pack(r, kind):
        d = {
            "symbol": r["Symbol"], "cmp": safe(r["CMP"]), "chg": safe(r["Day Chg %"]),
            "score": safe(r.get("tech_score", r["LONG SCORE"])),
            "signal": r["LONG SIGNAL"], "rsi": safe(r["RSI"]),
            "pattern": r.get("Pattern Long"),
            "sl": safe(r.get("SL Long")), "t1": safe(r.get("T1 Long")), "t2": safe(r.get("T2 Long")),
        }
        if kind == "fno":
            d.update({"action": r.get("fo_action"), "confidence": r.get("fo_confidence"),
                      "trend": r.get("fo_trend"), "buildup": bool(r.get("long_buildup"))})
        elif kind == "a":
            d.update({"roe": safe(r.get("roe")), "buildup": bool(r.get("long_buildup"))})
        else:
            d.update({"blended": safe(r.get("blended_score")), "fund": safe(r.get("fund_score")),
                      "roe": safe(r.get("roe")), "pe": safe(r.get("pe_ttm")),
                      "pe5y": safe(r.get("pe_5y_median"))})
        return d

    payload = {
        "generated_at": datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
        "universe_scanned": len(results),
        "fno_universe": len(fo_universe),
        "notes": notes,
        "fno": [pack(r, "fno") for r in fno],
        "tier_a": [pack(r, "a") for r in tier_a],
        "tier_b": [pack(r, "b") for r in tier_b],
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
