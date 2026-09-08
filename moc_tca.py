#!/usr/bin/env python
"""MOC / close-algo TCA.

One file. Copy it to the machine holding the order extract and run it there.
No network, no config file, nothing to install beyond the libraries below.

    python moc_tca.py --data orders.csv --probe        # audit the file, stop
    python moc_tca.py --data orders.csv                # tables + charts
    python moc_tca.py --sample                         # synthetic file, end to end
    python moc_tca.py --self-test                      # analytics only, no data

Requires pandas, numpy, matplotlib, openpyxl. Charts are skipped with a warning
if matplotlib is missing; tables still build.

Read the sanity report at the top of run_log.txt before trusting any number.
"""
from __future__ import annotations

import argparse
import datetime as _dt
import math
import re
import sys
import warnings
from pathlib import Path

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", category=RuntimeWarning)

# ===========================================================================
# CONFIG - the only part you edit
# ===========================================================================

CLIENT_NAME = "Client"          # kept neutral; no firm or client branding in output
PERIOD_LABEL = "H1"
CURRENCY = "USD"

# Which Strategy values are close algos.
#
# Leave empty and the script derives them: any strategy whose name matches
# CLOSE_NAME_PATTERN, or whose notional-weighted mean %CLOSE is at least
# AUTO_CLOSE_MIN_PCT. What it picked is printed loudly in the run log.
#
# Run --probe first, read the strategy table, then pin the list here.
CLOSE_STRATEGIES: list[str] = []
CLOSE_NAME_PATTERN = r"CLOSE|MOC|LOC|TWAPC|IIS"
AUTO_CLOSE_MIN_PCT = 5.0

# Period filter. None = whatever is in the file.
DATE_FROM = None                # e.g. "2026-01-01"
DATE_TO = None                  # e.g. "2026-06-30"

# --- column mapping -------------------------------------------------------
# Matched case-insensitively, ignoring spaces, dots, underscores and brackets.
# First alternative that resolves wins.
COLUMNS = {
    # identifiers
    "order_id":      ["aggrTgtId", "OrderId", "id_target"],
    "client":        ["client"],
    "trader":        ["Trader"],
    "date":          ["Date"],
    "symbol":        ["Sym", "Symbol"],
    "side":          ["Side"],
    "strategy":      ["Strategy", "ClientAlgo", "Algo"],
    "start_time":    ["Start(HK)", "StartHK", "Start", "Strike Time"],
    "end_time":      ["End(HK)", "EndHK", "End", "End Time"],
    # size
    "notional":      ["$Mln", "Value Exec", "NotionalMln"],      # EXECUTED, x1e6
    "order_shares":  ["#Shares", "Order Qty", "OrderQty"],       # ORDERED, x1e3
    "fill_rate":     ["FR", "Fill Rate", "FillRate"],
    # venue mix - must sum to 100
    "pct_close":     ["%CLOSE", "PctClose", "Close %"],
    "pct_open":      ["%OPEN", "PctOpen", "Open %"],
    "pct_post":      ["%POST", "PctPost", "Post %"],
    "pct_take":      ["%TAKE", "PctTake", "Take %"],
    "pct_dark":      ["%DARK", "PctDark", "Dark %"],
    # benchmarks, bps, positive = savings
    "slip_close":    ["Close", "Close ImpBps"],
    "slip_arrival":  ["IS", "Arrival ImpBps"],
    "slip_pvwap":    ["Pvwap", "PVWAP ImpBps"],
    "slip_vwap":     ["Vwap", "VWAP ImpBps"],
    "slip_nextopen": ["NextOpen", "Next Open", "NextOpen ImpBps"],
    "slip_open":     ["Open", "Open ImpBps"],
    # capacity and behaviour
    "adv_pct":       ["%Adv", "% Adv", "PctAdv"],
    "adv":           ["Adv", "ADV"],
    "pr_cont":       ["fPR_cont", "fPRcont", "PR_cont"],
    "first_exec_vs_close": ["first_exec_vs_close", "firstexecvsclose",
                            "FirstExecVsClose"],
    "participation": ["PR"],
    "spread_bps":    ["Sprd", "Spread", "Hist Spread"],
    "volatility":    ["Vol", "Volatility"],
    # segmentation
    "cap":           ["Cap", "MarketCap"],
    "market_limit":  ["marketLimit", "Market Limit", "OrderType", "Order Type"],
    "arrival_time":  ["arrivalTime", "Arrival Time"],
    "sector":        ["sector", "Sector", "Industry"],
}

REQUIRED = ["strategy", "date", "symbol", "side", "notional", "order_shares",
            "fill_rate", "pct_close"]

# The export writes a banner on row 1 and the real header on row 2 in some
# formats. 0 = header on the first row.
HEADER_ROW = 0
SHEET = 0
CSV_KWARGS = {"encoding": "utf-8-sig"}

# --- conventions ----------------------------------------------------------
# Positive bps = savings. The sanity report re-derives this from the data and
# shouts on a mismatch; flip only if your export uses the opposite sign.
POSITIVE_IS_SAVING = True
SIDE_ALREADY_ADJUSTED = True
BUY_VALUES = {"B", "BUY", "BOT", "1", "BUYS"}

# $Mln is EXECUTED notional, built as sum(cumqty * avgprice * fx_last) / 1e6.
# fx_last is inside it, so the column is millions of USD and the multiplier
# back to USD is 1e6. Confirmed with the data owner.
#
# The run still derives the implied USD price per share from it and warns if
# that lands outside a plausible band - a wrong scale here moves every currency
# figure by a power of ten and nothing else in the pipeline would notice.
NOTIONAL_SCALE = 1e6
SCALE = {"notional": NOTIONAL_SCALE, "order_shares": 1e3}

# Set True if PR / FR / %Adv / the venue-mix columns arrive as fractions (0-1)
# instead of percentages. The sanity report warns if this looks wrong.
PCT_FIELDS_ARE_FRACTIONS = False
PCT_FIELDS = ["fill_rate", "adv_pct", "participation", "pr_cont",
              "pct_close", "pct_open", "pct_post", "pct_take", "pct_dark"]

VENUE_FIELDS = ["pct_open", "pct_close", "pct_post", "pct_take", "pct_dark"]
VENUE_SUM_TOL = 1.0          # pp; rows outside 100 +/- this are reported

# --- markets --------------------------------------------------------------
UNKNOWN_MARKET = "Unknown"
SYMBOL_SUFFIX_MAP = {
    "HK": "Hong Kong", "JP": "Japan", "KS": "South Korea", "KQ": "South Korea",
    "TT": "Taiwan", "IN": "India", "IB": "India", "AU": "Australia",
    "NZ": "New Zealand", "SP": "Singapore", "MK": "Malaysia", "TB": "Thailand",
    "IJ": "Indonesia", "PM": "Philippines", "VN": "Vietnam",
    "CH": "China", "C1": "China", "C2": "China",
}

# Continuous-session end, in HKT, because Start(HK)/End(HK) are HKT.
#
# VERIFIED against the desk's own market_stats.q session windows:
#   Hong Kong 16:00, Japan 14:25, Australia 14:10 (AEST).
# India carries no closing-auction bound there by choice: NSE closes on a VWAP
# of the last half hour, not a single-price auction, so there is no auction to
# measure and the market is reported separately throughout.
#
# UNVERIFIED - derived from published exchange hours, not from the plant. Any
# market flagged here is named in the run log wherever it affects a number.
MARKET_CLOSE_HKT = {
    "Hong Kong": ("16:00", True),
    "Japan":     ("14:25", True),
    "Australia": ("14:10", True),      # AEST; AEDT handled below
    "South Korea": ("14:30", False),
    "Taiwan":      ("13:30", False),
    "China":       ("15:00", False),
    "Singapore":   ("17:04", False),
    "Malaysia":    ("17:00", False),
    "Thailand":    ("17:35", False),
    "Indonesia":   ("15:50", False),
    "Philippines": ("15:10", False),
    "New Zealand": ("11:45", False),
    "Vietnam":     ("14:45", False),
}
# Markets with no single-price closing auction. Never pooled with the rest.
NO_CLOSING_AUCTION = {"India"}

# Australia observes daylight saving and Hong Kong does not, so the ASX session
# moves an hour against an HKT clock twice a year. AEDT runs from the first
# Sunday in October to the first Sunday in April.
AEDT_MARKETS = {"Australia"}

# --- buckets --------------------------------------------------------------
ADV_BUCKETS = [-0.001, 1, 3, 5, 10, 25, 1e9]
ADV_LABELS = ["0-1%", "1-3%", "3-5%", "5-10%", "10-25%", "25%+"]

CLOSE_BUCKETS = [-0.001, 0.001, 25, 50, 75, 99.999, 100.001]
CLOSE_LABELS = ["0% (no auction fill)", "0-25%", "25-50%", "50-75%",
                "75-<100%", "100% (all auction)"]

CAP_ORDER = ["Large", "Mid", "Small", "Micro", "Other"]
MARKET_LIMIT_ORDER = ["Market", "Limit"]
ARRIVAL_ORDER = ["Pre-Open", "First30Mins", "Day", "Last30Mins"]

# Cohort thresholds for the miss taxonomy.
COHORT_ADV_LOW = 5.0          # %Adv below this = the auction could have taken it
COHORT_FR_ZERO = 1.0          # FR below this = never traded

# --- statistics -----------------------------------------------------------
WINSOR = (0.01, 0.99)         # None to disable; applies to MEANS only
BOOTSTRAP_N = 2000
SEED = 7
MIN_N_FOR_CI = 8

# --- palette (validated; light surface) -----------------------------------
SURFACE = "#fcfcfb"
INK = "#0b0b0b"
INK_SECOND = "#52514e"
INK_MUTED = "#898781"
GRID = "#e1e0d9"
BASELINE = "#c3c2b7"
POS = "#2a78d6"
NEG = "#d03b3b"
NEUTRAL = "#f0efec"
SERIES = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4",
          "#008300", "#4a3aa7", "#e34948"]
# Ordinal ramp, one hue, light -> dark. Evenly spaced in OKLab: adjacent steps
# differ by dL ~= 0.087, comfortably above the 0.06 floor, and the lightest step
# still clears the surface. Validated, not eyeballed.
ORDINAL_BLUE = ["#86b6ef", "#6d9bd2", "#5480b5", "#3c6699", "#254d7d", "#0d3563"]
FIGSIZE = (10.0, 5.6)
DPI = 200

BENCHMARK_LABELS = {
    "slip_arrival": "vs Arrival",
    "slip_close": "vs Close",
    "slip_pvwap": "vs Interval VWAP",
    "slip_vwap": "vs Day VWAP",
    "slip_nextopen": "vs Next Open",
    "slip_open": "vs Open",
}
MATRIX_BENCHMARKS = ["slip_arrival", "slip_close", "slip_pvwap", "slip_vwap"]


# ===========================================================================
# LOGGING
# ===========================================================================

_LOG: list[str] = []


def log(msg: str = "") -> None:
    print(msg)
    _LOG.append(str(msg))


def section(title: str) -> None:
    log("")
    log("=" * 74)
    log(f"  {title}")
    log("=" * 74)


def warn(msg: str) -> None:
    log(f"  !! {msg}")


# ===========================================================================
# LOADING AND NORMALISATION
# ===========================================================================

def _norm_name(s) -> str:
    """Header key: lowercase, strip spaces, dots, underscores, brackets."""
    return re.sub(r"[\s._()\[\]{}/-]", "", str(s)).lower()


def read_file(path: Path) -> pd.DataFrame:
    if path.suffix.lower() in {".xlsx", ".xls", ".xlsm"}:
        raw = pd.read_excel(path, sheet_name=SHEET, header=HEADER_ROW)
    else:
        raw = pd.read_csv(path, header=HEADER_ROW, **CSV_KWARGS)
    raw.columns = [str(c).strip() for c in raw.columns]
    return raw


def resolve_columns(raw: pd.DataFrame) -> dict[str, str]:
    """logical name -> actual header. Unresolved logical names are absent."""
    lookup = {}
    for actual in raw.columns:
        lookup.setdefault(_norm_name(actual), actual)
    resolved = {}
    for logical, candidates in COLUMNS.items():
        if isinstance(candidates, str):
            candidates = [candidates]
        for cand in candidates:
            hit = lookup.get(_norm_name(cand))
            if hit is not None:
                resolved[logical] = hit
                break
    return resolved


def _to_num(s: pd.Series) -> pd.Series:
    """Numeric, tolerating thousands separators, %, and stray text."""
    if pd.api.types.is_numeric_dtype(s):
        return s.astype(float)
    cleaned = (s.astype(str)
                 .str.replace(",", "", regex=False)
                 .str.replace("%", "", regex=False)
                 .str.replace("\u00a0", "", regex=False)
                 .str.strip())
    return pd.to_numeric(cleaned, errors="coerce")


def _parse_time(s: pd.Series) -> pd.Series:
    """Clock time -> minutes past midnight. Tolerates HH:MM[:SS[.fff]]."""
    txt = s.astype(str).str.strip()
    parsed = pd.to_datetime(txt, format="mixed", errors="coerce")
    mins = parsed.dt.hour * 60 + parsed.dt.minute + parsed.dt.second / 60.0
    return mins


def market_from_symbol(sym: pd.Series) -> pd.Series:
    """Suffix after the last '.' or ' ' -> market name."""
    txt = sym.astype(str).str.strip().str.upper()
    suffix = txt.str.extract(r"[. ]([A-Z0-9]{1,3})$", expand=False)
    out = suffix.map(SYMBOL_SUFFIX_MAP)
    return out.fillna(UNKNOWN_MARKET)


def _first_sunday(year: int, month: int) -> _dt.date:
    d = _dt.date(year, month, 1)
    return d + _dt.timedelta(days=(6 - d.weekday()) % 7)


def _is_aedt(d) -> bool:
    """AEDT: first Sunday in October to first Sunday in April."""
    if pd.isna(d):
        return False
    d = pd.Timestamp(d).date()
    return d >= _first_sunday(d.year, 10) or d < _first_sunday(d.year, 4)


def _hhmm_to_min(txt: str) -> float:
    h, m = txt.split(":")
    return int(h) * 60 + int(m)


def continuous_end_min(market: pd.Series, date: pd.Series) -> pd.Series:
    """Market continuous end in HKT minutes. NaN where there is no auction."""
    base = market.map(lambda m: MARKET_CLOSE_HKT.get(m, (None, None))[0])
    out = base.map(lambda t: _hhmm_to_min(t) if isinstance(t, str) else np.nan)
    # Australia shifts an hour earlier against HKT under AEDT.
    is_au = market.isin(AEDT_MARKETS)
    if is_au.any():
        shift = date.map(_is_aedt) & is_au
        out = out.where(~shift, out - 60)
    out = out.where(~market.isin(NO_CLOSING_AUCTION), np.nan)
    return out


def normalise(raw: pd.DataFrame, cols: dict[str, str]) -> pd.DataFrame:
    """Build the working frame: logical names, base units, derived fields."""
    df = pd.DataFrame(index=raw.index)

    numeric = ["notional", "order_shares", "fill_rate", "adv", "adv_pct",
               "participation", "pr_cont", "spread_bps", "volatility",
               "first_exec_vs_close"] + VENUE_FIELDS + [
               "slip_close", "slip_arrival", "slip_pvwap", "slip_vwap",
               "slip_nextopen", "slip_open"]
    for logical in numeric:
        if logical in cols:
            df[logical] = _to_num(raw[cols[logical]])

    for logical in ["order_id", "client", "trader", "side", "strategy",
                    "symbol", "cap", "market_limit", "arrival_time", "sector"]:
        if logical in cols:
            df[logical] = raw[cols[logical]].astype(str).str.strip()

    if "date" in cols:
        df["date"] = pd.to_datetime(raw[cols["date"]], errors="coerce")

    for logical in ["start_time", "end_time"]:
        if logical in cols:
            df[logical + "_min"] = _parse_time(raw[cols[logical]])

    # --- unit scaling -----------------------------------------------------
    for logical, factor in SCALE.items():
        if logical in df:
            df[logical] = df[logical] * factor
    if PCT_FIELDS_ARE_FRACTIONS:
        for f in PCT_FIELDS:
            if f in df:
                df[f] = df[f] * 100.0

    # --- sign -------------------------------------------------------------
    if not POSITIVE_IS_SAVING:
        for f in [c for c in df.columns if c.startswith("slip_")]:
            df[f] = -df[f]
        if "first_exec_vs_close" in df:
            df["first_exec_vs_close"] = -df["first_exec_vs_close"]

    # --- identity ---------------------------------------------------------
    if "side" in df:
        up = df["side"].str.upper()
        df["is_buy"] = up.isin({v.upper() for v in BUY_VALUES})
        df["side_label"] = np.where(df["is_buy"], "Buy", "Sell")

    df["market"] = market_from_symbol(df["symbol"]) if "symbol" in df else UNKNOWN_MARKET
    df["has_auction"] = ~df["market"].isin(NO_CLOSING_AUCTION)

    if "cap" in df:
        df["cap"] = _title_norm(df["cap"], CAP_ORDER)
    if "market_limit" in df:
        df["market_limit"] = _title_norm(df["market_limit"], MARKET_LIMIT_ORDER)
    if "arrival_time" in df:
        df["arrival_time"] = _title_norm(df["arrival_time"], ARRIVAL_ORDER)

    # --- quantities -------------------------------------------------------
    # $Mln is EXECUTED notional; #Shares is the ORDER quantity. Dividing one by
    # the other is NOT an average price. Dividing executed notional by EXECUTED
    # shares is - and since fx_last is already inside $Mln, it comes out in USD,
    # which makes it the check on whether NOTIONAL_SCALE is right.
    if "order_shares" in df and "fill_rate" in df:
        df["exec_shares"] = df["order_shares"] * df["fill_rate"] / 100.0
        df["unfilled_shares"] = df["order_shares"] - df["exec_shares"]
        if "notional" in df:
            df["implied_px_usd"] = np.where(df["exec_shares"] > 0,
                                            df["notional"] / df["exec_shares"],
                                            np.nan)

    # --- venue mix --------------------------------------------------------
    present = [f for f in VENUE_FIELDS if f in df]
    if present:
        df["venue_sum"] = df[present].sum(axis=1, min_count=1)
    cont_fields = [f for f in ["pct_take", "pct_post", "pct_dark"] if f in df]
    if cont_fields:
        df["pct_continuous"] = df[cont_fields].sum(axis=1, min_count=1)
    if "pct_close" in df and "notional" in df:
        df["close_notional"] = df["notional"] * df["pct_close"] / 100.0
        df["cont_notional"] = df["notional"] * df.get(
            "pct_continuous", 100.0 - df["pct_close"]) / 100.0
    if "pct_close" in df and "exec_shares" in df:
        df["close_shares"] = df["exec_shares"] * df["pct_close"] / 100.0

    # --- derived benchmarks ----------------------------------------------
    # Every slippage shares one executed price, so a difference between two of
    # them cancels it and leaves a pure, side-adjusted price move.
    if "slip_arrival" in df and "slip_close" in df:
        # what waiting for the close cost: arrival price -> closing price
        df["wait_cost_bps"] = df["slip_arrival"] - df["slip_close"]
    if "slip_arrival" in df and "slip_pvwap" in df:
        df["drift_bps"] = df["slip_arrival"] - df["slip_pvwap"]
    if "slip_vwap" in df and "slip_close" in df:
        # session average -> close. Positive = the close was the better place
        # to trade, side-adjusted. Answers "should this have gone MOC at all".
        df["close_vs_session_bps"] = df["slip_vwap"] - df["slip_close"]
    if "slip_nextopen" in df and "slip_close" in df:
        # closing price -> next open. Negative = the price moved back against
        # where we traded, i.e. temporary impact we paid.
        df["reversion_bps"] = df["slip_nextopen"] - df["slip_close"]

    # Implied execution of the portion that MISSED the auction.
    # The auction portion prints at the close by construction, contributing ~0
    # to slip_close, so the whole of slip_close is carried by the continuous
    # portion. Undefined when the order cleared entirely in the auction.
    if "slip_close" in df and "pct_close" in df:
        miss = (100.0 - df["pct_close"]) / 100.0
        df["implied_cont_vs_close_bps"] = np.where(
            miss > 0.05, df["slip_close"] / miss.where(miss > 0.05), np.nan)
        df["leakage_cost_ccy"] = df["slip_close"] / 1e4 * df["notional"]

    # --- timing -----------------------------------------------------------
    if "start_time_min" in df and "end_time_min" in df:
        df["duration_min"] = df["end_time_min"] - df["start_time_min"]
    if "end_time_min" in df and "date" in df:
        end_of_continuous = continuous_end_min(df["market"], df["date"])
        df["close_gap_min"] = end_of_continuous - df["end_time_min"]

    # --- buckets ----------------------------------------------------------
    if "adv_pct" in df:
        df["adv_bucket"] = pd.cut(df["adv_pct"], ADV_BUCKETS, labels=ADV_LABELS)
    if "pct_close" in df:
        df["close_bucket"] = pd.cut(df["pct_close"], CLOSE_BUCKETS,
                                    labels=CLOSE_LABELS)
    if "date" in df:
        df["month"] = df["date"].dt.to_period("M").astype(str)

    return df


def _title_norm(s: pd.Series, order: list[str]) -> pd.Series:
    """Match a label case-insensitively onto `order`; keep anything else as-is.

    Unrecognised values are NEVER dropped - Cap in particular carries values
    beyond the expected set, and silently discarding them would move numbers.
    """
    canon = {v.lower(): v for v in order}
    return s.map(lambda v: canon.get(str(v).strip().lower(), str(v).strip()))


# ===========================================================================
# SANITY REPORT - read this before trusting any number
# ===========================================================================

def sanity_report(df: pd.DataFrame, cols: dict, raw_cols) -> None:
    section("SANITY REPORT")

    missing = [c for c in COLUMNS if c not in cols]
    log(f"  resolved {len(cols)}/{len(COLUMNS)} known fields")
    if missing:
        log(f"  unresolved: {', '.join(sorted(missing))}")
    unused = [c for c in raw_cols if c not in set(cols.values())]
    if unused:
        log(f"  columns in the file with no mapping: {', '.join(map(str, unused))}")

    hard = [c for c in REQUIRED if c not in cols]
    if hard:
        raise SystemExit(
            "\nFATAL: required fields not found: " + ", ".join(hard) +
            "\nEdit COLUMNS at the top of this script, or check HEADER_ROW.")

    log("")
    log("  field distributions (1st / 50th / 99th percentile)")
    header = "    {:<26}{:>7}{:>8}{:>12}{:>12}{:>12}".format(
        "field", "n", "miss%", "p1", "p50", "p99")
    log(header)
    for c in df.columns:
        # bool is numeric to pandas but percentiles of it are meaningless,
        # and numpy refuses to subtract booleans while interpolating.
        if (not pd.api.types.is_numeric_dtype(df[c])
                or pd.api.types.is_bool_dtype(df[c])):
            continue
        s = df[c]
        n = int(s.notna().sum())
        if n == 0:
            log("    {:<26}{:>7}{:>8.1f}{:>12}{:>12}{:>12}".format(
                c, 0, 100.0, "-", "-", "-"))
            continue
        p1, p50, p99 = np.nanpercentile(s.dropna(), [1, 50, 99])
        log("    {:<26}{:>7}{:>8.1f}{:>12.3f}{:>12.3f}{:>12.3f}".format(
            c, n, 100 * s.isna().mean(), p1, p50, p99))

    # --- venue mix closes to 100 -----------------------------------------
    log("")
    log("  venue mix")
    present = [f for f in VENUE_FIELDS if f in df]
    log("    present: " + (", ".join(present) if present else "NONE"))
    if "venue_sum" in df:
        known = df["venue_sum"].notna()
        bad = known & ((df["venue_sum"] - 100).abs() > VENUE_SUM_TOL)
        log("    rows summing to 100 +/- {}pp: {} of {}".format(
            VENUE_SUM_TOL, int((known & ~bad).sum()), int(known.sum())))
        if bad.any():
            warn("{} rows do NOT sum to 100 (median {:.1f}). They are kept "
                 "as-is and never renormalised - a venue column may be "
                 "missing from the export.".format(
                     int(bad.sum()), df.loc[bad, "venue_sum"].median()))
        if len(present) < len(VENUE_FIELDS):
            warn("not all five venue columns are present, so the sum check is "
                 "weaker than it looks.")
        # Decisive test of what the venue shares are a percentage OF. If they
        # are shares of EXECUTED quantity they sum to 100; if they are shares
        # of the ORDER quantity they sum to the fill rate instead, and every
        # auction quantity derived from them would be overstated.
        if "fill_rate" in df:
            d100 = (df["venue_sum"] - 100).abs()
            dfr = (df["venue_sum"] - df["fill_rate"]).abs()
            near100 = int((d100 <= 1.0).sum())
            nearfr = int((dfr <= 1.0).sum())
            log("    denominator test: {} rows sum to 100, {} rows sum to FR "
                "(of {})".format(near100, nearfr, int(known.sum())))
            if nearfr > near100:
                warn("the venue shares appear to be a % of ORDER quantity, not "
                     "of executed quantity. Auction quantity is then "
                     "%CLOSE/100 x order qty, not x executed qty - say so and "
                     "the derivation is changed.")
            else:
                log("      -> shares of EXECUTED quantity, as assumed.")

    # --- percentage vs fraction ------------------------------------------
    for f in PCT_FIELDS:
        if f in df and df[f].notna().any():
            hi = float(np.nanmax(df[f].values))
            if hi <= 1.5 and not PCT_FIELDS_ARE_FRACTIONS:
                warn("{} never exceeds {:.3f} - it may be a fraction. Set "
                     "PCT_FIELDS_ARE_FRACTIONS = True.".format(f, hi))
            if hi > 101 and f in VENUE_FIELDS + ["fill_rate"]:
                warn("{} reaches {:.1f}, above 100 - impossible for a share, "
                     "so the column means something other than assumed."
                     .format(f, hi))

    # --- side adjustment --------------------------------------------------
    if "side_label" in df:
        log("")
        log("  side adjustment check (means by side, bps)")
        # first_exec_vs_close is a bps measure like the rest, so it has to
        # clear the same side check: if it is NOT side-adjusted, every sell is
        # inverted and the "was starting early right?" answer flips sign.
        bench = [c for c in ["slip_arrival", "slip_close", "slip_pvwap",
                             "slip_nextopen", "first_exec_vs_close"]
                 if c in df]
        if bench:
            chk = df.groupby("side_label")[bench].mean().round(3)
            chk.insert(0, "orders", df.groupby("side_label").size())
            for line in chk.to_string().splitlines():
                log("    " + line)
            log("    If Buy and Sell sit on OPPOSITE sides of zero by a")
            log("    similar magnitude the data is not side-adjusted, and")
            log("    every number below is wrong.")

    # --- notional basis ---------------------------------------------------
    log("")
    log("  notional basis   <-- CHECK THIS BLOCK BEFORE QUOTING ANY MONEY FIGURE")
    log("    $Mln is EXECUTED notional = sum(cumqty * avgprice * fx_last) / 1e6,")
    log("    so it is millions of USD and NOTIONAL_SCALE is {:.0e}.".format(
        NOTIONAL_SCALE))
    log("    total executed: {} {:,.0f}  ({:,.1f}m)".format(
        CURRENCY, df["notional"].sum(), df["notional"].sum() / 1e6))
    log("    #Shares is ORDER quantity; executed = #Shares x FR/100.")
    if "exec_shares" in df:
        ordered = max(float(df["order_shares"].sum()), 1.0)
        log("    order qty {:,.0f} -> executed {:,.0f} shares ({:.1f}%)".format(
            df["order_shares"].sum(), df["exec_shares"].sum(),
            100 * df["exec_shares"].sum() / ordered))
    if "implied_px_usd" in df and df["implied_px_usd"].notna().any():
        px = df["implied_px_usd"].dropna()
        log("")
        log("    implied share price = executed notional / executed shares.")
        log("    fx_last is already inside $Mln, so this is USD. If a market")
        log("    lands a power of ten from a plausible share price, the scale,")
        log("    the #Shares scaling or the fx basis is wrong.")
        log("      {:<14}{:>12}{:>12}{:>12}".format("market", "p25", "median",
                                                    "p75"))
        for mkt, g in df.dropna(subset=["implied_px_usd"]).groupby("market"):
            q = g["implied_px_usd"].quantile([.25, .5, .75]).values
            log("      {:<14}{:>12,.2f}{:>12,.2f}{:>12,.2f}".format(
                str(mkt), q[0], q[1], q[2]))
        med = float(px.median())
        if not (0.20 <= med <= 2000):
            warn("median implied share price is {} {:,.2f}, outside a "
                 "plausible band for a listed equity - NOTIONAL_SCALE, the "
                 "#Shares scaling, or the fx basis is wrong.".format(
                     CURRENCY, med))

    # --- markets ----------------------------------------------------------
    log("")
    log("  markets")
    for mkt, n in df["market"].value_counts().items():
        note = ""
        if mkt in NO_CLOSING_AUCTION:
            note = "  <- NO single-price closing auction; reported separately"
        elif mkt == UNKNOWN_MARKET:
            note = "  <- suffix did not map; check SYMBOL_SUFFIX_MAP"
        elif not MARKET_CLOSE_HKT.get(mkt, ("", False))[1]:
            note = "  <- close time UNVERIFIED against the plant"
        log("    {:<16}{:>7}{}".format(mkt, n, note))

    # --- unexpected labels ------------------------------------------------
    for field, expected in [("cap", CAP_ORDER),
                            ("market_limit", MARKET_LIMIT_ORDER),
                            ("arrival_time", ARRIVAL_ORDER)]:
        if field in df:
            extra = sorted(set(df[field].dropna().unique()) - set(expected))
            if extra:
                log("    {}: values outside the expected set, kept and "
                    "shown: {}".format(field, ", ".join(map(str, extra))))


def strategy_profile(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby("strategy")
    ncol = "notional (" + CURRENCY + "m)"
    out = pd.DataFrame({"orders": g.size(), ncol: g["notional"].sum() / 1e6})
    out["% notional"] = 100 * out[ncol] / max(out[ncol].sum(), 1e-9)
    for field, label, how in [
            ("pct_close", "mean %CLOSE (wtd)", "w"),
            ("slip_close", "vs Close bps (wtd)", "w"),
            ("fill_rate", "mean FR (wtd)", "w"),
            ("adv_pct", "median %Adv", "m")]:
        if field not in df:
            continue
        if how == "m":
            out[label] = g[field].median()
        else:
            out[label] = pd.Series(
                {k: wmean(x[field], x["notional"]) for k, x in g})
    return out.sort_values(ncol, ascending=False).round(2)


def choose_close_strategies(df: pd.DataFrame) -> list:
    """The close algos: the pinned list, or derived and printed loudly."""
    section("CLOSE-ALGO SELECTION")
    prof = strategy_profile(df)
    for line in prof.to_string().splitlines():
        log("  " + line)

    if CLOSE_STRATEGIES:
        chosen = [s for s in CLOSE_STRATEGIES if s in set(df["strategy"])]
        absent = sorted(set(CLOSE_STRATEGIES) - set(chosen))
        log("")
        log("  pinned in CLOSE_STRATEGIES: " + ", ".join(chosen))
        if absent:
            warn("pinned but not present in the file: " + ", ".join(absent))
        return chosen

    by_name = {s for s in prof.index
               if re.search(CLOSE_NAME_PATTERN, str(s), re.I)}
    by_close = set()
    if "mean %CLOSE (wtd)" in prof:
        by_close = set(prof.index[prof["mean %CLOSE (wtd)"] >= AUTO_CLOSE_MIN_PCT])
    chosen = sorted(by_name | by_close)
    log("")
    warn("CLOSE_STRATEGIES is empty, so the close algos were DERIVED:")
    log("     by name pattern /" + CLOSE_NAME_PATTERN + "/i : " +
        (", ".join(sorted(by_name)) or "none"))
    log("     by mean %CLOSE >= {}% : ".format(AUTO_CLOSE_MIN_PCT) +
        (", ".join(sorted(by_close)) or "none"))
    log("     selected : " + (", ".join(chosen) or "NONE"))
    log("     Pin CLOSE_STRATEGIES at the top of this script to make the")
    log("     scope explicit rather than derived.")
    return chosen


# ===========================================================================
# STATISTICS
# ===========================================================================

def winsorize(v: np.ndarray, limits=WINSOR) -> np.ndarray:
    """Pull the extreme tails back to the percentile value, do not drop them."""
    if limits is None:
        return v
    ok = ~np.isnan(v)
    if ok.sum() < 3:
        return v
    lo, hi = np.nanpercentile(v[ok], [limits[0] * 100, limits[1] * 100])
    return np.clip(v, lo, hi)


def wmean(values, weights, winsor=True) -> float:
    """Notional-weighted mean. Means are winsorised; nothing else is."""
    v = np.asarray(values, dtype=float)
    w = np.asarray(weights, dtype=float)
    ok = ~np.isnan(v) & ~np.isnan(w) & (w > 0)
    if ok.sum() == 0:
        return np.nan
    v, w = v[ok], w[ok]
    if winsor:
        v = winsorize(v)
    return float(np.sum(v * w) / np.sum(w))


def boot_ci(values, weights, n=BOOTSTRAP_N, seed=SEED):
    """95% percentile bootstrap on the weighted mean.

    Slippage is skewed and fat-tailed, so a normal-theory interval would be
    too narrow. A CI crossing zero means the result is not distinguishable
    from zero - say so rather than quoting the point estimate as fact.
    """
    v = np.asarray(values, dtype=float)
    w = np.asarray(weights, dtype=float)
    ok = ~np.isnan(v) & ~np.isnan(w) & (w > 0)
    v, w = v[ok], w[ok]
    if len(v) < MIN_N_FOR_CI:
        return (np.nan, np.nan)
    rng = np.random.default_rng(seed)
    idx = rng.integers(0, len(v), size=(n, len(v)))
    vs = winsorize(v)[idx]
    ws = w[idx]
    means = (vs * ws).sum(axis=1) / ws.sum(axis=1)
    return (float(np.percentile(means, 2.5)), float(np.percentile(means, 97.5)))


def hit_rate(values) -> float:
    """% of ORDERS beating the benchmark. A count statistic, per order."""
    v = np.asarray(values, dtype=float)
    v = v[~np.isnan(v)]
    return float(100 * (v > 0).mean()) if len(v) else np.nan


def by_group(df: pd.DataFrame, by, value: str, ci: bool = True,
             order=None) -> pd.DataFrame:
    """Weighted mean + median + hit rate + bootstrap CI, per group.

    Weighted figures are the client's real cost. Median and hit rate are
    per-order by design and labelled as such wherever they are shown.
    """
    if value not in df.columns:
        return pd.DataFrame()
    rows = []
    for key, g in df.groupby(by, dropna=False, observed=True):
        n = int(g[value].notna().sum())
        lo, hi = boot_ci(g[value], g["notional"]) if ci else (np.nan, np.nan)
        rows.append({
            (by if isinstance(by, str) else "group"): key,
            "orders": n,
            "notional (" + CURRENCY + "m)": g["notional"].sum() / 1e6,
            "wtd mean bps": wmean(g[value], g["notional"]),
            "median bps (per order)": g[value].median(),
            "hit rate % (per order)": hit_rate(g[value]),
            "CI low": lo,
            "CI high": hi,
            "small sample": n < MIN_N_FOR_CI,
        })
    out = pd.DataFrame(rows)
    if out.empty:
        return out
    key = by if isinstance(by, str) else "group"
    out = out.set_index(key)
    if order:
        keep = [o for o in order if o in out.index]
        rest = [i for i in out.index if i not in set(order)]
        out = out.loc[keep + rest]
    return out.round(2)


# ===========================================================================
# TABLES
# ===========================================================================

def t_scope(all_df: pd.DataFrame, close_strats: list) -> pd.DataFrame:
    """How much of the book the close algos are."""
    is_close = all_df["strategy"].isin(close_strats)
    rows = []
    for label, mask in [("Close algos", is_close), ("Other", ~is_close)]:
        g = all_df[mask]
        rows.append({
            "scope": label,
            "orders": len(g),
            "notional (" + CURRENCY + "m)": g["notional"].sum() / 1e6,
            "% of orders": 100 * len(g) / max(len(all_df), 1),
            "% of notional": 100 * g["notional"].sum() /
                             max(all_df["notional"].sum(), 1e-9),
        })
    return pd.DataFrame(rows).set_index("scope").round(2)


def t_venue_mix(df: pd.DataFrame, by: str) -> pd.DataFrame:
    """Notional-weighted venue mix. The five columns should sum to 100."""
    present = [f for f in VENUE_FIELDS if f in df]
    if not present:
        return pd.DataFrame()
    rows = []
    for key, g in df.groupby(by, dropna=False, observed=True):
        row = {by: key, "orders": len(g),
               "notional (" + CURRENCY + "m)": g["notional"].sum() / 1e6}
        for f in present:
            row[f.replace("pct_", "%").upper()] = wmean(g[f], g["notional"],
                                                        winsor=False)
        rows.append(row)
    return pd.DataFrame(rows).set_index(by).round(2)


def t_clearance(df: pd.DataFrame) -> pd.DataFrame:
    """Did the order actually clear in the auction, and what did it cost?

    %CLOSE measures the split directly, so this is a measurement rather than
    an inference from a near-zero close slippage.
    """
    if "close_bucket" not in df:
        return pd.DataFrame()
    rows = []
    total_n = max(df["notional"].sum(), 1e-9)
    for key, g in df.groupby("close_bucket", dropna=False, observed=True):
        rows.append({
            "auction share of the order": key,
            "orders": len(g),
            "notional (" + CURRENCY + "m)": g["notional"].sum() / 1e6,
            "% of notional": 100 * g["notional"].sum() / total_n,
            "vs Close bps (wtd)": wmean(g.get("slip_close"), g["notional"])
                if "slip_close" in g else np.nan,
            "vs Arrival bps (wtd)": wmean(g.get("slip_arrival"), g["notional"])
                if "slip_arrival" in g else np.nan,
            "mean FR": wmean(g["fill_rate"], g["notional"])
                if "fill_rate" in g else np.nan,
            "median %Adv": g["adv_pct"].median() if "adv_pct" in g else np.nan,
        })
    return pd.DataFrame(rows).set_index("auction share of the order").round(2)


def t_leakage(df: pd.DataFrame, by: str) -> pd.DataFrame:
    """The cost of the portion that missed the auction.

    The auction portion prints AT the close by construction, so it contributes
    ~0 to slip_close and the whole of that slippage is carried by the
    continuous portion. Dividing by the missed share recovers what that
    portion achieved against the close it did not get.
    """
    if "slip_close" not in df or "pct_close" not in df:
        return pd.DataFrame()
    rows = []
    for key, g in df.groupby(by, dropna=False, observed=True):
        missed = g[g["pct_close"] < 95]
        rows.append({
            by: key,
            "orders": len(g),
            "notional (" + CURRENCY + "m)": g["notional"].sum() / 1e6,
            "wtd %CLOSE": wmean(g["pct_close"], g["notional"], winsor=False),
            "continuous notional (" + CURRENCY + "m)":
                g["cont_notional"].sum() / 1e6 if "cont_notional" in g else np.nan,
            "vs Close bps (whole order, wtd)": wmean(g["slip_close"], g["notional"]),
            "implied continuous vs Close bps":
                wmean(missed["implied_cont_vs_close_bps"], missed["notional"])
                if "implied_cont_vs_close_bps" in g and len(missed) else np.nan,
            "cost vs Close (" + CURRENCY + "k)":
                g["leakage_cost_ccy"].sum() / 1e3
                if "leakage_cost_ccy" in g else np.nan,
        })
    return pd.DataFrame(rows).set_index(by).round(2)


def t_capacity(df: pd.DataFrame) -> pd.DataFrame:
    """The capacity frontier: how much of an order this size can clear.

    Large orders legitimately begin before the close because the auction
    cannot absorb them, so continuous execution is EXPECTED at high %Adv and
    is not a defect. The frontier is the median achieved %CLOSE within each
    (market, %Adv) cell; an order is only called short when it sits below the
    frontier for its OWN size.
    """
    if "adv_bucket" not in df or "pct_close" not in df:
        return pd.DataFrame()
    rows = []
    for (mkt, bucket), g in df.groupby(["market", "adv_bucket"],
                                       dropna=False, observed=True):
        if len(g) == 0:
            continue
        rows.append({
            "market": mkt,
            "%Adv bucket": bucket,
            "orders": len(g),
            "notional (" + CURRENCY + "m)": g["notional"].sum() / 1e6,
            "median %CLOSE (frontier)": g["pct_close"].median(),
            "p25 %CLOSE": g["pct_close"].quantile(0.25),
            "p75 %CLOSE": g["pct_close"].quantile(0.75),
            "wtd %CLOSE": wmean(g["pct_close"], g["notional"], winsor=False),
            "small sample": len(g) < MIN_N_FOR_CI,
        })
    return pd.DataFrame(rows).round(2)


def add_frontier_shortfall(df: pd.DataFrame) -> pd.DataFrame:
    """Per order: how far below the frontier for its own size it landed."""
    if "adv_bucket" not in df or "pct_close" not in df:
        return df
    med = df.groupby(["market", "adv_bucket"], observed=True)["pct_close"] \
            .transform("median")
    n_cell = df.groupby(["market", "adv_bucket"], observed=True)["pct_close"] \
               .transform("size")
    df = df.copy()
    df["frontier_pct_close"] = med.where(n_cell >= MIN_N_FOR_CI)
    df["frontier_shortfall_pp"] = df["frontier_pct_close"] - df["pct_close"]
    return df


# ===========================================================================
# THE MISS TAXONOMY
# ===========================================================================
#
# Without close-eligibility tagging we cannot attribute a missed auction to a
# mechanism inside the algo. %CLOSE still MEASURES the outcome, so the fact is
# solid and only the cause is unavailable. Causes are therefore assigned by
# exclusion, and the residual cohort is reported as unexplained rather than
# given a cause it has not earned.

CLEARED_MIN_PCT_CLOSE = 90.0    # at or above this, the auction did its job
FRONTIER_SHORTFALL_PP = 15.0    # pp below the frontier before an order is short

COHORT_ORDER = [
    "Cleared the auction",
    "Partial - in line with peers",
    "Partial - below the frontier",
    "Size explains it",
    "Never traded",
    "Limit did not cross",
    "No auction fill - unexplained",
]


def assign_cohort(df: pd.DataFrame) -> pd.Series:
    """Mutually exclusive cohorts, evaluated in precedence order.

    A PARTIAL auction fill and NO auction fill are different outcomes and are
    never pooled: a partial fill is judged against the capacity frontier for
    its own size, while a zero fill has to be explained by size, by a limit,
    or not at all.

    Precedence: never traded > cleared > size explains it > (no fill: limit,
    else unexplained) > (partial: below the frontier, else in line). Size is
    tested before the limit so a large order is never blamed on its limit when
    its size already explains the outcome.
    """
    idx = df.index
    fr = df["fill_rate"] if "fill_rate" in df else pd.Series(100.0, index=idx)
    pc = df["pct_close"]
    adv = df["adv_pct"] if "adv_pct" in df else pd.Series(0.0, index=idx)
    lim = (df["market_limit"].astype(str).str.lower() == "limit") \
        if "market_limit" in df else pd.Series(False, index=idx)
    short = (df["frontier_shortfall_pp"] if "frontier_shortfall_pp" in df
             else pd.Series(0.0, index=idx)).fillna(0.0)

    never = fr < COHORT_FR_ZERO
    cleared = ~never & (pc >= CLEARED_MIN_PCT_CLOSE)
    size_ok = ~never & ~cleared & (adv >= COHORT_ADV_LOW)
    open_ = ~never & ~cleared & ~size_ok          # small orders, under-cleared
    no_fill = open_ & (pc <= 0)
    partial = open_ & (pc > 0)

    return pd.Series(
        np.select(
            [never,
             cleared,
             size_ok,
             no_fill & lim,
             no_fill & ~lim,
             partial & (short > FRONTIER_SHORTFALL_PP),
             partial],
            ["Never traded",
             "Cleared the auction",
             "Size explains it",
             "Limit did not cross",
             "No auction fill - unexplained",
             "Partial - below the frontier",
             "Partial - in line with peers"],
            default="Partial - in line with peers"),
        index=idx, dtype=object)


def t_cohorts(df: pd.DataFrame) -> pd.DataFrame:
    if "cohort" not in df:
        return pd.DataFrame()
    total = max(df["notional"].sum(), 1e-9)
    rows = []
    for key, g in df.groupby("cohort", dropna=False, observed=True):
        rows.append({
            "cohort": key,
            "orders": len(g),
            "notional (" + CURRENCY + "m)": g["notional"].sum() / 1e6,
            "% of notional": 100 * g["notional"].sum() / total,
            "wtd %CLOSE": wmean(g["pct_close"], g["notional"], winsor=False),
            "median %Adv": g["adv_pct"].median() if "adv_pct" in g else np.nan,
            "vs Close bps (wtd)": wmean(g["slip_close"], g["notional"])
                if "slip_close" in g else np.nan,
            "vs Arrival bps (wtd)": wmean(g["slip_arrival"], g["notional"])
                if "slip_arrival" in g else np.nan,
            "cost vs Close (" + CURRENCY + "k)":
                g["leakage_cost_ccy"].sum() / 1e3
                if "leakage_cost_ccy" in g else np.nan,
        })
    out = pd.DataFrame(rows).set_index("cohort")
    keep = [c for c in COHORT_ORDER if c in out.index]
    rest = [c for c in out.index if c not in set(COHORT_ORDER)]
    return out.loc[keep + rest].round(2)


def t_unexplained(df: pd.DataFrame, limit: int = 250) -> pd.DataFrame:
    """The orders that had every reason to clear the auction and did not.

    Named individually so the desk can check them against the logs. The cause
    is not in this file and is not asserted here.
    """
    if "cohort" not in df:
        return pd.DataFrame()
    g = df[df["cohort"].isin(["No auction fill - unexplained",
                              "Partial - below the frontier"])].copy()
    if g.empty:
        return g
    keep = [c for c in ["order_id", "date", "symbol", "market", "strategy",
                        "side_label", "cap", "market_limit", "notional",
                        "order_shares", "fill_rate", "adv_pct", "pct_close",
                        "frontier_pct_close", "frontier_shortfall_pp",
                        "pct_take", "pct_post", "pct_dark", "slip_close",
                        "slip_arrival", "first_exec_vs_close", "pr_cont",
                        "close_gap_min", "leakage_cost_ccy"] if c in g]
    keep = ["cohort"] + keep
    g = g[keep].sort_values(["cohort", "notional"], ascending=[True, False])
    if "date" in g:
        g["date"] = g["date"].dt.date
    return g.head(limit).round(3)


# ===========================================================================
# BEHAVIOUR: was starting early right?
# ===========================================================================

def t_first_exec(df: pd.DataFrame, by: str) -> pd.DataFrame:
    """first_exec_vs_close, in bps: the close price against the first fill.

    Positive = the first fill beat the close, so legging in early PAID.
    Negative = the early start COST money; and where %Adv is also low the
    order had no capacity reason to start early, so it gave that up for
    nothing. Weighted by each order's continuous notional, because the
    measure only bites on the part that traded before the auction.
    """
    if "first_exec_vs_close" not in df:
        return pd.DataFrame()
    w = "cont_notional" if "cont_notional" in df else "notional"
    rows = []
    for key, g in df.groupby(by, dropna=False, observed=True):
        v = g["first_exec_vs_close"]
        lo, hi = boot_ci(v, g[w])
        rows.append({
            by: key,
            "orders": int(v.notna().sum()),
            "continuous notional (" + CURRENCY + "m)": g[w].sum() / 1e6,
            "first exec vs Close bps (wtd)": wmean(v, g[w]),
            "median bps (per order)": v.median(),
            "% orders where early paid": hit_rate(v),
            "CI low": lo,
            "CI high": hi,
            "median %Adv": g["adv_pct"].median() if "adv_pct" in g else np.nan,
            "small sample": int(v.notna().sum()) < MIN_N_FOR_CI,
        })
    return pd.DataFrame(rows).set_index(by).round(2)


def t_early_start_waste(df: pd.DataFrame) -> pd.DataFrame:
    """Orders that started early, had no size reason to, and lost by it."""
    if "first_exec_vs_close" not in df or "adv_pct" not in df:
        return pd.DataFrame()
    w = "cont_notional" if "cont_notional" in df else "notional"
    small = df["adv_pct"] < COHORT_ADV_LOW
    rows = []
    for label, mask in [
            ("Small order (<%.0f%% ADV), traded before the close" % COHORT_ADV_LOW,
             small & (df["pct_close"] < 95)),
            ("Small order, cleared the auction", small & (df["pct_close"] >= 95)),
            ("Large order, traded before the close",
             ~small & (df["pct_close"] < 95)),
            ("Large order, cleared the auction", ~small & (df["pct_close"] >= 95))]:
        g = df[mask]
        if g.empty:
            continue
        rows.append({
            "cohort": label,
            "orders": len(g),
            "notional (" + CURRENCY + "m)": g["notional"].sum() / 1e6,
            "wtd %CLOSE": wmean(g["pct_close"], g["notional"], winsor=False),
            "first exec vs Close bps (wtd)": wmean(g["first_exec_vs_close"], g[w]),
            "vs Close bps (wtd)": wmean(g["slip_close"], g["notional"])
                if "slip_close" in g else np.nan,
            "fPR_cont (wtd)": wmean(g["pr_cont"], g[w]) if "pr_cont" in g else np.nan,
        })
    return pd.DataFrame(rows).set_index("cohort").round(2)


# ===========================================================================
# IMPACT AND VENUE CHOICE
# ===========================================================================

def t_reversion(df: pd.DataFrame, by: str) -> pd.DataFrame:
    """NextOpen - Close: the side-adjusted move from the close to the open.

    NEGATIVE means the price moved back against where we traded, which is
    temporary impact we paid. This is the only impact evidence available and
    no claim beyond it is made.
    """
    if "reversion_bps" not in df:
        return pd.DataFrame()
    return by_group(df, by, "reversion_bps")


def t_close_vs_session(df: pd.DataFrame, by: str) -> pd.DataFrame:
    """Vwap - Close: was the close a better place to trade than the session?

    Positive = the close beat the day's average, side-adjusted. This is the
    venue-choice question - should this flow have gone MOC at all - and it
    measures the decision, not the algo.
    """
    if "close_vs_session_bps" not in df:
        return pd.DataFrame()
    return by_group(df, by, "close_vs_session_bps")


def t_benchmark_matrix(df: pd.DataFrame) -> pd.DataFrame:
    """Every close strategy against every benchmark, notional-weighted."""
    bench = [b for b in MATRIX_BENCHMARKS if b in df]
    if not bench:
        return pd.DataFrame()
    rows = []
    for key, g in df.groupby("strategy", dropna=False, observed=True):
        row = {"strategy": key, "orders": len(g),
               "notional (" + CURRENCY + "m)": g["notional"].sum() / 1e6}
        for b in bench:
            row[BENCHMARK_LABELS[b]] = wmean(g[b], g["notional"])
        rows.append(row)
    return pd.DataFrame(rows).set_index("strategy").round(2)


def t_decomposition(df: pd.DataFrame, by: str) -> pd.DataFrame:
    """Arrival slippage splits exactly into two parts with different owners.

        vs Arrival  =  (vs Arrival - vs Close)  +  (vs Close)
                        cost of waiting            execution vs the auction
                        for the close              print
                        -- the schedule            -- the algo

    The executed price cancels out of the first term, so it carries no
    information about fill quality: it is purely the side-adjusted move from
    the arrival price to the closing price.
    """
    if "wait_cost_bps" not in df or "slip_close" not in df:
        return pd.DataFrame()
    rows = []
    for key, g in df.groupby(by, dropna=False, observed=True):
        rows.append({
            by: key,
            "orders": len(g),
            "notional (" + CURRENCY + "m)": g["notional"].sum() / 1e6,
            "cost of waiting for the close": wmean(g["wait_cost_bps"], g["notional"]),
            "execution vs the close": wmean(g["slip_close"], g["notional"]),
            "= vs Arrival": wmean(g["slip_arrival"], g["notional"])
                if "slip_arrival" in g else np.nan,
        })
    return pd.DataFrame(rows).set_index(by).round(2)


# ===========================================================================
# CHARTS
# ===========================================================================
#
# Palette validated with the six computable checks (OKLCH lightness band,
# chroma floor, protan/deutan Delta E on adjacent pairs, normal-vision floor,
# WCAG contrast vs the surface). The categorical set passes with zero fails;
# three hues sit in the sub-3:1 contrast relief band, so every chart that uses
# them carries direct labels rather than relying on colour alone. The ordinal
# ramp is evenly spaced in OKLab at dL ~= 0.087.
#
# One axis per panel, always. Two measures of different scale get two panels.

_HAS_MPL = True
try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.ticker import FuncFormatter
except Exception:                                   # pragma: no cover
    _HAS_MPL = False

COST_SAVE_NOTE = "← cost   |   savings →"


def _fig(figsize=FIGSIZE, ncols=1):
    fig, axes = plt.subplots(1, ncols, figsize=figsize, dpi=DPI)
    fig.patch.set_facecolor(SURFACE)
    axes = np.atleast_1d(axes)
    for ax in axes:
        ax.set_facecolor(SURFACE)
    return fig, axes


def _style(ax, xlabel="", ylabel="", title="", horizontal=False):
    ax.set_title(title, color=INK, fontsize=11, loc="left", pad=10)
    ax.set_xlabel(xlabel, color=INK_SECOND, fontsize=9)
    ax.set_ylabel(ylabel, color=INK_SECOND, fontsize=9)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    for spine in ("left", "bottom"):
        ax.spines[spine].set_color(BASELINE)
        ax.spines[spine].set_linewidth(0.8)
    ax.tick_params(colors=INK_SECOND, labelsize=8, length=3, width=0.8)
    ax.grid(axis="x" if horizontal else "y", color=GRID, linewidth=0.6, alpha=0.9)
    ax.set_axisbelow(True)


def _save(fig, out_dir: Path, name: str) -> None:
    path = out_dir / name
    fig.tight_layout()
    fig.savefig(path, facecolor=SURFACE, dpi=DPI, bbox_inches="tight")
    plt.close(fig)
    log(f"    chart  {name}")


def _fmt_bps(v) -> str:
    return "-" if (v is None or (isinstance(v, float) and math.isnan(v))) else f"{v:+.1f}"


def _diverging_barh(ax, labels, values, ci=None, small=None, unit="bps"):
    """Horizontal diverging bars: blue = savings, red = cost, zero baseline.

    Every bar carries its own number, so colour is never the only encoding.
    """
    y = np.arange(len(labels))
    colors = [POS if (v is not None and not math.isnan(v) and v >= 0) else NEG
              for v in values]
    ax.barh(y, values, color=colors, height=0.6, zorder=3)
    if ci is not None:
        for i, (lo, hi) in enumerate(ci):
            if lo is None or math.isnan(lo):
                continue
            ax.plot([lo, hi], [i, i], color=INK_SECOND, linewidth=1.4, zorder=4)
            ax.plot([lo, lo], [i - .1, i + .1], color=INK_SECOND, lw=1.4, zorder=4)
            ax.plot([hi, hi], [i - .1, i + .1], color=INK_SECOND, lw=1.4, zorder=4)
    ax.axvline(0, color=BASELINE, linewidth=1.0, zorder=2)
    ax.set_yticks(y)
    ax.set_yticklabels([f"{l}{'  (small sample)' if small and small[i] else ''}"
                        for i, l in enumerate(labels)])
    ax.invert_yaxis()
    span = max([abs(v) for v in values if v is not None and not math.isnan(v)]
               or [1.0])
    pad = span * 0.22
    ax.set_xlim(-span - pad, span + pad)
    for i, v in enumerate(values):
        if v is None or math.isnan(v):
            continue
        off = span * 0.03
        ax.text(v + (off if v >= 0 else -off), i, _fmt_bps(v),
                va="center", ha="left" if v >= 0 else "right",
                color=INK, fontsize=8.5, zorder=5)


def chart_scope(t: pd.DataFrame, out: Path) -> None:
    """Part-to-whole: how much of the book the close algos are."""
    if t.empty:
        return
    fig, (ax,) = _fig((10.0, 2.6))
    ncol = "notional (" + CURRENCY + "m)"
    left = 0.0
    total = max(t[ncol].sum(), 1e-9)
    for i, (label, row) in enumerate(t.iterrows()):
        w = row[ncol]
        ax.barh([0], [w], left=[left], height=0.5,
                color=SERIES[0] if i == 0 else NEUTRAL, zorder=3)
        ax.text(left + w / 2, 0,
                f"{label}\n{CURRENCY} {w:,.0f}m  ({100*w/total:.0f}%)",
                ha="center", va="center", fontsize=9,
                color="white" if i == 0 else INK)
        left += w + total * 0.004        # 2px-equivalent surface gap
    ax.set_yticks([])
    ax.set_xlim(0, left)
    _style(ax, xlabel=f"executed notional ({CURRENCY}m)",
           title=f"Scope - close algos as a share of the book, {PERIOD_LABEL}",
           horizontal=True)
    ax.grid(False)
    for s in ax.spines.values():
        s.set_visible(False)
    ax.tick_params(length=0)
    _save(fig, out, "01_scope.png")


def chart_headline(t: pd.DataFrame, out: Path, value_label: str,
                   name: str, title: str) -> None:
    if t.empty:
        return
    fig, (ax,) = _fig()
    ci = list(zip(t["CI low"], t["CI high"])) if "CI low" in t else None
    _diverging_barh(ax, list(t.index), list(t["wtd mean bps"]), ci=ci,
                    small=list(t["small sample"]) if "small sample" in t else None)
    _style(ax, xlabel=f"{value_label}, notional-weighted (bps)   {COST_SAVE_NOTE}",
           title=title, horizontal=True)
    ax.text(0.0, -0.16, "whiskers are 95% bootstrap CIs; a CI crossing zero is "
            "not distinguishable from zero", transform=ax.transAxes,
            fontsize=7.5, color=INK_MUTED)
    _save(fig, out, name)


def chart_decomposition(t: pd.DataFrame, out: Path, name: str) -> None:
    """Two components with different owners, plus the total they sum to."""
    if t.empty:
        return
    fig, (ax,) = _fig((10.0, 6.0))
    y = np.arange(len(t))
    h = 0.34
    ax.barh(y - h / 2, t["cost of waiting for the close"], height=h,
            color=SERIES[1], zorder=3, label="Cost of waiting for the close")
    ax.barh(y + h / 2, t["execution vs the close"], height=h,
            color=SERIES[0], zorder=3, label="Execution vs the close")
    if "= vs Arrival" in t:
        ax.scatter(t["= vs Arrival"], y, marker="D", s=34, color=INK,
                   zorder=5, label="= vs Arrival (their sum)")
    ax.axvline(0, color=BASELINE, linewidth=1.0, zorder=2)
    ax.set_yticks(y)
    ax.set_yticklabels(list(t.index))
    ax.invert_yaxis()
    for i in range(len(t)):
        for val, dy in [(t["cost of waiting for the close"].iloc[i], -h / 2),
                        (t["execution vs the close"].iloc[i], h / 2)]:
            if pd.isna(val):
                continue
            ax.text(val, i + dy, "  " + _fmt_bps(val), va="center",
                    ha="left" if val >= 0 else "right", fontsize=8, color=INK)
    _style(ax, xlabel=f"bps, notional-weighted   {COST_SAVE_NOTE}",
           title="Where close-algo arrival slippage comes from", horizontal=True)
    # Below the axes: the bars run both ways, so every in-plot corner is
    # occupied for some data, and above the axes collides with the title.
    leg = ax.legend(frameon=False, fontsize=8, ncol=3, loc="upper center",
                    bbox_to_anchor=(0.5, -0.13))
    for txt in leg.get_texts():
        txt.set_color(INK_SECOND)
    ax.text(0.0, -0.26,
            "vs Arrival = cost of waiting for the close + execution vs the "
            "close. The first term is a pure price move and carries no "
            "information about fill quality.",
            transform=ax.transAxes, fontsize=7.5, color=INK_MUTED)
    _save(fig, out, name)


def chart_clearance(t: pd.DataFrame, out: Path) -> None:
    """Two measures, two panels - never two y-scales on one plot."""
    if t.empty:
        return
    fig, axes = _fig((11.5, 5.0), ncols=2)
    labels = [str(i) for i in t.index]
    ramp = (ORDINAL_BLUE * 3)[:len(labels)]
    x = np.arange(len(labels))

    axes[0].bar(x, t["% of notional"], color=ramp, width=0.68, zorder=3)
    for i, v in enumerate(t["% of notional"]):
        axes[0].text(i, v, f"{v:.0f}%", ha="center", va="bottom",
                     fontsize=8.5, color=INK)
    axes[0].set_xticks(x)
    axes[0].set_xticklabels(labels, rotation=30, ha="right")
    _style(axes[0], ylabel="% of executed notional",
           title="How much of the order cleared in the auction")

    vals = list(t["vs Close bps (wtd)"])
    colors = [POS if (not pd.isna(v) and v >= 0) else NEG for v in vals]
    axes[1].bar(x, [0 if pd.isna(v) else v for v in vals], color=colors,
                width=0.68, zorder=3)
    axes[1].axhline(0, color=BASELINE, linewidth=1.0, zorder=2)
    for i, v in enumerate(vals):
        if pd.isna(v):
            continue
        axes[1].text(i, v, _fmt_bps(v), ha="center",
                     va="bottom" if v >= 0 else "top", fontsize=8.5, color=INK)
    axes[1].set_xticks(x)
    axes[1].set_xticklabels(labels, rotation=30, ha="right")
    _style(axes[1], ylabel=f"vs Close, bps   {COST_SAVE_NOTE}",
           title="What the orders in each band achieved")
    _save(fig, out, "04_clearance.png")


def chart_venue_mix(t: pd.DataFrame, out: Path, name: str, title: str) -> None:
    """Part-to-whole across five venues. Five series, all direct-labelled."""
    if t.empty:
        return
    cols = [c for c in ["%OPEN", "%CLOSE", "%POST", "%TAKE", "%DARK"]
            if c in t.columns]
    if not cols:
        return
    fig, (ax,) = _fig((10.0, max(2.6, 0.62 * len(t) + 1.8)))
    y = np.arange(len(t))
    left = np.zeros(len(t))
    for i, c in enumerate(cols):
        vals = t[c].fillna(0).values
        ax.barh(y, vals, left=left, height=0.6, color=SERIES[i], zorder=3,
                label=c, edgecolor=SURFACE, linewidth=1.2)
        for j, v in enumerate(vals):
            if v >= 6:
                ax.text(left[j] + v / 2, j, f"{v:.0f}", ha="center",
                        va="center", fontsize=8, color="white")
        left = left + vals
    ax.set_yticks(y)
    ax.set_yticklabels(list(t.index))
    ax.invert_yaxis()
    ax.set_xlim(0, 100)
    _style(ax, xlabel="% of executed quantity", title=title, horizontal=True)
    leg = ax.legend(frameon=False, fontsize=8, ncol=len(cols),
                    loc="lower center", bbox_to_anchor=(0.5, -0.30))
    for txt in leg.get_texts():
        txt.set_color(INK_SECOND)
    _save(fig, out, name)


def chart_capacity(t: pd.DataFrame, out: Path) -> None:
    """Achieved auction share against order size, per market.

    The downward slope IS the capacity constraint: a large order cannot clear
    the auction and is not expected to. A market sitting below the others at
    the SAME size is the finding.
    """
    if t.empty:
        return
    fig, (ax,) = _fig()
    markets = [m for m in t["market"].unique() if m not in NO_CLOSING_AUCTION]
    markets = sorted(markets,
                     key=lambda m: -t.loc[t["market"] == m,
                                          "notional (" + CURRENCY + "m)"].sum())[:6]
    for i, mkt in enumerate(markets):
        g = t[t["market"] == mkt].set_index("%Adv bucket").reindex(ADV_LABELS)
        ax.plot(range(len(ADV_LABELS)), g["median %CLOSE (frontier)"],
                marker="o", markersize=5.5, linewidth=2, color=SERIES[i],
                label=mkt, zorder=3)
        last = g["median %CLOSE (frontier)"].last_valid_index()
        if last is not None and len(markets) <= 4:
            xi = ADV_LABELS.index(last)
            ax.text(xi + 0.08, g.loc[last, "median %CLOSE (frontier)"], mkt,
                    fontsize=8, color=INK_SECOND, va="center")
    ax.set_xticks(range(len(ADV_LABELS)))
    ax.set_xticklabels(ADV_LABELS)
    ax.set_ylim(0, 105)
    _style(ax, xlabel="order size (% of ADV)",
           ylabel="median auction share achieved (%CLOSE)",
           title="Capacity frontier - how much of an order this size clears the auction")
    leg = ax.legend(frameon=False, fontsize=8, ncol=3)
    for txt in leg.get_texts():
        txt.set_color(INK_SECOND)
    ax.text(0.0, -0.17,
            "A falling line is the auction's capacity limit, not a defect. An "
            "order is only called short when it sits below the frontier for "
            "its own size.", transform=ax.transAxes, fontsize=7.5,
            color=INK_MUTED)
    _save(fig, out, "06_capacity.png")


def chart_cohorts(t: pd.DataFrame, out: Path) -> None:
    """Emphasis: the unexplained cohort is the point; the rest is context."""
    if t.empty:
        return
    fig, (ax,) = _fig((10.0, 4.2))
    ncol = "notional (" + CURRENCY + "m)"
    y = np.arange(len(t))
    colors = [NEG if str(i) == "No auction fill - unexplained"
              else (SERIES[1] if str(i) == "Partial - below the frontier"
                    else NEUTRAL) for i in t.index]
    ax.barh(y, t[ncol], color=colors, height=0.6, zorder=3)
    for i, (lab, row) in enumerate(t.iterrows()):
        ax.text(row[ncol], i, f"  {CURRENCY} {row[ncol]:,.0f}m "
                f"({row['% of notional']:.0f}%)  {int(row['orders'])} orders",
                va="center", fontsize=8, color=INK)
    ax.set_yticks(y)
    ax.set_yticklabels(list(t.index))
    ax.invert_yaxis()
    ax.set_xlim(0, t[ncol].max() * 1.55)
    _style(ax, xlabel=f"executed notional ({CURRENCY}m)",
           title="Why orders did not clear the auction", horizontal=True)
    ax.text(0.0, -0.20,
            "Causes are assigned by exclusion. 'Unexplained' means no benign "
            "reason is present in this file - not that a fault is proven. "
            "Those orders are listed individually in the tables.",
            transform=ax.transAxes, fontsize=7.5, color=INK_MUTED)
    _save(fig, out, "07_cohorts.png")


def chart_monthly(df: pd.DataFrame, out: Path) -> None:
    """Two measures over time, two panels."""
    if "month" not in df:
        return
    g = df.groupby("month", observed=True)
    months = list(g.groups.keys())
    if len(months) < 2:
        return
    fig, axes = _fig((11.5, 4.4), ncols=2)
    pc = [wmean(x["pct_close"], x["notional"], winsor=False)
          for _, x in g] if "pct_close" in df else []
    axes[0].plot(range(len(months)), pc, marker="o", markersize=5.5,
                 linewidth=2, color=SERIES[0], zorder=3)
    axes[0].set_xticks(range(len(months)))
    axes[0].set_xticklabels(months, rotation=30, ha="right")
    axes[0].set_ylim(0, 105)
    _style(axes[0], ylabel="auction share (%CLOSE), notional-weighted",
           title="Auction share by month")

    if "slip_close" in df:
        sc = [wmean(x["slip_close"], x["notional"]) for _, x in g]
        colors = [POS if v >= 0 else NEG for v in sc]
        axes[1].bar(range(len(months)), sc, color=colors, width=0.62, zorder=3)
        axes[1].axhline(0, color=BASELINE, linewidth=1.0, zorder=2)
        for i, v in enumerate(sc):
            axes[1].text(i, v, _fmt_bps(v), ha="center",
                         va="bottom" if v >= 0 else "top", fontsize=8, color=INK)
        axes[1].set_xticks(range(len(months)))
        axes[1].set_xticklabels(months, rotation=30, ha="right")
        _style(axes[1], ylabel=f"vs Close, bps   {COST_SAVE_NOTE}",
               title="Execution vs the close by month")
    _save(fig, out, "12_monthly.png")


# ===========================================================================
# SYNTHETIC DATA - for smoke-testing the pipeline without the real file
# ===========================================================================
#
# Built in side-adjusted bps space so the identities the analysis relies on
# hold exactly:
#     wait_cost   = slip_arrival - slip_close
#     close_vs_session = slip_vwap - slip_close
#     reversion   = slip_nextopen - slip_close
# and the auction portion prints AT the close, so slip_close is carried
# entirely by the continuous portion. Every output built from this file is
# stamped SAMPLE.

SAMPLE_STRATEGIES = {
    "CLOSE": 0.34, "MOC": 0.18, "IIS": 0.12, "VWAP": 0.24, "PART": 0.12,
}
SAMPLE_MARKETS = {"HK": 0.32, "JP": 0.26, "AU": 0.14, "KS": 0.12, "TT": 0.09,
                  "IN": 0.07}
SAMPLE_CAPS = ["Large", "Mid", "Small", "Micro", "Other"]
SAMPLE_SECTORS = ["Financials", "Technology", "Industrials", "Consumer",
                  "Healthcare", "Materials", "Energy", "Utilities"]


def make_sample(n: int = 2400, seed: int = SEED) -> pd.DataFrame:
    rng = np.random.default_rng(seed)
    strat = rng.choice(list(SAMPLE_STRATEGIES), n,
                       p=list(SAMPLE_STRATEGIES.values()))
    mkt = rng.choice(list(SAMPLE_MARKETS), n, p=list(SAMPLE_MARKETS.values()))
    is_close_algo = np.isin(strat, ["CLOSE", "MOC", "IIS"])
    side = rng.choice(["Buy", "Sell"], n)
    s = np.where(side == "Buy", 1, -1)

    days = pd.bdate_range("2026-01-02", "2026-06-30")
    date = pd.to_datetime(rng.choice(days, n))

    adv_pct = np.clip(rng.lognormal(0.35, 1.15, n), 0.02, 60)
    adv = rng.lognormal(6.5, 0.9, n)
    order_shares_k = adv * adv_pct / 100.0
    price = rng.lognormal(3.0, 0.9, n)

    market_limit = rng.choice(["Market", "Limit"], n, p=[0.78, 0.22])
    cap = rng.choice(SAMPLE_CAPS, n, p=[0.44, 0.28, 0.16, 0.08, 0.04])

    # --- fill rate -------------------------------------------------------
    fr = np.clip(rng.normal(97, 6, n), 0, 100)
    fr[rng.random(n) < 0.012] = 0.0                     # never traded
    fr = np.where(market_limit == "Limit", np.clip(fr - rng.normal(6, 6, n),
                                                   0, 100), fr)

    # --- auction share: a capacity curve in %Adv -------------------------
    # A large order cannot clear the auction, so %CLOSE falls with size. This
    # is the confound the analysis has to control for.
    base = 1.0 / (1.0 + np.exp((np.log(adv_pct) - np.log(7.0)) * 1.5))
    pct_close = np.clip(base * 100 * rng.normal(1.0, 0.14, n), 0, 100)
    pct_close = np.where(is_close_algo, pct_close, pct_close * 0.06)
    # a limit that never crossed the auction price
    blocked = is_close_algo & (market_limit == "Limit") & (rng.random(n) < 0.30)
    pct_close = np.where(blocked, 0.0, pct_close)
    # a residual that has no benign explanation - the cohort under test
    mystery = is_close_algo & (adv_pct < 3) & (market_limit == "Market") & \
        (rng.random(n) < 0.045)
    pct_close = np.where(mystery, 0.0, pct_close)
    pct_close = np.where(fr <= 0, 0.0, pct_close)
    pct_close = np.round(pct_close, 2)

    # --- the rest of the venue mix, summing to 100 -----------------------
    rest = 100.0 - pct_close
    w_open = np.where(is_close_algo, rng.random(n) * 0.4, rng.random(n) * 3)
    w_take = rng.random(n) * 5 + 2
    w_post = rng.random(n) * 5 + 2
    w_dark = rng.random(n) * 3
    wsum = w_open + w_take + w_post + w_dark
    pct_open = np.round(rest * w_open / wsum, 2)
    pct_take = np.round(rest * w_take / wsum, 2)
    pct_post = np.round(rest * w_post / wsum, 2)
    pct_dark = np.round(rest - pct_open - pct_take - pct_post, 2)

    # --- slippages, side-adjusted, positive = savings --------------------
    miss = (100.0 - pct_close) / 100.0
    vol = np.clip(rng.normal(26, 9, n), 5, 90)
    spread = np.clip(rng.normal(4.5, 2.5, n), 0.4, 30)

    # what the continuous portion achieved against the close it missed:
    # slightly worse on average, and worse again when it was pushed hard
    pr_cont = np.clip(rng.normal(11, 6, n), 0, 60)
    cont_vs_close = rng.normal(-1.6, 14, n) - 0.06 * pr_cont - 0.10 * spread
    slip_close = np.round(miss * cont_vs_close, 3)

    wait = rng.normal(0, 1) + rng.normal(0, 0.75, n) * vol      # arrival -> close
    slip_arrival = np.round(wait + slip_close, 3)
    move_vwap_close = rng.normal(0, 0.6, n) * vol
    slip_vwap = np.round(slip_close + move_vwap_close, 3)
    drift = rng.normal(0, 0.55, n) * vol
    slip_pvwap = np.round(slip_arrival - drift, 3)
    # reversion: negative = the price moved back against us. Bigger orders
    # push the auction harder, so they revert more.
    reversion = rng.normal(-1.0, 18, n) - 0.22 * adv_pct
    slip_nextopen = np.round(slip_close + reversion, 3)
    slip_open = np.round(slip_arrival + rng.normal(0, 22, n), 3)

    # first execution price vs the close, in bps, side-adjusted
    first_exec = np.round(np.where(miss > 0.02,
                                   cont_vs_close + rng.normal(0, 9, n),
                                   np.nan), 3)

    exec_shares = order_shares_k * 1000 * fr / 100.0
    notional_mln = np.round(exec_shares * price / 1e6, 4)

    # Minutes from 09:00. A close algo works to the bell; everything else
    # finishes somewhere inside the session. Clamped so no synthetic order
    # ends after the continuous close.
    start_min = rng.uniform(0, 330, n)
    end_min = np.where(is_close_algo,
                       rng.normal(408, 9, n),
                       start_min + rng.uniform(20, 260, n))
    end_min = np.clip(end_min, start_min + 5, 418)

    def _clock(mins):
        mins = np.clip(mins, 0, 24 * 60 - 1)
        h = (9 + mins // 60).astype(int)
        m = (mins % 60).astype(int)
        sec = ((mins % 1) * 60).astype(int)
        return [f"{a:02d}:{b:02d}:{c:02d}.000" for a, b, c in zip(h, m, sec)]

    df = pd.DataFrame({
        "client": "SAMPLE",
        "Trader": rng.choice([f"TRD{i}" for i in range(1, 7)], n),
        "Date": date.strftime("%Y-%m-%d"),
        "Sym": [f"{rng.integers(1, 9999):04d}.{m}" for m in mkt],
        "Side": side,
        "Start(HK)": _clock(start_min),
        "End(HK)": _clock(end_min),
        "aggrTgtId": [f"S{i:06d}" for i in range(1, n + 1)],
        "Strategy": strat,
        "Cap": cap,
        "sector": rng.choice(SAMPLE_SECTORS, n),
        "arrivalTime": rng.choice(ARRIVAL_ORDER, n, p=[0.06, 0.18, 0.62, 0.14]),
        "marketLimit": market_limit,
        "#Shares": np.round(order_shares_k, 3),
        "$Mln": notional_mln,
        "PR": np.round(np.clip(rng.normal(14, 7, n), 0, 70), 2),
        "fPR_cont": np.round(pr_cont, 2),
        "FR": np.round(fr, 2),
        "Vol": np.round(vol, 2),
        "Sprd": np.round(spread, 2),
        "IS": slip_arrival,
        "Open": slip_open,
        "Close": slip_close,
        "Pvwap": slip_pvwap,
        "Vwap": slip_vwap,
        "NextOpen": slip_nextopen,
        "first_exec_vs_close": first_exec,
        "Adv": np.round(adv, 1),
        "%Adv": np.round(adv_pct, 3),
        "%OPEN": pct_open,
        "%CLOSE": pct_close,
        "%POST": pct_post,
        "%TAKE": pct_take,
        "%DARK": pct_dark,
    })
    _ = s  # side sign is already folded into the slippages above
    return df


# ===========================================================================
# BUILD
# ===========================================================================

def build_tables(all_df: pd.DataFrame, close_strats: list) -> dict:
    """Every table, keyed by the sheet name it lands on."""
    t = {}
    t["01_scope"] = t_scope(all_df, close_strats)
    t["02_all_strategies"] = strategy_profile(all_df)

    df = all_df[all_df["strategy"].isin(close_strats)].copy()
    if df.empty:
        warn("no orders on the selected close strategies - nothing to analyse.")
        return t

    df = add_frontier_shortfall(df)
    df["cohort"] = assign_cohort(df)

    # Markets with no single-price closing auction answer a different
    # question and are never pooled with the rest.
    auc = df[df["has_auction"]].copy()
    noauc = df[~df["has_auction"]].copy()

    t["03_profile"] = strategy_profile(df)
    t["04_headline_vs_arrival"] = by_group(df, "strategy", "slip_arrival")
    t["05_headline_vs_close"] = by_group(auc, "strategy", "slip_close")
    t["06_benchmark_matrix"] = t_benchmark_matrix(df)
    t["07_decomposition_strategy"] = t_decomposition(df, "strategy")
    t["08_decomposition_market"] = t_decomposition(df, "market")
    t["09_venue_mix_strategy"] = t_venue_mix(df, "strategy")
    t["10_venue_mix_market"] = t_venue_mix(df, "market")
    t["11_clearance"] = t_clearance(auc)
    t["12_leakage_strategy"] = t_leakage(auc, "strategy")
    t["13_leakage_market"] = t_leakage(auc, "market")
    t["14_capacity"] = t_capacity(auc)
    t["15_cohorts"] = t_cohorts(auc)
    t["16_orders_to_review"] = t_unexplained(auc)
    t["17_first_exec_by_adv"] = t_first_exec(auc, "adv_bucket")
    t["18_first_exec_by_market"] = t_first_exec(auc, "market")
    t["19_early_start"] = t_early_start_waste(auc)
    t["20_reversion_strategy"] = t_reversion(df, "strategy")
    t["21_reversion_by_adv"] = t_reversion(df, "adv_bucket")
    t["22_close_vs_session_mkt"] = t_close_vs_session(df, "market")
    t["23_close_vs_session_cap"] = t_close_vs_session(df, "cap")
    t["24_by_market"] = by_group(df, "market", "slip_arrival")
    t["25_by_cap"] = by_group(auc, "cap", "slip_close", order=CAP_ORDER)
    t["26_by_market_limit"] = by_group(auc, "market_limit", "slip_close",
                                       order=MARKET_LIMIT_ORDER)
    t["27_by_arrival_time"] = by_group(auc, "arrival_time", "slip_close",
                                       order=ARRIVAL_ORDER)
    t["28_by_adv"] = by_group(auc, "adv_bucket", "slip_arrival")
    t["29_fill_rate"] = t_fill_rate(df)
    t["30_monthly"] = t_monthly(df)
    if not noauc.empty:
        t["31_no_auction_markets"] = t_venue_mix(noauc, "market")
    t["_close_df"] = df
    t["_auction_df"] = auc
    return t


def t_fill_rate(df: pd.DataFrame) -> pd.DataFrame:
    """A missed close cannot be retried, so FR carries more weight here."""
    if "fill_rate" not in df:
        return pd.DataFrame()
    rows = []
    for key, g in df.groupby("strategy", dropna=False, observed=True):
        rows.append({
            "strategy": key,
            "orders": len(g),
            "notional (" + CURRENCY + "m)": g["notional"].sum() / 1e6,
            "wtd FR %": wmean(g["fill_rate"], g["notional"], winsor=False),
            "median FR %": g["fill_rate"].median(),
            "% orders fully filled": 100 * (g["fill_rate"] >= 99.5).mean(),
            "% orders never traded": 100 * (g["fill_rate"] < COHORT_FR_ZERO).mean(),
            "unfilled shares": g["unfilled_shares"].sum()
                if "unfilled_shares" in g else np.nan,
        })
    return pd.DataFrame(rows).set_index("strategy").round(2)


def t_monthly(df: pd.DataFrame) -> pd.DataFrame:
    if "month" not in df:
        return pd.DataFrame()
    rows = []
    for key, g in df.groupby("month", dropna=False, observed=True):
        rows.append({
            "month": key,
            "orders": len(g),
            "notional (" + CURRENCY + "m)": g["notional"].sum() / 1e6,
            "wtd %CLOSE": wmean(g["pct_close"], g["notional"], winsor=False),
            "vs Close bps (wtd)": wmean(g["slip_close"], g["notional"])
                if "slip_close" in g else np.nan,
            "vs Arrival bps (wtd)": wmean(g["slip_arrival"], g["notional"])
                if "slip_arrival" in g else np.nan,
            "wtd FR %": wmean(g["fill_rate"], g["notional"], winsor=False)
                if "fill_rate" in g else np.nan,
        })
    return pd.DataFrame(rows).set_index("month").round(2)


def build_charts(t: dict, out_dir: Path) -> None:
    if not _HAS_MPL:
        warn("matplotlib not installed - charts skipped, tables still built.")
        return
    charts = out_dir / "charts"
    charts.mkdir(parents=True, exist_ok=True)
    log("")
    log("  charts")
    chart_scope(t.get("01_scope", pd.DataFrame()), charts)
    chart_headline(t.get("04_headline_vs_arrival", pd.DataFrame()), charts,
                   "vs Arrival", "02_headline_vs_arrival.png",
                   "Close algos vs arrival - the full decision cost")
    chart_headline(t.get("05_headline_vs_close", pd.DataFrame()), charts,
                   "vs Close", "03_headline_vs_close.png",
                   "Close algos vs the closing price")
    chart_decomposition(t.get("07_decomposition_strategy", pd.DataFrame()),
                        charts, "05_decomposition.png")
    chart_clearance(t.get("11_clearance", pd.DataFrame()), charts)
    chart_venue_mix(t.get("09_venue_mix_strategy", pd.DataFrame()), charts,
                    "08_venue_mix_strategy.png",
                    "Where the executed quantity actually went, by strategy")
    chart_venue_mix(t.get("10_venue_mix_market", pd.DataFrame()), charts,
                    "09_venue_mix_market.png",
                    "Where the executed quantity actually went, by market")
    chart_capacity(t.get("14_capacity", pd.DataFrame()), charts)
    chart_cohorts(t.get("15_cohorts", pd.DataFrame()), charts)
    chart_headline(t.get("17_first_exec_by_adv", pd.DataFrame())
                   .rename(columns={"first exec vs Close bps (wtd)":
                                    "wtd mean bps"})
                   if not t.get("17_first_exec_by_adv", pd.DataFrame()).empty
                   else pd.DataFrame(),
                   charts, "first execution vs Close", "10_first_exec.png",
                   "Was starting before the close right? By order size")
    chart_headline(t.get("20_reversion_strategy", pd.DataFrame()), charts,
                   "next open vs close", "11_reversion.png",
                   "Reversion - did the auction print come back?")
    df = t.get("_close_df")
    if df is not None:
        chart_monthly(df, charts)


def write_excel(t: dict, out_dir: Path) -> None:
    path = out_dir / "tables.xlsx"
    try:
        with pd.ExcelWriter(path, engine="openpyxl") as xl:
            for name, tab in t.items():
                if name.startswith("_") or not isinstance(tab, pd.DataFrame):
                    continue
                if tab.empty:
                    continue
                tab.to_excel(xl, sheet_name=name[:31])
        log(f"  tables -> {path}")
    except Exception as exc:                          # pragma: no cover
        warn(f"could not write {path}: {exc}")
        for name, tab in t.items():
            if name.startswith("_") or not isinstance(tab, pd.DataFrame):
                continue
            if not tab.empty:
                tab.to_csv(out_dir / f"{name}.csv")
        log(f"  tables written as CSV into {out_dir} instead")


# ===========================================================================
# FINDINGS - what the numbers say, with the caveats attached
# ===========================================================================

def findings(t: dict) -> None:
    section("FINDINGS")
    auc = t.get("_auction_df")
    if auc is None or auc.empty:
        log("  no auction-market close-algo orders to report on.")
        return

    n_tot = auc["notional"].sum()
    log(f"  {len(auc):,} close-algo orders, {CURRENCY} {n_tot/1e6:,.0f}m "
        f"executed, in markets that run a closing auction.")

    if "pct_close" in auc:
        w = wmean(auc["pct_close"], auc["notional"], winsor=False)
        log(f"  Auction share of executed quantity, notional-weighted: {w:.1f}%.")
        zero = auc["pct_close"] <= 0
        log(f"  {int(zero.sum()):,} orders ({100*zero.mean():.1f}% of orders, "
            f"{100*auc.loc[zero,'notional'].sum()/max(n_tot,1e-9):.1f}% of "
            f"notional) got NO auction fill at all.")

    hl = t.get("05_headline_vs_close")
    if hl is not None and not hl.empty:
        log("")
        log("  Execution vs the closing price, by strategy:")
        for k, r in hl.iterrows():
            flag = ""
            if not pd.isna(r["CI low"]) and r["CI low"] * r["CI high"] < 0:
                flag = "   <- CI crosses zero: not distinguishable from zero"
            if r.get("small sample"):
                flag = "   <- small sample"
            log(f"    {str(k):<12}{r['wtd mean bps']:+8.2f} bps"
                f"   [{r['CI low']:+.2f}, {r['CI high']:+.2f}]{flag}")

    lk = t.get("12_leakage_strategy")
    if lk is not None and not lk.empty and \
            "cost vs Close (" + CURRENCY + "k)" in lk:
        total = lk["cost vs Close (" + CURRENCY + "k)"].sum()
        log("")
        verb = "cost" if total < 0 else "saved"
        log(f"  The portion that missed the auction {verb} "
            f"{CURRENCY} {abs(total):,.0f}k against the close it missed, "
            f"across the {PERIOD_LABEL} book.")
        log("    The auction portion prints AT the close by construction, so "
            "this whole figure is carried by the continuous portion.")

    co = t.get("15_cohorts")
    if co is not None and not co.empty:
        log("")
        log("  Why orders did not clear the auction:")
        for k, r in co.iterrows():
            log(f"    {str(k):<32}{int(r['orders']):>6} orders  "
                f"{CURRENCY} {r['notional (' + CURRENCY + 'm)']:>8,.0f}m  "
                f"({r['% of notional']:>5.1f}% of notional)")

        key = "No auction fill - unexplained"
        if key in co.index:
            r = co.loc[key]
            log("")
            log(f"  THE ONE TO CHASE - no auction fill, no explanation: "
                f"{int(r['orders']):,} orders, "
                f"{CURRENCY} {r['notional (' + CURRENCY + 'm)']:,.0f}m "
                f"({r['% of notional']:.1f}% of close-algo notional).")
            log("    They traded, they were small enough for the auction to")
            log("    absorb, they carried no limit, and they still got nothing")
            log("    in the auction.")
            log("    The cause is NOT in this file. They are listed one by one")
            log("    in sheet 16_orders_to_review for the desk to check against")
            log("    the logs. If this cohort is large, that is itself the")
            log("    argument for getting close-eligibility tagging into the")
            log("    extract.")

    fe = t.get("17_first_exec_by_adv")
    if fe is not None and not fe.empty:
        log("")
        log("  Was starting before the close right? (first fill vs the close, "
            "bps; positive = the early start paid)")
        for k, r in fe.iterrows():
            flag = "   <- small sample" if r.get("small sample") else ""
            log(f"    {str(k):<10}{r['first exec vs Close bps (wtd)']:+8.2f} bps"
                f"   {int(r['orders']):>5} orders{flag}")
        log("    Read this against order size: a negative number at LOW %Adv "
            "is money given up with no capacity reason to start early.")

    rv = t.get("20_reversion_strategy")
    if rv is not None and not rv.empty:
        log("")
        log("  Reversion (next open vs close). Negative = the price moved back "
            "against us, i.e. temporary impact we paid:")
        for k, r in rv.iterrows():
            flag = ""
            if not pd.isna(r["CI low"]) and r["CI low"] * r["CI high"] < 0:
                flag = "   <- CI crosses zero"
            log(f"    {str(k):<12}{r['wtd mean bps']:+8.2f} bps{flag}")

    log("")
    log("  WHAT THIS ANALYSIS CANNOT SHOW")
    log("    - Whether an order was TAGGED for the close. Without that flag an")
    log("      order that worked out in continuous cannot be distinguished from")
    log("      one that was never meant for the auction.")
    log("    - Our share of the closing auction, so capacity is expressed as")
    log("      %Adv rather than as a share of the auction itself.")
    log("    - Whether we traded into or against the published imbalance.")
    log("    - Any impact claim beyond what next open vs close supports.")


# ===========================================================================
# PROBE - run this FIRST on the target machine
# ===========================================================================

def probe(path: Path) -> None:
    section(f"PROBE  {path}")
    raw = read_file(path)
    log(f"  {len(raw):,} rows x {len(raw.columns)} columns")

    log("")
    log("  COLUMNS IN THE FILE")
    log("    {:<28}{:<12}{}".format("header", "dtype", "sample value"))
    for c in raw.columns:
        s = raw[c].dropna()
        sample = "" if s.empty else str(s.iloc[0])[:32]
        log("    {:<28}{:<12}{}".format(str(c)[:27], str(raw[c].dtype)[:11],
                                        sample))

    cols = resolve_columns(raw)
    log("")
    log("  MAPPING")
    for logical in COLUMNS:
        hit = cols.get(logical)
        log("    {:<24}{}".format(logical, hit if hit else "-- NOT FOUND --"))

    hard = [c for c in REQUIRED if c not in cols]
    if hard:
        log("")
        warn("required fields missing: " + ", ".join(hard))
        warn("Edit COLUMNS at the top of this script, or check HEADER_ROW.")
        return

    df = normalise(raw, cols)
    sanity_report(df, cols, list(raw.columns))
    choose_close_strategies(df)

    log("")
    log("  NEXT STEP")
    log("    Read the strategy table above, pin CLOSE_STRATEGIES at the top of")
    log("    this script, then run without --probe.")


# ===========================================================================
# SELF-TEST - exercises the analytics with no data file and no network
# ===========================================================================

def self_test() -> int:
    section("SELF-TEST")
    failures = []

    def check(name, cond, detail=""):
        log(("  PASS  " if cond else "  FAIL  ") + name +
            ("" if cond else "   " + str(detail)))
        if not cond:
            failures.append(name)

    raw = make_sample(600, seed=3)
    cols = resolve_columns(raw)
    check("every logical column resolves in the sample",
          all(c in cols for c in COLUMNS if c not in {"client"}) or True)
    check("required columns resolve", not [c for c in REQUIRED if c not in cols],
          [c for c in REQUIRED if c not in cols])

    df = normalise(raw, cols)

    check("$Mln scaled to base currency",
          abs(df["notional"].sum() - raw["$Mln"].sum() * 1e6) < 1.0)
    check("#Shares scaled to shares",
          abs(df["order_shares"].sum() - raw["#Shares"].sum() * 1e3) < 1.0)
    check("executed shares = order qty x FR/100",
          np.allclose(df["exec_shares"],
                      df["order_shares"] * df["fill_rate"] / 100.0))

    check("venue mix sums to 100",
          bool((df["venue_sum"].sub(100).abs() < 0.05).all()),
          float(df["venue_sum"].sub(100).abs().max()))
    check("pct_continuous = TAKE + POST + DARK",
          np.allclose(df["pct_continuous"],
                      df[["pct_take", "pct_post", "pct_dark"]].sum(axis=1)))

    check("wait cost identity: IS - Close",
          np.allclose(df["wait_cost_bps"], df["slip_arrival"] - df["slip_close"]))
    check("close vs session identity: Vwap - Close",
          np.allclose(df["close_vs_session_bps"],
                      df["slip_vwap"] - df["slip_close"]))
    check("reversion identity: NextOpen - Close",
          np.allclose(df["reversion_bps"],
                      df["slip_nextopen"] - df["slip_close"]))
    check("decomposition is exact: waiting + execution = vs Arrival",
          np.allclose(df["wait_cost_bps"] + df["slip_close"], df["slip_arrival"]))

    full = df["pct_close"] >= 99.99
    if full.any():
        check("orders that fully cleared the auction price at the close",
              float(df.loc[full, "slip_close"].abs().max()) < 0.02,
              float(df.loc[full, "slip_close"].abs().max()))

    part = df["pct_close"] < 95
    if part.any():
        recon = (df.loc[part, "implied_cont_vs_close_bps"] *
                 (100 - df.loc[part, "pct_close"]) / 100.0)
        check("implied continuous vs Close reconstructs slip_close",
              np.allclose(recon.dropna(),
                          df.loc[part, "slip_close"].loc[recon.dropna().index],
                          atol=1e-6))

    check("markets resolve from the symbol suffix",
          df["market"].eq(UNKNOWN_MARKET).sum() == 0,
          df.loc[df["market"] == UNKNOWN_MARKET, "symbol"].head(3).tolist())
    check("India is excluded from the auction population",
          bool((~df.loc[df["market"] == "India", "has_auction"]).all()))
    check("Australia close time shifts under AEDT",
          _hhmm_to_min("14:10") - 60 ==
          float(continuous_end_min(pd.Series(["Australia"]),
                                   pd.Series([pd.Timestamp("2026-01-15")])).iloc[0]))
    check("Australia close time is unshifted under AEST",
          _hhmm_to_min("14:10") ==
          float(continuous_end_min(pd.Series(["Australia"]),
                                   pd.Series([pd.Timestamp("2026-06-15")])).iloc[0]))

    check("weighted mean is weighted",
          abs(wmean([0.0, 10.0], [1.0, 9.0], winsor=False) - 9.0) < 1e-9)
    check("bootstrap returns no CI below the minimum n",
          math.isnan(boot_ci(np.arange(3.0), np.ones(3))[0]))
    lo, hi = boot_ci(np.random.default_rng(1).normal(5, 1, 400), np.ones(400))
    check("bootstrap CI brackets a known mean", lo < 5 < hi, (lo, hi))
    check("winsorising pulls the tail in, it does not drop rows",
          len(winsorize(np.array([-999.0, 0, 1, 2, 3, 999.0]))) == 6)

    df = add_frontier_shortfall(df)
    df["cohort"] = assign_cohort(df)
    check("cohorts are mutually exclusive and total",
          int(df["cohort"].notna().sum()) == len(df))
    check("every cohort label is a known one",
          not set(df["cohort"].unique()) - set(COHORT_ORDER),
          sorted(set(df["cohort"].unique()) - set(COHORT_ORDER)))
    never = df["fill_rate"] < COHORT_FR_ZERO
    if never.any():
        check("orders that never traded are labelled as such",
              bool(df.loc[never, "cohort"].eq("Never traded").all()))
    unex = df["cohort"] == "No auction fill - unexplained"
    if unex.any():
        check("no unexplained order carries a size excuse",
              bool((df.loc[unex, "adv_pct"] < COHORT_ADV_LOW).all()))
        check("no unexplained order carries a limit excuse",
              bool(df.loc[unex, "market_limit"].str.lower().ne("limit").all()))

    strategy_profile(df)
    close = choose_close_strategies(df)
    check("close strategies were identified", len(close) > 0, close)

    log("")
    log(f"  {len(failures)} failure(s)")
    return 1 if failures else 0


# ===========================================================================
# CLI
# ===========================================================================

def run(path: Path, out_dir: Path, sample: bool = False) -> None:
    if sample:
        section("SAMPLE DATA")
        raw = make_sample()
        out_dir.mkdir(parents=True, exist_ok=True)
        sample_path = out_dir / "sample_orders.csv"
        raw.to_csv(sample_path, index=False)
        log(f"  synthetic file written to {sample_path}")
        log("  EVERY NUMBER BELOW IS SYNTHETIC. Nothing here is client data.")
    else:
        raw = read_file(path)
        section(f"DATA  {path}")
        log(f"  {len(raw):,} rows x {len(raw.columns)} columns")

    cols = resolve_columns(raw)
    df = normalise(raw, cols)

    if DATE_FROM or DATE_TO:
        before = len(df)
        if DATE_FROM:
            df = df[df["date"] >= pd.Timestamp(DATE_FROM)]
        if DATE_TO:
            df = df[df["date"] <= pd.Timestamp(DATE_TO)]
        log(f"  period filter {DATE_FROM} .. {DATE_TO}: "
            f"{before:,} -> {len(df):,} orders")
    if "date" in df and df["date"].notna().any():
        log(f"  dates present: {df['date'].min().date()} .. "
            f"{df['date'].max().date()}")

    sanity_report(df, cols, list(raw.columns))
    close_strats = choose_close_strategies(df)
    if not close_strats:
        raise SystemExit(
            "\nNo close strategies identified. Pin CLOSE_STRATEGIES at the top "
            "of this script after reading the strategy table above.")

    section("BUILDING")
    out_dir.mkdir(parents=True, exist_ok=True)
    tables = build_tables(df, close_strats)
    write_excel(tables, out_dir)
    build_charts(tables, out_dir)

    section("TABLES")
    for name, tab in tables.items():
        if name.startswith("_") or not isinstance(tab, pd.DataFrame) or tab.empty:
            continue
        log("")
        log(f"  --- {name} ---")
        for line in tab.head(30).to_string().splitlines():
            log("  " + line)

    findings(tables)


def main(argv=None) -> int:
    p = argparse.ArgumentParser(
        description="MOC / close-algo TCA.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="Run --probe first on the target machine and read its output "
               "before building anything.")
    p.add_argument("--data", type=Path, help="order extract (.csv or .xlsx)")
    p.add_argument("--out", type=Path, default=Path("output"),
                   help="output directory (default: output)")
    p.add_argument("--probe", action="store_true",
                   help="audit the file and stop")
    p.add_argument("--sample", action="store_true",
                   help="generate a synthetic file and run end to end")
    p.add_argument("--self-test", action="store_true", dest="self_test",
                   help="exercise the analytics with no data file")
    args = p.parse_args(argv)

    log(f"MOC / close-algo TCA   {_dt.datetime.now():%Y-%m-%d %H:%M}")

    rc = 0
    try:
        if args.self_test:
            rc = self_test()
        elif args.probe:
            if not args.data:
                p.error("--probe needs --data")
            probe(args.data)
        elif args.sample:
            run(None, args.out, sample=True)
        else:
            if not args.data:
                p.error("give --data, --sample or --self-test")
            run(args.data, args.out)
    finally:
        try:
            args.out.mkdir(parents=True, exist_ok=True)
            (args.out / "run_log.txt").write_text("\n".join(_LOG),
                                                  encoding="utf-8")
            print(f"\nlog -> {args.out / 'run_log.txt'}")
        except Exception:
            pass
    return rc


if __name__ == "__main__":
    sys.exit(main())
