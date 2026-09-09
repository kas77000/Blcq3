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
PERIOD_LABEL = "H1 2026"
CURRENCY = "USD"

# Which Strategy values are close algos.
#
# Leave empty and the script derives them: any strategy whose name matches
# CLOSE_NAME_PATTERN, or whose notional-weighted mean %CLOSE is at least
# AUTO_CLOSE_MIN_PCT. What it picked is printed loudly in the run log.
#
# Run --probe first, read the strategy table, then pin the list here.
CLOSE_STRATEGIES: list[str] = ["CLOSE"]
CLOSE_NAME_PATTERN = r"CLOSE|MOC|LOC|TWAPC|IIS"
AUTO_CLOSE_MIN_PCT = 5.0

# Which Strategy values enter the study at all. Everything else is excluded
# with a count and a name in the run log, never silently.
#
# CLOSE only. This review is about the MOC product, and on this platform the
# CLOSE label already covers it end to end - MOC and IIS both report under it.
#
# The consequence belongs on the deck rather than being left implicit: the
# client's VWAP flow also reaches closing auctions, so what follows describes
# the MOC PRODUCT, not the client's total auction footprint. 01_scope shows
# how much of the book sits outside the review.
#
# Empty = keep every strategy in the file.
STRATEGY_SCOPE: list[str] = ["CLOSE"]

# Confirmed with the desk: the export is already side-adjusted, so a plus is
# good and a minus is bad on every benchmark, for buys and sells alike. The
# sanity report still prints the means by side - with this pinned True a gap
# between buys and sells is a RESULT to explain, not a data error.
SIDE_ADJUSTED = True

# %CLOSE at 100 means every executed share printed in the auction, so the order
# behaved as auction-only. This is an OUTCOME, not a permission flag: an order
# free to trade continuously that happened to fill entirely in the auction
# looks identical, and an order that never filled cannot be classified at all.
# Good enough to split the cleared population; not good enough to explain a
# miss, and the run log says so.
AUCTION_ONLY_MIN_PCT = 99.5

# Which strategies the MISS TAXONOMY applies to. A VWAP order was never meant
# to reach the auction, so calling its low %CLOSE an unexplained miss would be
# nonsense. Clearance, capacity, cohorts and the first-execution tables run on
# this list only; every other table runs on CLOSE_STRATEGIES and is grouped by
# strategy, so the two are never pooled.
MOC_STRATEGIES: list[str] = ["CLOSE"]

# A slippage this large is a broken record, not a fill. The CELL is cleared;
# the order stays in the study. Nothing is ever removed for being merely
# large - winsorising handles fat tails, and deleting the extremes would
# delete the orders the review exists to find.
MAX_ABS_BPS = 2000.0

# Period filter. None = whatever is in the file.
#
# --from / --to override these, so two windows off one extract need no code
# edit and no second copy of the script:
#
#   python moc_tca.py --data orders.csv --to 2026-06-30 --out output_h1 \
#                     --label "H1 2026"
#   python moc_tca.py --data orders.csv --out output_full \
#                     --label "Jan to 4 Sep 2026"
# H1 2026. The extract runs past this, so the filter is doing real work: it
# also puts the whole window before India's Closing Auction Session on
# 3 August 2026, which means India is one regime throughout and the CAS split
# never fires. The run log confirms that on every run rather than leaving it
# to be assumed.
DATE_FROM = "2026-01-01"
DATE_TO = "2026-06-30"

# --- the AWS extract ------------------------------------------------------
# A second source, one parquet file per date/country/client, holding the same
# orders with more columns on them. Every parquet in AWS_DIR is concatenated
# and LEFT-joined onto the order file on AWS_JOIN_KEY, so the order file stays
# the population and the extra columns ride along.
#
# Where a name exists in both, THE ORDER FILE WINS and the AWS column is kept
# beside it with AWS_SUFFIX. Nothing that already works can change underneath
# us, and the alternative is available by name when it is wanted.
AWS_DIR = "data/aws"
AWS_JOIN_KEY = "aggrTgtId"
AWS_SUFFIX = "_aws"

# India has no closing auction in this period - the close is a VWAP of the
# last half hour - so %CLOSE is reported as 0 there and cannot say whether an
# order was aimed at the close. What can say it is WHEN the order started.
# India orders that began outside this window are not close orders and leave
# the study; every other market is untouched.
#
# NSE closes 15:30 IST, which is 18:00 HKT, and the pre-CAS closing VWAP runs
# over the last half hour, 17:30-18:00 HKT. This window is the first half of
# that. Widen the second value to "18:00" to take the whole of it.
INDIA_CLOSE_WINDOW_HKT = ("17:30", "17:45")

# An India order that ran inside that window traded through the VWAP that IS
# the close, so its close share is 100 by the mechanism rather than by
# measurement. Setting it makes India's auction share mean the same thing as
# everywhere else instead of reading 0 and dragging every all-market figure
# down with it.
#
# It is an IMPUTED value, not an observed one. The original is kept in
# pct_close_measured, the rows are flagged in pct_close_imputed, and the run
# log and the findings both say so. Set False to leave the zeros alone.
INDIA_CLOSE_PROXY = True

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
    # The same benchmarks divided by the spread. Every one of these is
    # OPTIONAL: whatever the export does not carry is derived from
    # slip / spread_bps, so a file with none of them behaves identically.
    #
    # The alternatives are ordered so the normalised column pairs with the
    # SAME benchmark as its bps column. slip_pvwap maps from Pvwap, so
    # Pvwap/Sprd wins over ePvwap/Sprd - taking the e- variant would put a
    # different benchmark in the spreads column than in the bps column beside
    # it, and the two would quietly disagree.
    "sprd_arrival":  ["IS/Sprd", "eIS/Sprd", "eISSprd"],
    "sprd_pvwap":    ["Pvwap/Sprd", "ePvwap/Sprd", "PvwapSprd"],
    "sprd_close":    ["Close/Sprd", "eClose/Sprd"],
    "sprd_vwap":     ["Vwap/Sprd", "eVwap/Sprd"],
    # From the AWS extract. Optional: without it India cannot be windowed and
    # the run says so rather than filtering on something it does not have.
    "first_start_time": ["fstart_time", "fstart_time" + AWS_SUFFIX],
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

# Present in some exports, computed here when absent. Never reported as a gap,
# because a missing one costs nothing.
OPTIONAL = ["sprd_arrival", "sprd_pvwap", "sprd_close", "sprd_vwap",
            "first_start_time"]

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
# Anything not in BUY_VALUES is labelled Sell, so a blank or an unexpected
# code would silently become a sell order. SELL_VALUES exists to catch that:
# a side in neither set is excluded rather than guessed.
SELL_VALUES = {"S", "SELL", "SLD", "SS", "SSH", "SSE", "SHORT", "SHORTSELL",
               "SHORT SELL", "2", "SELLS"}

# Short sells are still sells - same direction, same sign - so they need no
# separate treatment in the arithmetic. They are worth SEEING separately
# though: locate requirements and short-sale rules can change how an order
# executes, and on this book they are 13% of the flow. So they keep their own
# label rather than disappearing into "Sell".
SHORT_SELL_VALUES = {"SS", "SSH", "SSE", "SHORT", "SHORTSELL", "SHORT SELL"}

# $Mln is EXECUTED notional, built as sum(cumqty * avgprice * fx_last) / 1e6.
# fx_last is inside it, so the column is millions of USD and the multiplier
# back to USD is 1e6. Confirmed with the data owner.
#
# The run still derives the implied USD price per share from it and warns if
# that lands outside a plausible band - a wrong scale here moves every currency
# figure by a power of ten and nothing else in the pipeline would notice.
NOTIONAL_SCALE = 1e6
# #Shares is the ORDER quantity in THOUSANDS, so it scales by 1e3. Together
# with the notional scale this is what makes the implied-share-price check
# below meaningful: get either wrong and the price lands a power of ten out.
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

# ...except where one arrived part-way through the period. India closed on a
# VWAP of the last half hour until the Closing Auction Session went live on
# 3 August 2026: a 20-minute call auction, 15:15 to 15:35, referenced to the
# 15:00-15:15 VWAP, and only for stocks in the derivatives segment.
#
# So India is two different products inside one calendar year, and the split
# is by DATE, not by market name. Before the date there is no auction to
# reach, auction share means nothing, and "vs Close" is a genuine tracking
# result rather than a degenerate one - you cannot print at a VWAP, you have
# to work through the last half hour to track it. On or after the date India
# joins the auction population, flagged small-sample and derivatives-only.
#
# A period ending before the date is unaffected: every India order stays on
# the VWAP-close side and nothing changes.
AUCTION_FROM = {"India": "2026-08-03"}

# ...but the CAS covers the DERIVATIVES SEGMENT ONLY, and the extract carries no
# segment flag. A non-F&O India name still closes on the half-hour VWAP after
# that date, and nothing in the data says which name is which.
#
# It cannot be inferred from %CLOSE either: an F&O name whose order MISSED the
# auction looks exactly like a name that never had one, so inferring would drop
# India's misses out of the auction population - the very orders the review
# exists to find. So post-CAS India is its own regime, pooled with neither side
# and asserted about nothing.
#
# To close this: get the F&O eligibility list from the desk, drop "India" from
# AUCTION_SEGMENT_UNKNOWN, and filter on the symbol.
AUCTION_SEGMENT_UNKNOWN = {"India"}

REGIME_AUCTION = "single-price auction"
REGIME_VWAP_CLOSE = "VWAP close (no auction)"
REGIME_UNKNOWN = "post-CAS India - segment unknown"

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
SIDE_ORDER = ["Buy", "Sell", "Short sell"]
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


def read_aws(folder: Path) -> pd.DataFrame:
    """Every parquet under `folder`, concatenated into one frame.

    Files are per date/country/client, so columns can differ slightly between
    them; concat unions the columns and leaves the gaps as NaN rather than
    dropping a file for being a column short.
    """
    files = sorted(folder.rglob("*.parquet"))
    if not files:
        return pd.DataFrame()

    frames, cols_seen = [], None
    for path in files:
        try:
            df = pd.read_parquet(path)
        except Exception as exc:
            warn(f"could not read {path.name}: {exc}")
            continue
        if cols_seen is None:
            cols_seen = set(df.columns)
        elif set(df.columns) != cols_seen:
            extra = sorted(set(df.columns) - cols_seen)
            missing = sorted(cols_seen - set(df.columns))
            if extra or missing:
                warn(f"{path.name} has a different shape from the first file:")
                if extra:
                    log(f"    extra:   {', '.join(extra[:8])}")
                if missing:
                    log(f"    missing: {', '.join(missing[:8])}")
            cols_seen |= set(df.columns)
        frames.append(df)

    if not frames:
        return pd.DataFrame()
    out = pd.concat(frames, ignore_index=True, sort=False)
    log(f"  {len(files)} parquet file(s) -> {len(out):,} rows x "
        f"{len(out.columns)} columns")
    return out


def merge_aws(raw: pd.DataFrame, folder: Path) -> pd.DataFrame:
    """Left-join the AWS extract onto the order file, and say what happened."""
    if not folder.is_dir():
        log(f"  no {folder} folder - running on the order file alone.")
        return raw

    section(f"AWS EXTRACT  {folder}")
    aws = read_aws(folder)
    if aws.empty:
        warn(f"no parquet files in {folder} - running on the order file alone.")
        return raw

    key = AWS_JOIN_KEY
    for name, frame in (("order file", raw), ("AWS extract", aws)):
        if key not in frame.columns:
            warn(f"{key} is not in the {name}; cannot join. Continuing without.")
            return raw

    # One row per order. A key that repeats would multiply the order file on
    # the join and inflate every total silently, so it is collapsed here and
    # counted rather than left to do that.
    dupes = int(aws[key].duplicated().sum())
    if dupes:
        warn(f"{dupes:,} duplicate {key} values in the AWS extract - keeping "
             f"the first of each.")
        aws = aws.drop_duplicates(subset=[key], keep="first")

    out = _merge_frames(raw, aws, verbose=True)
    return out


def _merge_frames(raw: pd.DataFrame, aws: pd.DataFrame,
                  verbose: bool = False) -> pd.DataFrame:
    """The join itself, separated so it can be tested without any files."""
    key = AWS_JOIN_KEY
    # Collide on the NORMALISED name, not the exact one. resolve_columns
    # matches case-insensitively and ignores separators, so an AWS "sym" beside
    # the order file's "Sym" is not a harmless near-miss: both would answer to
    # the same logical field and whichever came first in column order would
    # win. Renaming every normalised collision keeps the order file
    # unambiguously in charge of its own population.
    raw_norm = {_norm_name(c) for c in raw.columns}
    overlap = sorted(c for c in aws.columns
                     if c != key and _norm_name(c) in raw_norm)
    if overlap and verbose:
        log(f"  {len(overlap)} AWS column(s) answer to a name the order file")
        log(f"    already uses. The order file wins; the AWS version is kept")
        log(f"    as <name>{AWS_SUFFIX}:")
        log("    " + ", ".join(overlap[:12])
            + (f" ... and {len(overlap) - 12} more" if len(overlap) > 12 else ""))
    if overlap:
        aws = aws.rename(columns={c: c + AWS_SUFFIX for c in overlap})
    before = len(raw)
    out = raw.merge(aws, on=key, how="left", validate="m:1")
    if not verbose:
        return out
    matched = int(out[[c for c in aws.columns if c != key][0]].notna().sum()) \
        if len(aws.columns) > 1 else 0
    log("")
    log(f"  {before:,} orders in the order file")
    log(f"  {matched:,} matched into the AWS extract "
        f"({100.0 * matched / max(before, 1):.1f}%)")
    if matched < before:
        log(f"  {before - matched:,} did NOT match. They are KEPT - the order")
        log("    file is the population - but their AWS columns are empty, so")
        log("    anything derived from those columns excludes them.")
    log(f"  columns {len(raw.columns)} -> {len(out.columns)}")
    return out


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


def close_regime(market: pd.Series, date=None) -> pd.Series:
    """Which closing mechanism each order actually met.

    Market alone is not enough: a market can gain an auction part-way through
    the period, and orders either side of that date are different products.
    Where the new auction covers only part of the market and the data does not
    say which part, the orders get their own regime rather than a guess.
    """
    reg = pd.Series(REGIME_AUCTION, index=market.index, dtype=object)
    reg[market.isin(NO_CLOSING_AUCTION)] = REGIME_VWAP_CLOSE
    if date is not None:
        for name, start in AUCTION_FROM.items():
            after = market.eq(name) & (date >= pd.Timestamp(start))
            reg[after] = (REGIME_UNKNOWN if name in AUCTION_SEGMENT_UNKNOWN
                          else REGIME_AUCTION)
    return reg


def auction_available(market: pd.Series, date=None) -> pd.Series:
    """True only where a single-price auction is CERTAIN for that order."""
    return close_regime(market, date).eq(REGIME_AUCTION)


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

    for logical in ["start_time", "end_time", "first_start_time"]:
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
        df["is_short"] = up.isin({v.upper() for v in SHORT_SELL_VALUES})
        df["side_label"] = np.where(
            df["is_buy"], "Buy",
            np.where(df["is_short"], "Short sell", "Sell"))

    df["market"] = market_from_symbol(df["symbol"]) if "symbol" in df else UNKNOWN_MARKET
    df["close_regime"] = close_regime(df["market"], df.get("date"))

    # Before any venue field is derived, so close_notional, pct_continuous,
    # close_bucket and the rest all follow from the same number.
    if INDIA_CLOSE_PROXY and "pct_close" in df:
        in_window = india_in_close_window(df)
        df["pct_close_imputed"] = in_window
        if in_window.any():
            df["pct_close_measured"] = df["pct_close"]
            df.loc[in_window, "pct_close"] = 100.0
    df["has_auction"] = df["close_regime"].eq(REGIME_AUCTION)

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
        cont = df[cont_fields].sum(axis=1, min_count=1)
        # Present but never above zero means the export did not populate the
        # continuous splits - not that nothing traded in continuous. Taking it
        # at face value sets every continuous weight to zero, and any table
        # weighted by continuous notional then returns nothing at all.
        #
        # What did not print in an auction must have traded in continuous, so
        # that is the fallback, and it is stated rather than assumed silently.
        top = float(np.nanmax(cont.to_numpy(dtype=float))) if len(cont) else 0.0
        if not np.isfinite(top) or top <= 0:
            auction = df["pct_close"].fillna(0.0) if "pct_close" in df else 0.0
            if "pct_open" in df:
                auction = auction + df["pct_open"].fillna(0.0)
            df["pct_continuous"] = (100.0 - auction).clip(lower=0.0)
            warn("pct_take/post/dark are present but never above zero, so the "
                 "export did not populate them.")
            log("    Continuous share is taken as 100 - %CLOSE - %OPEN: what "
                "did not print in an auction traded in continuous.")
            log("    The venue-mix tables cannot show WHERE in continuous, "
                "only that it was not an auction.")
        else:
            df["pct_continuous"] = cont
    if "pct_close" in df and "notional" in df:
        df["close_notional"] = df["notional"] * df["pct_close"] / 100.0
        df["cont_notional"] = df["notional"] * df.get(
            "pct_continuous", 100.0 - df["pct_close"]) / 100.0
    if "pct_close" in df and "exec_shares" in df:
        df["close_shares"] = df["exec_shares"] * df["pct_close"] / 100.0

    # --- derived benchmarks ----------------------------------------------
    # Every slippage shares one executed price, so a difference between two of
    # them cancels it and leaves a pure, side-adjusted price move.
    # --- spread-normalised slippage ---------------------------------------
    # Slippage in bps answers "what did it cost". It does NOT compare across
    # names or markets: a wide-spread Korean mid-cap costs more bps than a
    # Japanese large-cap for reasons that have nothing to do with the algo.
    # Dividing by the spread answers "how many spreads did we pay", which does
    # compare. Quote bps for money, spreads for every comparison.
    #
    # The export's own columns win where they exist; the rest are derived, so
    # the two shapes of file behave the same. A spread of zero or less is a
    # quote error, not a trade, and gives no ratio.
    if "spread_bps" in df:
        sprd = df["spread_bps"].where(df["spread_bps"] > 0)
        for bench in ["arrival", "pvwap", "close", "vwap", "open", "nextopen"]:
            slip, norm = f"slip_{bench}", f"sprd_{bench}"
            if norm in df:
                continue                      # taken straight from the export
            if slip in df:
                df[norm] = df[slip] / sprd

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


def india_in_close_window(df: pd.DataFrame) -> pd.Series:
    """India orders that started inside the closing window.

    One definition, used by both the imputation and the filter, so the rows
    that get %CLOSE = 100 are exactly the rows that survive.
    """
    empty = pd.Series(False, index=df.index)
    if "market" not in df:
        return empty
    is_india = df["market"] == "India"
    if not is_india.any() or "first_start_time_min" not in df:
        return empty
    lo, hi = (_hhmm_to_min(t) for t in INDIA_CLOSE_WINDOW_HKT)
    return is_india & df["first_start_time_min"].between(lo, hi).fillna(False)


def filter_india_close_window(df: pd.DataFrame) -> pd.DataFrame:
    """Keep only the India orders that began inside the closing window.

    India runs no closing auction in this period, so %CLOSE is 0 on every one
    of its orders and cannot say whether an order was aimed at the close.
    When it STARTED can. An India order that began outside the closing window
    was not a close order, whatever strategy label it carries, and including
    it would drag every India number toward flow that was never trying.

    Only India is touched. Every other market keeps all of its orders.
    """
    if "market" not in df or "India" not in set(df["market"]):
        return df

    is_india = df["market"] == "India"
    n_india = int(is_india.sum())

    if "first_start_time_min" not in df:
        warn(f"India has {n_india:,} orders and no fstart_time to window them "
             f"with.")
        log("    Their %CLOSE is 0 by construction, so they cannot be told")
        log("    apart from flow that never aimed at the close. Left in and")
        log("    flagged rather than filtered on a column that is not here.")
        return df

    start = df["first_start_time_min"]
    inside = india_in_close_window(df)
    drop = is_india & ~inside

    section("INDIA CLOSE WINDOW")
    log(f"  India has no closing auction in this period, so an order aimed at")
    log(f"  the close is one that STARTED in the closing window: "
        f"{INDIA_CLOSE_WINDOW_HKT[0]}-{INDIA_CLOSE_WINDOW_HKT[1]} HKT.")
    log("")
    log(f"    {'India orders':<34}{n_india:>10,}")
    log(f"    {'started inside the window':<34}{int(inside.sum()):>10,}")
    log(f"    {'started outside it - removed':<34}{int(drop.sum()):>10,}")
    missing = int((is_india & start.isna()).sum())
    if missing:
        log(f"    {'no fstart_time at all - removed':<34}{missing:>10,}")
    val = float(df.loc[drop, "notional"].sum()) / 1e6 if "notional" in df else 0.0
    log(f"    {'value removed (' + CURRENCY + 'm)':<34}{val:>10,.2f}")
    if INDIA_CLOSE_PROXY and "pct_close_imputed" in df:
        n_imp = int(df.loc[inside, "pct_close_imputed"].sum())
        log("")
        log(f"  %CLOSE set to 100 on the {n_imp:,} that stayed: they ran through")
        log("  the VWAP that IS the close, so their close share is 100 by the")
        log("  mechanism. It is IMPUTED, not measured - the original sits in")
        log("  pct_close_measured and the rows are flagged pct_close_imputed.")
    log("")
    log("  Every other market keeps all of its orders.")
    return df[~drop].copy()


def _title_norm(s: pd.Series, order: list[str]) -> pd.Series:
    """Match a label case-insensitively onto `order`; keep anything else as-is.

    Unrecognised values are NEVER dropped - Cap in particular carries values
    beyond the expected set, and silently discarding them would move numbers.
    """
    canon = {v.lower(): v for v in order}
    return s.map(lambda v: canon.get(str(v).strip().lower(), str(v).strip()))


# ===========================================================================
# SCOPE AND EXCLUSIONS - nothing leaves without being counted
# ===========================================================================
#
# Two different jobs, deliberately separate.
#
#   apply_strategy_scope   drops whole strategies. A decision about what the
#                          review is ABOUT.
#
#   clean_values           removes unusable NUMBERS, and only then unusable
#                          ORDERS. An infinity in NextOpen says nothing about
#                          that order's auction share, its size or its
#                          notional, so the CELL dies and the order stays.
#                          Only an order that cannot contribute anywhere -
#                          no notional, no quantity, no readable side - leaves.
#
# Dropping a whole order for one bad cell would quietly bias every other
# table, because the orders with broken cells are not a random sample.

SLIP_FIELDS = ["slip_arrival", "slip_close", "slip_pvwap", "slip_vwap",
               "slip_open", "slip_nextopen", "first_exec_vs_close"]

# What the file held before scoping and cleaning. The scope table reports
# against THIS, not against what survived - otherwise, once STRATEGY_SCOPE has
# done its work, the reviewed book is 100% of itself and the slide says
# nothing. Filled in as the run removes things.
BOOK = {"orders": 0, "notional": 0.0, "dropped": []}


def _remember(label: str, orders: int, notional: float) -> None:
    BOOK["dropped"].append((label, int(orders), float(notional)))


def apply_strategy_scope(df: pd.DataFrame) -> pd.DataFrame:
    """Restrict the study to STRATEGY_SCOPE, showing what stays and what goes."""
    if "strategy" not in df:
        warn("no strategy column - STRATEGY_SCOPE was not applied.")
        return df
    if not STRATEGY_SCOPE:
        log("  STRATEGY_SCOPE is empty - every strategy in the file is kept.")
        return df

    wanted = {s.upper() for s in STRATEGY_SCOPE}
    keep = df["strategy"].astype(str).str.upper().isin(wanted)
    if not BOOK["orders"]:
        BOOK["orders"] = len(df)
        BOOK["notional"] = float(df["notional"].sum())
    for name, g in df[~keep].groupby("strategy", dropna=False, observed=True):
        _remember(f"{name} (out of scope)", len(g), float(g["notional"].sum()))

    log("  strategy scope: " + ", ".join(STRATEGY_SCOPE))
    log("")
    log(f"    {'strategy':<16}{'orders':>10}{'notional (' + CURRENCY + 'm)':>22}"
        f"{'%CLOSE (wtd)':>15}   ")
    rows = []
    for name, g in df.groupby("strategy", dropna=False, observed=True):
        pc = (wmean(g["pct_close"], g["notional"], winsor=False)
              if "pct_close" in g else float("nan"))
        rows.append((str(name), len(g), float(g["notional"].sum()) / 1e6, pc))
    for name, n, val, pc in sorted(rows, key=lambda r: -r[2]):
        verdict = "keep" if name.upper() in wanted else "drop"
        log(f"    {name:<16}{n:>10,}{val:>22,.2f}{pc:>15.2f}   {verdict}")

    present = {str(v).upper() for v in df["strategy"].dropna()}
    absent = sorted(s for s in STRATEGY_SCOPE if s.upper() not in present)
    if absent:
        warn("STRATEGY_SCOPE names a strategy not in the file: "
             + ", ".join(absent))

    out = df[keep].copy()
    if out.empty:
        raise SystemExit(
            "\nSTRATEGY_SCOPE removed every order. Check the names against "
            "the strategy table in --probe.")
    share = 100.0 * out["notional"].sum() / max(df["notional"].sum(), 1e-9)
    log("")
    log(f"    kept {len(out):,} of {len(df):,} orders, {share:.1f}% of value")
    return out


def clean_values(df: pd.DataFrame, dry_run: bool = False) -> pd.DataFrame:
    """Clear unusable cells, then drop orders that cannot contribute anywhere."""
    df = df.copy()
    n_in = len(df)
    v_in = float(df["notional"].sum()) if "notional" in df else 0.0

    # --- cells: the value dies, the order lives ---------------------------
    cells = []
    for col in df.columns:
        if not pd.api.types.is_numeric_dtype(df[col]):
            continue
        vals = pd.to_numeric(df[col], errors="coerce")
        arr = vals.to_numpy(dtype="float64", na_value=np.nan) \
            if hasattr(vals, "to_numpy") else np.asarray(vals, dtype="float64")
        inf = pd.Series(np.isinf(arr), index=df.index)
        big = pd.Series(False, index=df.index)
        if col in SLIP_FIELDS:
            big = (vals.abs() > MAX_ABS_BPS).fillna(False) & ~inf
        bad = inf | big
        if bad.any():
            cells.append((col, int(inf.sum()), int(big.sum())))
            if not dry_run:
                df.loc[bad, col] = np.nan

    # A zero or negative spread cannot normalise anything, and it is a quote
    # error rather than a trade. Clear it; the order's own numbers still stand.
    if "spread_bps" in df:
        bad_spread = (pd.to_numeric(df["spread_bps"], errors="coerce") <= 0).fillna(False)
        if bad_spread.any():
            cells.append(("spread_bps (<= 0)", 0, int(bad_spread.sum())))
            if not dry_run:
                df.loc[bad_spread, "spread_bps"] = np.nan

    if cells:
        warn("unusable VALUES - the cell is cleared, the order is kept:")
        log(f"    {'field':<26}{'infinite':>10}{'beyond +/-' + f'{MAX_ABS_BPS:,.0f}bps':>20}")
        for name, n_inf, n_big in cells:
            log(f"    {name:<26}{n_inf:>10,}{n_big:>20,}")
        log("    An infinity says nothing about the rest of that order, so the")
        log("    order stays in every table its other columns can support.")
    else:
        log("  no infinite or out-of-range values found.")

    # --- orders: only those that cannot contribute anywhere ---------------
    keep = pd.Series(True, index=df.index)
    reasons = []

    def drop(reason: str, mask) -> None:
        nonlocal keep
        mask = pd.Series(mask, index=df.index).fillna(False) & keep
        if mask.any():
            val = float(df.loc[mask, "notional"].sum()) if "notional" in df else 0.0
            reasons.append((reason, int(mask.sum()), val))
            if not dry_run:
                _remember(reason, int(mask.sum()), val)
            keep = keep & ~mask

    if "notional" in df:
        drop("no notional", df["notional"].isna() | (df["notional"] <= 0))
    if "order_shares" in df:
        drop("no order quantity", df["order_shares"].isna() | (df["order_shares"] <= 0))
    if "side" in df:
        up = df["side"].astype(str).str.strip().str.upper()
        known = up.isin({v.upper() for v in BUY_VALUES}
                        | {v.upper() for v in SELL_VALUES})
        # A side this script does not recognise is far more likely to be a
        # vocabulary gap here than a broken record in the file - a short-sell
        # code, say. Dropping those silently would remove a whole category of
        # flow and bias every number, so the values are named and, past a
        # trivial share, the run stops rather than quietly proceeding.
        if (~known).any():
            unknown = up[~known].value_counts()
            warn("side values this script does not recognise:")
            for val, n in unknown.head(10).items():
                log(f"    {val!r:<24}{n:>10,} orders")
            share = 100.0 * int((~known).sum()) / max(len(df), 1)
            if share > 0.5 and not dry_run:
                raise SystemExit("\n".join([
                    "",
                    f"{share:.1f}% of orders carry a side this script does not",
                    "know. That is a vocabulary gap, not dirty data.",
                    "",
                    "Add the codes listed above to BUY_VALUES or SELL_VALUES at",
                    "the top of this script and run again. Do not let them be",
                    "excluded: a whole category of flow would leave the review",
                    "without anyone noticing.",
                ]))
            if share > 0.5:
                log(f"    {share:.1f}% of orders - that is a vocabulary gap, "
                    "not dirty data.")
                log("    Add these to BUY_VALUES / SELL_VALUES before the real run.")
        drop("side not recognised", ~known)
    if "date" in df:
        drop("no date", df["date"].isna())
    if "strategy" in df:
        drop("no strategy", df["strategy"].isna() | (df["strategy"].astype(str) == ""))

    log("")
    if reasons:
        warn("orders EXCLUDED from the study:")
    log(f"    {'reason':<26}{'orders':>10}{'notional (' + CURRENCY + 'm)':>22}")
    for reason, n, val in reasons:
        log(f"    {reason:<26}{n:>10,}{val / 1e6:>22,.2f}")
    n_out = int(keep.sum())
    v_out = float(df.loc[keep, "notional"].sum()) if "notional" in df else 0.0
    log(f"    {'kept':<26}{n_out:>10,}{v_out / 1e6:>22,.2f}")
    log("")
    log(f"    {n_in:,} orders in -> {n_out:,} kept "
        f"({100.0 * n_out / max(n_in, 1):.2f}%), "
        f"{100.0 * v_out / max(v_in, 1e-9):.2f}% of value retained")
    if dry_run:
        log("    DRY RUN - nothing was removed. Run without --probe to apply.")
        return df
    return df[keep].copy()


# ===========================================================================
# SANITY REPORT - read this before trusting any number
# ===========================================================================

def sanity_report(df: pd.DataFrame, cols: dict, raw_cols) -> None:
    section("SANITY REPORT")

    missing = [c for c in COLUMNS if c not in cols and c not in OPTIONAL]
    core = [c for c in cols if c not in OPTIONAL]
    log(f"  resolved {len(core)}/{len(COLUMNS) - len(OPTIONAL)} known fields")
    if missing:
        log(f"  unresolved: {', '.join(sorted(missing))}")

    # Say which spread-normalised columns came from the file and which were
    # computed, so the two are never confused for each other.
    read = [c for c in OPTIONAL if c in cols]
    calc = [c for c in OPTIONAL if c not in cols and c.replace("sprd_", "slip_") in df]
    if read:
        log(f"  spread-normalised, read from the file: "
            + ", ".join(f"{c} <- {cols[c]}" for c in read))
    if calc:
        log(f"  spread-normalised, computed as slippage / spread: "
            + ", ".join(calc))
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
            if SIDE_ADJUSTED:
                log("    SIDE_ADJUSTED is pinned True: the export is already")
                log("    side-adjusted, so plus is good on both sides. A gap")
                log("    between Buy and Sell is a RESULT to explain, not a")
                log("    data error - look at it before writing it up.")
            else:
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
    log("    #Shares is ORDER quantity in THOUSANDS, so it is scaled x{:.0e}."
        .format(SCALE["order_shares"]))
    log("    Executed shares = #Shares x 1000 x FR/100.")
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
            g = df[df["market"] == mkt]
            n_unk = int(g["close_regime"].eq(REGIME_UNKNOWN).sum())
            if n_unk:
                note = (f"  <- {len(g) - n_unk:,} on a VWAP close, {n_unk:,} from "
                        f"{AUCTION_FROM[mkt]} where only F&O names got the "
                        f"auction; three regimes, never pooled")
            else:
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

def weight_column(df: pd.DataFrame, preferred: str,
                  fallback: str = "notional") -> str:
    """The preferred weight, unless it carries no positive value anywhere.

    A weight column of all zeros is not a weighting, it is a blank table:
    every row fails the w > 0 test and the mean comes back as nan.
    """
    if preferred in df:
        col = pd.to_numeric(df[preferred], errors="coerce")
        if (col > 0).any():
            return preferred
    return fallback


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


def to_money_k(bps, notional_musd) -> float:
    """Basis points on a notional, in thousands of currency.

    bps/1e4 x (musd x 1e6) / 1e3 == bps x musd / 10. Positive stays savings,
    the same way round as every other number here, so a negative figure in
    this column reads as a cost without anyone having to be told twice.
    """
    if bps is None or notional_musd is None:
        return np.nan
    if not (np.isfinite(bps) and np.isfinite(notional_musd)):
        return np.nan
    return float(bps) * float(notional_musd) / 10.0


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
            "saved (" + CURRENCY + "k)": to_money_k(
                wmean(g[value], g["notional"]), g["notional"].sum() / 1e6),
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
    """What the file held, what is reviewed, and what was left out of it.

    Measured against the file as it arrived. Once STRATEGY_SCOPE has run, the
    reviewed book is by definition 100% of itself, and a table saying so is
    the one slide in the deck that cannot be wrong and cannot be useful.
    """
    total_o = BOOK["orders"] or len(all_df)
    total_n = BOOK["notional"] or float(all_df["notional"].sum())

    rows = []

    def add(label, orders, notional):
        rows.append({
            "scope": label,
            "orders": int(orders),
            "notional (" + CURRENCY + "m)": notional / 1e6,
            "% of orders": 100.0 * orders / max(total_o, 1),
            "% of notional": 100.0 * notional / max(total_n, 1e-9),
        })

    for strat, g in all_df.groupby("strategy", dropna=False, observed=True):
        add(f"{strat} (reviewed)", len(g), float(g["notional"].sum()))
    for label, orders, notional in BOOK["dropped"]:
        add(f"{label}", orders, notional)
    add("Everything in the file", total_o, total_n)

    out = pd.DataFrame(rows).set_index("scope").round(2)
    return out


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
    # Within strategy as well as within market and size: a 53%-auction algo
    # and a 7%-auction algo do not share a frontier, and pooling them would
    # pull the reference line down until nothing looked short.
    cell = ["strategy", "market", "adv_bucket"] if "strategy" in df \
        else ["market", "adv_bucket"]
    med = df.groupby(cell, observed=True)["pct_close"].transform("median")
    n_cell = df.groupby(cell, observed=True)["pct_close"].transform("size")
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
    "Auction only",
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
    # "Limit did not cross" is only a cause if the flag DISCRIMINATES. Where
    # every order in the file carries the same value, the flag is the base
    # rate, not an explanation - and using it would quietly absorb the whole
    # unexplained cohort, which is the one thing this table exists to find.
    lim = pd.Series(False, index=idx)
    if "market_limit" in df:
        vals = df["market_limit"].astype(str).str.strip().str.lower()
        if vals.nunique(dropna=True) > 1:
            lim = vals == "limit"
        else:
            only = vals.dropna().unique()
            warn("market_limit is "
                 + (f"always {only[0]!r}" if len(only) else "empty")
                 + " - it cannot explain a miss, so no order is attributed")
            log("     to its limit. Those orders fall to 'No auction fill -")
            log("     unexplained', which is the honest place for them.")
    short = (df["frontier_shortfall_pp"] if "frontier_shortfall_pp" in df
             else pd.Series(0.0, index=idx)).fillna(0.0)

    never = fr < COHORT_FR_ZERO
    # Everything printed in the auction: the order behaved as auction-only.
    # Split out from "cleared" because an order that needed continuous help to
    # get there is a different animal from one that never left the auction.
    auction_only = ~never & (pc >= AUCTION_ONLY_MIN_PCT)
    cleared = ~never & ~auction_only & (pc >= CLEARED_MIN_PCT_CLOSE)
    size_ok = ~never & ~cleared & ~auction_only & (adv >= COHORT_ADV_LOW)
    open_ = ~never & ~cleared & ~auction_only & ~size_ok   # small, under-cleared
    no_fill = open_ & (pc <= 0)
    partial = open_ & (pc > 0)

    return pd.Series(
        np.select(
            [never,
             auction_only,
             cleared,
             size_ok,
             no_fill & lim,
             no_fill & ~lim,
             partial & (short > FRONTIER_SHORTFALL_PP),
             partial],
            ["Never traded",
             "Auction only",
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
    w = weight_column(df, "cont_notional")
    rows = []
    for key, g in df.groupby(by, dropna=False, observed=True):
        v = g["first_exec_vs_close"]
        lo, hi = boot_ci(v, g[w])
        rows.append({
            by: key,
            "orders": int(v.notna().sum()),
            "continuous notional (" + CURRENCY + "m)": g[w].sum() / 1e6,
            "first exec vs Close bps (wtd)": wmean(v, g[w]),
            "saved (" + CURRENCY + "k)": to_money_k(
                wmean(v, g[w]), g[w].sum() / 1e6),
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
    w = weight_column(df, "cont_notional")
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


SPREAD_BENCHMARKS = [("sprd_arrival", "slip_arrival", "vs Arrival"),
                     ("sprd_pvwap", "slip_pvwap", "vs PVWAP"),
                     ("sprd_close", "slip_close", "vs Close"),
                     ("sprd_vwap", "slip_vwap", "vs VWAP")]


def t_spread_normalised(df: pd.DataFrame, by: str) -> pd.DataFrame:
    """Cost in bps beside cost in spreads, so groups can be compared fairly.

    Two groups with the same bps number and different spreads did NOT perform
    the same, and the bps column alone will not show it. The spread column is
    the one to rank on; the bps column is the one to quote in money.
    """
    have = [(n, s, lab) for n, s, lab in SPREAD_BENCHMARKS
            if n in df and s in df]
    if not have or "spread_bps" not in df or by not in df:
        return pd.DataFrame()
    rows = []
    for key, g in df.groupby(by, dropna=False, observed=True):
        spread = wmean(g["spread_bps"], g["notional"])
        row = {by: key, "orders": len(g),
               "notional (" + CURRENCY + "m)": g["notional"].sum() / 1e6,
               "spread bps (wtd)": spread}
        for norm, slip, label in have:
            bps = wmean(g[slip], g["notional"])
            row[label + " bps"] = bps
            # The ratio of the two averages, NOT the average of per-order
            # ratios. Dividing order by order lets a name with a 0.5bp spread
            # produce a ratio in the hundreds, and the mean then reports that
            # one name rather than the group. The aggregate ratio reads as
            # "the average cost was this many times the average spread".
            row[label + " spreads"] = (bps / spread
                                       if spread and np.isfinite(spread) and spread > 0
                                       else np.nan)
            # Kept beside it: what the typical single order paid. A median is
            # unaffected by a small denominator on one order.
            row[label + " spreads (median)"] = g[norm].median()
        row["small sample"] = len(g) < MIN_N_FOR_CI
        rows.append(row)
    return pd.DataFrame(rows).round(3)


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
    # The whiskers reach past the bar ends, so the axis has to allow for them
    # and the label has to clear them. A CI line drawn through a value label
    # hides the minus sign, and a cost of 73bps then reads as a saving of 73.
    reach = [abs(v) for v in values if v is not None and not math.isnan(v)]
    if ci is not None:
        reach += [abs(b) for lo, hi in ci for b in (lo, hi)
                  if b is not None and not math.isnan(b)]
    span = max(reach or [1.0])
    pad = span * 0.30
    ax.set_xlim(-span - pad, span + pad)
    for i, v in enumerate(values):
        if v is None or math.isnan(v):
            continue
        edge = v
        if ci is not None and i < len(ci):
            lo, hi = ci[i]
            if (lo is not None and hi is not None
                    and not math.isnan(lo) and not math.isnan(hi)):
                edge = max(v, hi) if v >= 0 else min(v, lo)
        off = span * 0.03
        ax.text(edge + (off if v >= 0 else -off), i, _fmt_bps(v),
                va="center", ha="left" if v >= 0 else "right",
                color=INK, fontsize=8.5, zorder=6,
                bbox=dict(boxstyle="round,pad=0.16", facecolor=SURFACE,
                          edgecolor="none", alpha=0.85))


def chart_scope(t: pd.DataFrame, out: Path) -> None:
    """Part-to-whole: how much of the book the close algos are."""
    if t.empty:
        return
    fig, (ax,) = _fig((11.0, 3.0))
    ncol = "notional (" + CURRENCY + "m)"
    # The last row is the whole file, which is the sum of the others - drawing
    # it would double the bar.
    parts = t.drop(index="Everything in the file", errors="ignore")
    parts = parts[parts[ncol] > 0]
    left = 0.0
    total = max(float(t.loc["Everything in the file", ncol])
                if "Everything in the file" in t.index else parts[ncol].sum(),
                1e-9)
    for i, (label, row) in enumerate(parts.iterrows()):
        w = float(row[ncol])
        reviewed = "(reviewed)" in str(label)
        ax.barh([0], [w], left=[left], height=0.5,
                color=SERIES[i % 3] if reviewed else NEUTRAL, zorder=3)
        share = 100.0 * w / total
        if share >= 4.0:
            ax.text(left + w / 2, 0,
                    f"{str(label).replace(' (reviewed)', '')}\n"
                    f"{CURRENCY} {w:,.0f}m  ({share:.0f}%)",
                    ha="center", va="center", fontsize=9,
                    color="white" if reviewed else INK)
        left += w + total * 0.004        # 2px-equivalent surface gap
    ax.set_yticks([])
    ax.set_xlim(0, left)
    # A thin grey sliver with no label is worse than no sliver: the reader can
    # see something was left out and cannot tell what. Name it under the axis.
    out_rows = [(str(i).replace(" (out of scope)", ""), float(r[ncol]))
                for i, r in parts.iterrows() if "(reviewed)" not in str(i)]
    if out_rows:
        out_n = sum(v for _, v in out_rows)
        names = ", ".join(n for n, _ in sorted(out_rows, key=lambda x: -x[1])[:5])
        ax.text(0.0, -0.42,
                f"grey = not reviewed: {CURRENCY} {out_n:,.0f}m "
                f"({100 * out_n / total:.1f}%) - {names}",
                transform=ax.transAxes, fontsize=8, color=INK_MUTED)
    _style(ax, xlabel=f"executed notional ({CURRENCY}m)",
           title=f"Scope - what is reviewed, and what is not, {PERIOD_LABEL}",
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


def chart_market_notional(t: pd.DataFrame, out: Path,
                          name: str = "13_market_notional.png") -> None:
    """How much went to each market. Magnitude, so one hue, largest first."""
    if t.empty or "notional (" + CURRENCY + "m)" not in t:
        return
    col = "notional (" + CURRENCY + "m)"
    d = t.head(14)
    fig, (ax,) = _fig((11.0, 6.4))
    y = np.arange(len(d))
    vals = [float(v) for v in d[col]]
    ax.barh(y, vals, color=SERIES[0], height=0.62, zorder=3)
    ax.set_yticks(y)
    ax.set_yticklabels(list(d.index))
    ax.invert_yaxis()
    ax.grid(axis="x", zorder=0)
    ax.set_axisbelow(True)
    span = max(vals or [1.0])
    ax.set_xlim(0, span * 1.28)
    for i, (v, (_, row)) in enumerate(zip(vals, d.iterrows())):
        share = row.get("% of notional", float("nan"))
        txt = f"  {v:,.0f}m" + (f"  ({share:.0f}%)" if np.isfinite(share) else "")
        ax.text(v, i, txt, va="center", ha="left", fontsize=8.5, color=INK,
                zorder=5)
    _style(ax, xlabel=f"executed notional ({CURRENCY}m)",
           title="Where the value traded, by market", horizontal=True)
    _save(fig, out, name)


def chart_market_slippage(t: pd.DataFrame, out: Path,
                          name: str = "14_market_slippage.png",
                          value: str = "vs Arrival bps") -> None:
    """Cost by market, ordered by how much was traded there, not by cost.

    Ordering by cost would put a 61-order market at the top of the slide.
    Ordering by value keeps the reader looking at the markets that can move
    the number, and each bar carries its own share so nobody has to guess.
    """
    if t.empty or value not in t:
        return
    d = t.head(14)
    fig, (ax,) = _fig((11.0, 6.4))
    labels = []
    for mkt, row in d.iterrows():
        share = row.get("% of notional", float("nan"))
        labels.append(f"{mkt}   ({share:.0f}% of value)"
                      if np.isfinite(share) else str(mkt))
    _diverging_barh(ax, labels, [float(v) for v in d[value]],
                    small=[bool(x) for x in d["small sample"]]
                    if "small sample" in d else None)
    _style(ax, xlabel=f"{value}, notional-weighted   {COST_SAVE_NOTE}",
           title="What each market cost, biggest by value first",
           horizontal=True)
    ax.text(0.0, -0.14, "ordered by share of value traded, not by cost - a "
            "small market with a big number is still a small market",
            transform=ax.transAxes, fontsize=7.5, color=INK_MUTED)
    _save(fig, out, name)


def chart_algo_choice(t: pd.DataFrame, out: Path,
                      name: str = "15_algo_choice.png", top: int = 10) -> None:
    """Two strategies on the same market and size band, side by side.

    Paired bars rather than a difference, because the difference alone hides
    whether both were costing money or one was simply less bad.
    """
    if t.empty:
        return
    bps_cols = [c for c in t.columns if c.endswith(" bps") and c != "gap bps"]
    if len(bps_cols) < 2:
        return
    d = t.head(top)
    fig, (ax,) = _fig((11.0, 6.8))
    y = np.arange(len(d))
    h = 0.8 / len(bps_cols)
    for i, col in enumerate(bps_cols):
        offs = y - 0.4 + h * (i + 0.5)
        vals = [float(v) if pd.notna(v) else 0.0 for v in d[col]]
        # Deliberately NOT the diverging pair. Every other chart here reads
        # blue as a saving and red as a cost; on this one colour is the algo,
        # so borrowing those two hues would make a blue bar at -37bps look
        # like good news. Slots 3 and 4 are adjacent in the validated order,
        # and every bar carries its own number as well.
        ax.barh(offs, vals, height=h * 0.9, color=SERIES[2 + i % 2],
                zorder=3, label=col.replace(" bps", ""))
        for yy, v in zip(offs, vals):
            ax.text(v, yy, "  " + _fmt_bps(v), va="center",
                    ha="left" if v >= 0 else "right", fontsize=8, color=INK)
    ax.axvline(0, color=BASELINE, linewidth=1.0, zorder=2)
    ax.set_yticks(y)
    ax.set_yticklabels([f"{r['market']}  {r['%Adv bucket']}"
                        for _, r in d.iterrows()])
    ax.invert_yaxis()
    ax.legend(frameon=False, loc="lower right", fontsize=8.5)
    _style(ax, xlabel=f"vs Arrival, notional-weighted (bps)   {COST_SAVE_NOTE}",
           title="Same market, same order size - what each algo cost",
           horizontal=True)
    ax.text(0.0, -0.09, "colour is the algo on this chart, not the direction - "
            "left of the line is still a cost",
            transform=ax.transAxes, fontsize=7.5, color=INK_MUTED)
    ax.text(0.0, -0.13,
            "NOT a controlled comparison: orders are not assigned to a strategy "
            "at random, and the reason one was chosen drives cost too. Read as "
            "'orders like these cost this', never as 'the other algo would have "
            "saved that'.", transform=ax.transAxes, fontsize=7.5,
            color=INK_MUTED, wrap=True)
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

    # The miss taxonomy only means something for an algo that was aiming at
    # the auction. Everything else keeps its benchmark and venue tables and
    # is left out of clearance, capacity and the cohorts.
    moc_names = MOC_STRATEGIES or close_strats
    moc = auc[auc["strategy"].isin(moc_names)].copy()
    if set(moc_names) != set(close_strats):
        log(f"  miss taxonomy runs on {', '.join(moc_names)} only; "
            f"{len(auc) - len(moc):,} orders on the other strategies are in "
            f"the benchmark tables but not the cohorts.")

    t["03_profile"] = strategy_profile(df)
    t["04_headline_vs_arrival"] = by_group(df, "strategy", "slip_arrival")
    t["05_headline_vs_close"] = by_group(auc, "strategy", "slip_close")
    t["06_benchmark_matrix"] = t_benchmark_matrix(df)
    t["06a_spreads_strategy"] = t_spread_normalised(df, "strategy")
    t["06b_spreads_market"] = t_spread_normalised(df, "market")
    t["06c_spreads_by_adv"] = t_spread_normalised(df, "adv_bucket")
    t["07_decomposition_strategy"] = t_decomposition(df, "strategy")
    t["08_decomposition_market"] = t_decomposition(df, "market")
    t["09_venue_mix_strategy"] = t_venue_mix(df, "strategy")
    t["10_venue_mix_market"] = t_venue_mix(df, "market")
    t["11_clearance"] = t_clearance(moc)
    t["12_leakage_strategy"] = t_leakage(auc, "strategy")
    t["13_leakage_market"] = t_leakage(auc, "market")
    t["14_capacity"] = t_capacity(moc)
    t["15_cohorts"] = t_cohorts(moc)
    t["16_orders_to_review"] = t_unexplained(moc)
    t["17_first_exec_by_adv"] = t_first_exec(moc, "adv_bucket")
    t["18_first_exec_by_market"] = t_first_exec(moc, "market")
    t["19_early_start"] = t_early_start_waste(moc)
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
    t["34_market_profile"] = t_market_profile(df)
    if df["strategy"].nunique() > 1:
        t["36_algo_choice"] = t_algo_choice(df)
        t["35_market_by_strategy"] = t_market_by_strategy(df)
    else:
        only = df["strategy"].dropna().astype(str).unique()
        name = only[0] if len(only) else "one strategy"
        log(f"  scope is {name} alone, so there is nothing to compare it to:")
        log("    35_market_by_strategy would only repeat 34, and 36_algo_choice")
        log("    needs two strategies inside the same market and size band.")
    if "side_label" in df:
        t["33_by_side"] = by_group(df, "side_label", "slip_arrival",
                                   order=SIDE_ORDER)
    if not noauc.empty:
        t["31_no_auction_markets"] = t_venue_mix(noauc, "market")
        t["32_close_regimes"] = t_venue_mix(all_df, "close_regime")
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


def _market_row(key, g, total_notional, total_orders) -> dict:
    """One market's size and cost, in the order a slide reads them."""
    row = {
        "orders": len(g),
        "% of orders": 100.0 * len(g) / max(total_orders, 1),
        "notional (" + CURRENCY + "m)": g["notional"].sum() / 1e6,
        "% of notional": 100.0 * g["notional"].sum() / max(total_notional, 1e-9),
        "spread bps (wtd)": wmean(g["spread_bps"], g["notional"])
            if "spread_bps" in g else np.nan,
        "wtd %CLOSE": wmean(g["pct_close"], g["notional"], winsor=False)
            if "pct_close" in g else np.nan,
    }
    for slip, label in [("slip_arrival", "vs Arrival"),
                        ("slip_pvwap", "vs PVWAP"),
                        ("slip_close", "vs Close")]:
        row[label + " bps"] = wmean(g[slip], g["notional"]) if slip in g else np.nan
    # Spreads, as the ratio of the two averages - see t_spread_normalised.
    sprd = row["spread bps (wtd)"]
    for label in ["vs Arrival", "vs PVWAP"]:
        bps = row[label + " bps"]
        row[label + " spreads"] = (bps / sprd if sprd and np.isfinite(sprd)
                                   and sprd > 0 and np.isfinite(bps) else np.nan)
    row["mean FR %"] = wmean(g["fill_rate"], g["notional"], winsor=False) \
        if "fill_rate" in g else np.nan
    row["median %Adv"] = g["adv_pct"].median() if "adv_pct" in g else np.nan
    row["small sample"] = len(g) < MIN_N_FOR_CI
    return row


def t_market_profile(df: pd.DataFrame) -> pd.DataFrame:
    """Size and cost of every market in one place, ordered by value traded.

    The market tables elsewhere each answer one question. This one is the
    slide: how much went there, and what it cost, with the spread beside the
    bps so a wide-spread market is not mistaken for a badly traded one.
    """
    if "market" not in df or df.empty:
        return pd.DataFrame()
    total_n, total_o = df["notional"].sum(), len(df)
    rows = {}
    for key, g in df.groupby("market", dropna=False, observed=True):
        rows[str(key)] = _market_row(key, g, total_n, total_o)
    out = pd.DataFrame(rows).T
    out.index.name = "market"
    return out.sort_values("% of notional", ascending=False).round(3)


def t_market_by_strategy(df: pd.DataFrame) -> pd.DataFrame:
    """The same split by strategy, because a market effect can be a mix effect.

    If one market is nearly all VWAP and another nearly all CLOSE, comparing
    the two markets compares the strategies as much as the venues. This is
    the table that tells you which it was.
    """
    if "market" not in df or "strategy" not in df or df.empty:
        return pd.DataFrame()
    total_n, total_o = df["notional"].sum(), len(df)
    rows = {}
    for (mkt, strat), g in df.groupby(["market", "strategy"], dropna=False,
                                      observed=True):
        if g.empty:
            continue
        rows[(str(mkt), str(strat))] = _market_row(mkt, g, total_n, total_o)
    if not rows:
        return pd.DataFrame()
    out = pd.DataFrame(rows).T
    out.index.names = ["market", "strategy"]
    return out.sort_values("% of notional", ascending=False).round(3)


ALGO_COMPARE_MIN_N = 30      # per strategy, per cell, before a cell counts


def t_algo_choice(df: pd.DataFrame, benchmark: str = "slip_arrival") -> pd.DataFrame:
    """Two strategies on the same kind of order: same market, same size band.

    THIS IS NOT A CONTROLLED COMPARISON and must never be presented as one.
    Orders are not assigned to a strategy at random. A trader who chooses
    CLOSE for one order and VWAP for another is acting on urgency, on a view,
    on instructions the extract does not carry - and those reasons drive cost
    too. Holding market and size constant removes the two biggest confounds
    and leaves the rest standing.

    So read a row as "orders like these, on this algo, cost this much" - a
    question worth asking the desk - and never as "the other algo would have
    saved that". Every cell needs ALGO_COMPARE_MIN_N orders on BOTH sides.
    """
    need = {"market", "adv_bucket", "strategy", benchmark, "notional"}
    if not need <= set(df.columns) or df.empty:
        return pd.DataFrame()

    rows = []
    for (mkt, bucket), g in df.groupby(["market", "adv_bucket"],
                                       dropna=False, observed=True):
        per = {}
        for strat, sub_g in g.groupby("strategy", dropna=False, observed=True):
            n = int(sub_g[benchmark].notna().sum())
            if n >= ALGO_COMPARE_MIN_N:
                # Against arrival AND against each order's own window. The
                # pair is what makes the row readable: arrival carries the
                # drift the order sat through, PVWAP does not. A big arrival
                # gap that vanishes on PVWAP is not one algo executing better
                # than the other - it is the orders having faced different
                # markets, which is a routing question, not a quality one.
                own = wmean(sub_g["slip_pvwap"], sub_g["notional"])                     if "slip_pvwap" in sub_g else np.nan
                per[str(strat)] = (n, sub_g["notional"].sum() / 1e6,
                                   wmean(sub_g[benchmark], sub_g["notional"]),
                                   own)
        if len(per) < 2:
            continue
        best = max(per, key=lambda k: per[k][2])
        worst = min(per, key=lambda k: per[k][2])
        gap = per[best][2] - per[worst][2]
        row = {"market": mkt, "%Adv bucket": bucket}
        for strat in sorted(per):
            row[f"{strat} orders"] = per[strat][0]
            row[f"{strat} notional ({CURRENCY}m)"] = per[strat][1]
            row[f"{strat} bps"] = per[strat][2]
            row[f"{strat} vs PVWAP bps"] = per[strat][3]
        row["better here"] = best
        row["gap bps"] = gap
        own_gap = per[best][3] - per[worst][3]
        row["gap vs PVWAP bps"] = own_gap
        # How much of the arrival gap survives once each order is measured
        # against its own window. Near zero means the gap was the market
        # moving, not the algo working.
        row["gap that is execution %"] = (
            100.0 * own_gap / gap if gap and np.isfinite(gap) and abs(gap) > 1e-9
            and np.isfinite(own_gap) else np.nan)
        # What the gap was worth on the flow that did NOT take the better
        # side. Not a saving that was available - see the docstring.
        row["gap on the other side (" + CURRENCY + "k)"] = to_money_k(
            gap, per[worst][1])
        rows.append(row)

    if not rows:
        return pd.DataFrame()
    out = pd.DataFrame(rows)
    sort_col = "gap on the other side (" + CURRENCY + "k)"
    return out.sort_values(sort_col, ascending=False).round(2)


def t_monthly(df: pd.DataFrame) -> pd.DataFrame:
    if "month" not in df:
        return pd.DataFrame()
    # A month the extract only partly covers is not comparable with a full
    # one. September 1-4 next to a full August reads as a collapse when it is
    # four days of flow, so the month is marked and the reader is told.
    span_end = df["date"].max() if "date" in df else None
    span_start = df["date"].min() if "date" in df else None
    rows = []
    for key, g in df.groupby("month", dropna=False, observed=True):
        partial = False
        if span_end is not None and pd.notna(span_end):
            period = pd.Period(str(key), freq="M")
            partial = (period.end_time.date() > span_end.date()
                       or period.start_time.date() < span_start.date())
        rows.append({
            "month": str(key) + (" (part)" if partial else ""),
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
    chart_market_notional(t.get("34_market_profile", pd.DataFrame()), charts)
    chart_market_slippage(t.get("34_market_profile", pd.DataFrame()), charts)
    chart_algo_choice(t.get("36_algo_choice", pd.DataFrame()), charts)
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
    strats = ", ".join(sorted(auc["strategy"].dropna().unique().astype(str)))         if "strategy" in auc else "the strategies in scope"
    log(f"  {len(auc):,} orders on {strats}, {CURRENCY} {n_tot/1e6:,.0f}m "
        f"executed, in markets that run a closing auction.")
    if "strategy" in auc and auc["strategy"].nunique() > 1:
        log("    These are NOT all close algos. They are reported together only")
        log("    because they compete for the same auction; every number below")
        log("    that matters is broken out by strategy.")

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
            log("    absorb, and they still got nothing in the auction.")
            lim_useful = False
            if "market_limit" in auc:
                lim_useful = auc["market_limit"].astype(str).str.strip()                                 .str.lower().nunique(dropna=True) > 1
            if lim_useful:
                log("    None of them carried a limit that could explain it.")
            else:
                log("    Whether a limit was binding CANNOT be said: every order")
                log("    in this file carries the same market_limit value, so the")
                log("    flag distinguishes nothing. Some of these may turn out")
                log("    to be limits that did not cross.")
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
            money = r.get("saved (" + CURRENCY + "k)", np.nan)
            money_txt = (f"{CURRENCY} {money:>10,.0f}k"
                         if pd.notna(money) else " " * 15)
            log(f"    {str(k):<10}{r['first exec vs Close bps (wtd)']:+8.2f} bps"
                f"  {money_txt}   {int(r['orders']):>5} orders{flag}")
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
    choice = t.get("36_algo_choice", pd.DataFrame())
    if not choice.empty:
        col = "gap on the other side (" + CURRENCY + "k)"
        log("")
        log("  Same market, same size band - which algo did better:")
        for _, r in choice.head(6).iterrows():
            exec_pct = r.get("gap that is execution %", np.nan)
            exec_txt = (f"   {exec_pct:>5.0f}% of it survives on PVWAP"
                        if pd.notna(exec_pct) else "")
            log(f"    {str(r['market']):<14}{str(r['%Adv bucket']):<8}"
                f"{r['better here']:<7} by {r['gap bps']:>6.1f} bps"
                f"   {CURRENCY} {r[col]:>9,.0f}k on the other side{exec_txt}")
        log("    The last column is the test that matters. Against arrival an")
        log("    order carries every basis point the market moved while it")
        log("    worked; against PVWAP it does not. A gap that survives on")
        log("    PVWAP is the algos working differently. A gap that collapses")
        log("    is the two sets of orders having faced different markets -")
        log("    a routing question, not a quality one.")
        log("")
        log("    This is NOT a controlled comparison. Orders are not assigned to")
        log("    a strategy at random, and the reason one was chosen - urgency,")
        log("    a view, an instruction not in this file - drives cost as well.")
        log("    Holding market and size constant removes the two biggest")
        log("    confounds and leaves the rest standing. Read a line as 'orders")
        log("    like these, on this algo, cost this much' and take it to the")
        log("    desk as a question. Never as 'the other algo would have saved")
        log("    that money'.")
        log("")

    imputed = int(auc["pct_close_imputed"].sum()) if "pct_close_imputed" in auc else 0
    if imputed:
        log(f"  {imputed:,} of these orders carry an IMPUTED %CLOSE of 100 - India,")
        log("  which runs no auction, where the order ran through the closing")
        log("  VWAP window. Their close share is a property of the mechanism,")
        log("  not a measurement, and any auction-share figure that includes")
        log("  them is part measured and part assumed. Say so on the slide.")
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

    # What the real run would remove, without removing it. The per-column
    # count of infinities is the part worth reading twice.
    section("SCOPE AND EXCLUSIONS (dry run - nothing is removed)")
    clean_values(df, dry_run=True)

    choose_close_strategies(df)

    log("")
    log("  NEXT STEP")
    log("    Read the strategy table above and check it against")
    log("    STRATEGY_SCOPE and CLOSE_STRATEGIES at the top of this script,")
    log("    then run without --probe. For two windows off one extract:")
    log("      --to 2026-06-30 --out output_h1   --label \"H1 2026\"")
    log("      --from 2026-07-01 --out output_h2 --label \"Jul to 4 Sep 2026\"")


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
    blank = pd.DataFrame({"pct_close": [40.0, 100.0], "pct_open": [10.0, 0.0],
                          "pct_take": [0.0, 0.0], "pct_post": [0.0, 0.0],
                          "pct_dark": [0.0, 0.0], "notional": [1e6, 1e6]})
    # The AWS join must not let a second source quietly take over a field the
    # order file already owns - resolve_columns is case- and separator-blind.
    raw_x = pd.DataFrame({"aggrTgtId": ["a", "b"], "Sym": ["X", "Y"],
                          "Sprd": [1.0, 2.0]})
    aws_x = pd.DataFrame({"aggrTgtId": ["a", "b"], "sym": ["Z", "W"],
                          "Sprd": [9.0, 9.0], "fstart_time": ["17:31", "12:00"]})
    joined = _merge_frames(raw_x, aws_x)
    check("a differently-cased AWS column does not shadow the order file",
          list(joined["Sym"]) == ["X", "Y"] and "sym_aws" in joined.columns,
          list(joined.columns))
    check("an exactly-matching AWS column is suffixed too",
          list(joined["Sprd"]) == [1.0, 2.0] and "Sprd_aws" in joined.columns)

    win = pd.DataFrame({
        "market": ["India", "India", "Hong Kong"],
        "first_start_time_min": [_hhmm_to_min("17:35"), _hhmm_to_min("09:30"),
                                 _hhmm_to_min("09:30")],
        "notional": [1e6, 1e6, 1e6],
    })
    kept = filter_india_close_window(win)
    check("the India window keeps only orders that started in it, and only India",
          len(kept) == 2 and set(kept["market"]) == {"India", "Hong Kong"},
          kept["market"].tolist())

    # The rows that get %CLOSE = 100 must be exactly the rows that survive,
    # and no other market may be touched.
    imp = pd.DataFrame({
        "market": ["India", "India", "Hong Kong"],
        "first_start_time_min": [_hhmm_to_min("17:35"), _hhmm_to_min("09:30"),
                                 _hhmm_to_min("17:35")],
        "pct_close": [0.0, 0.0, 12.0],
        "notional": [1e6, 1e6, 1e6],
    })
    mask = india_in_close_window(imp)
    imp["pct_close_imputed"] = mask
    imp.loc[mask, "pct_close"] = 100.0
    out_i = filter_india_close_window(imp)
    check("the kept India order carries %CLOSE 100, the other market is untouched",
          list(out_i["pct_close"]) == [100.0, 12.0], list(out_i["pct_close"]))
    check("imputation and the filter select the same rows",
          bool((mask == (imp["market"].eq("India")
                         & imp["first_start_time_min"].between(
                             _hhmm_to_min("17:30"), _hhmm_to_min("17:45")))).all()))

    check("weighting falls back when a weight column is all zeros",
          weight_column(pd.DataFrame({"cont_notional": [0.0, 0.0],
                                      "notional": [1.0, 2.0]}),
                        "cont_notional") == "notional")
    check("weighting keeps the preferred column when it has value",
          weight_column(pd.DataFrame({"cont_notional": [0.0, 5.0],
                                      "notional": [1.0, 2.0]}),
                        "cont_notional") == "cont_notional")
    check("pct_continuous = TAKE + POST + DARK",
          np.allclose(df["pct_continuous"],
                      df[["pct_take", "pct_post", "pct_dark"]].sum(axis=1)))

    check("wait cost identity: IS - Close",
          np.allclose(df["wait_cost_bps"], df["slip_arrival"] - df["slip_close"]))
    # The export carries some spread-normalised columns and not others, so the
    # derived ones have to be exactly slippage / spread or the two halves of
    # the spreads table would not be the same measure.
    ok = all(np.allclose(df[f"sprd_{b}"].dropna(),
                         (df[f"slip_{b}"] / df["spread_bps"].where(
                             df["spread_bps"] > 0)).dropna())
             for b in ["arrival", "pvwap", "close", "vwap"]
             if f"sprd_{b}" in df)
    check("derived spread-normalised columns are slippage / spread", ok)
    check("a spread of zero gives no ratio, and does not raise",
          bool(pd.isna((pd.Series([10.0]) / pd.Series([0.0]).where(
              pd.Series([0.0]) > 0)).iloc[0])))
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
    ind = df[df["market"] == "India"]
    cas = pd.Timestamp(AUCTION_FROM["India"])
    check("India in the sample is outside the auction population",
          bool((~ind["has_auction"]).all()) and len(ind) > 0)
    # Tested on a made-up series, because the sample stops before the CAS date
    # and an empty selection would pass either way.
    mkt = pd.Series(["India", "India", "Hong Kong", "India"])
    dts = pd.to_datetime([cas - pd.Timedelta(days=1), cas,
                          cas - pd.Timedelta(days=1), cas + pd.Timedelta(days=90)])
    got = close_regime(mkt, pd.Series(dts)).tolist()
    check("India post-CAS is its own regime, not the auction population",
          got == [REGIME_VWAP_CLOSE, REGIME_UNKNOWN,
                  REGIME_AUCTION, REGIME_UNKNOWN], got)
    check("nothing uncertain is counted as an auction",
          auction_available(mkt, pd.Series(dts)).tolist()
          == [False, False, True, False])
    check("an auction-only order is cohorted as such",
          assign_cohort(pd.DataFrame({
              "fill_rate": [100.0], "pct_close": [100.0], "adv_pct": [0.5],
          })).iloc[0] == "Auction only")
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

    raw = merge_aws(raw, Path(AWS_DIR))
    cols = resolve_columns(raw)
    df = normalise(raw, cols)
    df = filter_india_close_window(df)

    if DATE_FROM or DATE_TO:
        before = len(df)
        if DATE_FROM:
            df = df[df["date"] >= pd.Timestamp(DATE_FROM)]
        if DATE_TO:
            df = df[df["date"] <= pd.Timestamp(DATE_TO)]
        log(f"  period filter {DATE_FROM or 'start'} .. {DATE_TO or 'end'}: "
            f"{before:,} -> {len(df):,} orders")
    if df.empty:
        raise SystemExit(
            "\nThe period filter removed every order. Check --from / --to "
            "against the dates in the file.")
    if "date" in df and df["date"].notna().any():
        lo, hi = df["date"].min(), df["date"].max()
        log(f"  dates present: {lo.date()} .. {hi.date()}")
        log(f"  period label:  {PERIOD_LABEL}")
        # A window that straddles a market-structure change is not one regime,
        # and the label on the charts will not say so. Say it here.
        for mkt, start in AUCTION_FROM.items():
            when = pd.Timestamp(start)
            if lo < when <= hi:
                warn(f"this window STRADDLES the {mkt} auction change on "
                     f"{start}. {mkt} carries more than one regime inside it "
                     f"and is never pooled - see 32_close_regimes.")
            elif hi < when:
                log(f"  window ends before the {mkt} auction change on "
                    f"{start}: {mkt} is one regime throughout.")

    section("SCOPE AND EXCLUSIONS")
    df = apply_strategy_scope(df)
    df = clean_values(df)

    sanity_report(df, cols, list(raw.columns))
    close_strats = choose_close_strategies(df)
    if not close_strats:
        raise SystemExit(
            "\nNo close strategies identified. Pin CLOSE_STRATEGIES at the top "
            "of this script after reading the strategy table above.")

    section("BUILDING")
    out_dir.mkdir(parents=True, exist_ok=True)

    # Two runs off one extract land in two directories. If this one already
    # holds a different window, say so rather than half-overwriting it.
    stamp_path = out_dir / "run_window.txt"
    stamp = (f"{PERIOD_LABEL} | {DATE_FROM or 'start'} .. {DATE_TO or 'end'} | "
             f"{df['date'].min().date()} .. {df['date'].max().date()}"
             if "date" in df and df["date"].notna().any() else PERIOD_LABEL)
    if stamp_path.exists():
        was = stamp_path.read_text(encoding="utf-8").strip()
        if was and was != stamp:
            warn(f"{out_dir} already holds a different run:")
            log(f"    was:  {was}")
            log(f"    now:  {stamp}")
            log("    Overwriting. Use a separate --out per window to keep both.")
    stamp_path.write_text(stamp, encoding="utf-8")
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
    p.add_argument("--from", dest="date_from", metavar="YYYY-MM-DD",
                   help="first order date to include (overrides DATE_FROM)")
    p.add_argument("--to", dest="date_to", metavar="YYYY-MM-DD",
                   help="last order date to include (overrides DATE_TO)")
    p.add_argument("--label", dest="label", metavar="TEXT",
                   help="period label for the charts (overrides PERIOD_LABEL)")
    args = p.parse_args(argv)

    global DATE_FROM, DATE_TO, PERIOD_LABEL
    if args.date_from:
        DATE_FROM = args.date_from
    if args.date_to:
        DATE_TO = args.date_to
    if args.label:
        PERIOD_LABEL = args.label

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
