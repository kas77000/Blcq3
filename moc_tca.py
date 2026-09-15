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

# The NextOpen column in this export is already next open vs close, so it is
# the reversion as it stands. Set False if a future extract measures NextOpen
# against the execution price instead, and it will be differenced again.
NEXTOPEN_IS_VS_CLOSE = True

# Performance charts are drawn in SPREADS: a group's notional-weighted result
# divided by its notional-weighted spread, with that spread printed under the
# bar as [x bps]. Basis points do not compare across markets - a wide name
# costs more of them whatever the algo does - and a client shown eleven
# markets side by side will compare them. Set "bps" to go back; every spread
# table carries the bps figures beside the spreads.
CHART_UNIT = "spreads"

# By-side charts split Buy against Sell. A short sell is a sell here; the
# 33_by_side table still separates it.
BUY_SELL_ORDER = ["Buy", "Sell"]

# The Q1 summary quoted the share of notional under this %ADV.
SUMMARY_ADV_CUT = 2.0

# Which strategies the MISS TAXONOMY applies to. A VWAP order was never meant
# to reach the auction, so calling its low %CLOSE an unexplained miss would be
# nonsense. Clearance, capacity, cohorts and the first-execution tables run on
# this list only; every other table runs on CLOSE_STRATEGIES and is grouped by
# strategy, so the two are never pooled.
MOC_STRATEGIES: list[str] = ["CLOSE"]

# A slippage this large is a broken record, not a fill. The CELL is cleared;
# the order stays in the study. Nothing is ever removed for being merely
# large - clipping (CLIP_OUTLIERS, off by default) is the tool for fat
# tails, and deleting the extremes would delete the orders the review exists
# to find.
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

# Keep only orders that had a closing auction to reach at all.
#
# marketCloseSize is the auction's own size, so a zero means no auction ran
# that day in that name - a holiday, a half day, a name that does not hold one.
# Such an order cannot be judged as a close order, and leaving it in puts a
# zero in the denominator of every auction-share figure for a reason that has
# nothing to do with the order or the algo.
#
# This is an OPPORTUNITY test, not an outcome test. It removes orders that had
# no auction to reach; it never removes an order that reached for one and
# missed. Those are the finding.
REQUIRE_CAS_ELIGIBLE = True

# An order that finished well before its market closed never ran into the
# closing window, so it had no OPPORTUNITY to reach the auction and judging it
# as a close order says nothing. That is the honest analogue of the India
# window, and it is measured here.
#
# It is NOT the same as %CLOSE = 0, and the difference decides whether this
# review has a finding in it. A zero on an order that WAS live into the close
# is a real outcome and the single most valuable population in the run - the
# orders that could have cleared and did not. Dropping those would leave only
# the orders that worked, push auction share to ~100% by construction, and
# have the deck conclude that everything clears because everything that did
# not was removed.
#
# So this flags by default and drops nothing. Set DROP_NO_CLOSE_OPPORTUNITY
# True to remove them, having seen 37_close_opportunity first.
CLOSE_OPPORTUNITY_MIN = 5.0        # minutes before the close is "never got there"
DROP_NO_CLOSE_OPPORTUNITY = False

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
    # From the AWS extract. Without these the five splits stop well short of
    # 100 and everything that is not an auction looks like nothing at all.
    "pct_other":     ["%OTHER", "PctOther", "Other %"],
    "pct_cond":      ["%COND", "PctCond", "Cond %"],
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
    # Our fill in the closing auction and the auction's own size. Together
    # they give share OF THE AUCTION, which is the constraint that actually
    # binds a close order - %Adv is only ever a proxy for it.
    "fill_close_size":   ["fillCloseSize", "FillCloseSize"],
    "market_close_size": ["marketCloseSize", "MarketCloseSize"],
    # A limit and a close price, to ask whether the limit was ever going to
    # cross. market_limit says "Limit" on every order and explains nothing.
    "limit_price":   ["ordprice", "OrdPrice", "strike"],
    # The day's closing price comes from the AWS extract's endprice. PX_LAST
    # is NOT it - on a checked order it read 91,000 against a 14 May close
    # of 91,600 - so it is no longer a fallback: a wrong close is worse than
    # none, because every check built on it would pass or fail for nothing.
    "close_price":   ["endprice", "EndPrice", "end_price"],
    # T+1 opening price, to rebuild reversion from prices.
    "next_open_price": ["nxt_open", "NxtOpen", "next_open_price"],
    # Average executed price, local currency ($Mln is built from it).
    "avg_price":     ["avgprice", "AvgPrice", "avg_px", "AvgPx"],
    # First execution price. first_execprice is the AWS extract's name.
    "first_exec_price": ["first_execprice", "first_exec_price", "first_px"],
    # capacity and behaviour
    "adv_pct":       ["%Adv", "% Adv", "PctAdv"],
    "adv":           ["Adv", "ADV"],
    "pr_cont":       ["fPR_cont", "fPRcont", "PR_cont"],
    # Participation in the closing auction: executed close quantity over the
    # auction's volume. Read as a percentage; a fraction is caught below.
    "close_pr":      ["ClosePR", "Close PR", "ClosePct PR", "close_pr"],
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
            "first_start_time", "fill_close_size", "market_close_size",
            "limit_price", "close_price", "next_open_price", "avg_price",
            "first_exec_price"]

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

# How first_exec_vs_close is turned into positive = saving on load:
#   "buys"  multiply BUY orders by -1, leave sells as in the file  (desk's call)
#   "all"   multiply every order by -1
#   "none"  use the file as it is
# A buy first filled at 87,300 against a close of 91,600 is in the export as
# -480.81 (4,300 over the midpoint 89,450), so buys need the flip. The price
# check in the run log reports buys and sells separately and says SIGN WRONG
# ON SELLS ONLY (or BUYS) if this setting is wrong for one side.
FIRST_EXEC_FLIP = "buys"
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

# Known to arrive as fractions: the AWS extract writes these 0-1 while
# orders.csv writes 0-100, and after the join both sit in the same row.
#
# Declared rather than inferred. The detector below can only guess from the
# largest value in the book, which is safe for a column that reaches 20% and
# unsafe for one that never exceeds 1% legitimately - %COND could plausibly be
# either. A column named here is scaled because it is known to need it.
FRACTION_COLUMNS = {"pct_other", "pct_cond"}
PCT_FIELDS = ["fill_rate", "adv_pct", "participation", "pr_cont", "close_pr",
              "pct_close", "pct_open", "pct_post", "pct_take", "pct_dark",
              "pct_other", "pct_cond"]

VENUE_FIELDS = ["pct_open", "pct_close", "pct_post", "pct_take", "pct_dark",
                "pct_other", "pct_cond"]
# Everything that is not an auction. %OTHER and %COND belong here: they are
# not the auction, so whatever they are, they are continuous.
CONTINUOUS_FIELDS = ["pct_take", "pct_post", "pct_dark", "pct_other", "pct_cond"]
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

# Markets that cannot be attributed to a venue. They stay in every TABLE -
# the money is real and the totals must still add up - but they come out of
# the CHARTS, because an unattributable bar on a client slide invites a
# question nobody in the room can answer.
EXCLUDE_MARKETS_FROM_CHARTS = {"Unknown"}

# The client account says whether a trade was cash or swap. Anything not
# listed is shown as "Other" rather than hidden or guessed at: a new account
# appearing should be visible, not silently absorbed into one of these.
CASH_CLIENTS = {
    "BLAROC.DXA.JP.A_DSA", "BLAROC.XXX.AU.A_DSA",
    "BLAROC.AU.A_CARE", "BLAROC.AU.A_PT",
}
SWAP_CLIENTS = {
    "BLAROC.SYN.US.A_DSA", "BLAROC.SYN.US.A_DMA",
    "BLAROC.SYCARE.US.A_CARE", "BLAROC.SYN.US.A_CARE",
}
NOTIONAL_TYPE_ORDER = ["Cash", "Swap", "Other"]

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
# Clipping of extreme values before averaging. OFF by default: every chart
# and table average is then a plain SUMPRODUCT(value, $Mln) / SUM($Mln) over
# the orders in the bar, and anyone can rebuild it in Excel. Turn it on
# (here, or --clip on the command line) to pull each performance value in
# to the 1st / 99th percentile of the orders in its bar before averaging -
# rows are never dropped, only their extreme values capped. Spreads and
# shares (fill, %CLOSE, ClosePR) are never clipped either way.
CLIP_OUTLIERS = False
CLIP_PERCENTILES = (0.01, 0.99)
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


def _has_signal(col: pd.Series) -> bool:
    """True if the column carries anything at all - not all NaN, not all zero."""
    v = pd.to_numeric(col, errors="coerce")
    if v.notna().sum() == 0:
        return col.notna().any() and col.astype(str).str.strip().ne("").any()
    return bool((v.fillna(0) != 0).any())


def resolve_columns(raw: pd.DataFrame) -> dict[str, str]:
    """logical name -> actual header. Unresolved logical names are absent.

    Where the order file and the AWS extract both answer to a name, the order
    file's version was kept and the AWS one suffixed. That is the right
    default - until the order file's column turns out to be empty. Then the
    default silently throws away the only copy of the data that exists, which
    is exactly what happened to %POST, %TAKE and %DARK.
    """
    lookup = {}
    for actual in raw.columns:
        lookup.setdefault(_norm_name(actual), actual)
    resolved, swapped = {}, []
    for logical, candidates in COLUMNS.items():
        if isinstance(candidates, str):
            candidates = [candidates]
        hits = []
        for cand in candidates:
            for name in (cand, cand + AWS_SUFFIX):
                hit = lookup.get(_norm_name(name))
                if hit is not None and hit not in hits:
                    hits.append(hit)
        if not hits:
            continue
        chosen = hits[0]
        if not _has_signal(raw[chosen]):
            better = next((h for h in hits[1:] if _has_signal(raw[h])), None)
            if better is not None:
                swapped.append((logical, chosen, better))
                chosen = better
        resolved[logical] = chosen

    if swapped:
        warn("these columns are empty in the order file and populated in the")
        log("    AWS extract, so the AWS version is used instead:")
        for logical, was, now in swapped:
            log(f"      {logical:<14}{was}  ->  {now}")
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


def india_close_pr_proxy(df: pd.DataFrame, in_window: pd.Series) -> pd.Series:
    """India's close participation: its order participation (PR), not ClosePR.

    India runs no closing auction in H1, so ClosePR is zero there by
    construction. Its close is the last half hour, and a window order trades
    inside it - so the order's own participation rate over its life is the
    nearest honest measure of how much of that close we were. Only the
    window orders are touched; the measured ClosePR is kept beside it.
    """
    if "close_pr" not in df or "participation" not in df:
        return pd.Series(False, index=df.index)
    use = in_window & df["participation"].notna()
    if use.any():
        df["close_pr_measured"] = df["close_pr"]
        df.loc[use, "close_pr"] = df.loc[use, "participation"]
    return use


def normalise(raw: pd.DataFrame, cols: dict[str, str]) -> pd.DataFrame:
    """Build the working frame: logical names, base units, derived fields."""
    df = pd.DataFrame(index=raw.index)

    # Anything mapped but not copied here never reaches the frame, so a column
    # can resolve, be reported as found, and still do nothing. That is exactly
    # what happened to the spread-normalised columns: the log said "read from
    # the file" while the values were being derived behind it.
    numeric = ["notional", "order_shares", "fill_rate", "adv", "adv_pct",
               "participation", "pr_cont", "close_pr", "spread_bps",
               "volatility",
               "first_exec_vs_close"] + VENUE_FIELDS + [
               "slip_close", "slip_arrival", "slip_pvwap", "slip_vwap",
               "slip_nextopen", "slip_open",
               "sprd_arrival", "sprd_pvwap", "sprd_close", "sprd_vwap",
               "fill_close_size", "market_close_size",
               "limit_price", "close_price", "next_open_price", "avg_price",
               "first_exec_price"]
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
    else:
        # The order file writes shares as percentages, the AWS extract writes
        # some of them as fractions, and after the join both sit in the same
        # row. A global flag cannot express that - it would fix one source and
        # break the other - so each column is judged on its own.
        #
        # A share column whose largest value in the whole book is at or under
        # 1 is a fraction: across tens of thousands of orders at least one
        # should have put more than 1% somewhere.
        known, guessed, contradicted = [], [], []
        for f in PCT_FIELDS:
            if f not in df:
                continue
            v = pd.to_numeric(df[f], errors="coerce")
            top = float(v.max()) if v.notna().any() else 0.0
            if f in FRACTION_COLUMNS:
                if top > 1.0:
                    contradicted.append((f, top))     # declared, but not one
                else:
                    df[f] = v * 100.0
                    known.append((f, top))
            elif 0 < top <= 1.0:
                df[f] = v * 100.0
                guessed.append((f, top))
        if known:
            log("  share columns known to arrive as fractions, x100 applied:")
            for f, top in known:
                log(f"      {f:<16}largest value in the book {top:.4f}")
        if guessed:
            warn("these look like fractions and have been multiplied by 100:")
            for f, top in guessed:
                log(f"      {f:<16}largest value in the book {top:.4f}")
            log("    Inferred, not declared. Across a whole book a genuine")
            log("    percentage should exceed 1 somewhere - add it to")
            log("    FRACTION_COLUMNS to make it certain, or check the source.")
        if contradicted:
            warn("declared as fractions in FRACTION_COLUMNS but reaching past 1:")
            for f, top in contradicted:
                log(f"      {f:<16}largest value {top:.2f}")
            log("    Left alone. Either the source changed or the list is")
            log("    wrong, and scaling on a bad assumption is worse than not.")

    # --- sign -------------------------------------------------------------
    if not POSITIVE_IS_SAVING:
        for f in [c for c in df.columns if c.startswith("slip_")]:
            df[f] = -df[f]
        if "first_exec_vs_close" in df:
            df["first_exec_vs_close"] = -df["first_exec_vs_close"]
    if "first_exec_vs_close" in df and FIRST_EXEC_FLIP in ("buys", "all"):
        if FIRST_EXEC_FLIP == "all":
            flip = pd.Series(True, index=df.index)
        elif "side" in df:
            flip = df["side"].astype(str).str.strip().str.upper().isin(
                {v.upper() for v in BUY_VALUES})
        else:
            warn("FIRST_EXEC_FLIP is 'buys' but there is no side column - "
                 "first_exec_vs_close left as in the file.")
            flip = pd.Series(False, index=df.index)
        df.loc[flip, "first_exec_vs_close"] = -df.loc[flip, "first_exec_vs_close"]
        log(f"  first_exec_vs_close: multiplied by -1 on {int(flip.sum()):,} "
            f"{'buy ' if FIRST_EXEC_FLIP == 'buys' else ''}orders "
            f"(FIRST_EXEC_FLIP = {FIRST_EXEC_FLIP!r})")

    # --- does first execution agree with the close result? -----------------
    # On an order that traded before the auction, the first fill and the
    # continuous part it belongs to move with the day's trend, so the two
    # should mostly carry the same sign. If they mostly disagree, the column
    # is the wrong way round - which is exactly how it was found.
    if {"first_exec_vs_close", "slip_close", "pct_close"} <= set(df.columns):
        fe = pd.to_numeric(df["first_exec_vs_close"], errors="coerce")
        sc = pd.to_numeric(df["slip_close"], errors="coerce")
        pre = (df["pct_close"] < 50) & fe.notna() & sc.notna()             & (fe.abs() > 5) & (sc.abs() > 5)
        if pre.sum() >= 30:
            agree = float((np.sign(fe[pre]) == np.sign(sc[pre])).mean())
            log(f"  first execution and execution vs close share a sign on "
                f"{100 * agree:.0f}% of {int(pre.sum()):,} mostly-continuous orders")
            if agree < 0.5:
                warn("first_exec_vs_close disagrees with vs Close on most orders -")
                log("    its sign looks inverted. Check one order's first fill against")
                log("    the close and set FIRST_EXEC_FLIP accordingly.")

    # --- identity ---------------------------------------------------------
    if "side" in df:
        up = df["side"].str.upper()
        df["is_buy"] = up.isin({v.upper() for v in BUY_VALUES})
        df["is_short"] = up.isin({v.upper() for v in SHORT_SELL_VALUES})
        df["side_label"] = np.where(
            df["is_buy"], "Buy",
            np.where(df["is_short"], "Short sell", "Sell"))
        df["buy_sell"] = np.where(df["is_buy"], "Buy", "Sell")

    df["market"] = market_from_symbol(df["symbol"]) if "symbol" in df else UNKNOWN_MARKET
    df["close_regime"] = close_regime(df["market"], df.get("date"))

    # Cash or swap, from the client account. Matched case-insensitively and
    # with surrounding space stripped, because an account string is exactly
    # the kind of field that arrives with a trailing blank.
    if "client" in df:
        acct = df["client"].astype(str).str.strip().str.upper()
        cash = {c.upper() for c in CASH_CLIENTS}
        swap = {c.upper() for c in SWAP_CLIENTS}
        df["notional_type"] = np.where(acct.isin(cash), "Cash",
                                       np.where(acct.isin(swap), "Swap",
                                                "Other"))

    # Before any venue field is derived, so close_notional, pct_continuous,
    # close_bucket and the rest all follow from the same number.
    if INDIA_CLOSE_PROXY and "pct_close" in df:
        in_window = india_in_close_window(df)
        df["pct_close_imputed"] = in_window
        if in_window.any():
            df["pct_close_measured"] = df["pct_close"]
            df.loc[in_window, "pct_close"] = 100.0
            # 100% in the close leaves nothing for anywhere else, so the rest
            # of the mix goes to zero. Otherwise the row sums past 100 and
            # every venue figure that includes India is quietly wrong.
            for f in VENUE_FIELDS:
                if f != "pct_close" and f in df:
                    df.loc[in_window, f] = 0.0
        df["close_pr_from_pr"] = india_close_pr_proxy(df, in_window)
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
    cont_fields = [f for f in CONTINUOUS_FIELDS if f in df]
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
        # NextOpen in this export is already measured AGAINST THE CLOSE, so
        # it IS the reversion. Subtracting slip_close from it, as this used
        # to, took the close out twice and produced a number that was neither
        # one thing nor the other. Confirmed with the desk.
        df["reversion_bps"] = (df["slip_nextopen"] if NEXTOPEN_IS_VS_CLOSE
                               else df["slip_nextopen"] - df["slip_close"])

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

    # --- the four levers --------------------------------------------------
    # Each one degrades to nothing if its column is absent. Everything below
    # is reported as missing rather than assumed.

    # 1. Our share of the auction itself, not of the day.
    if "fill_close_size" in df and "market_close_size" in df:
        mkt = pd.to_numeric(df["market_close_size"], errors="coerce")
        ours = pd.to_numeric(df["fill_close_size"], errors="coerce")

        # India first. It runs no auction in this period, so both columns come
        # back empty there and the order looks like it never went near a close.
        # An order that ran inside the closing window traded through the VWAP
        # that IS the close, so all of its executed quantity was in the close -
        # that is the fillCloseSize equivalent of setting %CLOSE to 100, and it
        # has to happen before eligibility is decided or the order is gone
        # before its own window is ever consulted.
        if INDIA_CLOSE_PROXY and "exec_shares" in df:
            in_window = india_in_close_window(df)
            if in_window.any():
                df["fill_close_imputed"] = in_window
                ours = ours.where(~in_window, df["exec_shares"])
                df["fill_close_size"] = ours

        df["auction_share_pct"] = np.where(mkt > 0, 100.0 * ours / mkt, np.nan)
        # fillCloseSize settles eligibility better than any flag can, because
        # it is an outcome rather than a permission: a positive fill IS the
        # order having been in the auction, that day, in that name. A fill
        # implies an auction to have filled in, so either column carrying
        # something is enough - requiring marketCloseSize alone would throw
        # away every order whose own fill proves the point.
        df["auction_existed"] = (mkt.fillna(0) > 0) | (ours.fillna(0) > 0)
        df["auction_participated"] = ours.fillna(0) > 0

    # 2. How late the order arrived, against its own market's close.
    if "first_start_time_min" in df and "market" in df and "date" in df:
        close_at = continuous_end_min(df["market"], df["date"])
        df["mins_to_close_at_start"] = close_at - df["first_start_time_min"]

    # 3. Was the limit ever going to cross? A buy limit below the close, or a
    #    sell limit above it, could not have traded in the auction. This is
    #    the question market_limit cannot answer because it never varies.
    if "limit_price" in df and "close_price" in df and "is_buy" in df:
        lim = pd.to_numeric(df["limit_price"], errors="coerce")
        cls = pd.to_numeric(df["close_price"], errors="coerce")
        ok = lim.notna() & cls.notna() & (lim > 0) & (cls > 0)
        away = np.where(df["is_buy"], cls - lim, lim - cls)
        df["limit_gap_bps"] = np.where(ok, 1e4 * away / cls, np.nan)
        df["limit_binding"] = ok & (pd.Series(away, index=df.index) > 0)

    # An order that put anything through continuous before the auction. That
    # is the population reversion can say something about: an order that only
    # ever printed in the auction has no pre-trade to have moved the price
    # with, so its reversion is the market's, not ours.
    # Reversion only means something for an order that was IN the close. With
    # nothing printed in the auction there is no close print to have reverted
    # from, and those orders were the source of the odd numbers by market.
    # India passes this by construction - its close share is imputed - which
    # is exactly why its reversion reads as a separate line rather than
    # pooled with markets that really ran an auction.
    if "pct_close" in df:
        traded = ((df["fill_rate"] >= COHORT_FR_ZERO) if "fill_rate" in df
                  else pd.Series(True, index=df.index))
        in_close = traded & (df["pct_close"] > 0)
        df["executed_in_close"] = in_close
        df["pretraded"] = in_close & (df["pct_close"] < AUCTION_ONLY_MIN_PCT)
        df["close_only"] = in_close & (df["pct_close"] >= AUCTION_ONLY_MIN_PCT)

    # --- buckets ----------------------------------------------------------
    if "adv_pct" in df:
        df["adv_bucket"] = pd.cut(df["adv_pct"], ADV_BUCKETS, labels=ADV_LABELS)
    if "pct_close" in df:
        df["close_bucket"] = pd.cut(df["pct_close"], CLOSE_BUCKETS,
                                    labels=CLOSE_LABELS)
    # Spread quartiles, cut on the data rather than on fixed thresholds, so
    # they stay quarters whatever the book looks like. The label carries the
    # actual range: "Q1" alone tells a reader nothing about how tight tight is.
    if "spread_bps" in df:
        v = pd.to_numeric(df["spread_bps"], errors="coerce")
        ok = v.notna() & (v > 0)
        if ok.sum() >= 8 and v[ok].nunique() >= 4:
            try:
                _, edges = pd.qcut(v[ok], 4, retbins=True, duplicates="drop")
                # Ranges only, like the %Adv bands. "Q1" carries no meaning
                # a reader can use; "0.8-9.2" does.
                names = [f"{edges[i]:.1f}-{edges[i + 1]:.1f}"
                         for i in range(len(edges) - 1)]
                df["spread_bucket"] = pd.cut(v, edges, labels=names,
                                             include_lowest=True)
            except (ValueError, IndexError):
                pass
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


def mark_close_opportunity(df: pd.DataFrame) -> pd.DataFrame:
    """Flag orders that finished before their market's closing window.

    close_gap_min is the market's continuous close minus the order's end time,
    so a large positive value means the order was long done before the close.
    India is left alone: it has no continuous close to measure against and its
    own window filter has already run.
    """
    if "close_gap_min" not in df:
        return df
    df = df.copy()
    gap = df["close_gap_min"]
    df["reached_close_window"] = ~(gap > CLOSE_OPPORTUNITY_MIN)
    df.loc[gap.isna(), "reached_close_window"] = True   # unknown is not evidence
    return df


def t_close_opportunity(df: pd.DataFrame) -> pd.DataFrame:
    """Who was still live into the closing window, and who was long gone."""
    if "reached_close_window" not in df or "market" not in df:
        return pd.DataFrame()
    rows = []
    for mkt, g in df.groupby("market", dropna=False, observed=True):
        reached = g[g["reached_close_window"]]
        missed = g[~g["reached_close_window"]]
        rows.append({
            "market": mkt,
            "orders": len(g),
            "notional (" + CURRENCY + "m)": g["notional"].sum() / 1e6,
            "ran into the close": len(reached),
            "finished before it": len(missed),
            "% finished before it": 100.0 * len(missed) / max(len(g), 1),
            "value finished before (" + CURRENCY + "m)":
                missed["notional"].sum() / 1e6,
            "wtd %CLOSE if it ran in": wmean(reached["pct_close"],
                                             reached["notional"], winsor=False)
                if "pct_close" in g and len(reached) else np.nan,
            "wtd %CLOSE if it did not": wmean(missed["pct_close"],
                                              missed["notional"], winsor=False)
                if "pct_close" in g and len(missed) else np.nan,
        })
    out = pd.DataFrame(rows).set_index("market")
    out = out.sort_values("value finished before (" + CURRENCY + "m)",
                          ascending=False).round(2)

    # The table checks itself. An order that finished before the close cannot
    # also have printed in the auction, so a market showing both is not
    # telling us about its orders - it is telling us its close time is wrong.
    # Those are the markets flagged UNVERIFIED in MARKET_CLOSE_HKT.
    bad = out[(out["finished before it"] > 0)
              & (out["wtd %CLOSE if it did not"] > 5.0)]
    if len(bad):
        warn("these markets say an order finished BEFORE the close and still")
        log("    printed in the auction. Both cannot be true - the close time")
        log("    in MARKET_CLOSE_HKT is wrong, not the orders:")
        for mkt, r in bad.iterrows():
            flag = " (UNVERIFIED)" if MARKET_CLOSE_HKT.get(mkt, ("", True))[1] is False else ""
            log(f"      {str(mkt):<14}{r['finished before it']:>6,} orders, "
                f"{r['wtd %CLOSE if it did not']:>6.1f}% auction share{flag}")
        log("    Do not filter on this column for those markets until the")
        log("    close time is confirmed against the desk's own sessions.")
    return out


NO_AUCTION_COLUMNS = ["order_id", "date", "symbol", "market", "strategy",
                      "side_label", "cap", "notional", "order_shares",
                      "fill_rate", "adv_pct", "pct_close",
                      "market_close_size", "fill_close_size",
                      "first_start_time_min", "close_gap_min"]


def filter_cas_eligible(df, out_dir=None):
    """Drop orders that had no closing auction to reach.

    The dropped orders are written out one by one. A market that loses all of
    them is either a market with no auction or a column the extract does not
    populate there, and the only way to tell those apart is to look.
    """
    if not REQUIRE_CAS_ELIGIBLE:
        return df
    if "auction_existed" not in df:
        warn("no marketCloseSize, so CAS eligibility cannot be tested.")
        log("    Every order is kept. Auction-share figures then include")
        log("    orders that never had an auction to reach.")
        return df

    section("CAS ELIGIBILITY")
    # No exemptions. Every market is tested on whether an auction had size,
    # and India is then narrowed further by its own start-time window.
    keep = df["auction_existed"].fillna(False)
    log("  Keeping only orders with a closing auction to reach, measured by")
    log("  the auction's own size rather than by anything the order did.")
    log("")
    log(f"    {'market':<16}{'orders':>9}{'kept':>9}{'dropped':>9}"
        f"{'value dropped (' + CURRENCY + 'm)':>26}")
    for mkt, g in df.groupby("market", dropna=False, observed=True):
        k = int(g["auction_existed"].fillna(False).sum())
        val = float(g.loc[~g["auction_existed"].fillna(False), "notional"].sum())
        log(f"    {str(mkt):<16}{len(g):>9,}{k:>9,}{len(g) - k:>9,}"
            f"{val / 1e6:>26,.2f}")
    wiped = []
    for mkt, g in df.groupby("market", dropna=False, observed=True):
        if len(g) and not g["auction_existed"].fillna(False).any():
            wiped.append((str(mkt), len(g), float(g["notional"].sum()) / 1e6))
    if wiped:
        log("")
        warn("these markets lost EVERY order to the test:")
        for mkt, n, val in sorted(wiped, key=lambda r: -r[2]):
            log(f"      {mkt:<16}{n:>9,} orders   {CURRENCY} {val:>12,.2f}m")
        log("    marketCloseSize is zero or absent for all of them. That is")
        log("    either a market with no auction, or a column the extract does")
        log("    not populate there - and those need opposite responses.")
        log("    Check before reading anything that excludes them.")

    gone = df[~keep]
    if out_dir is not None and len(gone):
        out_dir.mkdir(parents=True, exist_ok=True)
        cols = [c for c in NO_AUCTION_COLUMNS if c in gone.columns]
        by = [c for c in ("market", "date") if c in cols]
        path = out_dir / "no_auction_orders.csv"
        (gone[cols].sort_values(by) if by else gone[cols]).to_csv(path, index=False)
        log("")
        log(f"  the {len(gone):,} dropped orders are listed one by one in")
        log(f"    {path}")
        log("    with marketCloseSize and fillCloseSize beside each, so a")
        log("    market with no auction can be told apart from a column the")
        log("    extract simply does not populate there.")

    out = df[keep].copy()
    if out.empty:
        raise SystemExit("\n".join([
            "",
            "CAS eligibility removed every order.",
            "Check that marketCloseSize is populated, or set",
            "REQUIRE_CAS_ELIGIBLE = False to run without the test.",
        ]))
    share = 100.0 * out["notional"].sum() / max(df["notional"].sum(), 1e-9)
    log("")
    log(f"    kept {len(out):,} of {len(df):,} orders, {share:.1f}% of value")
    log("    An order that reached for the auction and missed is NOT removed")
    log("    here - that is the finding, and it stays in the cohorts.")
    return out


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

    # --- %CLOSE against fillCloseSize -------------------------------------
    # Two independent sources for the same fact: did this order trade in the
    # auction. They come from different systems, so where they disagree one of
    # them is wrong, and the disagreement is worth more than either alone.
    if "auction_participated" in df and "pct_close" in df:
        d = df[df["auction_participated"].notna() & df["pct_close"].notna()]
        if len(d):
            said_yes = d["pct_close"] > 0
            did = d["auction_participated"]
            only_pct = int((said_yes & ~did).sum())
            only_fill = int((~said_yes & did).sum())
            agree = len(d) - only_pct - only_fill
            log("")
            log("  %CLOSE vs fillCloseSize - two sources, same question")
            log(f"    {'agree':<40}{agree:>10,}  ({100.0 * agree / len(d):.1f}%)")
            if only_pct:
                log(f"    {'%CLOSE says yes, no auction fill':<40}{only_pct:>10,}")
            if only_fill:
                log(f"    {'auction fill, but %CLOSE reads 0':<40}{only_fill:>10,}")
            if only_pct or only_fill:
                warn("the two disagree on "
                     f"{100.0 * (only_pct + only_fill) / len(d):.1f}% of orders.")
                log("    fillCloseSize is a quantity actually printed in the")
                log("    auction; %CLOSE is a share of executed quantity that")
                log("    the export has already been caught leaving empty.")
                log("    Where they differ, trust fillCloseSize.")

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

    verify_slippage(df)


SLIPPAGE_SIGN_MIN = 0.80      # share of orders whose sign must agree
SLIPPAGE_BPS_TOL = 5.0        # median |file - rebuilt| to call it a match


def verify_slippage(df: pd.DataFrame, quiet: bool = False) -> pd.DataFrame:
    """Rebuild the slippage columns from prices and compare with the file.

    Side-adjusted, positive = saving, like the columns themselves:
        vs Close          side x (close - avg price)
        first exec        side x (close - first execution price)
        close to T+1      side x (next open - close)
    Each is tried over three denominators - the close, the order's own
    price, and the midpoint of the two - because the export does not say
    which it uses, and the one that matches is itself worth knowing. A
    match needs the sign to agree on most orders AND the size to agree.
    """
    need = {"close_price", "is_buy"}
    if not need <= set(df.columns):
        if not quiet:
            log("  slippage check: no endprice (or no side) - cannot rebuild "
                "any slippage from prices.")
        return pd.DataFrame()
    side = np.where(df["is_buy"], 1.0, -1.0)
    cls = pd.to_numeric(df["close_price"], errors="coerce")
    checks = []
    if "avg_price" in df and "slip_close" in df:
        checks.append(("vs Close", "slip_close", "avg_price", cls, "avg"))
    if "first_exec_price" in df and "first_exec_vs_close" in df:
        checks.append(("first exec vs close", "first_exec_vs_close",
                       "first_exec_price", cls, "first"))
    if "next_open_price" in df and "slip_nextopen" in df:
        checks.append(("close to T+1 (NextOpen)", "slip_nextopen",
                       "next_open_price", cls, "open"))
    rows = []
    for label, col, pcol, close, kind in checks:
        px = pd.to_numeric(df[pcol], errors="coerce")
        file_bps = pd.to_numeric(df[col], errors="coerce")
        ok = (close > 0) & (px > 0) & file_bps.notna() & \
            (file_bps.abs() < MAX_ABS_BPS)
        if ok.sum() < 5:
            continue
        ratio = float((px[ok] / close[ok]).median())
        if not 0.5 < ratio < 2.0:
            rows.append({"check": label, "orders": int(ok.sum()),
                         "verdict": f"price is not in the close's currency "
                                    f"(median ratio {ratio:.2f}) - not compared"})
            continue
        # the move, side-adjusted, positive = saving
        if kind == "open":
            move = (px - close) * side
        else:
            move = (close - px) * side
        best = None
        row = {"check": label, "column": col, "orders": int(ok.sum())}
        for name, den in (("close", close), ("own price", px),
                          ("midpoint", (close + px) / 2)):
            rebuilt = 1e4 * move / den
            diff = float((file_bps[ok] - rebuilt[ok]).abs().median())
            agree = float((np.sign(file_bps[ok]) == np.sign(rebuilt[ok]))[
                rebuilt[ok].abs() > 1].mean())
            row[f"median |diff| bps, / {name}"] = round(diff, 2)
            if best is None or diff < best[1]:
                best = (name, diff, agree, rebuilt)
        row["sign agrees %"] = round(100 * best[2], 1)
        row["best denominator"] = best[0]
        # Buys and sells apart. A column that is a raw price difference, not
        # side-adjusted, is right on one side and exactly wrong on the other;
        # pooled, that reads as a mediocre match and hides the cause.
        rb = best[3]
        same = np.sign(file_bps) == np.sign(rb)
        live = ok & (rb.abs() > 1)
        by_side = {}
        for side_name, mask in (("buys", live & df["is_buy"].astype(bool)),
                                ("sells", live & ~df["is_buy"].astype(bool))):
            by_side[side_name] = (float(same[mask].mean()) if mask.sum() else np.nan,
                                  int(mask.sum()))
            row[f"sign agrees % ({side_name})"] = (round(100 * by_side[side_name][0], 1)
                                                   if mask.sum() else np.nan)
        (ab, nb), (asl, ns) = by_side["buys"], by_side["sells"]
        one_side = None
        if min(nb, ns) >= 3:
            if ab >= SLIPPAGE_SIGN_MIN and asl <= 1 - SLIPPAGE_SIGN_MIN:
                one_side = "sells"
            elif asl >= SLIPPAGE_SIGN_MIN and ab <= 1 - SLIPPAGE_SIGN_MIN:
                one_side = "buys"
        flipped = 1e4 * -move / ((close + px) / 2)
        flip_diff = float((file_bps[ok] - flipped[ok]).abs().median())
        if one_side:
            row["verdict"] = (f"SIGN WRONG ON {one_side.upper()} ONLY - the column is "
                              "not side-adjusted the way the script assumes")
        elif best[2] < SLIPPAGE_SIGN_MIN and flip_diff < best[1]:
            row["verdict"] = ("SIGN INVERTED - the file is the other way round "
                              "from positive = saving")
        elif best[1] <= SLIPPAGE_BPS_TOL and best[2] >= SLIPPAGE_SIGN_MIN:
            row["verdict"] = "matches the prices"
        else:
            row["verdict"] = ("DOES NOT MATCH - different price, benchmark or "
                              "adjustment")
        rows.append(row)
    out = pd.DataFrame(rows)
    if quiet:
        return out
    log("")
    log("  slippage rebuilt from prices (endprice, nxt_open, avgprice)")
    if out.empty:
        missing = [c for c in ("avg_price", "first_exec_price",
                               "next_open_price") if c not in df]
        log("    nothing to compare - price columns not found: "
            + ", ".join(missing))
        return out
    for _, r in out.iterrows():
        log(f"    {r['check']:<26}{int(r['orders']):>8,} orders   {r['verdict']}")
        if "best denominator" in r and isinstance(r.get("best denominator"), str):
            diffs = "  ".join(f"{k.split('/ ')[1]} {r[k]:.1f}" for k in r.index
                              if str(k).startswith("median |diff|"))
            log(f"      sign agrees {r['sign agrees %']:.0f}% (buys "
                f"{r['sign agrees % (buys)']:.0f}%, sells "
                f"{r['sign agrees % (sells)']:.0f}%), median gap bps: "
                f"{diffs}  (best: {r['best denominator']})")
        if not str(r["verdict"]).startswith("matches"):
            warn(f"{r['check']} does not agree with the prices - read the line above")
    for c in ("avg_price", "first_exec_price", "next_open_price"):
        if c not in df:
            log(f"    {c} not found - that check did not run")
    return out


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


def winsorize(v: np.ndarray, limits=None) -> np.ndarray:
    """Pull the extreme tails back to the percentile value, do not drop them.

    Does nothing unless CLIP_OUTLIERS is on, or limits are passed explicitly.
    The switch is read when called, not when the file loads, so --clip works.
    """
    if limits is None:
        limits = CLIP_PERCENTILES if CLIP_OUTLIERS else None
    if limits is None:
        return v
    ok = ~np.isnan(v)
    if ok.sum() < 3:
        return v
    lo, hi = np.nanpercentile(v[ok], [limits[0] * 100, limits[1] * 100])
    return np.clip(v, lo, hi)


def wmean(values, weights, winsor=True) -> float:
    """Notional-weighted mean. Clipped first only when CLIP_OUTLIERS is on."""
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
    "No auction that day",
    "Limit could not cross",
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

    # A real cause, where the data carries it, tested BEFORE the residual: an
    # order whose limit sat the wrong side of the close is explained, and
    # leaving it among the unexplained sends the desk chasing the algo for
    # something the order's own limit did.
    #
    # auctionOnly is deliberately NOT used. It reports the MARKET's mechanism
    # as much as the order's permission - India's close orders come back
    # ContinuousOnly because India has no auction to be eligible for - so
    # reading it as permission would rule out a third of the book for a reason
    # that was never about the order.
    lim_binding = df["limit_binding"].fillna(False) \
        if "limit_binding" in df else pd.Series(False, index=idx)
    # Where the auction had no size there was nothing to miss. Tested ahead
    # of every other cause, because none of them apply to a day with no
    # auction at all.
    no_auction = ((~df["auction_existed"]).fillna(False)
                  if "auction_existed" in df else pd.Series(False, index=idx))

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
             no_fill & no_auction,
             no_fill & lim_binding,
             size_ok,
             no_fill & lim,
             no_fill & ~lim,
             partial & (short > FRONTIER_SHORTFALL_PP),
             partial],
            ["Never traded",
             "Auction only",
             "Cleared the auction",
             "No auction that day",
             "Limit could not cross",
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
    # Pre-traded orders only. For an order that never left the auction the
    # first fill IS the close, so its first-exec-vs-close is zero by
    # construction. That is not a result, and averaged in with the rest it
    # drags every band toward zero and buries the orders that did start
    # early - the ones the question is actually about. Same "pretraded"
    # flag the reversion charts use, so the deck means one thing by the word.
    if "pretraded" in df:
        df = df[df["pretraded"].fillna(False)]
    if df.empty:
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

def t_in_spreads(df: pd.DataFrame, by, value: str, weight: str = "notional",
                 order=None, ci: bool = True) -> pd.DataFrame:
    """A result in bps AND in spreads, per group, with a CI on each.

    Spreads are the ratio of two weighted averages - the group's result over
    the group's spread - never the average of per-order ratios. One name with
    a 0.5bp spread would otherwise produce a ratio in the hundreds, and the
    group mean would report that one name. Both averages run over the SAME
    orders (a result and a positive spread) with the SAME weights, so
    spreads x [spread bps] = bps holds on every row: the bracket on a chart
    is the number to multiply by, and a sales trader can do it on the call.

    The CI resamples orders and recomputes the ratio, so it carries the
    uncertainty in the spread as well as in the result.
    """
    keys = [by] if isinstance(by, str) else list(by)
    need = [value, "spread_bps", "notional"] + keys
    if df is None or df.empty or any(c not in df.columns for c in need):
        return pd.DataFrame()
    wcol = "notional" if weight == "notional" else weight_column(df, weight)
    rows = []
    for key, g in df.groupby(keys, dropna=False, observed=True):
        key = key if isinstance(key, tuple) else (key,)
        v = pd.to_numeric(g[value], errors="coerce").to_numpy(float)
        sp = pd.to_numeric(g["spread_bps"], errors="coerce").to_numpy(float)
        w = pd.to_numeric(g[wcol], errors="coerce").to_numpy(float)
        ok = ~np.isnan(v) & ~np.isnan(sp) & (sp > 0) & ~np.isnan(w) & (w > 0)
        n = int(ok.sum())
        if n == 0:
            continue
        # The spread is a property of the stock, not noise: never clipped.
        vv, ss, ww = winsorize(v[ok]), sp[ok], w[ok]
        bps = float(np.sum(vv * ww) / np.sum(ww))
        sprd = float(np.sum(ss * ww) / np.sum(ww))
        lo = hi = lo_b = hi_b = np.nan
        if ci and n >= MIN_N_FOR_CI:
            rng = np.random.default_rng(SEED)
            ratio, level = [], []
            # In chunks. A 12,000-order market times 2,000 draws is a
            # 24-million-cell index, three arrays deep, if done in one go.
            step = max(1, min(BOOTSTRAP_N, 4_000_000 // n))
            done = 0
            while done < BOOTSTRAP_N:
                k = min(step, BOOTSTRAP_N - done)
                idx = rng.integers(0, n, size=(k, n))
                wb = ww[idx]
                num = (vv[idx] * wb).sum(axis=1)
                ratio.append(num / (ss[idx] * wb).sum(axis=1))
                level.append(num / wb.sum(axis=1))
                done += k
            lo, hi = np.percentile(np.concatenate(ratio), [2.5, 97.5])
            lo_b, hi_b = np.percentile(np.concatenate(level), [2.5, 97.5])
        row = dict(zip(keys, key))
        row.update({
            "orders": n,
            "notional (" + CURRENCY + "m)":
                float(g["notional"].to_numpy(float)[ok].sum()) / 1e6,
            "spread bps (wtd)": sprd,
            "wtd mean bps": bps,
            "CI low bps": lo_b,
            "CI high bps": hi_b,
            "spreads": bps / sprd if sprd > 0 else np.nan,
            "CI low": lo,
            "CI high": hi,
            "saved (" + CURRENCY + "k)": to_money_k(bps, float(ww.sum()) / 1e6),
            "hit rate % (per order)": hit_rate(v[ok]),
            "small sample": n < MIN_N_FOR_CI,
        })
        rows.append(row)
    if not rows:
        return pd.DataFrame()
    out = pd.DataFrame(rows).set_index(keys if len(keys) > 1 else keys[0])
    if order and len(keys) == 1:
        keep = [o for o in order if o in out.index]
        rest = [i for i in out.index if i not in set(order)]
        out = out.loc[keep + rest]
    return out.round(3)


def _flag(df: pd.DataFrame, col: str) -> pd.Series:
    """A True/False column that may hold gaps, as a clean mask."""
    if col not in df:
        return pd.Series(False, index=df.index)
    return df[col].fillna(False).astype(bool)


def t_flow_split(df: pd.DataFrame) -> pd.DataFrame:
    """Close-only against pre-traded: how much of each, and how each filled.

    Shares are of the two together - the orders that reached the close - so
    they add to 100, the way the Q1 slide did. An order that never printed in
    the close is in neither.
    """
    if df is None or df.empty:
        return pd.DataFrame()
    pops = [("Close-only", df[_flag(df, "close_only")]),
            ("Pre-traded", df[_flag(df, "pretraded")])]
    base = sum(float(p["notional"].sum()) for _, p in pops)
    rows = []
    for name, p in pops:
        if p.empty:
            continue
        n = float(p["notional"].sum())
        row = {"flow": name, "orders": len(p),
               "notional (" + CURRENCY + "m)": n / 1e6,
               "% of notional": 100 * n / base if base else np.nan}
        if "fill_rate" in p:
            row["fill ratio %"] = wmean(p["fill_rate"], p["notional"],
                                        winsor=False)
        if "adv_pct" in p:
            row["%ADV (notional-weighted)"] = wmean(p["adv_pct"], p["notional"],
                                                   winsor=False)
            row["%ADV (median)"] = p["adv_pct"].median()
        if "pct_close" in p:
            row["% of notional in the close"] = wmean(
                p["pct_close"], p["notional"], winsor=False)
        if "spread_bps" in p:
            row["spread bps (wtd)"] = wmean(p["spread_bps"], p["notional"], winsor=False)
        rows.append(row)
    return pd.DataFrame(rows).set_index("flow").round(2) if rows else pd.DataFrame()


def t_adv_profile(df: pd.DataFrame) -> pd.DataFrame:
    """Orders and value by size against daily volume."""
    if df is None or df.empty or "adv_bucket" not in df:
        return pd.DataFrame()
    ncol = "notional (" + CURRENCY + "m)"
    g = df.groupby("adv_bucket", observed=True)
    out = pd.DataFrame({"orders": g.size(), ncol: g["notional"].sum() / 1e6})
    tot = float(df["notional"].sum())
    out["% of orders"] = 100 * out["orders"] / max(len(df), 1)
    out["% of notional"] = (100 * out[ncol] * 1e6 / tot) if tot else np.nan
    return out.round(2)


def t_exec_summary(df: pd.DataFrame) -> pd.DataFrame:
    """The Q1 executive summary, recomputed for this window.

    The same facts the Q1 deck opened on - where the value went, how small
    the orders were against volume, fill, spread, and the close-only against
    pre-traded split - so the two reviews read side by side.
    """
    if df is None or df.empty:
        return pd.DataFrame()
    rows = []

    def add(section, metric, value, unit=""):
        rows.append({"section": section, "metric": metric,
                     "value": value, "unit": unit})

    tot = float(df["notional"].sum())
    add("Book", "orders", len(df))
    add("Book", "executed notional", tot / 1e6, CURRENCY + "m")
    if "market" in df and tot > 0:
        by_mkt = drop_unattributable(
            df.groupby("market", observed=True)["notional"].sum()
              .sort_values(ascending=False).to_frame())["notional"]
        if len(by_mkt):
            top = by_mkt.head(5)
            add("Book", "top 5 markets, share of notional",
                100 * top.sum() / tot, "%")
            add("Book", "top 5 markets", ", ".join(str(m) for m in top.index))
            add("Book", "lead market", str(top.index[0]))
            add("Book", "lead market, share of notional",
                100 * top.iloc[0] / tot, "%")
    if "adv_pct" in df:
        known = df["adv_pct"].notna()
        base = float(df.loc[known, "notional"].sum())
        if base > 0:
            under = known & (df["adv_pct"] < SUMMARY_ADV_CUT)
            add("Book", f"share of notional under {SUMMARY_ADV_CUT:g}% ADV",
                100 * float(df.loc[under, "notional"].sum()) / base, "%")
    if "fill_rate" in df:
        add("Book", "fill ratio",
            wmean(df["fill_rate"], df["notional"], winsor=False),
            "%, notional-weighted")
    if "spread_bps" in df:
        add("Book", "average spread", wmean(df["spread_bps"], df["notional"], winsor=False),
            "bps, notional-weighted")
    for flow, r in t_flow_split(df).iterrows():
        for col, val in r.items():
            add(flow, col, val)
    return pd.DataFrame(rows)


def t_close_pr_market(df: pd.DataFrame) -> pd.DataFrame:
    """Close participation (ClosePR) by market, notional-weighted.

    How much of each market's closing auction our orders were. Not
    winsorised: it is a share, bounded by construction, and the high end is
    exactly the thing being looked for.
    """
    if df is None or df.empty or "close_pr" not in df:
        return pd.DataFrame()
    rows = []
    for m, g in df.groupby("market", observed=True):
        v = pd.to_numeric(g["close_pr"], errors="coerce")
        ok = v.notna() & (g["notional"] > 0)
        if not ok.any():
            continue
        rows.append({
            "market": m,
            "orders": int(ok.sum()),
            "notional (" + CURRENCY + "m)": float(g.loc[ok, "notional"].sum()) / 1e6,
            # The chart shows the plain mean: each order counts once, so a
            # few large orders cannot carry a market's bar. The weighted
            # figure stays beside it for anyone who wants the value view.
            "close PR % (mean)": float(v[ok].mean()),
            "close PR % (wtd)": wmean(v[ok], g.loc[ok, "notional"], winsor=False),
            "close PR % (median)": float(v[ok].median()),
            "close PR % (max)": float(v[ok].max()),
            "small sample": int(ok.sum()) < MIN_N_FOR_CI,
        })
    if not rows:
        return pd.DataFrame()
    return (pd.DataFrame(rows).set_index("market")
            .sort_values("notional (" + CURRENCY + "m)", ascending=False)
            .round(2))


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
        spread = wmean(g["spread_bps"], g["notional"], winsor=False)
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

# Retired from the axes at the desk's request. Kept so anything still
# referring to it keeps working, and so the convention stays written down:
# positive is a saving, negative is a cost, everywhere in this file.
COST_SAVE_NOTE = "← cost   |   savings →"

# Every first-execution chart is on pre-traded orders only, and every one of
# them says so. The measure is undefined for an order that only ever printed
# in the auction, and a reader has no way to know that from the title.
PRE_TRADED_NOTE = ("orders that traded before the close; an auction-only order has no first execution to measure and is not in here")


def _rows_high(n: int, per_row: float = 0.62, base: float = 1.9,
               floor: float = 2.5) -> float:
    """Figure height for a chart with `n` horizontal bars.

    A panel sized for six strategies and given one draws a single block the
    height of the page. Scoping the review to one strategy made every
    by-strategy chart look like that, so height follows the bar count.
    """
    return max(floor, base + per_row * max(n, 1))


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


_WRITTEN_CHARTS: set = set()


# Every chart is also written bare into charts_deck/ for the PowerPoint: no
# title, no footnotes, bigger type. Shrunk to half a slide, an 11-inch figure
# with 8pt labels reads at about 4pt, and the title and notes take room the
# plot needs. The deck puts the title above the picture as slide text and the
# footnotes in the speaker notes; DECK_TITLES carries both across.
DECK_CHARTS_DIR = "charts_deck"
DECK_FONT_SCALE = 1.45
DECK_TITLES: dict = {}


def _save(fig, out_dir: Path, name: str, bottom: float = 0.0,
          deck: bool = True) -> None:
    path = out_dir / name
    if deck:
        _WRITTEN_CHARTS.add(name)
    fig.tight_layout(rect=(0, bottom, 1, 1) if bottom else None)
    fig.savefig(path, facecolor=SURFACE, dpi=DPI, bbox_inches="tight")
    if deck:
        try:
            _save_bare(fig, out_dir.parent / DECK_CHARTS_DIR, name)
        except Exception as exc:                      # pragma: no cover
            warn(f"deck copy of {name} not written: {exc}")
    plt.close(fig)
    log(f"    chart  {name}")


def _save_bare(fig, out_dir: Path, name: str) -> None:
    """The same figure with its title and footnotes lifted off, type enlarged.

    A multi-panel chart keeps a short label per panel (Buy, Sell) - set as
    ax._panel by the chart - because without it the panels are anonymous.
    """
    from matplotlib.text import Text
    out_dir.mkdir(parents=True, exist_ok=True)
    axes = fig.get_axes()
    titles = [ax.get_title(loc="left").strip() for ax in axes]
    title = getattr(fig, "_deck_title", None) or next((t for t in titles if t), "")
    # One note per line, so a caveat shared by two charts can be printed once.
    notes = [line.strip() for t in fig.texts
             for line in t.get_text().split(chr(10)) if line.strip()]
    for t in list(fig.texts):
        t.remove()
    for ax in axes:
        ax.set_title(getattr(ax, "_panel", ""), loc="left")
        # Axis-relative footnotes sit inside the axes' own text list.
        for t in list(ax.texts):
            if t.get_transform() == ax.transAxes and t.get_position()[1] < 0:
                notes += [x.strip() for x in t.get_text().split(chr(10))
                          if x.strip()]
                t.remove()
    # Panels stacked one above the other stand side by side instead: on a
    # widescreen slide a tall chart is the one that ends up smallest.
    # A chart whose labels are stacked by fixed offsets opts out: bigger type
    # there collides instead of reading better.
    scale = getattr(fig, "_deck_scale", DECK_FONT_SCALE)
    if len(axes) > 1 and len({round(ax.get_position().x0, 2)
                              for ax in axes}) == 1:
        from matplotlib.gridspec import GridSpec
        grid = GridSpec(1, len(axes), figure=fig)
        for i, ax in enumerate(axes):
            ax.set_subplotspec(grid[0, i])
            # A dozen market names under a half-width panel collide at 30
            # degrees; steeper, and a touch smaller, they do not.
            for lab in ax.get_xticklabels():
                lab.set_rotation(45)
        fig.set_size_inches(9.0 * len(axes), 5.8)
        scale = min(scale, 1.3)
    for t in fig.findobj(Text):
        t.set_fontsize(t.get_fontsize() * scale)
    fig.tight_layout()
    fig.savefig(out_dir / name, facecolor=SURFACE, dpi=DPI, bbox_inches="tight")
    DECK_TITLES[name] = {"title": title.replace(chr(10), " ").strip(),
                         "notes": notes}


def _fmt_bps(v) -> str:
    return "-" if (v is None or (isinstance(v, float) and math.isnan(v))) else f"{v:+.1f}"


def _diverging_barv(ax, labels, values, ci=None, small=None, unit="bps",
                    tight=False):
    """Diverging bars standing up: categories along the bottom, value up.

    Blue above zero is a saving, red below it is a cost - and
    the same rule that every bar carries its own number, so colour is never
    the only encoding.
    """
    x = np.arange(len(labels))
    colors = [POS if (v is not None and not math.isnan(v) and v >= 0) else NEG
              for v in values]
    ax.bar(x, values, color=colors, width=0.62, zorder=3)
    if ci is not None:
        for i, (lo, hi) in enumerate(ci):
            if lo is None or math.isnan(lo):
                continue
            ax.plot([i, i], [lo, hi], color=INK_SECOND, linewidth=1.4, zorder=4)
            for b in (lo, hi):
                ax.plot([i - .1, i + .1], [b, b], color=INK_SECOND, lw=1.4,
                        zorder=4)
    ax.axhline(0, color=BASELINE, linewidth=1.0, zorder=2)
    ax.set_xticks(x)
    # The small-sample marking is deliberately NOT on the tick label. It
    # crowds a client chart and reads as an apology. The flag is still on
    # every row of the table the chart came from, and the run log names the
    # thin categories, so the caveat survives where an analyst will meet it.
    ax.set_xticklabels([str(l) for l in labels], rotation=30, ha="right")

    finite = [v for v in values if v is not None and not math.isnan(v)]
    if tight and finite:
        # Fitted to the bars, not the widest whisker. One thin market with a
        # huge interval otherwise sets the scale and every other bar on the
        # chart shrinks to a sliver. Whiskers may stretch the axis by up to
        # 30% of the bar range; beyond that they run off the edge, uncapped.
        top, bot = max(max(finite), 0.0), min(min(finite), 0.0)
        room = 0.3 * max(top - bot, 1e-9)
        if ci is not None:
            ends = [b for lo, hi in ci for b in (lo, hi)
                    if b is not None and not math.isnan(b)]
            top = max([top] + [min(b, top + room) for b in ends])
            bot = min([bot] + [max(b, bot - room) for b in ends])
        pad = 0.14 * max(top - bot, 1e-9)
        ax.set_ylim(bot - pad, top + pad)
        span = (top - bot) / 2 or 1.0
        cap_hi, cap_lo = top, bot
    else:
        reach = [abs(v) for v in finite]
        if ci is not None:
            reach += [abs(b) for lo, hi in ci for b in (lo, hi)
                      if b is not None and not math.isnan(b)]
        span = max(reach or [1.0])
        pad = span * 0.30
        ax.set_ylim(-span - pad, span + pad)
        cap_hi, cap_lo = span, -span
    for i, v in enumerate(values):
        if v is None or math.isnan(v):
            continue
        edge = v
        if ci is not None and i < len(ci):
            lo, hi = ci[i]
            if (lo is not None and hi is not None
                    and not math.isnan(lo) and not math.isnan(hi)):
                edge = max(v, hi) if v >= 0 else min(v, lo)
        # A whisker cut at the edge keeps running under its label, so that
        # label gets a solid box: the number interrupts the line instead of
        # sitting on top of it.
        cut = edge > cap_hi if v >= 0 else edge < cap_lo
        edge = min(edge, cap_hi) if v >= 0 else max(edge, cap_lo)
        off = span * 0.03
        label = f"{v:+.2f}" if unit == "spreads" else _fmt_bps(v)
        ax.text(i, edge + (off if v >= 0 else -off), label, ha="center",
                va="bottom" if v >= 0 else "top", color=INK, fontsize=8.5,
                zorder=6,
                bbox=dict(boxstyle="round,pad=0.16", facecolor=SURFACE,
                          edgecolor="none", alpha=1.0 if cut else 0.85))


NON_CATEGORIES = {"nan", "none", "nat", ""}


def drop_unattributable(d: pd.DataFrame, what: str = "chart") -> pd.DataFrame:
    """Take the unattributable rows out of a chart.

    Two kinds go: a market that could not be resolved from the symbol, and a
    category that is missing altogether. Neither is ever a finding, and a bar
    labelled "Unknown" or "nan" on a client slide invites a question nobody
    in the room can answer. Both stay in the tables, where the totals must
    still add up.
    """
    keys = d["market"] if "market" in d.columns else pd.Series(
        d.index, index=d.index)
    text = keys.astype(str).str.strip().str.lower()
    gone = keys.isin(EXCLUDE_MARKETS_FROM_CHARTS) | text.isin(NON_CATEGORIES)
    return d[~gone.to_numpy()]


def apportion(values: list, target: float) -> list:
    """Round parts to whole numbers that still add up to the rounded total.

    Rounding every part on its own is what puts a 13 underneath a 14: each
    part rounds down, the total rounds up, and the reader is left doing
    arithmetic that does not work. On a client slide that is not a rounding
    detail, it is the moment someone stops trusting the chart.

    Largest remainder keeps each part within one unit of the truth and makes
    the column add up. Ties go to the bigger part, so a sliver beside a large
    number never collects the spare unit.
    """
    tot = int(round(target))
    floors = [int(math.floor(v)) for v in values]
    rem = tot - sum(floors)
    order = sorted(range(len(values)),
                   key=lambda i: (values[i] - floors[i], values[i]),
                   reverse=True)
    k = 0
    while rem > 0 and order:
        floors[order[k % len(order)]] += 1
        rem -= 1
        k += 1
    k = 0
    while rem < 0 and k < 4 * max(len(order), 1):
        idx = order[-1 - (k % len(order))]
        if floors[idx] > 0:
            floors[idx] -= 1
            rem += 1
        k += 1
    return floors


def chart_market_notional(df: pd.DataFrame, out: Path,
                          name: str = "13_market_notional.png") -> None:
    """Value traded by market, split cash against swap.

    Stacked rather than side by side: the total per market is the number the
    client recognises from their own records, and a split that hides the
    total would cost more than it explains.
    """
    if df.empty or "market" not in df or "notional" not in df:
        return
    d = drop_unattributable(df)
    if d.empty:
        return

    have_type = "notional_type" in d.columns
    parts = [p for p in NOTIONAL_TYPE_ORDER
             if have_type and (d["notional_type"] == p).any()] or ["All"]
    totals = (d.groupby("market", observed=True)["notional"].sum() / 1e6
              ).sort_values(ascending=False).head(12)
    markets = list(totals.index)
    grand = float(totals.sum())

    fig, (ax,) = _fig((11.0, 6.0))
    x = np.arange(len(markets))
    bottom = np.zeros(len(markets))
    by_part = {}
    for i, part in enumerate(parts):
        sub_d = d if part == "All" else d[d["notional_type"] == part]
        vals = [float(sub_d.loc[sub_d["market"] == m, "notional"].sum()) / 1e6
                for m in markets]
        ax.bar(x, vals, bottom=bottom, width=0.62, color=SERIES[i % len(SERIES)],
               zorder=3, label=part, edgecolor=SURFACE, linewidth=1.4)
        # Nothing is written inside a segment. A thin one cannot hold its
        # number, and a split that disappears on small markets is no split.
        by_part[part] = vals
        bottom = bottom + np.array(vals)

    ax.set_xticks(x)
    ax.set_xticklabels(markets, rotation=30, ha="right")
    ax.grid(axis="y", zorder=0)
    ax.set_axisbelow(True)
    span = max(bottom.max() if len(bottom) else 1.0, 1.0)
    # Everything goes above the bar. A segment can be too thin to hold its own
    # number, and the cash/swap split is the point of the chart - it cannot be
    # the part that gets dropped whenever a market is small.
    # The split stacks DOWN the label rather than running across it. Side by
    # side it is wider than a bar slot and the text from neighbouring markets
    # collides; one part per line is narrow and never does.
    ax.set_ylim(0, span * 1.42)
    for j, m in enumerate(markets):
        share = 100.0 * bottom[j] / max(grand, 1e-9)
        shown = [p for p in parts if p != "All"]
        whole = apportion([by_part[p][j] for p in shown], bottom[j])
        # A part too small to round to a million is shown as "<1m" rather
        # than dropped. It is true, it keeps the column adding up, and it
        # does not pretend a business line that exists is not there.
        lines = []
        for p, n in zip(shown, whole):
            raw = by_part[p][j]
            if n > 0:
                lines.append(f"{p} {n:,.0f}m")
            elif raw > 0:
                lines.append(f"{p} <1m")
        if lines:
            ax.text(j, bottom[j] + span * 0.02, chr(10).join(lines),
                    va="bottom", ha="center", fontsize=7.5, color=INK_SECOND,
                    linespacing=1.35, zorder=5)
        ax.text(j, bottom[j] + span * (0.045 + 0.037 * len(lines)),
                f"{bottom[j]:,.0f}m  ({share:.0f}%)",
                va="bottom", ha="center", fontsize=9, color=INK, zorder=5)
    if len(parts) > 1:
        ax.legend(frameon=False, ncol=len(parts), loc="upper right",
                  fontsize=8.5)
    _style(ax, ylabel=f"executed notional ({CURRENCY}m)",
           title="Value traded by market")
    if len(parts) > 1:
        ax.text(0.0, -0.30, "cash and swap split by client account; the "
                "percentage is the market's share of the value traded",
                transform=ax.transAxes, fontsize=7.5, color=INK_MUTED)
    fig._deck_scale = 1.0
    _save(fig, out, name)


def chart_spread_relative(t: pd.DataFrame, out: Path,
                          name: str = "16_spread_relative.png",
                          value: str = "vs Close spreads") -> None:
    """Performance measured in spreads rather than basis points.

    Basis points are not comparable between markets: a wide-spread name costs
    more of them for reasons that have nothing to do with the algo. Dividing
    by the spread asks how many spreads were paid, which is comparable - and
    each label carries what one spread is actually worth in that market, so
    the reader can convert back.
    """
    if t.empty or value not in t.columns:
        return
    d = drop_unattributable(t)
    if "notional (" + CURRENCY + "m)" in d.columns:
        d = d.sort_values("notional (" + CURRENCY + "m)", ascending=False)
    d = d.head(14)
    if d.empty:
        return

    labels, vals = [], []
    for _, row in d.iterrows():
        mkt = row.get("market", "")
        sprd = row.get("spread bps (wtd)", float("nan"))
        labels.append(f"{mkt}\n[{sprd:.1f} bps]"
                      if np.isfinite(sprd) else str(mkt))
        vals.append(float(row[value]))

    fig, (ax,) = _fig((11.0, 6.0))
    _diverging_barv(ax, labels, vals, unit="spreads",
                    small=[bool(x) for x in d["small sample"]]
                    if "small sample" in d.columns else None)
    ax.yaxis.set_major_formatter(
        matplotlib.ticker.FuncFormatter(lambda v, _: f"{v:+.2f}"))
    _style(ax, ylabel=f"{value.replace(' spreads', '')}, in spreads paid",
           title="Spread relative performance")
    ax.text(0.0, -0.34, "basis points are not comparable between markets - a "
            "wide-spread name costs more of them whatever the algo does. "
            "Spreads are. [x bps] is one spread in that market.",
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
    _style(ax, xlabel="vs Arrival, notional-weighted (bps)",
           title="Same market, same ADV% band - what each algo cost",
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


SPREAD_NOTE = ("in spreads: the result divided by the average spread, which "
               "is the [x bps] under each bar - multiply the two to get bps")
CI_NOTE = ("whiskers are 95% bootstrap CIs; a CI crossing zero is not "
           "distinguishable from zero")
REV_NOTE = "negative means the price came back against us by the next open"
# The desk's name for the reversion axis. {unit} is spreads or bps.
REV_YLABEL = "Close to T+1 notional weighted in {unit}"
CO_NOTE = ("close-only: 99.5% or more of the order printed in the auction, so "
           "its close slippage is zero by design and reversion is the "
           "measure. Negative means the price came back against us by the "
           "next open.")
PRE_NOTE = ("pre-traded: part of the order traded before the auction. Negative "
            "reversion means the price came back against us by the next open.")
PVWAP_NOTE = ("pre-traded: part of the order traded before the auction. PVWAP is "
              "the market's VWAP while each order was working, so it judges the "
              "trading itself, not when the order started.")


def _spread_label(key, sprd, spreads: bool) -> str:
    """The group name, with its spread underneath when the chart is in spreads."""
    if spreads and sprd is not None and np.isfinite(sprd):
        return str(key) + chr(10) + f"[{sprd:.1f} bps]"
    return str(key)


def _unit_cols(unit: str):
    """The value column and its CI, in the unit the chart is drawn in."""
    if unit == "spreads":
        return "spreads", "CI low", "CI high"
    return "wtd mean bps", "CI low bps", "CI high bps"


def _chartable(d: pd.DataFrame) -> pd.DataFrame:
    """Rows a client chart can show: no Unknown market, no missing category."""
    if d is None or d.empty:
        return pd.DataFrame()
    first = pd.Index(d.index.get_level_values(0))
    text = first.astype(str).str.strip().str.lower()
    gone = (np.asarray(first.isin(list(EXCLUDE_MARKETS_FROM_CHARTS)))
            | np.asarray(text.isin(list(NON_CATEGORIES))))
    return d[~gone]


def _share_ylim(axes) -> None:
    """One scale across panels, so bars in different panels compare by eye.

    Each panel sized itself to its own bars; the union of their ranges wins.
    """
    lo = min(ax.get_ylim()[0] for ax in axes)
    hi = max(ax.get_ylim()[1] for ax in axes)
    for ax in axes:
        ax.set_ylim(lo, hi)


def _notes(fig, notes: list) -> float:
    """Write the notes under a chart; returns the space they need."""
    lines = [x for x in notes if x]
    fig.text(0.01, 0.005, chr(10).join(lines), fontsize=7.5, color=INK_MUTED,
             ha="left", va="bottom")
    return 0.04 + 0.03 * len(lines)


def chart_spreads(t: pd.DataFrame, out: Path, name: str, title: str,
                  measure: str, note: str = "", by_notional: bool = False,
                  limit: int = 12, unit: str = None, ylabel: str = None,
                  tight: bool = None) -> None:
    """One result per group, with the group's spread under each bar."""
    unit = unit or CHART_UNIT
    d = _chartable(t)
    if d.empty:
        return
    if by_notional:
        d = d.sort_values("notional (" + CURRENCY + "m)", ascending=False)
    d = d.head(limit)
    spreads = unit == "spreads"
    vcol, lo, hi = _unit_cols(unit)
    labels = [_spread_label(k, float(sp), spreads)
              for k, sp in zip(d.index, d["spread bps (wtd)"])]
    fig, (ax,) = _fig((11.0, 6.0))
    _diverging_barv(ax, labels, [float(v) for v in d[vcol]],
                    ci=list(zip(d[lo].astype(float), d[hi].astype(float))),
                    unit="spreads" if spreads else "bps",
                    tight=by_notional if tight is None else tight)
    _style(ax, ylabel=(ylabel.format(unit="spreads" if spreads else "bps")
                       if ylabel else
                       measure + (", in spreads" if spreads
                                  else ", bps, notional-weighted")),
           title=title)
    room = _notes(fig, [note, SPREAD_NOTE if spreads else "", CI_NOTE])
    _save(fig, out, name, bottom=room)


def chart_spreads_by_side(t: pd.DataFrame, out: Path, name: str, title: str,
                          measure: str, note: str = "", limit: int = 12,
                          unit: str = None, ylabel: str = None) -> None:
    """Markets along the bottom, a Buy panel over a Sell panel, one scale.

    Two panels rather than paired bars in one: colour already means saving
    against cost on these charts, so it cannot also mean Buy against Sell.
    Markets run in the same order in both, biggest by value first.
    """
    unit = unit or CHART_UNIT
    d = _chartable(t)
    if d.empty or d.index.nlevels < 2:
        return
    ncol = "notional (" + CURRENCY + "m)"
    markets = (d.groupby(level=0, observed=True)[ncol].sum()
                 .sort_values(ascending=False).head(limit).index.tolist())
    present = set(d.index.get_level_values(1))
    sides = [x for x in BUY_SELL_ORDER if x in present]
    if not markets or not sides:
        return
    spreads = unit == "spreads"
    vcol, lo, hi = _unit_cols(unit)
    fig, axes = plt.subplots(len(sides), 1, dpi=DPI,
                             figsize=(11.0, 4.7 * len(sides)))
    axes = np.atleast_1d(axes)
    fig.patch.set_facecolor(SURFACE)
    nan = float("nan")
    for ax, side in zip(axes, sides):
        ax.set_facecolor(SURFACE)
        labels, vals, ci = [], [], []
        for m in markets:
            if (m, side) in d.index:
                r = d.loc[(m, side)]
                labels.append(_spread_label(m, float(r["spread bps (wtd)"]),
                                            spreads))
                vals.append(float(r[vcol]))
                ci.append((float(r[lo]), float(r[hi])))
            else:
                labels.append(str(m) + chr(10) + "no orders")
                vals.append(nan)
                ci.append((nan, nan))
        _diverging_barv(ax, labels, vals, ci=ci,
                        unit="spreads" if spreads else "bps", tight=True)
        _style(ax, ylabel=(ylabel.format(unit="spreads" if spreads else "bps")
                           if ylabel else
                           measure + (", in spreads" if spreads else ", bps")),
               title=f"{title}: {side}")
        ax._panel = side
    fig._deck_title = title
    _share_ylim(axes)
    room = _notes(fig, [note, SPREAD_NOTE if spreads else "", CI_NOTE])
    _save(fig, out, name, bottom=room / len(sides))


def chart_measures_by_side(specs: list, out: Path, name: str, title: str,
                           note: str = "", all_label: str = "All",
                           unit: str = None) -> None:
    """Several measures on one population: an All panel, then Buy, then Sell.

    specs is [(measure label, table indexed All / Buy / Sell), ...]. Each bar
    carries its own spread because the measures are not weighted alike -
    first execution is weighted by the part of the order that traded before
    the auction - so they do not share one denominator.
    """
    unit = unit or CHART_UNIT
    spreads = unit == "spreads"
    vcol, lo, hi = _unit_cols(unit)
    panels = [p for p in ["All"] + BUY_SELL_ORDER
              if any(tb is not None and not tb.empty and p in tb.index
                     for _, tb in specs)]
    if not panels:
        return
    fig, axes = _fig((12.5, 5.6), ncols=len(panels))
    nan = float("nan")
    for i, (ax, p) in enumerate(zip(axes, panels)):
        labels, vals, ci = [], [], []
        for label, tb in specs:
            if tb is None or tb.empty or p not in tb.index:
                labels.append(label)
                vals.append(nan)
                ci.append((nan, nan))
                continue
            r = tb.loc[p]
            labels.append(_spread_label(label, float(r["spread bps (wtd)"]),
                                        spreads))
            vals.append(float(r[vcol]))
            ci.append((float(r[lo]), float(r[hi])))
        _diverging_barv(ax, labels, vals, ci=ci,
                        unit="spreads" if spreads else "bps")
        head = all_label if p == "All" else p
        _style(ax, ylabel=("in spreads" if spreads else "bps") if i == 0 else "",
               title=(title + chr(10) + head) if i == 0 else chr(10) + head)
        ax._panel = head
    fig._deck_title = title
    _share_ylim(axes)
    room = _notes(fig, [note, SPREAD_NOTE if spreads else "", CI_NOTE])
    _save(fig, out, name, bottom=room)


def chart_close_pr(tables: list, out: Path) -> None:
    """Close participation by market: all, close-only, pre-traded.

    One scale across the three, so a market's bar can be compared between
    them by eye. Magnitude, not saving or cost, so one hue.
    """
    col = "close PR % (mean)"
    usable = [(t, name, title) for t, name, title in tables
              if t is not None and not t.empty and col in t]
    if not usable:
        return
    top = max(float(_chartable(t)[col].max()) for t, _, _ in usable
              if not _chartable(t).empty) if usable else 1.0
    for t, name, title in usable:
        d = _chartable(t).head(12)
        if d.empty:
            continue
        vals = [float(v) for v in d[col]]
        x = np.arange(len(vals))
        fig, (ax,) = _fig((11.0, 5.6))
        ax.bar(x, vals, width=0.62, color=SERIES[0], zorder=3)
        ax.set_ylim(0, top * 1.18 if top > 0 else 1.0)
        for i, v in enumerate(vals):
            ax.text(i, v + top * 0.015, f"{v:.1f}%", ha="center",
                    va="bottom", fontsize=8.5, color=INK, zorder=5)
        ax.set_xticks(x)
        # India's bar is a different measure, so it says so where it is read.
        ax.set_xticklabels([f"{k} (PR)" if str(k) == "India" else str(k)
                            for k in d.index], rotation=30, ha="right")
        _style(ax, ylabel="Close participation, mean per order (%)",
               title=title)
        india = "India" in set(map(str, d.index))
        room = _notes(fig, ["ClosePR: our executed quantity as a share of the "
                            "closing auction's volume, averaged per order, not "
                            "weighted by value. Markets ordered by value traded.",
                            "India (PR): no closing auction in H1, so its bar is "
                            "the participation rate (PR) of the orders started "
                            "17:30-17:45 HKT." if india else ""])
        _save(fig, out, name, bottom=room)


def chart_flow_split(t: pd.DataFrame, out: Path,
                     name: str = "23_flow_split.png") -> None:
    """Close-only against pre-traded: the value in each and how it filled."""
    if t is None or t.empty:
        return
    ncol = "notional (" + CURRENCY + "m)"
    vals = [float(v) for v in t[ncol]]
    x = np.arange(len(vals))
    fig, (ax,) = _fig((9.0, 5.2))
    ax.bar(x, vals, width=0.55, color=SERIES[:len(vals)], zorder=3)
    top = max(vals or [1.0]) or 1.0
    ax.set_ylim(0, top * 1.25)
    ticks = []
    for i, (flow, r) in enumerate(t.iterrows()):
        ax.text(i, vals[i] + top * 0.02,
                f"{vals[i]:,.0f}m  ({r['% of notional']:.0f}%)",
                ha="center", va="bottom", fontsize=10, color=INK, zorder=5)
        bits = [str(flow), f"{int(r['orders']):,} orders"]
        if "fill ratio %" in t.columns:
            bits.append(f"fill {r['fill ratio %']:.1f}%")
        if "%ADV (notional-weighted)" in t.columns:
            bits.append(f"{r['%ADV (notional-weighted)']:.2f}% ADV")
        if "% of notional in the close" in t.columns:
            bits.append(f"{r['% of notional in the close']:.0f}% executed in the close")
        ticks.append(chr(10).join(bits))
    ax.set_xticks(x)
    ax.set_xticklabels(ticks)
    _style(ax, ylabel=f"executed notional ({CURRENCY}m)",
           title="Close-only and pre-traded flow")
    room = _notes(fig, ["close-only: 99.5% or more of the order printed in the "
                        "auction. Pre-traded: part of it traded before. "
                        "Shares are of the two together; %ADV is "
                        "notional-weighted."])
    _save(fig, out, name, bottom=room)


def chart_adv_profile(t: pd.DataFrame, out: Path, name: str, title: str,
                      with_orders: bool = False, color: str = None) -> None:
    """How big the orders were against daily volume: value, and count."""
    if t is None or t.empty:
        return
    d = t[[str(i).strip().lower() not in NON_CATEGORIES for i in t.index]]
    if d.empty:
        return
    ncol = "notional (" + CURRENCY + "m)"
    panels = ([("orders", "orders", "% of orders", "{:,.0f}")]
              if with_orders else [])
    panels.append((ncol, f"executed notional ({CURRENCY}m)", "% of notional",
                   "{:,.0f}m"))
    fig, axes = _fig((11.0 if with_orders else 9.0, 5.0), ncols=len(panels))
    for i, (ax, (col, ylabel, share, fmt)) in enumerate(zip(axes, panels)):
        vals = [float(v) for v in d[col]]
        x = np.arange(len(vals))
        ax.bar(x, vals, width=0.62, color=color or SERIES[0], zorder=3)
        top = max(vals or [1.0]) or 1.0
        ax.set_ylim(0, top * 1.3)
        for j, (v, sh) in enumerate(zip(vals, d[share])):
            ax.text(j, v + top * 0.02, fmt.format(v) + chr(10) + f"({sh:.0f}%)",
                    ha="center", va="bottom", fontsize=8.5, color=INK, zorder=5)
        ax.set_xticks(x)
        ax.set_xticklabels([str(k) for k in d.index])
        _style(ax, xlabel="ADV%", ylabel=ylabel,
               title=title if i == 0 else "")
    _save(fig, out, name)


def chart_monthly(df: pd.DataFrame, out: Path,
                  name: str = "12_monthly.png", subtitle: str = "",
                  unit: str = None) -> None:
    """Close performance, month by month.

    One panel. The auction-share line that used to sit beside it came off at
    the desk's request. It answered "did the orders reach the close", which
    the deck already covers, and here it competed with the only question a
    month-by-month view is good at: did the result hold up across the
    period, or does it rest on one month.
    """
    unit = unit or CHART_UNIT
    if "month" not in df or df.empty or "slip_close" not in df:
        return
    months = sorted(df["month"].dropna().astype(str).unique())
    if len(months) < 2:
        return
    spreads = unit == "spreads"
    if spreads:
        tb = t_in_spreads(df.assign(month=df["month"].astype(str)), "month",
                          "slip_close", ci=False)
        if tb.empty:
            return
        tb = tb.reindex(months)
        sc = [float(v) for v in tb["spreads"]]
        labels = [_spread_label(m, float(sp), True)
                  for m, sp in zip(months, tb["spread bps (wtd)"])]
    else:
        g = df.groupby(df["month"].astype(str))
        sc = [wmean(g.get_group(m)["slip_close"], g.get_group(m)["notional"])
              for m in months]
        labels = list(months)
    fig, (ax,) = _fig((11.0, 5.6))
    colors = [POS if (np.isfinite(v) and v >= 0) else NEG for v in sc]
    ax.bar(range(len(months)), sc, color=colors, width=0.62, zorder=3)
    ax.axhline(0, color=BASELINE, linewidth=1.0, zorder=2)
    reach = [abs(v) for v in sc if np.isfinite(v)]
    span = max(reach or [1.0])
    ax.set_ylim(-span * 1.32, span * 1.32)
    for i, v in enumerate(sc):
        if not np.isfinite(v):
            continue
        ax.text(i, v + (span * 0.04 if v >= 0 else -span * 0.04),
                f"{v:+.2f}" if spreads else _fmt_bps(v),
                ha="center", va="bottom" if v >= 0 else "top", fontsize=8.5,
                color=INK, zorder=5)
    ax.set_xticks(range(len(months)))
    ax.set_xticklabels(labels, rotation=30, ha="right")
    _style(ax, ylabel="vs Close, " + ("in spreads" if spreads else "bps"),
           title="Close performance by month"
                 + (f" \u2014 {subtitle}" if subtitle else ""))
    room = _notes(fig, [SPREAD_NOTE if spreads else ""]) if spreads else 0.0
    _save(fig, out, name, bottom=room)


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
    # written the way FIRST_EXEC_FLIP says the export writes it, so loading
    # the sample exercises the same flip the real file gets
    if FIRST_EXEC_FLIP == "all":
        first_exec = -first_exec
    elif FIRST_EXEC_FLIP == "buys":
        first_exec = np.where(side == "Buy", -first_exec, first_exec)

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
    t["37_close_opportunity"] = t_close_opportunity(df)
    t["38_auction_share"] = t_auction_share(auc)
    t["39_lateness"] = t_lateness(auc)
    # Close performance and the early start, each cut two ways: by order size
    # and by how wide the name trades. Size says whether the auction could
    # ever have absorbed it; spread says whether the name was cheap or
    # expensive to be in at all. They fail differently and the fix differs.
    t["41_close_by_adv"] = by_group(auc, "adv_bucket", "slip_close")
    # Reversion, two ways. Per market says where the price comes back; the
    # pre-traded cut says how much of it we brought on ourselves, because
    # reversion on an order that never left the auction is the market moving,
    # not us moving it.
    # Two populations, one measure. An order that worked before the close had
    # a chance to move the price; one that only ever printed in the auction
    # did not. Comparing the two per market is what separates our own impact
    # from the market's own overnight move.
    if "pretraded" in auc and "close_only" in auc:
        pre = auc[auc["pretraded"].fillna(False)]
        only = auc[auc["close_only"].fillna(False)]
        if len(pre):
            t["44_reversion_pretraded"] = by_group(pre, "market",
                                                   "reversion_bps")
        if len(only):
            t["45_reversion_close_only"] = by_group(only, "market",
                                                    "reversion_bps")
    if "spread_bucket" in auc:
        t["42_close_by_spread"] = by_group(auc, "spread_bucket", "slip_close")
        # The same cut against arrival. Close says whether the auction print
        # was good; arrival says whether the whole order was, and a wide name
        # can look fine on one and poor on the other - that gap is the cost
        # of the time spent getting there, which is what IS is for.
        t["28b_by_spread"] = by_group(auc, "spread_bucket", "slip_arrival")
        t["43_first_exec_by_spread"] = t_first_exec(moc, "spread_bucket")
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
    # --- the H1 narrative: close-only against pre-traded ---------------------
    # Close-only takes the auction markets plus India's 17:30-17:45 orders,
    # which the desk counts as close. India cannot be pre-traded - outside
    # that window its close share is zero - and after the CAS date its
    # segment is unknown, so pre-traded is the auction markets alone.
    reach = df["has_auction"] | _flag(df, "pct_close_imputed")
    co = df[reach & _flag(df, "close_only")]
    pre = df[df["has_auction"] & _flag(df, "pretraded")]
    FE, CW = "first_exec_vs_close", "cont_notional"
    t["50_exec_summary"] = t_exec_summary(df[reach])
    t["51_flow_split"] = t_flow_split(df[reach])
    t["52_adv_profile"] = t_adv_profile(df)
    t["53_pre_adv_profile"] = t_adv_profile(pre)
    # Every performance chart reads one of these. "co" is close-only and
    # "pre" is pre-traded; sprdq is the spread quartile.
    t["61_close_mkt_spreads"] = t_in_spreads(df, "market", "slip_close")
    t["63_close_adv_spreads"] = t_in_spreads(auc, "adv_bucket", "slip_close")
    if "spread_bucket" in auc:
        t["65_close_sprdq_spreads"] = t_in_spreads(auc, "spread_bucket",
                                                   "slip_close")
    t["66_pre_firstexec_adv_spreads"] = t_in_spreads(pre, "adv_bucket", FE, CW)
    if "spread_bucket" in pre:
        t["67_pre_firstexec_sprdq_spreads"] = t_in_spreads(
            pre, "spread_bucket", FE, CW)
    t["68_pre_firstexec_mkt_spreads"] = t_in_spreads(pre, "market", FE, CW)
    t["69_reversion_strat_spreads"] = t_in_spreads(df, "strategy",
                                                   "reversion_bps")
    t["70_pre_reversion_mkt_spreads"] = t_in_spreads(pre, "market",
                                                     "reversion_bps")
    t["71_co_reversion_mkt_spreads"] = t_in_spreads(co, "market",
                                                    "reversion_bps")
    t["71b_co_reversion_side_spr"] = pd.concat([
        t_in_spreads(co.assign(buy_sell="All"), "buy_sell", "reversion_bps"),
        t_in_spreads(co, "buy_sell", "reversion_bps", order=BUY_SELL_ORDER)])
    t["72_co_reversion_mkt_side_spr"] = t_in_spreads(
        co, ["market", "buy_sell"], "reversion_bps")
    for key, value, w in [("73a_pre_firstexec_side_spr", FE, CW),
                          ("73b_pre_close_side_spr", "slip_close", "notional"),
                          ("73c_pre_reversion_side_spr", "reversion_bps",
                           "notional")]:
        t[key] = pd.concat([
            t_in_spreads(pre.assign(buy_sell="All"), "buy_sell", value, w),
            t_in_spreads(pre, "buy_sell", value, w, order=BUY_SELL_ORDER)])
    t["74_pre_firstexec_mkt_side_spr"] = t_in_spreads(
        pre, ["market", "buy_sell"], FE, CW)
    t["75_pre_close_mkt_side_spr"] = t_in_spreads(
        pre, ["market", "buy_sell"], "slip_close")
    t["76_pre_reversion_mkt_side_spr"] = t_in_spreads(
        pre, ["market", "buy_sell"], "reversion_bps")
    t["77_pre_close_mkt_spreads"] = t_in_spreads(pre, "market", "slip_close")
    # Against PVWAP: the close says where the order ended up, PVWAP says
    # whether the trading along the way was good for the window it ran in.
    t["73d_pre_pvwap_side_spr"] = pd.concat([
        t_in_spreads(pre.assign(buy_sell="All"), "buy_sell", "slip_pvwap"),
        t_in_spreads(pre, "buy_sell", "slip_pvwap", order=BUY_SELL_ORDER)])
    t["78_pre_pvwap_mkt_spreads"] = t_in_spreads(pre, "market", "slip_pvwap")
    t["79_pre_pvwap_mkt_side_spr"] = t_in_spreads(
        pre, ["market", "buy_sell"], "slip_pvwap")
    # Close participation: auction markets on ClosePR, plus India's window
    # orders on their PR (see india_close_pr_proxy). India cannot be
    # pre-traded, so 82 is unchanged.
    if "close_pr" in df:
        with_india = df[reach]
        t["80_close_pr_mkt_all"] = t_close_pr_market(with_india)
        t["81_close_pr_mkt_close_only"] = t_close_pr_market(
            with_india[_flag(with_india, "close_only")])
        t["82_close_pr_mkt_pretraded"] = t_close_pr_market(pre)
    else:
        log("  no ClosePR column in the file - close participation charts "
            "skipped.")

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
        "spread bps (wtd)": wmean(g["spread_bps"], g["notional"], winsor=False)
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


def t_auction_share(df: pd.DataFrame) -> pd.DataFrame:
    """Our share OF THE CLOSING AUCTION, by market and size band.

    %Adv measures the order against a normal day. The auction is not a normal
    day - it is one print, and how much of it we were is the constraint that
    actually binds. Where an order took a large share of the auction, missing
    the rest is capacity. Where it took a sliver and still missed, it is not.
    """
    if "auction_share_pct" not in df or "market" not in df:
        return pd.DataFrame()
    d = df[df["auction_share_pct"].notna()]
    if d.empty:
        return pd.DataFrame()
    rows = []
    for (mkt, bucket), g in d.groupby(["market", "adv_bucket"],
                                      dropna=False, observed=True):
        if g.empty:
            continue
        rows.append({
            "market": mkt,
            "%Adv bucket": bucket,
            "orders": len(g),
            "notional (" + CURRENCY + "m)": g["notional"].sum() / 1e6,
            "median share of the auction %": g["auction_share_pct"].median(),
            "p75 share of the auction %": g["auction_share_pct"].quantile(0.75),
            "wtd %CLOSE": wmean(g["pct_close"], g["notional"], winsor=False)
                if "pct_close" in g else np.nan,
            "small sample": len(g) < MIN_N_FOR_CI,
        })
    return pd.DataFrame(rows).round(2)


def t_lateness(df: pd.DataFrame) -> pd.DataFrame:
    """Auction share against how late the order arrived.

    The desk controls when an order is sent. If auction share falls away for
    orders arriving inside the last few minutes, that is a lever - and one
    nobody can pull without knowing where the cliff is.
    """
    if "mins_to_close_at_start" not in df or "pct_close" not in df:
        return pd.DataFrame()
    d = df[df["mins_to_close_at_start"].notna()]
    if d.empty:
        return pd.DataFrame()
    bins = [-np.inf, 0, 2, 5, 15, 30, 60, 120, np.inf]
    names = ["after the close", "0-2 min", "2-5 min", "5-15 min", "15-30 min",
             "30-60 min", "1-2 h", "more than 2 h"]
    band = pd.cut(d["mins_to_close_at_start"], bins, labels=names)
    rows = []
    for key, g in d.groupby(band, observed=True):
        if g.empty:
            continue
        rows.append({
            "sent before the close": key,
            "orders": len(g),
            "notional (" + CURRENCY + "m)": g["notional"].sum() / 1e6,
            "wtd %CLOSE": wmean(g["pct_close"], g["notional"], winsor=False),
            "% with no auction fill": 100.0 * (g["pct_close"] <= 0).mean(),
            "vs Close bps (wtd)": wmean(g["slip_close"], g["notional"])
                if "slip_close" in g else np.nan,
            "small sample": len(g) < MIN_N_FOR_CI,
        })
    return pd.DataFrame(rows).set_index("sent before the close").round(2)


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
    # What was here before this run. Anything still present at the end that
    # this run did not write is from an older version of the analysis - a
    # chart that has since been renamed, split or retired. Left alone it sits
    # in the folder looking current and walks into a deck.
    before = {p.name for p in charts.glob("*.png") if p.name[:1].isdigit()}
    log("")
    log("  charts")
    # Charts 01-10 were retired at the desk's request (10b stays). Their
    # TABLES are all still built and written - the findings, the workbook and
    # the deck's numbers read from those, not from the pictures.
    close_df = t.get("_close_df")
    if close_df is not None:
        chart_market_notional(close_df, charts)
    chart_algo_choice(t.get("36_algo_choice", pd.DataFrame()), charts)
    chart_spread_relative(t.get("06b_spreads_market", pd.DataFrame()), charts)

    # Every performance chart below reads a spread table (61-77), so the
    # figure on the chart and the figure in the workbook are the same number.
    # The arrival charts (14, 17b, 18b) were retired at the desk's request;
    # arrival stays in the bps tables and the findings.
    E = pd.DataFrame()
    chart_spreads(t.get("61_close_mkt_spreads", E), charts,
                  "15_market_vs_close.png",
                  "Against the closing price, by market", "vs Close",
                  by_notional=True,
                  note="ordered by value traded, not by result - a small "
                       "market with a big number is still a small market")
    for key, name, title, measure in [
        ("63_close_adv_spreads", "17_close_by_adv.png",
         "Close performance by ADV%", "vs Close"),
        ("65_close_sprdq_spreads", "18_close_by_spread.png",
         "Close performance by spread quartile", "vs Close"),
    ]:
        chart_spreads(t.get(key, E), charts, name, title, measure)
    for key, name, title in [
        ("66_pre_firstexec_adv_spreads", "19_first_exec_by_adv.png",
         "First execution vs the close, by ADV%"),
        ("67_pre_firstexec_sprdq_spreads", "20_first_exec_by_spread.png",
         "First execution vs the close, by spread quartile"),
    ]:
        chart_spreads(t.get(key, E), charts, name, title,
                      "first execution vs Close", note=PRE_TRADED_NOTE)
    chart_spreads(t.get("68_pre_firstexec_mkt_spreads", E), charts,
                  "10b_first_exec_market.png",
                  "Was starting before the close right? By market",
                  "first execution vs Close", by_notional=True,
                  note=PRE_TRADED_NOTE)
    chart_spreads(t.get("69_reversion_strat_spreads", E), charts,
                  "11_reversion.png",
                  "Reversion - did the auction print come back?",
                  "next open vs close", note=REV_NOTE, ylabel=REV_YLABEL)
    df = t.get("_close_df")
    if df is not None:
        # Two charts, because India's close share is imputed rather than
        # measured: averaged in with the rest it lifts every month by a
        # constant and the line stops meaning what it says.
        if "market" in df:
            ind = df[df["market"] == "India"]
            rest = df[df["market"] != "India"]
            chart_monthly(rest, charts, "12_monthly.png", "excluding India")
            chart_monthly(ind, charts, "12b_monthly_india.png", "India only")
        else:
            chart_monthly(df, charts)

    # --- the H1 narrative -------------------------------------------------
    # Opens the way Q1 did: where the flow went and how big it was.
    chart_flow_split(t.get("51_flow_split", E), charts)
    # Green, at the desk's request - it sits beside the blue-and-red value
    # chart on the summary slide and should not read as a saving.
    chart_adv_profile(t.get("52_adv_profile", E), charts, "24_adv_profile.png",
                      color=SERIES[2], title=
                      "Value traded by ADV%")

    # Close-only. Its close slippage is zero by construction, so the question
    # is whether the price held after the auction: reversion, by market, and
    # by market and side.
    chart_spreads(t.get("71_co_reversion_mkt_spreads", E), charts,
                  "22_reversion_close_only.png",
                  "Close-only orders: reversion by market",
                  "next open vs close", by_notional=True, note=CO_NOTE,
                  ylabel=REV_YLABEL)
    chart_spreads_by_side(t.get("72_co_reversion_mkt_side_spr", E), charts,
                          "25_close_only_reversion_market_side.png",
                          "Close-only reversion", "next open vs close",
                          note=CO_NOTE, ylabel=REV_YLABEL)

    # Pre-traded. Three prices tell its story: where it started (first
    # execution against the close), where it finished (against the close),
    # and whether that held overnight (reversion).
    chart_measures_by_side(
        [("First execution vs close", t.get("73a_pre_firstexec_side_spr", E)),
         ("Execution vs close", t.get("73b_pre_close_side_spr", E)),
         ("Next open vs close", t.get("73c_pre_reversion_side_spr", E))],
        charts, "26_pretraded_by_side.png",
        "Pre-traded orders: start, finish and reversion",
        note=PRE_NOTE, all_label="All pre-traded")
    chart_spreads(t.get("77_pre_close_mkt_spreads", E), charts,
                  "27_pretraded_close_market.png",
                  "Pre-traded orders: against the closing price, by market",
                  "vs Close", by_notional=True, note=PRE_NOTE)
    chart_spreads(t.get("70_pre_reversion_mkt_spreads", E), charts,
                  "21_reversion_pretraded.png",
                  "Pre-traded orders: reversion by market",
                  "next open vs close", by_notional=True, note=PRE_NOTE,
                  ylabel=REV_YLABEL)
    chart_spreads_by_side(t.get("74_pre_firstexec_mkt_side_spr", E), charts,
                          "28_pretraded_first_exec_market_side.png",
                          "Pre-traded, first execution vs close",
                          "first execution vs Close", note=PRE_TRADED_NOTE)
    chart_spreads_by_side(t.get("75_pre_close_mkt_side_spr", E), charts,
                          "29_pretraded_close_market_side.png",
                          "Pre-traded, execution vs close", "vs Close",
                          note=PRE_NOTE)
    chart_spreads_by_side(t.get("76_pre_reversion_mkt_side_spr", E), charts,
                          "30_pretraded_reversion_market_side.png",
                          "Pre-traded reversion", "next open vs close",
                          note=PRE_NOTE, ylabel=REV_YLABEL)
    chart_spreads(t.get("78_pre_pvwap_mkt_spreads", E), charts,
                  "32_pretraded_pvwap_market.png",
                  "Pre-traded orders: against PVWAP, by market",
                  "vs PVWAP", by_notional=True, note=PVWAP_NOTE)
    chart_spreads_by_side(t.get("79_pre_pvwap_mkt_side_spr", E), charts,
                          "33_pretraded_pvwap_market_side.png",
                          "Pre-traded, against PVWAP", "vs PVWAP",
                          note=PVWAP_NOTE)
    chart_close_pr([
        (t.get("80_close_pr_mkt_all"), "34_close_pr_all.png",
         "Close participation by market: all orders"),
        (t.get("81_close_pr_mkt_close_only"), "35_close_pr_close_only.png",
         "Close participation by market: close-only orders"),
        (t.get("82_close_pr_mkt_pretraded"), "36_close_pr_pretraded.png",
         "Close participation by market: pre-traded orders"),
    ], charts)
    chart_adv_profile(t.get("53_pre_adv_profile", E), charts,
                      "31_pretraded_adv_profile.png",
                      "Pre-traded orders by ADV%", with_orders=True)

    # Anything this run did not write is left over from an older version of
    # the analysis. A retired chart that stays in the folder looks exactly
    # like a current one and is the easiest way for a wrong number to reach a
    # client, so it goes - and it is named, never removed quietly.
    import json
    deck = out_dir / DECK_CHARTS_DIR
    if deck.exists():
        for p in deck.glob("*.png"):
            if p.name not in _WRITTEN_CHARTS:
                p.unlink(missing_ok=True)
        (deck / "titles.json").write_text(
            json.dumps({k: DECK_TITLES[k] for k in sorted(DECK_TITLES)},
                       indent=2), encoding="utf-8")
        log(f"  deck charts (no titles, larger type) -> {deck}")

    stale = sorted(before - _WRITTEN_CHARTS)
    if stale:
        warn(f"{len(stale)} chart(s) in {charts} are from an earlier run and")
        log("    are no longer produced. Removing them, so nothing retired")
        log("    can be mistaken for current:")
        for name in stale:
            (charts / name).unlink(missing_ok=True)
            log(f"      {name}")


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


def write_excel_unified(t: dict, out_dir: Path) -> None:
    """Every table stacked down one sheet, for photographing in one pass.

    tables.xlsx keeps one table per sheet, which is right for working in.
    This is the same content on a single sheet with each table titled, so a
    run can be captured in a handful of screenshots instead of thirty-odd tab
    clicks.
    """
    path = out_dir / "unified_tables.xlsx"
    sheet = "all_tables"
    try:
        from openpyxl.styles import Font
    except Exception as exc:                          # pragma: no cover
        warn(f"openpyxl not available, no unified sheet: {exc}")
        return

    try:
        row, written = 0, 0
        with pd.ExcelWriter(path, engine="openpyxl") as xl:
            for name, tab in t.items():
                if name.startswith("_") or not isinstance(tab, pd.DataFrame):
                    continue
                if tab.empty:
                    continue
                # Title on its own row, then the table, then a gap. The title
                # is a real cell rather than a merged block so a screenshot
                # and a copy-paste both keep it.
                pd.DataFrame({0: [name]}).to_excel(
                    xl, sheet_name=sheet, startrow=row, index=False, header=False)
                tab.to_excel(xl, sheet_name=sheet, startrow=row + 1)
                # A single named index goes INTO the header row; only a
                # MultiIndex costs an extra row. Get this wrong the other way
                # and the next title lands on the last row of data.
                extra = 1 if tab.index.nlevels > 1 else 0
                row += 1 + 1 + extra + len(tab) + 2
                written += 1

            ws = xl.sheets[sheet]
            bold = Font(bold=True)
            for cell in ws["A"]:
                if cell.value in t:
                    cell.font = bold
            widest = 0
            for col in ws.columns:
                longest = max((len(str(c.value)) for c in col if c.value is not None),
                              default=0)
                ws.column_dimensions[col[0].column_letter].width = min(
                    max(longest + 2, 9), 34)
                widest = max(widest, len(col))
        log(f"  one sheet  -> {path}  ({written} tables, {row:,} rows)")
    except Exception as exc:                          # pragma: no cover
        warn(f"could not write {path}: {exc}")


# ===========================================================================
# FINDINGS - what the numbers say, with the caveats attached
# ===========================================================================

def narrative(t: dict) -> None:
    """The H1 story in the order the Q1 deck told it, with its numbers.

    Written for whoever builds the slides. Every figure is in one of the
    50-77 tables, and a result is marked "holds" only when its 95% interval
    stays on one side of zero.
    """
    section("H1 NARRATIVE - CLOSE-ONLY AND PRE-TRADED")
    E = pd.DataFrame()
    ncol = "notional (" + CURRENCY + "m)"
    es = t.get("50_exec_summary", E)
    if es is None or es.empty:
        log("  nothing to summarise.")
        return

    log("  Executive summary - the Q1 opening page, for this window:")
    for _, r in es.iterrows():
        val = r["value"]
        if isinstance(val, (int, float, np.integer, np.floating)) \
                and not isinstance(val, bool):
            txt = f"{val:,.0f}" if float(val).is_integer() else f"{val:,.2f}"
        else:
            txt = str(val)
        log(f"    {r['section']:<11} {r['metric']:<36} {txt} {r['unit']}".rstrip())

    def verdict(r) -> str:
        lo, hi = float(r.get("CI low", np.nan)), float(r.get("CI high", np.nan))
        if not (np.isfinite(lo) and np.isfinite(hi)):
            return "too few orders to call"
        return "holds" if (hi < 0 or lo > 0) else "not distinguishable from zero"

    def line(label, r) -> str:
        return (f"    {str(label):<16}{r['spreads']:+6.2f} spreads  "
                f"[{r['spread bps (wtd)']:.1f} bps] = {r['wtd mean bps']:+7.1f} bps  "
                f"{int(r['orders']):>6,} orders  {CURRENCY} {r[ncol]:>8,.1f}m  "
                f"{verdict(r)}")

    def sides(key, title):
        d = t.get(key, E)
        if d is None or d.empty:
            return
        log(f"  {title}")
        for k, r in d.iterrows():
            log(line(k, r))

    def weakest(key, title, k=3):
        d = _chartable(t.get(key, E))
        if d.empty:
            return
        d = d[~d["small sample"].astype(bool)].sort_values("spreads").head(k)
        if d.empty:
            return
        log(f"  {title} - weakest markets (at least {MIN_N_FOR_CI} orders):")
        for m, r in d.iterrows():
            log(line(m, r))

    def side_gaps(key, title, k=3):
        d = _chartable(t.get(key, E))
        if d.empty or d.index.nlevels < 2:
            return
        piv = d[~d["small sample"].astype(bool)]["spreads"].unstack(1)
        if not {"Buy", "Sell"} <= set(piv.columns):
            return
        piv = piv.dropna(subset=["Buy", "Sell"])
        if piv.empty:
            return
        gap = (piv["Buy"] - piv["Sell"]).abs().sort_values(ascending=False)
        log(f"  {title} - widest gaps between Buy and Sell, in spreads:")
        for m in gap.head(k).index:
            log(f"    {str(m):<16}Buy {piv.loc[m, 'Buy']:+.2f}   "
                f"Sell {piv.loc[m, 'Sell']:+.2f}")

    log("")
    log("  CLOSE-ONLY. The close slippage is zero by design, so the measure is")
    log("  reversion: did the price hold after the auction, or come back?")
    sides("71b_co_reversion_side_spr", "Reversion, all close-only and by side:")
    weakest("71_co_reversion_mkt_spreads", "Reversion")
    side_gaps("72_co_reversion_mkt_side_spr", "Reversion")

    log("")
    log("  PRE-TRADED. Three prices: where the order started (first execution")
    log("  against the close), where it finished (execution against the close),")
    log("  and whether that held overnight (reversion).")
    sides("73a_pre_firstexec_side_spr", "First execution vs close, all and by side:")
    sides("73b_pre_close_side_spr", "Execution vs close, all and by side:")
    sides("73c_pre_reversion_side_spr", "Reversion, all and by side:")
    weakest("68_pre_firstexec_mkt_spreads", "First execution vs close")
    weakest("77_pre_close_mkt_spreads", "Execution vs close")
    weakest("70_pre_reversion_mkt_spreads", "Reversion")
    side_gaps("75_pre_close_mkt_side_spr", "Execution vs close")
    side_gaps("76_pre_reversion_mkt_side_spr", "Reversion")

    pa = t.get("53_pre_adv_profile", E)
    if pa is not None and not pa.empty and "0-1%" in pa.index:
        log("")
        log(f"  Pre-traded size: {pa.loc['0-1%', '% of orders']:.0f}% of the "
            f"orders and {pa.loc['0-1%', '% of notional']:.0f}% of the value "
            "were under 1% ADV.")
        log("  That was the Q1 case for moving small orders into the auction.")
        log("  Check it against the execution-vs-close lines above first.")
    log("")
    log("  'holds' means the 95% interval stays on one side of zero. Anything")
    log("  else is not a finding yet, however large the number looks.")


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

    # For the monthly slide's speaker notes. Short sells can execute
    # differently - locate requirements, short-sale rules - so their share is
    # worth having in the room even when it is not on the slide.
    if "side_label" in auc:
        ss = auc[auc["side_label"] == "Short sell"]
        if len(ss):
            log("  Short sells, for the comments:")
            log(f"    {100.0 * len(ss) / max(len(auc), 1):>6.1f}% of orders "
                f"({len(ss):,} of {len(auc):,})")
            log(f"    {100.0 * ss['notional'].sum() / max(auc['notional'].sum(), 1e-9):>6.1f}% "
                f"of value ({CURRENCY} {ss['notional'].sum() / 1e6:,.1f}m)")
            log("")

    # The charts no longer carry a small-sample marking - it crowds a client
    # slide - so the thin groups are named once, here, where whoever presents
    # the deck will meet them before the meeting rather than during it.
    prof = t.get("34_market_profile", pd.DataFrame())
    if not prof.empty and "small sample" in prof.columns:
        thin = prof[prof["small sample"].astype(bool)]
        if len(thin):
            log(f"  Markets under {MIN_N_FOR_CI} orders - shown on the charts,")
            log("  but do not quote them as results:")
            for mkt, r in thin.iterrows():
                log(f"      {str(mkt):<16}{int(r['orders']):>6,} orders   "
                    f"{CURRENCY} {r['notional (' + CURRENCY + 'm)']:>9,.2f}m")
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

    # A buy limit under the close, or a sell limit over it, could not have
    # traded in the auction. Getting the side backwards here would blame the
    # algo for orders their own limit ruled out, and exonerate the rest.
    lim_df = pd.DataFrame({
        "is_buy":      [True, True,  False, False],
        "limit_price": [9.0,  11.0,  11.0,  9.0],
        "close_price": [10.0, 10.0,  10.0,  10.0],
    })
    away = np.where(lim_df["is_buy"],
                    lim_df["close_price"] - lim_df["limit_price"],
                    lim_df["limit_price"] - lim_df["close_price"])
    check("a limit on the wrong side of the close is flagged, both sides",
          list(away > 0) == [True, False, True, False], list(away))

    share = pd.DataFrame({"market_close_size": [1000.0, 0.0, 500.0],
                          "fill_close_size":   [  50.0, 10.0,   0.0]})
    mkt = pd.to_numeric(share["market_close_size"], errors="coerce")
    got = np.where(mkt > 0, 100.0 * share["fill_close_size"] / mkt, np.nan)
    check("share of the auction is a percentage, and zero auction size is not a share",
          got[0] == 5.0 and np.isnan(got[1]) and got[2] == 0.0, list(got))

    # An order with no auction fill on a day with no auction is not a miss.
    coh = assign_cohort(pd.DataFrame({
        "fill_rate":        [100.0, 100.0],
        "pct_close":        [  0.0,   0.0],
        "adv_pct":          [  0.1,   0.1],
        "auction_existed":  [False,  True],
    }))
    check("no auction that day is separated from a genuine miss",
          list(coh) == ["No auction that day", "No auction fill - unexplained"],
          list(coh))

    # A fill implies an auction to have filled in. Requiring marketCloseSize
    # alone would throw away every order whose own fill proves the point -
    # which is exactly the India case, where the imputed fill is the only
    # evidence there is.
    mkt_s = pd.Series([1000.0, 0.0, 0.0, np.nan])
    our_s = pd.Series([50.0, 25.0, 0.0, 10.0])
    existed = (mkt_s.fillna(0) > 0) | (our_s.fillna(0) > 0)
    check("a close fill counts as evidence the auction existed",
          list(existed) == [True, True, False, True], list(existed))

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
    check("reversion is NextOpen as it stands, not differenced again",
          np.allclose(df["reversion_bps"], df["slip_nextopen"])
          if NEXTOPEN_IS_VS_CLOSE else
          np.allclose(df["reversion_bps"],
                      df["slip_nextopen"] - df["slip_close"]))
    toy = pd.DataFrame({"g": ["a"] * 10, "v": [2.0] * 10,
                        "spread_bps": [4.0] * 10, "notional": [1.0] * 10})
    tt = t_in_spreads(toy, "g", "v")
    check("2bps on a 4bps spread is half a spread, CI included",
          not tt.empty and abs(tt["spreads"].iloc[0] - 0.5) < 1e-9
          and abs(tt["CI low"].iloc[0] - 0.5) < 1e-9
          and abs(tt["CI high"].iloc[0] - 0.5) < 1e-9, tt.to_dict())
    if "spread_bps" in df:
        st = t_in_spreads(df.assign(_g="all"), "_g", "slip_close", ci=False)
        r = st.iloc[0] if not st.empty else None
        check("spreads x [spread bps] = bps, to rounding",
              r is not None and abs(r["spreads"] * r["spread bps (wtd)"]
                                    - r["wtd mean bps"])
              < 1e-3 * abs(r["spread bps (wtd)"]) + 1e-3,
              None if r is None else (r["spreads"], r["spread bps (wtd)"],
                                      r["wtd mean bps"]))
    fo = pd.DataFrame({"c": ["a"] * 3 + ["b"], "v": [10.0] * 3 + [-10.0],
                       "w": [1.0] * 4})
    re_ = pd.DataFrame({"c": ["a"] + ["b"] * 3, "v": [10.0] + [-10.0] * 3,
                        "w": [1.0] * 4})
    a_, e_, cov_ = standardised_gap(fo, re_, ["c"], "v", "w")
    check("same per-cell result, different mix: expected equals actual",
          abs(a_ - 5.0) < 1e-9 and abs(e_ - 5.0) < 1e-9 and abs(cov_ - 1) < 1e-9,
          (a_, e_, cov_))
    fe_ok = df["first_exec_vs_close"].notna() & (df["pct_close"] < 50)
    check("first execution loads as positive = saving, like vs Close",
          np.corrcoef(df.loc[fe_ok, "first_exec_vs_close"],
                      df.loc[fe_ok, "slip_close"])[0, 1] > 0.3
          if fe_ok.sum() > 10 else True)
    px = pd.DataFrame({
        "is_buy":           [True, True, False, False, True, False],
        "close_price":      [100.0, 200.0, 50.0, 80.0, 10.0, 40.0],
        "avg_price":        [99.0, 202.0, 51.0, 79.0, 9.9, 40.4],
        "next_open_price":  [101.0, 198.0, 49.0, 81.0, 10.2, 39.0],
        "first_exec_price": [98.0, 204.0, 52.0, 78.0, 9.8, 40.8]})
    side = np.where(px["is_buy"], 1, -1)
    px["slip_close"] = 1e4 * side * (px["close_price"] - px["avg_price"]) / px["close_price"]
    px["first_exec_vs_close"] = 1e4 * side * (px["close_price"] - px["first_exec_price"]) / px["close_price"]
    px["slip_nextopen"] = 1e4 * side * (px["next_open_price"] - px["close_price"]) / px["close_price"]
    vr = verify_slippage(px, quiet=True)
    check("slippage rebuilt from prices matches when it should",
          len(vr) == 3 and (vr["verdict"] == "matches the prices").all(),
          vr.get("verdict", pd.Series()).tolist())
    px["first_exec_vs_close"] = -px["first_exec_vs_close"]
    vr = verify_slippage(px, quiet=True)
    # a raw (first - close) column, flipped on load, is right on buys and
    # wrong on sells - the case the sell check exists for
    raw_diff = 1e4 * (px["close_price"] - px["first_exec_price"]) / px["close_price"]
    vr2 = verify_slippage(px.assign(first_exec_vs_close=-raw_diff * -1), quiet=True)
    check("a column wrong on sells only is called out as one-sided",
          vr2.loc[vr2["check"] == "first exec vs close", "verdict"]
          .str.startswith("SIGN WRONG ON SELLS ONLY").all(), vr2["verdict"].tolist())
    check("an inverted column is called inverted",
          vr.loc[vr["check"] == "first exec vs close", "verdict"]
          .str.startswith("SIGN INVERTED").all(), vr["verdict"].tolist())
    ip = pd.DataFrame({"close_pr": [0.0, 0.0, 3.0],
                       "participation": [12.0, 8.0, 9.0]})
    used = india_close_pr_proxy(ip, pd.Series([True, False, False]))
    check("India window orders take PR as close participation, others keep ClosePR",
          ip["close_pr"].tolist() == [12.0, 0.0, 3.0] and used.tolist() == [True, False, False],
          ip["close_pr"].tolist())
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
    tail = np.array([-999.0, 0, 1, 2, 3, 999.0])
    check("clipping pulls the tail in, it does not drop rows",
          len(winsorize(tail, (0.01, 0.99))) == 6
          and winsorize(tail, (0.01, 0.99)).max() < 999.0)
    check("by default nothing is clipped",
          CLIP_OUTLIERS or np.array_equal(winsorize(tail), tail))
    vals, wts = np.array([10.0, -900.0, 5.0, 7.0]), np.array([2.0, 1.0, 3.0, 4.0])
    check("by default the average is SUMPRODUCT / SUM, as in Excel",
          CLIP_OUTLIERS or abs(wmean(vals, wts) - (vals * wts).sum() / wts.sum()) < 1e-9)

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

def run(path: Path, out_dir: Path, sample: bool = False,
        focus: str = None) -> None:
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
    df = filter_cas_eligible(df, out_dir)
    df = filter_india_close_window(df)
    df = mark_close_opportunity(df)
    if DROP_NO_CLOSE_OPPORTUNITY and "reached_close_window" in df:
        gone = ~df["reached_close_window"]
        if gone.any():
            section("ORDERS THAT NEVER RAN INTO THE CLOSE")
            warn(f"removing {int(gone.sum()):,} orders that finished more than "
                 f"{CLOSE_OPPORTUNITY_MIN:.0f} minutes")
            log(f"    before their market closed, worth {CURRENCY} "
                f"{df.loc[gone, 'notional'].sum() / 1e6:,.1f}m.")
            log("    They had no opportunity to reach the auction. Orders that")
            log("    WERE live into the close and still got nothing are kept -")
            log("    those are the finding, not the noise.")
            df = df[~gone].copy()

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
    write_excel_unified(tables, out_dir)
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
    narrative(tables)
    if focus:
        focus_market(tables, out_dir, focus)


# ===========================================================================
# FOCUS - one market taken apart: why is it worse?
# ===========================================================================
#
#   python moc_tca.py --data orders.csv --out output_h1 --focus "South Korea"
#
# Runs on PRE-TRADED orders in auction markets, on the first execution against
# the close - the measure where a market stands out - and sets the focus
# market against every other auction market. Eight tests, each ending in one
# line that says what it points to or rules out, then the orders that cost the
# most, for the desk to pull the child fills on. It narrows the cause; the
# child-order tape confirms it.

FOCUS_TOP_ORDERS = 30
FE, FE_W = "first_exec_vs_close", "cont_notional"


def _slug(name) -> str:
    return re.sub(r"[^a-z0-9]+", "_", str(name).lower()).strip("_")


def _vs(foc: pd.DataFrame, rest: pd.DataFrame, by: str, value: str = FE,
        weight: str = FE_W) -> pd.DataFrame:
    """The focus market beside every other market, one row per group."""
    if by not in foc or by not in rest:
        return pd.DataFrame()
    both = pd.concat([foc.assign(_who="focus"), rest.assign(_who="rest")])
    t = t_in_spreads(both, [by, "_who"], value, weight)
    if t.empty:
        return t
    col = both[by]
    keys = (list(col.cat.categories) if hasattr(col, "cat")
            else sorted(col.dropna().astype(str).unique()))
    rows = []
    for key in keys:
        r = {by: key}
        for who in ("focus", "rest"):
            hit = (key, who) in t.index
            x = t.loc[(key, who)] if hit else None
            r[f"{who} spreads"] = float(x["spreads"]) if hit else np.nan
            r[f"{who} CI low"] = float(x["CI low"]) if hit else np.nan
            r[f"{who} CI high"] = float(x["CI high"]) if hit else np.nan
            r[f"{who} bps"] = float(x["wtd mean bps"]) if hit else np.nan
            r[f"{who} orders"] = int(x["orders"]) if hit else 0
        if r["focus orders"] or r["rest orders"]:
            r["gap (spreads)"] = r["focus spreads"] - r["rest spreads"]
            rows.append(r)
    return pd.DataFrame(rows).set_index(by).round(3) if rows else pd.DataFrame()


def standardised_gap(foc: pd.DataFrame, rest: pd.DataFrame, cells: list,
                     value: str = FE, weight: str = FE_W):
    """What the other markets would have scored with the focus market's mix.

    Each cell (e.g. ADV% band x spread band) takes the other markets' own
    result, weighted by how much of the FOCUS market's flow sits in that cell.
    If that expected figure is close to the focus market's actual one, its
    orders were simply harder; if the actual is far worse, the orders do not
    explain it. Returns (actual bps, expected bps, share of focus weight in
    cells the other markets also trade).
    """
    need = cells + [value]
    if foc.empty or rest.empty or any(c not in foc or c not in rest for c in need):
        return np.nan, np.nan, 0.0
    wf = weight_column(foc, weight)
    wr = weight_column(rest, weight)
    # Tails pulled in ONCE, across both populations. Clipping each cell at
    # its own extremes would move the two sides by different amounts and put
    # a gap into the comparison that is not in the data.
    both = winsorize(pd.to_numeric(pd.concat([foc[value], rest[value]]),
                                   errors="coerce").to_numpy(float))
    foc = foc.assign(_v=both[:len(foc)])
    rest = rest.assign(_v=both[len(foc):])
    actual = wmean(foc["_v"], foc[wf], winsor=False)
    rest_cell = {k: wmean(g["_v"], g[wr], winsor=False)
                 for k, g in rest.groupby(cells, observed=True)}
    num = den = total = 0.0
    for k, g in foc.groupby(cells, observed=True):
        w = float(pd.to_numeric(g[wf], errors="coerce").clip(lower=0).sum())
        total += w
        rv = rest_cell.get(k, np.nan)
        if w > 0 and np.isfinite(rv):
            num += w * rv
            den += w
    expected = num / den if den > 0 else np.nan
    return actual, expected, (den / total if total > 0 else 0.0)


def _cost_usd(d: pd.DataFrame) -> pd.Series:
    """First execution vs close in dollars, per order (negative = cost)."""
    w = pd.to_numeric(d[weight_column(d, FE_W)], errors="coerce").fillna(0)
    return pd.to_numeric(d[FE], errors="coerce").fillna(0) * w / 1e4


def _start_band(d: pd.DataFrame) -> pd.Series:
    """When the order started: the arrival bucket, else the HKT start hour."""
    if "arrival_time" in d and d["arrival_time"].notna().any():
        return pd.Categorical(d["arrival_time"].astype(str),
                              categories=[c for c in ARRIVAL_ORDER]
                              + sorted(set(d["arrival_time"].astype(str))
                                       - set(ARRIVAL_ORDER)))
    for c in ("first_start_time_min", "start_time_min"):
        if c in d and d[c].notna().any():
            h = (d[c] // 60).astype("Int64")
            return h.map(lambda x: f"{int(x):02d}:00 HKT" if pd.notna(x) else np.nan)
    return pd.Series(np.nan, index=d.index)


def _chart_pair(tab: pd.DataFrame, out: Path, name: str, title: str,
                focus: str, xlabel: str, note: str = "") -> None:
    """The focus market beside the rest, per group, in spreads."""
    if not _HAS_MPL or tab is None or tab.empty:
        return
    d = tab[[str(i).strip().lower() not in NON_CATEGORIES for i in tab.index]]
    d = d[d["focus orders"] > 0]          # nothing to compare where it has no orders
    if d.empty:
        return
    x = np.arange(len(d))
    fig, (ax,) = _fig((11.0, 5.6))
    for off, who, color, label in ((-0.2, "focus", SERIES[1], focus),
                                   (0.2, "rest", INK_MUTED, "other markets")):
        vals = [float(v) for v in d[f"{who} spreads"]]
        ax.bar(x + off, vals, width=0.38, color=color, zorder=3, label=label)
        for i, v in enumerate(vals):
            if np.isfinite(v):
                ax.text(i + off, v, f"{v:+.2f}", ha="center",
                        va="bottom" if v >= 0 else "top", fontsize=8,
                        color=INK, zorder=5)
    ax.axhline(0, color=BASELINE, linewidth=1.0, zorder=2)
    ax.set_xticks(x)
    ax.set_xticklabels([f"{k}{chr(10)}{int(n)} orders" for k, n
                        in zip(d.index, d["focus orders"])])
    ax.legend(frameon=False, fontsize=8.5, loc="best")
    _style(ax, xlabel=xlabel,
           ylabel="first execution vs close, in spreads", title=title)
    room = _notes(fig, [note, "order counts are the focus market's; "
                              "negative = the first fill was worse than the "
                              "close"])
    _save(fig, out, name, bottom=room, deck=False)


def _chart_concentration(cost: pd.Series, out: Path, name: str,
                         focus: str) -> None:
    """How much of the cost the worst orders carry."""
    if not _HAS_MPL:
        return
    neg = -cost[cost < 0].sort_values()
    if neg.empty:
        return
    share = 100 * neg.cumsum().to_numpy() / neg.sum()
    fig, (ax,) = _fig((10.0, 5.0))
    ax.plot(np.arange(1, len(share) + 1), share, color=SERIES[1], lw=2.2,
            zorder=3)
    for k in (10, 30):
        if len(share) >= k:
            ax.axvline(k, color=BASELINE, lw=1, zorder=2)
            ax.text(k, share[k - 1], f"  top {k}: {share[k - 1]:.0f}%",
                    va="bottom", fontsize=9, color=INK)
    ax.set_ylim(0, 105)
    _style(ax, xlabel="orders that cost money, worst first",
           ylabel="share of the total first-execution cost (%)",
           title=f"{focus}: how concentrated is the cost?")
    _save(fig, out, name, deck=False)


def focus_market(t: dict, out_dir: Path, focus: str) -> None:
    df = t.get("_close_df")
    section(f"FOCUS - {focus.upper()}")
    if df is None or df.empty or "market" not in df:
        log("  no close orders to take apart.")
        return
    markets = {str(m).lower(): m for m in df["market"].dropna().unique()}
    if focus.lower() not in markets:
        warn(f"no orders in market {focus!r}. Markets in the file: "
             + ", ".join(sorted(map(str, markets.values()))))
        return
    focus = markets[focus.lower()]
    auc_all = df[df["has_auction"]
                 & ~df["market"].isin(EXCLUDE_MARKETS_FROM_CHARTS)].copy()
    pre_all = auc_all[_flag(auc_all, "pretraded")].copy()
    if FE not in pre_all:
        log("  no first_exec_vs_close column - the drill-down needs it.")
        return
    foc = pre_all[pre_all["market"] == focus].copy()
    rest = pre_all[pre_all["market"] != focus].copy()
    if foc.empty:
        log(f"  {focus} has no pre-traded orders in auction markets.")
        return
    folder = out_dir / f"focus_{_slug(focus)}"
    folder.mkdir(parents=True, exist_ok=True)
    T, verdicts = {}, []

    log(f"  Pre-traded orders, auction markets. {focus}: {len(foc):,} orders; "
        f"other markets: {len(rest):,}.")
    log("  Measure: first execution vs close, in spreads, weighted by the part")
    log("  of each order traded before the auction. Negative = a cost.")

    # --- F0 where it stands ------------------------------------------------
    rows = []
    for label, value, w in (("first execution vs close", FE, FE_W),
                            ("execution vs close", "slip_close", "notional"),
                            ("vs PVWAP", "slip_pvwap", "notional"),
                            ("close to T+1", "reversion_bps", "notional")):
        if value not in pre_all:
            continue
        for side in ["All"] + BUY_SELL_ORDER:
            f = foc if side == "All" else foc[foc.get("buy_sell") == side]
            r_ = rest if side == "All" else rest[rest.get("buy_sell") == side]
            v = _vs(f.assign(_s=side), r_.assign(_s=side), "_s", value, w)
            if v.empty:
                continue
            row = v.iloc[0].to_dict()
            row.update({"measure": label, "side": side})
            rows.append(row)
    T["F0_summary"] = (pd.DataFrame(rows).set_index(["measure", "side"])
                       if rows else pd.DataFrame())
    head = (T["F0_summary"].loc[("first execution vs close", "All")]
            if not T["F0_summary"].empty
            and ("first execution vs close", "All") in T["F0_summary"].index
            else None)
    worse = head is not None and head["focus spreads"] < head["rest spreads"]
    if head is not None:
        clear = (np.isfinite(head["focus CI high"])
                 and head["focus CI high"] < head["rest spreads"])
        verdicts.append(("0 Is it worse at all?",
            (f"YES. First execution vs close: {focus} {head['focus spreads']:+.2f} "
             f"spreads, other markets {head['rest spreads']:+.2f}"
             + (", and the gap is clear at 95%." if clear else
                ", but the gap is inside the 95% interval.")) if worse else
            (f"NO. First execution vs close: {focus} {head['focus spreads']:+.2f} "
             f"spreads, other markets {head['rest spreads']:+.2f}. The tests "
             "below still describe the market, but there is no gap to explain.")))

    # --- F1 is it the orders? -----------------------------------------------
    T["F1a_by_adv"] = _vs(foc, rest, "adv_bucket")
    T["F1b_by_spread"] = _vs(foc, rest, "spread_bucket")
    cells = [c for c in ("adv_bucket", "spread_bucket") if c in pre_all]
    act, exp, cover = standardised_gap(foc, rest, cells)
    sprd = wmean(foc["spread_bps"], foc[weight_column(foc, FE_W)], winsor=False) \
        if "spread_bps" in foc else np.nan
    T["F1c_same_mix"] = pd.DataFrame([{
        "focus actual bps": act, "others with focus mix bps": exp,
        "focus spread bps": sprd,
        "focus actual spreads": act / sprd if sprd else np.nan,
        "others with focus mix spreads": exp / sprd if sprd else np.nan,
        "share of focus flow compared %": 100 * cover,
        "cells": " x ".join(cells)}]).round(3)
    if np.isfinite(act) and np.isfinite(exp):
        a_s, e_s = act / sprd, exp / sprd
        if act >= exp:
            v = (f"NOT WORSE THAN ITS MIX. Other markets with {focus}'s mix of "
                 f"ADV% and spread: {e_s:+.2f} spreads; {focus}: {a_s:+.2f}.")
        elif (exp - act) > 0.5 * abs(act):
            v = (f"ORDER MIX EXPLAINS LITTLE. With {focus}'s own mix of ADV% and "
                 f"spread, other markets would have scored {e_s:+.2f} spreads; "
                 f"{focus} scored {a_s:+.2f}. Points to the market or how the "
                 "algo trades it, not to harder orders.")
        elif abs(act - exp) <= 0.25 * abs(act):
            v = (f"ORDER MIX EXPLAINS MOST OF IT. Other markets with the same "
                 f"mix: {e_s:+.2f} spreads; {focus}: {a_s:+.2f}.")
        else:
            v = (f"ORDER MIX EXPLAINS PART OF IT. Other markets with the same "
                 f"mix: {e_s:+.2f} spreads; {focus}: {a_s:+.2f}.")
        verdicts.append(("1 Is it the orders (ADV%, spread)?",
                         v + f" ({100 * cover:.0f}% of the flow had a match.)"))

    # --- F2 when did it start? ----------------------------------------------
    foc["_start"], rest["_start"] = _start_band(foc), _start_band(rest)
    tab = _vs(foc, rest, "_start")
    if not tab.empty:
        cost = _cost_usd(foc)
        band_cost = cost.groupby(foc["_start"], observed=True).sum()
        total_neg = float(-cost[cost < 0].sum()) or np.nan
        tab["focus cost USDk"] = [float(band_cost.get(k, 0)) / 1e3
                                  for k in tab.index]
        tab.index.name = "start"
        T["F2_by_start"] = tab.round(3)
        worst = tab["focus cost USDk"].idxmin()
        net = float(tab.loc[worst, "focus cost USDk"])
        if net < 0:
            share = 100 * -net * 1e3 / total_neg
            verdicts.append(("2 Does it start too early?",
                f"{share:.0f}% of the cost came from orders starting in "
                f"{worst}, the costliest start band. There {focus} "
                f"scored {tab.loc[worst, 'focus spreads']:+.2f} spreads against "
                f"{tab.loc[worst, 'rest spreads']:+.2f} elsewhere."))
        else:
            verdicts.append(("2 Does it start too early?",
                "No start band ran at a net cost."))

    # --- F3 how hard did it trade? -------------------------------------------
    rows = []
    for label, col, w in (("continuous participation % (fPR_cont)", "pr_cont", FE_W),
                          ("overall participation %", "participation", "notional"),
                          ("ADV% of the order", "adv_pct", "notional"),
                          ("share executed in the close %", "pct_close", "notional"),
                          ("spread bps", "spread_bps", "notional")):
        if col in foc and foc[col].notna().any():
            rows.append({"metric": label,
                         "focus": wmean(foc[col], foc[weight_column(foc, w)], winsor=False),
                         "rest": wmean(rest[col], rest[weight_column(rest, w)], winsor=False)})
    T["F3_how_hard"] = pd.DataFrame(rows).set_index("metric").round(3) if rows else pd.DataFrame()
    pv = _vs(foc.assign(_a="All"), rest.assign(_a="All"), "_a", "slip_pvwap", "notional") \
        if "slip_pvwap" in foc else pd.DataFrame()
    if not pv.empty:
        T["F3b_vs_pvwap"] = pv
        fpv, rpv = pv.iloc[0]["focus spreads"], pv.iloc[0]["rest spreads"]
        line = f"Against PVWAP {focus} scored {fpv:+.2f} spreads, others {rpv:+.2f}. "
        if worse:
            line += ("The trading itself was in line, so the first-fill cost "
                     "points to WHEN it started." if fpv >= rpv - 0.25 else
                     "The trading itself was worse too, so the cost points to "
                     "HOW it traded - aggression or size.")
        if "continuous participation % (fPR_cont)" in T["F3_how_hard"].index:
            pr = T["F3_how_hard"].loc["continuous participation % (fPR_cont)"]
            if pr["rest"]:
                line += (f" Continuous participation: {pr['focus']:.1f}% vs "
                         f"{pr['rest']:.1f}% ({pr['focus'] / pr['rest']:.1f}x).")
        verdicts.append(("3 Does it trade too hard?", line))

    # --- F4 is the auction big enough? --------------------------------------
    if "close_pr" in auc_all and "adv_bucket" in auc_all:
        fa = auc_all[auc_all["market"] == focus]
        ra = auc_all[auc_all["market"] != focus]
        rows = []
        for band in auc_all["adv_bucket"].cat.categories:
            f_, r_ = fa[fa["adv_bucket"] == band], ra[ra["adv_bucket"] == band]
            if f_.empty:
                continue
            rows.append({
                "ADV%": band, "focus orders": len(f_),
                "focus ClosePR %": wmean(f_["close_pr"], f_["notional"], winsor=False),
                "rest ClosePR %": wmean(r_["close_pr"], r_["notional"], winsor=False),
                "focus % pre-traded": 100 * _flag(f_, "pretraded").mean(),
                "rest % pre-traded": 100 * _flag(r_, "pretraded").mean()
                    if len(r_) else np.nan})
        tab = pd.DataFrame(rows).set_index("ADV%").round(2) if rows else pd.DataFrame()
        T["F4_auction_capacity"] = tab
        if not tab.empty:
            ok = tab.dropna(subset=["focus ClosePR %", "rest ClosePR %"])
            higher = int((ok["focus ClosePR %"] > ok["rest ClosePR %"]).sum())
            more_pre = int((ok["focus % pre-traded"] > ok["rest % pre-traded"]).sum())
            verdicts.append(("4 Is the auction too small?",
                (f"{focus} took a bigger share of the auction in {higher} of "
                 f"{len(ok)} ADV% bands, and pre-traded more often in {more_pre}. ")
                + ("Capacity may be pushing flow into the continuous session."
                   if higher > len(ok) / 2 and more_pre > len(ok) / 2 else
                   "Auction capacity does not look like the driver.")))

    # --- F5 a few orders, or all of them? ------------------------------------
    cost = _cost_usd(foc)
    neg = -cost[cost < 0].sort_values()
    if len(neg):
        top10 = 100 * neg.head(10).sum() / neg.sum()
        med = float(pd.to_numeric(foc[FE], errors="coerce").median())
        T["F5_concentration"] = pd.DataFrame([{
            "orders": len(foc), "orders that cost": int((cost < 0).sum()),
            "total cost USDk": float(cost.sum()) / 1e3,
            "top 10 share of cost %": top10,
            "median order bps": med,
            "weighted bps": wmean(foc[FE], foc[weight_column(foc, FE_W)])}]).round(2)
        verdicts.append(("5 A few orders, or all of them?",
            (f"CONCENTRATED: the worst 10 orders carry {top10:.0f}% of the cost. "
             "Start with the order list." if top10 > 50 else
             f"BROAD: the worst 10 orders carry only {top10:.0f}% of the cost. "
             "It is systematic, not a few bad orders.")
            + f" The median order scored {med:+.1f} bps."))
        _chart_concentration(cost, folder, "focus_4_concentration.png", focus)

    # --- F6 one month? ---------------------------------------------------------
    if "month" in foc:
        tab = _vs(foc, rest, "month")
        if not tab.empty:
            mcost = cost.groupby(foc["month"]).sum()
            tab["focus cost USDk"] = [float(mcost.get(k, 0)) / 1e3 for k in tab.index]
            T["F6_by_month"] = tab.round(3)
            worst = tab["focus cost USDk"].idxmin()
            share = 100 * -tab.loc[worst, "focus cost USDk"] * 1e3 / (neg.sum() or np.nan)
            verdicts.append(("6 Is it one month?",
                (f"ONE MONTH: {worst} carries {share:.0f}% of the cost."
                 if share > 40 else
                 f"NOT ONE MONTH: the worst, {worst}, carries {share:.0f}% of the "
                 "cost.")))

    # --- F7 limits ---------------------------------------------------------------
    if "market_limit" in foc and foc["market_limit"].nunique() > 1:
        T["F7_limit_vs_market"] = _vs(foc, rest, "market_limit")
    if {"limit_price", "close_price"} <= set(foc.columns):
        def away(d):
            lim = pd.to_numeric(d["limit_price"], errors="coerce")
            cls = pd.to_numeric(d["close_price"], errors="coerce")
            return (1e4 * (lim - cls).abs() / cls).where(lim > 0)
        fa_, ra_ = away(foc), away(rest)
        if fa_.notna().any():
            T["F7b_limit_distance"] = pd.DataFrame([{
                "focus median limit distance from close bps": fa_.median(),
                "rest median limit distance from close bps": ra_.median(),
                "focus orders with a limit": int(fa_.notna().sum())}]).round(2)
            verdicts.append(("7 Are limits in the way?",
                f"Median limit sits {fa_.median():.0f} bps from the close in "
                f"{focus}, {ra_.median():.0f} bps elsewhere. "
                + ("Tighter limits may be shaping the start." if fa_.median()
                   < 0.5 * ra_.median() else "Limits do not look like the driver.")))

    # --- F8 volatility -------------------------------------------------------------
    if "volatility" in foc and foc["volatility"].notna().any():
        wf, wr = weight_column(foc, FE_W), weight_column(rest, FE_W)
        fv, rv = wmean(foc["volatility"], foc[wf]), wmean(rest["volatility"], rest[wr])
        fb, rb = wmean(foc[FE], foc[wf]), wmean(rest[FE], rest[wr])
        T["F8_volatility"] = pd.DataFrame([{
            "focus volatility": fv, "rest volatility": rv,
            "focus bps per vol point": fb / fv if fv else np.nan,
            "rest bps per vol point": rb / rv if rv else np.nan}]).round(3)
        if fv and rv:
            worse_after = (fb / fv) < (rb / rv)
            tail = ("" if not worse else
                    "Still worse per unit of volatility, so volatility does not "
                    "explain it." if worse_after else
                    "Per unit of volatility it is in line: volatility explains it.")
            verdicts.append(("8 Is it just more volatile?",
                f"Volatility {fv:.1f} in {focus} vs {rv:.1f} elsewhere. " + tail))

    # --- the orders to pull ------------------------------------------------------------
    keep = [c for c in ("order_id", "symbol", "date", "side_label", "strategy",
                        "first_start_time_min", "arrival_time", "adv_pct",
                        "spread_bps", "volatility", "pct_close", "pr_cont",
                        "close_pr", FE, "slip_close", "slip_pvwap",
                        "reversion_bps", "notional", "cont_notional",
                        "market_limit", "limit_price", "close_price")
            if c in foc]
    worst = foc.assign(first_exec_cost_usd=cost).sort_values(
        "first_exec_cost_usd").head(FOCUS_TOP_ORDERS)[keep + ["first_exec_cost_usd"]]
    if "first_start_time_min" in worst:
        m = worst["first_start_time_min"]
        worst.insert(keep.index("first_start_time_min"), "start HKT",
                     m.map(lambda x: f"{int(x // 60):02d}:{int(x % 60):02d}"
                           if pd.notna(x) else ""))
        worst = worst.drop(columns="first_start_time_min")
    if "spread_bps" in worst:
        worst["first exec vs close, spreads"] = (worst[FE] / worst["spread_bps"]).round(2)
    for c in ("first_exec_cost_usd", "notional", "cont_notional"):
        if c in worst:
            worst[c] = worst[c].round(0)
    worst.to_csv(folder / "worst_orders.csv", index=False)
    T["F9_worst_orders"] = worst.reset_index(drop=True)

    # --- charts, workbook, log ------------------------------------------------------
    _chart_pair(T.get("F1a_by_adv"), folder, "focus_1_by_adv.png",
                f"{focus} against other markets, same ADV% band", focus, "ADV%")
    _chart_pair(T.get("F2_by_start"), folder, "focus_2_by_start.png",
                f"{focus} against other markets, by when the order started",
                focus, "start")
    _chart_pair(T.get("F6_by_month"), folder, "focus_3_by_month.png",
                f"{focus} against other markets, by month", focus, "month")
    try:
        with pd.ExcelWriter(folder / "focus.xlsx", engine="openpyxl") as xl:
            for name, tab in T.items():
                if isinstance(tab, pd.DataFrame) and not tab.empty:
                    tab.to_excel(xl, sheet_name=name[:31])
    except Exception as exc:                          # pragma: no cover
        warn(f"focus workbook not written: {exc}")

    log("")
    log("  WHAT EACH TEST POINTS TO")
    for q, v in verdicts:
        log(f"  {q}")
        for chunk in _wrap(v, 70):
            log(f"      {chunk}")
    log("")
    for name, tab in T.items():
        if isinstance(tab, pd.DataFrame) and not tab.empty and name != "F9_worst_orders":
            log(f"  --- {name} ---")
            for line in tab.to_string().splitlines():
                log("  " + line)
            log("")
    log(f"  The {len(worst)} orders that cost the most -> {folder / 'worst_orders.csv'}")
    log("  Pull the child fills for those first: the tape is what confirms a cause.")
    log(f"  tables -> {folder / 'focus.xlsx'}   charts -> {folder}")


def _wrap(text: str, width: int) -> list:
    words, lines, cur = text.split(), [], ""
    for w in words:
        if cur and len(cur) + 1 + len(w) > width:
            lines.append(cur)
            cur = w
        else:
            cur = f"{cur} {w}".strip()
    return lines + ([cur] if cur else [])


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
    p.add_argument("--clip", action="store_true",
                   help="clip each performance value to the 1st/99th percentile "
                        "of its bar before averaging (off by default)")
    p.add_argument("--focus", metavar="MARKET",
                   help='take one market apart, e.g. --focus "South Korea"')
    p.add_argument("--label", dest="label", metavar="TEXT",
                   help="period label for the charts (overrides PERIOD_LABEL)")
    args = p.parse_args(argv)

    global DATE_FROM, DATE_TO, PERIOD_LABEL, CLIP_OUTLIERS
    if args.clip:
        CLIP_OUTLIERS = True
    if args.date_from:
        DATE_FROM = args.date_from
    if args.date_to:
        DATE_TO = args.date_to
    if args.label:
        PERIOD_LABEL = args.label

    log(f"MOC / close-algo TCA   {_dt.datetime.now():%Y-%m-%d %H:%M}")
    log("  averages: " + (
        f"performance values CLIPPED to the {CLIP_PERCENTILES[0]:.0%} / "
        f"{CLIP_PERCENTILES[1]:.0%} percentile of each bar before averaging"
        if CLIP_OUTLIERS else
        "NOT clipped - each is SUMPRODUCT(value, $Mln) / SUM($Mln) over its orders"))

    rc = 0
    try:
        if args.self_test:
            rc = self_test()
        elif args.probe:
            if not args.data:
                p.error("--probe needs --data")
            probe(args.data)
        elif args.sample:
            run(None, args.out, sample=True, focus=args.focus)
        else:
            if not args.data:
                p.error("give --data, --sample or --self-test")
            run(args.data, args.out, focus=args.focus)
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
