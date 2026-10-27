#!/usr/bin/env python3
"""Build the YTD 2026 MOC TCA deck from a moc_tca.py run, in the H1 deck's style.

    python build_ytd_deck.py                          # reads output/, writes output/MOC_TCA_YTD_2026.pptx
    python build_ytd_deck.py --results output --template H1_deck.pptx

Reads <results>/tables.xlsx and the charts the run wrote - nothing else. Every
number on a slide is looked up in a table at build time, and every sentence
that takes a side (cost or saving, clear or not) is chosen from the sign and
the 95% interval of that number. Rerun moc_tca.py and the deck follows.

The look is the H1 deck's: 4:3, white, a bold title top left, two or three
large bullets, one wide chart (or two side by side), a centred
"Data Set : ... Filter : ..." line. No footer and no firm name: with
--template the slides go onto that file's own master instead, so the result
can be pasted straight into a template.

Writes <deck>.pptx and <deck>.evidence.md: each bullet with the table row it
rests on, so any number can be defended on the call. The wording is a first
draft to polish; the numbers are not.

Sign convention everywhere: positive = saving, negative = cost.
"""
from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path

import pandas as pd

try:
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Emu, Inches, Pt
except ImportError:
    print("ERROR: python-pptx is not installed.\n       pip install python-pptx",
          file=sys.stderr)
    raise SystemExit(2)

# ===========================================================================
# CONFIG
# ===========================================================================

PERIOD = "YTD 2026"
DATA_SET = "YTD 2026 (1 Jan - 23 Oct)"

FILTER_CLOSE_ONLY = "CloseFillRate > 99.5"
FILTER_PRE_TRADED = "CloseFillRate > 0 and < 99.5"
FILTER_AUCTION = "markets with a closing auction"

# A t-stat this size or larger is called clear. 2 is about 95%.
T_CLEAR = 2.0

# ===========================================================================
# LAYOUT - the H1 deck, measured off its photos
# ===========================================================================

SLIDE_W, SLIDE_H = Inches(10), Inches(7.5)          # 4:3
LEFT = Inches(0.6)
WIDTH = SLIDE_W - 2 * LEFT
TITLE_TOP = Inches(0.45)
BULLETS_TOP = Inches(1.15)
DATASET_TOP = Inches(6.5)
BOTTOM = Inches(7.0)
FONT = "Arial"


def _rgb(h: str) -> RGBColor:
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


BLACK, GREY, DARK = _rgb("#000000"), _rgb("#555555"), _rgb("#1a1a1a")


def textbox(slide, left, top, width, height, text, size, *, bold=False,
            color=BLACK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(left, top, width, height)
    frame = box.text_frame
    frame.word_wrap = True
    frame.vertical_anchor = anchor
    for i, line in enumerate(text.split("\n")):
        para = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        para.alignment = align
        run = para.add_run()
        run.text = line
        run.font.size, run.font.bold = Pt(size), bold
        run.font.color.rgb, run.font.name = color, FONT
    return box


def bullets(slide, items, size):
    """Round bullets, the text hanging off them, a gap between items - the H1
    deck's bullet block."""
    box = slide.shapes.add_textbox(LEFT, BULLETS_TOP, WIDTH, Inches(0.5))
    frame = box.text_frame
    frame.word_wrap = True
    from pptx.oxml.ns import qn
    for i, text in enumerate(items):
        para = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        para.space_after = Pt(9)
        # A real bullet with a hanging indent, so a wrapped line sits under
        # the text and not under the bullet.
        pPr = para._p.get_or_add_pPr()
        pPr.set("marL", str(Inches(0.32)))
        pPr.set("indent", str(-Inches(0.32)))
        for tag in ("a:buNone", "a:buChar", "a:buFont"):
            for el in pPr.findall(qn(tag)):
                pPr.remove(el)
        bu_font = pPr.makeelement(qn("a:buFont"), {"typeface": "Arial"})
        bu_char = pPr.makeelement(qn("a:buChar"), {"char": "•"})
        pPr.append(bu_font)
        pPr.append(bu_char)
        run = para.add_run()
        run.text = text
        run.font.size, run.font.color.rgb, run.font.name = Pt(size), DARK, FONT
    return box


def bullets_height(items, size) -> int:
    """Rough height of the bullet block, so the chart starts below it."""
    chars = int(88 * 16 / size)                     # characters per line
    lines = sum(max(1, math.ceil((len(t) + 4) / chars)) for t in items)
    return Inches(lines * size * 1.25 / 72 + len(items) * 9 / 72 + 0.1)


def charts_row(slide, paths, top, bottom):
    """Charts side by side at one height, the row centred across the slide."""
    from PIL import Image
    gap = Inches(0.2)
    ratios = []
    for p in paths:
        if p is not None and p.exists():
            with Image.open(p) as im:
                ratios.append(im.width / im.height)
        else:
            ratios.append(None)
    known = [r for r in ratios if r] or [1.8]
    fill = [r or sum(known) / len(known) for r in ratios]
    avail_w = WIDTH - gap * (len(paths) - 1)
    h = min(bottom - top, int(avail_w / sum(fill)))
    widths = [int(h * r) for r in fill]
    left = int((SLIDE_W - sum(widths) - gap * (len(paths) - 1)) / 2)
    top = int(top)                                  # right under the bullets
    for p, r, w in zip(paths, ratios, widths):
        if r:
            slide.shapes.add_picture(str(p), Emu(left), Emu(top),
                                     width=Emu(w), height=Emu(h))
        else:
            textbox(slide, Emu(left), Emu(top), Emu(w), Inches(0.5),
                    f"[chart not found: {p.name if p else '?'}]", 11, color=GREY)
        left += w + gap


# ===========================================================================
# READING THE RUN
# ===========================================================================

class Run:
    """tables.xlsx and the charts from one moc_tca.py output folder."""

    def __init__(self, folder: Path):
        self.folder = folder
        bare = folder / "charts_deck"
        self.charts = bare if bare.exists() else folder / "charts"
        self.titles = {}
        if (bare / "titles.json").exists():
            raw = json.loads((bare / "titles.json").read_text(encoding="utf-8"))
            self.titles = {Path(k).name: v for k, v in raw.items()}
        path = folder / "tables.xlsx"
        if not path.exists():
            raise SystemExit(f"ERROR: {path} not found - run moc_tca.py first.")
        self.xl = pd.ExcelFile(path)
        self.missing: list[str] = []

    def chart(self, name: str) -> Path | None:
        hits = list(self.charts.rglob(name))
        return hits[0] if hits else self.charts / name

    def table(self, name: str, keys: int = 1) -> pd.DataFrame | None:
        sheet = next((s for s in self.xl.sheet_names if s == name[:31]), None)
        if sheet is None:
            self.missing.append(name)
            return None
        df = pd.read_excel(self.xl, sheet)
        idx = list(df.columns[:keys])
        df[idx] = df[idx].ffill()               # a MultiIndex writes its outer label once
        df[idx[0]] = df[idx[0]].astype(str)
        return df.set_index(idx if keys > 1 else idx[0])

    def rows(self, name: str, keys: int = 1):
        df = self.table(name, keys)
        if df is None:
            return []
        out = []
        for k, r in df.iterrows():
            first = k[0] if isinstance(k, tuple) else k
            if str(first).strip().lower() in {"unknown", "nan", "none", ""}:
                continue
            out.append((k, r))
        return out

    def row(self, name: str, key, keys: int = 1):
        return next((r for k, r in self.rows(name, keys) if k == key), None)


def num(r, col) -> float:
    try:
        return float(r[col])
    except (TypeError, ValueError, KeyError):
        return math.nan


def holds(r) -> bool:
    """The 95% interval sits on one side of zero."""
    lo, hi = num(r, "CI low"), num(r, "CI high")
    return math.isfinite(lo) and math.isfinite(hi) and (lo > 0 or hi < 0)


def thin(r) -> bool:
    return str(r.get("small sample", False)).strip().lower() in ("true", "1")


def clear_t(t) -> bool:
    return math.isfinite(t) and abs(t) >= T_CLEAR


def sp(v) -> str:
    return f"{abs(v):.2f}"


def bps(v) -> str:
    return f"{abs(v):.1f} bps"


def cost_or_saving(v) -> str:
    return "a cost" if v < 0 else "a saving"


def cite(table, key, r) -> str:
    ci = (f"CI [{num(r, 'CI low'):+.2f}, {num(r, 'CI high'):+.2f}]"
          if math.isfinite(num(r, "CI low")) else "no CI (too few orders)")
    return (f"`{table}` [{key}]: {num(r, 'spreads'):+.2f} spreads x "
            f"[{num(r, 'spread bps (wtd)'):.1f} bps] = "
            f"{num(r, 'wtd mean bps'):+.1f} bps, {ci}, "
            f"{int(num(r, 'orders')):,} orders, "
            f"{'holds' if holds(r) else 'not distinguishable from zero'}")


def sized(rows, n_min=30):
    """Rows with enough orders to say something on a client slide."""
    return [(k, r) for k, r in rows if not thin(r) and num(r, "orders") >= n_min]


# ===========================================================================
# THE STORY
# ===========================================================================

def slide(title, b, charts, src, notes, dataset=None, appendix=False):
    return dict(title=title, bullets=[x for x in b if x], charts=charts,
                src=src, notes=notes, dataset=dataset, appendix=appendix)


def story(run: Run) -> list[dict]:
    S = []
    ds_all = f"Data Set : {DATA_SET}"
    ds_co = f"Data Set : {DATA_SET}   Filter : {FILTER_CLOSE_ONLY}"
    ds_pre = f"Data Set : {DATA_SET}   Filter : {FILTER_PRE_TRADED}"
    ds_auc = f"Data Set : {DATA_SET}   Filter : {FILTER_AUCTION}"

    # --- 1. overview ------------------------------------------------------------
    prof = run.rows("34_market_profile")
    size = run.table("52_adv_profile")
    b, src = [], []
    if prof:
        tot = sum(num(r, "notional (USDm)") for _, r in prof)
        ranked = sorted(prof, key=lambda kr: -num(kr[1], "notional (USDm)"))
        top5 = 100 * sum(num(r, "notional (USDm)") for _, r in ranked[:5]) / tot
        b.append(f"Executed US${tot / 1000:.1f}bn across {len(prof)} APAC markets.")
        a, c = ranked[0][0], ranked[1][0]
        b.append(f"{a} and {c} were the largest contributors, with the top five "
                 f"markets accounting for {top5:.0f}% of total executed notional.")
        src.append("`34_market_profile`: total " + f"{tot:,.1f}m; "
                   + ", ".join(f"{k} {num(r, '% of notional'):.1f}%" for k, r in ranked[:5]))
    es = run.table("50_exec_summary")
    fill = math.nan
    if es is not None and "metric" in es:
        hit = es[es["metric"] == "fill ratio"]
        fill = float(hit["value"].iloc[0]) if len(hit) else math.nan
    if size is not None and "0-1%" in size.index:
        u1 = num(size.loc["0-1%"], "% of orders")
        b.append(f"{u1:.0f}% of orders were below 1% ADV"
                 + (f", and the overall fill rate reached {fill:.1f}%." if math.isfinite(fill) else "."))
        src.append(f"`52_adv_profile` [0-1%]: {u1:.1f}% of orders; "
                   f"`50_exec_summary` fill ratio {fill:.2f}%")
    S.append(slide(f"{PERIOD} MOC TCA Overview", b,
                   ["13_market_notional.png", "24_adv_profile.png"], src,
                   "Scope: CLOSE strategy, 1 Jan to 23 Oct 2026. Left: value by market, "
                   "cash and swap by client account. Right: value by order size (% ADV). "
                   "India grows against H1: from 3 Aug its CAS names are in the "
                   "closing auction and those orders are kept.", ds_all))

    # --- 2. pre-traded vs close only -----------------------------------------------
    flow = run.table("51_flow_split")
    b, src = [], []
    if flow is not None and {"Close-only", "Pre-traded"} <= set(flow.index):
        co, pre = flow.loc["Close-only"], flow.loc["Pre-traded"]
        b.append(f"{num(co, '% of notional'):.0f}% of the value went straight into the "
                 f"auction: {int(num(co, 'orders')):,} orders, {num(co, 'fill ratio %'):.1f}% "
                 f"filled, at {num(co, '%ADV (notional-weighted)'):.2f}% ADV on average.")
        b.append(f"The other {num(pre, '% of notional'):.0f}% started before the close. "
                 f"Those orders were larger, at {num(pre, '%ADV (notional-weighted)'):.2f}% ADV, "
                 f"and still executed {num(pre, '% of notional in the close'):.0f}% of their "
                 f"value in the close.")
        src.append("`51_flow_split`: " + "; ".join(
            f"{k} {num(r, 'notional (USDm)'):,.0f}m ({num(r, '% of notional'):.1f}%), "
            f"{int(num(r, 'orders')):,} orders, fill {num(r, 'fill ratio %'):.2f}%, "
            f"{num(r, '%ADV (notional-weighted)'):.2f}% ADV, "
            f"{num(r, '% of notional in the close'):.0f}% in the close"
            for k, r in flow.iterrows()))
    S.append(slide("Pre-traded vs Close only", b, ["23_flow_split.png"], src,
                   "Close-only: 99.5% or more of the order printed in the auction. "
                   "Pre-traded: part of it traded before.\n\nFrom here on most charts "
                   "are in spreads: the result divided by the average spread, the "
                   "[x bps] under each bar. Multiply the two to get bps.", ds_all))

    # --- 3. where the cost came from ---------------------------------------------------
    d3 = run.table("91_decomp3_market")
    b, src = [], []
    if d3 is not None and "All" in d3.index:
        a = d3.loc["All"]
        tr, au, ex = (num(a, "pre-close trend bps"), num(a, "auction move bps"),
                      num(a, "execution vs close bps"))
        tot = num(a, "total vs Arrival bps")
        if not math.isfinite(tot):
            tot = tr + au + ex
        b.append(f"Against the arrival price, the close orders came in at {tot:+.1f} bps, "
                 f"{cost_or_saving(tot)}.")
        legs = {"the stock's move before the close": tr, "the auction move": au,
                "execution against the close": ex}
        big_leg = max(legs, key=lambda k: abs(legs[k]))
        rest = [k for k in legs if k != big_leg]
        b.append(f"The largest part was {big_leg}, {legs[big_leg]:+.1f} bps. "
                 f"{rest[0][0].upper() + rest[0][1:]} was {legs[rest[0]]:+.1f} bps and "
                 f"{rest[1]} {legs[rest[1]]:+.1f} bps.")
        big = sorted([(k, r) for k, r in d3.iterrows() if k != "All"
                      and not thin(r) and num(r, "orders") >= 30
                      and clear_t(num(r, "auction move t-stat"))
                      and num(r, "auction move bps") < 0],
                     key=lambda kr: num(kr[1], "auction move bps"))
        if big:
            k, r = big[0]
            b.append(f"{k} is where the auction cost most: "
                     f"{num(r, 'auction move bps'):+.1f} bps over {int(num(r, 'orders')):,} orders.")
        src.append(f"`91_decomp3_market` [All]: trend {tr:+.2f}, auction {au:+.2f}, "
                   f"execution {ex:+.2f}, total {tot:+.2f} bps")
        src += [f"`91_decomp3_market` [{k}]: auction {num(r, 'auction move bps'):+.2f} bps, "
                f"t {num(r, 'auction move t-stat'):.1f}, {int(num(r, 'orders')):,} orders"
                for k, r in big[:3]]
    S.append(slide("Where the Cost Came From", b, ["40_decomposition_market.png"], src,
                   "vs Arrival = pre-close trend (arrival -> last continuous trade) + "
                   "auction move (last continuous trade -> official close) + execution vs "
                   "close. The first two are the market; only the third is the algo. "
                   "Positive is a saving. Small markets (NZ, PH) swing on a handful of "
                   "orders - do not read them.", ds_auc))

    # --- 4. auction move -------------------------------------------------------------------
    am = sized(run.rows("90_auction_move_mkt"))
    dis = run.table("92_dislocation_share")
    b, src = [], []
    if am:
        neg = [(k, r) for k, r in am if num(r, "spreads") < 0]
        b.append(f"The close printed against our orders in {len(neg)} of {len(am)} markets, "
                 "by less than one spread in most of them.")
        worst = sorted([(k, r) for k, r in am if holds(r) and num(r, "spreads") < 0],
                       key=lambda kr: num(kr[1], "spreads"))
        if worst:
            k, r = worst[0]
            b.append(f"{k} moved most: {sp(num(r, 'spreads'))} spreads, about "
                     f"{bps(num(r, 'wtd mean bps'))}.")
        src += [cite("90_auction_move_mkt", k, r) for k, r in am]
    if dis is not None and "All" in dis.index:
        al = dis.loc["All"]
        top = dis.drop(index="All").sort_values("% orders > 1 spread", ascending=False)
        top = top[top["orders"] >= 30].head(2)
        b.append(f"Across the book, {num(al, '% orders > 1 spread'):.0f}% of orders were in "
                 f"names whose close moved more than one spread from the last trade"
                 + (f", most often in {' and '.join(top.index)}." if len(top) else "."))
        src.append(f"`92_dislocation_share` [All]: >1 sp {num(al, '% orders > 1 spread'):.1f}%, "
                   f">2 {num(al, '% orders > 2 spread'):.1f}%, >3 {num(al, '% orders > 3 spread'):.1f}%")
    S.append(slide("Auction Move, by Market", b,
                   ["41_auction_move_market.png", "49_dislocation_share_market.png"], src,
                   "Auction move = official close against the last continuous trade, "
                   "side-adjusted: negative means the close printed against our order. "
                   "Right: share of orders whose close moved more than 1, 2, 3 spreads, "
                   "either direction.", ds_auc))

    # --- 5. pre-close drift ------------------------------------------------------------------
    dr = run.table("93_preclose_drift_mkt")
    b, src = [], []
    if dr is not None and "All" in dr.index:
        a = dr.loc["All"]
        cols = [c for c in dr.columns if c.startswith("last ") and c.endswith(" bps")]
        parts = ", ".join(f"{num(a, c):+.1f} bps over the {c[:-4]}" for c in cols)
        against = all(num(a, c) < 0 for c in cols)
        b.append(("Before the auction, the price was already moving against the orders: "
                  if against else "Before the auction, the price moved ") + parts + ".")
        b.append(f"The auction then added {num(a, 'auction bps'):+.1f} bps.")
        src.append("`93_preclose_drift_mkt` [All]: " + ", ".join(
            f"{c} {num(a, c):+.2f} (t {num(a, c.replace(' bps', ' t-stat')):.1f})"
            for c in cols + ["auction bps"]))
    side = run.table("93b_preclose_drift_side")
    if side is not None and {"Buy", "Sell"} <= set(side.index):
        bu, se = side.loc["Buy"], side.loc["Sell"]
        b.append(f"Buys paid in the auction ({num(bu, 'auction bps'):+.1f} bps); sells did "
                 f"not ({num(se, 'auction bps'):+.1f} bps).")
        src.append(f"`93b_preclose_drift_side`: auction Buy {num(bu, 'auction bps'):+.2f}, "
                   f"Sell {num(se, 'auction bps'):+.2f} bps")
    S.append(slide("Pre-close Drift", b, ["42_pre_close_drift_market.png"], src,
                   "Each window ends at the continuous close and is side-adjusted: "
                   "negative means the price moved against the order. A move before the "
                   "auction is not proof of information leakage - published imbalances "
                   "move prices too.", ds_auc))

    # --- 6. close vs last VWAP ----------------------------------------------------------------
    vw = run.table("94_close_vs_vwap_mkt")
    b, src = [], []
    if vw is not None and "All" in vw.index:
        a = vw.loc["All"]
        b.append(f"On the typical name-day the close printed {num(a, 'median |close vs VWAP| bps'):.0f} bps "
                 f"from the last 30 minutes' VWAP, either way.")
        if "% name-days > 1 spread" in vw:
            b.append(f"It was more than one spread away on "
                     f"{num(a, '% name-days > 1 spread'):.0f}% of name-days.")
        top = vw.drop(index="All")
        top = top[top["name-days"] >= 30].head(2)
        if len(top):
            b.append(f"The close was furthest from the VWAP in {' and '.join(top.index)}.")
        src.append(f"`94_close_vs_vwap_mkt` [All]: median {num(a, 'median |close vs VWAP| bps'):.2f}, "
                   f"mean {num(a, 'mean |close vs VWAP| bps'):.2f} bps")
    S.append(slide("Close vs Last 30 Min VWAP", b, ["44_close_vs_vwap_market.png"], src,
                   "One count per name per day, not side-adjusted. The VWAP is the extract's "
                   "last30min_vwap_cont (continuous trading only). The email asked for 15 "
                   "minutes; the data carries 30.", ds_auc))

    # --- 7. close only participation ----------------------------------------------------------
    pr = sized(run.rows("81_close_pr_mkt_close_only"), 10)
    b, src = [], []
    if pr:
        lo = min(pr, key=lambda kr: num(kr[1], "close PR % (mean)"))
        hi = max(pr, key=lambda kr: num(kr[1], "close PR % (mean)"))
        b.append(f"Close-only orders were a small part of each auction: from "
                 f"{num(lo[1], 'close PR % (mean)'):.1f}% in {lo[0]} to "
                 f"{num(hi[1], 'close PR % (mean)'):.1f}% in {hi[0]} on average.")
        src += [f"`81_close_pr_mkt_close_only` [{k}]: mean {num(r, 'close PR % (mean)'):.2f}%, "
                f"{int(num(r, 'orders')):,} orders" for k, r in pr]
    S.append(slide("Close Only Participation, by Market", b,
                   ["35_close_pr_close_only.png"], src,
                   "Participation = our shares in the close over the official close auction "
                   "volume, averaged per order. India (PR): non-CAS names use the order's "
                   "participation rate.", ds_co))

    # --- 8. close only reversion by market ---------------------------------------------------
    rv = sized(run.rows("71_co_reversion_mkt_spreads"))
    allc = run.row("71b_co_reversion_side_spr", "All")
    b, src = [], []
    if allc is not None:
        b.append(f"Overall, close-only prints did not come back overnight: "
                 f"{num(allc, 'spreads'):+.2f} spreads to the next open"
                 + ("." if not holds(allc) else ", a result that holds."))
        src.append(cite("71b_co_reversion_side_spr", "All", allc))
    if rv:
        bad = sorted([(k, r) for k, r in rv if holds(r) and num(r, "spreads") < 0],
                     key=lambda kr: num(kr[1], "spreads"))
        b.append("No market shows a clear reversion against us." if not bad else
                 f"The clearest reversion against us was in {bad[0][0]}: "
                 f"{sp(num(bad[0][1], 'spreads'))} spreads.")
        src += [cite("71_co_reversion_mkt_spreads", k, r) for k, r in rv]
    S.append(slide("Close Only Reversion, by Market", b, ["22_reversion_close_only.png"], src,
                   "Close-only orders print at the close by design, so reversion to the next "
                   "open is the test. Negative = the price came back against us.", ds_co))

    # --- 9. close only reversion by side --------------------------------------------------------
    b, src = [], []
    for s_ in ("Buy", "Sell"):
        r = run.row("71b_co_reversion_side_spr", s_)
        if r is not None:
            b.append(f"{s_}s: {num(r, 'spreads'):+.2f} spreads to the next open"
                     + (", clear." if holds(r) else ", within noise."))
            src.append(cite("71b_co_reversion_side_spr", s_, r))
    S.append(slide("Close Only Reversion, by Side", b,
                   ["25_close_only_reversion_market_side.png"], src,
                   "Buy panel and sell panel on one scale.", ds_co))

    # --- 10. cost vs participation ---------------------------------------------------------------
    pb = run.rows("95a_closepr_auction_move")
    seg = run.table("103_impact_segments")
    b, src = [], []
    if pb:
        first, last = pb[0], pb[-1]
        b.append(f"The bigger our share of the auction, the more the close moved against "
                 f"us: {num(first[1], 'spreads'):+.2f} spreads below 5% participation, "
                 f"{num(last[1], 'spreads'):+.2f} at {last[0]}.")
        src += [cite("95a_closepr_auction_move", k, r) for k, r in pb]
    if seg is not None and "steepest cost" in seg:
        st = seg[seg["steepest cost"].astype(str).str.lower() == "true"]
        if len(st):
            r = st.iloc[0]
            b.append(f"The cost rises fastest between {st.index[0]} participation "
                     f"({num(r, 'slope bps per 1pp ClosePR'):+.2f} bps per point).")
            src.append(f"`103_impact_segments` [{st.index[0]}]: slope "
                       f"{num(r, 'slope bps per 1pp ClosePR'):+.3f}, t {num(r, 't-stat'):.1f}")
    S.append(slide("Auction Move, by Close Participation", b,
                   ["45a_closepr_auction_move.png", "51_impact_curve.png"], src,
                   "Participation = shares executed in the close / official close auction "
                   "volume. Execution vs close cannot show this - an auction fill prints at "
                   "the close - so the auction move is the measure. Right: fit with spread "
                   "and vol held constant.", ds_auc))

    # --- 11. pre-traded by side --------------------------------------------------------------------
    b, src = [], []
    first = run.row("73a_pre_firstexec_side_spr", "All")
    finish = run.row("73b_pre_close_side_spr", "All")
    rev = run.row("73c_pre_reversion_side_spr", "All")
    if first is not None and finish is not None:
        b.append(f"Pre-traded orders started {num(first, 'spreads'):+.2f} spreads against the "
                 f"close on their first fill and finished {num(finish, 'spreads'):+.2f} "
                 f"spreads against it.")
        src += [cite("73a_pre_firstexec_side_spr", "All", first),
                cite("73b_pre_close_side_spr", "All", finish)]
    if rev is not None:
        b.append(f"Overnight, the price moved {num(rev, 'spreads'):+.2f} spreads by the next open"
                 + (", within noise." if not holds(rev) else "."))
        src.append(cite("73c_pre_reversion_side_spr", "All", rev))
    rb = run.row("73c_pre_reversion_side_spr", "Buy")
    rs = run.row("73c_pre_reversion_side_spr", "Sell")
    if rb is not None and rs is not None:
        b.append(f"Next open: buys {num(rb, 'spreads'):+.2f}, sells {num(rs, 'spreads'):+.2f} spreads.")
        src += [cite("73c_pre_reversion_side_spr", "Buy", rb),
                cite("73c_pre_reversion_side_spr", "Sell", rs)]
    S.append(slide("Pre-traded Analysis, by Side", b, ["26_pretraded_by_side.png"], src,
                   "Start = first fill against the close. Finish = the whole execution "
                   "against the close. Next open = reversion. Positive is a saving.", ds_pre))

    # --- 12. pre-traded first fills by side and market ----------------------------------------------
    fe2 = {k: r for k, r in run.rows("74_pre_firstexec_mkt_side_spr", 2)}
    b, src = [], []
    for s_ in ("Buy", "Sell"):
        best = sorted([(k, r) for k, r in fe2.items() if k[1] == s_ and holds(r)
                       and not thin(r)], key=lambda kr: -num(kr[1], "spreads"))
        worst = [x for x in best if num(x[1], "spreads") < 0]
        if best and num(best[0][1], "spreads") > 0:
            (m, _), r = best[0]
            b.append(f"{s_}s started best in {m}: {num(r, 'spreads'):+.2f} spreads.")
            src.append(cite("74_pre_firstexec_mkt_side_spr", f"{m}, {s_}", r))
        if worst:
            (m, _), r = worst[-1]
            b.append(f"{s_}s started worst in {m}: {num(r, 'spreads'):+.2f} spreads.")
            src.append(cite("74_pre_firstexec_mkt_side_spr", f"{m}, {s_}", r))
    S.append(slide("Pre-traded Analysis, by Side/Market", b,
                   ["28_pretraded_first_exec_market_side.png"], src,
                   "First fill against the close, by market, buys and sells on one scale.",
                   ds_pre))

    # --- 13. pre-traded reversion by side and market --------------------------------------------------
    rv2 = {k: r for k, r in run.rows("76_pre_reversion_mkt_side_spr", 2)}
    b, src = [], []
    bad = sorted([(k, r) for k, r in rv2.items() if holds(r) and num(r, "spreads") < 0
                  and not thin(r)], key=lambda kr: num(kr[1], "spreads"))
    if bad:
        for (m, s_), r in bad[:2]:
            b.append(f"{m} {s_.lower()}s came back {sp(num(r, 'spreads'))} spreads by the next open.")
            src.append(cite("76_pre_reversion_mkt_side_spr", f"{m}, {s_}", r))
    else:
        b.append("No market and side shows a clear reversion against us.")
    S.append(slide("Pre-traded Reversion, by Side/Market", b,
                   ["30_pretraded_reversion_market_side.png"], src,
                   "Negative = the price came back against us by the next open. In H1 "
                   "Korea sells stood out here; check whether they still do.", ds_pre))

    # --- 14. pre-traded participation --------------------------------------------------------------------
    pp = sized(run.rows("82_close_pr_mkt_pretraded"), 10)
    b, src = [], []
    if pp:
        hi = sorted(pp, key=lambda kr: -num(kr[1], "close PR % (mean)"))[:3]
        b.append("Pre-traded orders took a bigger share of the auction, highest in "
                 + ", ".join(f"{k} ({num(r, 'close PR % (mean)'):.1f}%)" for k, r in hi) + ".")
        src += [f"`82_close_pr_mkt_pretraded` [{k}]: mean {num(r, 'close PR % (mean)'):.2f}%"
                for k, r in pp]
    S.append(slide("Pre-traded Orders | Close participation by market", b,
                   ["36_close_pr_pretraded.png"], src,
                   "Participation in the close for the part of pre-traded orders that "
                   "reached the auction, averaged per order.", ds_pre))

    # --- 15. buy vs sell ----------------------------------------------------------------------------------
    sg = run.table("98_side_asymmetry_mkt")
    b, src = [], []
    if sg is not None and "All" in sg.index and "auction move: buy - sell" in sg:
        a = sg.loc["All"]
        gap, t = num(a, "auction move: buy - sell"), num(a, "auction move: gap t-stat")
        b.append(f"In the auction, buys did {abs(gap):.1f} bps "
                 f"{'worse' if gap < 0 else 'better'} than sells"
                 + (", a gap that holds." if clear_t(t) else ", within noise."))
        src.append(f"`98_side_asymmetry_mkt` [All]: auction move buy "
                   f"{num(a, 'auction move: buy bps'):+.2f}, sell {num(a, 'auction move: sell bps'):+.2f}, "
                   f"gap {gap:+.2f}, t {t:.1f}")
        b.append("Part of this is the market: when prices rise into the close, buys pay "
                 "and sells gain, whatever the algo does.")
    S.append(slide("Buy vs Sell", b, ["48_side_gap_market.png", "46_closepr_by_side.png"], src,
                   "Left: buy minus sell on the auction move, by market, with +/- 1.96 "
                   "date-clustered standard errors. Right: the auction move by participation, "
                   "buys and sells.", ds_auc))

    # --- 16. volatility ------------------------------------------------------------------------------------
    vd = run.table("109_vol_dislocation")
    b, src = [], []
    if vd is not None and len(vd) >= 2 and "|auction move| bps (median)" in vd:
        order = sorted(vd.index, key=lambda x: float(str(x).split("-")[0]))
        lo, hi = vd.loc[order[0]], vd.loc[order[-1]]
        b.append(f"The close moves further in volatile names: a median "
                 f"{num(hi, '|auction move| bps (median)'):.0f} bps from the last trade in the "
                 f"most volatile quarter, {num(lo, '|auction move| bps (median)'):.0f} bps in the "
                 f"calmest.")
        src.append("`109_vol_dislocation`: " + ", ".join(
            f"{k} {num(vd.loc[k], '|auction move| bps (median)'):.1f}" for k in order))
    vf = run.table("110_vol_flow_vs_arrival", 2)
    if vf is not None:
        try:
            w = vf["wtd mean bps"].unstack()
            order = sorted(w.index, key=lambda x: float(str(x).split("-")[0]))
            top = w.loc[order[-1]]
            if {"Close only", "Pre-traded"} <= set(top.index):
                b.append(f"In the most volatile names, close-only orders came in at "
                         f"{top['Close only']:+.1f} bps against arrival, pre-traded at "
                         f"{top['Pre-traded']:+.1f} bps.")
                src.append(f"`110_vol_flow_vs_arrival` [{order[-1]}]: close only "
                           f"{top['Close only']:+.2f}, pre-traded {top['Pre-traded']:+.2f} bps")
        except Exception:
            pass
    S.append(slide("Volatility", b, ["56_vol_dislocation.png", "57_vol_close_vs_pretraded.png"],
                   src, "Vol = Parkinson daily volatility from the extract, in quartiles. "
                   "Right: does finishing the order before the auction pay off in calmer "
                   "names, and the auction in volatile ones?", ds_auc))

    # --- 17. India and the CAS ------------------------------------------------------------------------------
    ic = run.table("107_india_cas")
    b, src = [], []
    if ic is not None:
        if "CAS name" in ic.index:
            r = ic.loc["CAS name"]
            b.append(f"Since 3 August, India's F&O names close in an auction. Our orders "
                     f"in those names saw an auction move of {num(r, 'auction move bps'):+.1f} bps "
                     f"over {int(num(r, 'orders')):,} orders.")
        if "before CAS (VWAP close)" in ic.index:
            r = ic.loc["before CAS (VWAP close)"]
            b.append(f"Before the CAS, against a VWAP close, execution vs close was "
                     f"{num(r, 'vs Close bps'):+.1f} bps.")
        b.append("Under three months of data: read it as a first look.")
        src += [f"`107_india_cas` [{k}]: {int(num(r, 'orders')):,} orders, vs Close "
                f"{num(r, 'vs Close bps'):+.2f}, auction move {num(r, 'auction move bps'):+.2f} bps"
                for k, r in ic.iterrows()]
    S.append(slide("India Close, Before and After CAS", b, ["55_india_cas.png"], src,
                   "CAS names: an AUC or ACB condition code in qatt that day. Non-CAS names "
                   "keep the H1 rule (orders started 17:30-17:45 HKT).", ds_all))

    # --- 18. takeaways -----------------------------------------------------------------------------------------
    S.append(slide(f"{PERIOD} MOC: Takeaways", [
        "Most of the cost against arrival is the stock moving before the close, not the auction.",
        "Suggest: keep orders under 1% ADV close-only; the auction absorbs them at little cost.",
        "Suggest: cap or slice participation above the point where the auction move steepens.",
        "Suggest: review the market and side where pre-traded orders came back overnight.",
    ], [], [], "DRAFT - agree each line with the desk. Every Suggest: line must point to "
       "a slide above whose number holds."))

    # --- appendix -------------------------------------------------------------------------------------------------
    for title, charts, ds in [
        ("Appendix: Execution vs Close, by Market", ["15_market_vs_close.png"], ds_all),
        ("Appendix: Spread Relative Performance", ["16_spread_relative.png"], ds_all),
        ("Appendix: Execution vs Close, by ADV% and Spread",
         ["17_close_by_adv.png", "18_close_by_spread.png"], ds_auc),
        ("Appendix: Cost in Daily Volatility", ["52_cost_per_vol_market.png"], ds_auc),
        ("Appendix: Name Concentration and Market Cap",
         ["53_concentration_market.png", "54_cap_dislocation.png"], ds_auc),
        ("Appendix: Close vs Last 30 Min VWAP, Distribution",
         ["43_close_vs_vwap_distribution.png"], ds_auc),
        ("Appendix: Where the Cost Came From, by Side and Strategy",
         ["40b_decomposition_side.png", "40c_decomposition_strategy.png"], ds_auc),
    ]:
        S.append(slide(title, [], charts, [], "Backup. Show only if asked.", ds, True))
    return S


# ===========================================================================
# BUILD
# ===========================================================================

def agenda(slides: list) -> dict:
    """The contents slide: every main slide title with the page it is on, and
    the appendix as one line. Built from the story, so it cannot drift."""
    main = [(i + 2, sp["title"]) for i, sp in enumerate(slides) if not sp["appendix"]]
    first_app = next((i + 2 for i, sp in enumerate(slides) if sp["appendix"]), None)
    rows = [f"{n}.  {t}" for n, t in main]
    if first_app:
        rows.append(f"{first_app}.  Appendix")
    return dict(title="Agenda", bullets=[], charts=[], src=[], dataset=None,
                appendix=False, agenda=rows,
                notes="Contents. Page numbers are the slide numbers in this deck.")


def agenda_columns(slide, rows):
    """The titles in two columns, the deck's text style, no bullets."""
    half = math.ceil(len(rows) / 2)
    col_w = int((WIDTH - Inches(0.3)) / 2)
    for k, part in enumerate((rows[:half], rows[half:])):
        box = slide.shapes.add_textbox(LEFT + k * (col_w + Inches(0.3)), BULLETS_TOP,
                                       col_w, Inches(5.2))
        frame = box.text_frame
        frame.word_wrap = True
        for i, text in enumerate(part):
            para = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
            para.space_after = Pt(8)
            # Hanging indent: a wrapped title lines up under its text.
            pPr = para._p.get_or_add_pPr()
            pPr.set("marL", str(Inches(0.42)))
            pPr.set("indent", str(-Inches(0.42)))
            number, title = text.split(".  ", 1)
            run = para.add_run()
            run.text = f"{number}.	{title}"
            run.font.size, run.font.color.rgb, run.font.name = Pt(14), DARK, FONT


def build(run: Run, out: Path, template: Path | None) -> list:
    slides = story(run)
    slides = [agenda(slides)] + slides
    if template:
        prs = Presentation(str(template))
        layout = next((l for l in prs.slide_layouts if l.name.lower() == "blank"),
                      prs.slide_layouts[len(prs.slide_layouts) - 1])
    else:
        prs = Presentation()
        prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
        layout = prs.slide_layouts[6]

    for spec in slides:
        s = prs.slides.add_slide(layout)
        for ph in list(s.placeholders):           # an empty placeholder prints "Click to add"
            ph._element.getparent().remove(ph._element)
        textbox(s, LEFT, TITLE_TOP, WIDTH, Inches(0.6), spec["title"], 26, bold=True)
        if spec.get("agenda"):
            agenda_columns(s, spec["agenda"])
        n = len(spec["bullets"])
        size = 18 if not spec["charts"] else (16 if n <= 2 else 14)
        top = BULLETS_TOP
        if n:
            bullets(s, spec["bullets"], size)
            top = BULLETS_TOP + bullets_height(spec["bullets"], size)
        if spec["charts"]:
            bottom = DATASET_TOP - Inches(0.05) if spec["dataset"] else BOTTOM - Inches(0.1)
            charts_row(s, [run.chart(c) for c in spec["charts"]],
                       max(top, BULLETS_TOP), bottom)
        if spec["dataset"]:
            textbox(s, LEFT, DATASET_TOP, WIDTH, Inches(0.35), spec["dataset"], 13,
                    align=PP_ALIGN.CENTER)
        chart_notes = [f"{run.titles[c]['title']}: " + " ".join(run.titles[c]["notes"])
                       for c in spec["charts"] if run.titles.get(c, {}).get("notes")]
        s.notes_slide.notes_text_frame.text = (
            spec["notes"]
            + ("\n\nOn the charts:\n" + "\n".join(chart_notes) if chart_notes else "")
            + ("\n\nSources:\n" + "\n".join(spec["src"]) if spec["src"] else ""))
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out)
    return slides


def write_evidence(md: Path, pptx: Path, slides: list) -> None:
    saved = Presentation(str(pptx))
    text = " ".join(sh.text_frame.text for sl in saved.slides
                    for sh in sl.shapes if sh.has_text_frame)
    lines = [f"# Evidence for {pptx.name}", "",
             "Every bullet and the table row it rests on. Spreads x [spread bps] = bps. "
             "'holds' = the 95% interval stays on one side of zero. + = saving.", ""]
    for i, s in enumerate(slides, start=1):
        lines += [f"## {i}. {s['title']}", ""]
        for b in s["bullets"]:
            flag = "" if b in text else "  **(NOT FOUND IN SAVED DECK)**"
            lines.append(f"- {b}{flag}")
        if s["charts"]:
            lines.append(f"- charts: {', '.join(s['charts'])}")
        if s["src"]:
            lines += ["", "Backed by:", ""] + [f"- {x}" for x in s["src"]]
        lines.append("")
    md.write_text("\n".join(lines), encoding="utf-8")


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--results", type=Path, default=Path("output"),
                   help="the moc_tca.py --out folder (tables.xlsx + charts)")
    p.add_argument("--out", type=Path, default=None)
    p.add_argument("--template", type=Path, default=None,
                   help="a .pptx or .potx whose master the slides go on")
    a = p.parse_args(argv)

    run = Run(a.results)
    out = a.out or a.results / f"MOC_TCA_{PERIOD.replace(' ', '_')}.pptx"
    slides = build(run, out, a.template)
    write_evidence(out.with_suffix(".evidence.md"), out, slides)
    print(f"deck      -> {out}  ({len(slides)} slides)")
    print(f"evidence  -> {out.with_suffix('.evidence.md')}")
    if run.missing:
        print("missing tables (their bullets were left out): "
              + ", ".join(sorted(set(run.missing))))
    gone = sorted({c for s in slides for c in s["charts"] if not run.chart(c).exists()})
    if gone:
        print("missing charts: " + ", ".join(gone))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
