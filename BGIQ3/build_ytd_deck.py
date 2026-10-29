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
FILTER_HK = "Hong Kong"
FILTER_INDIA = "India from 3 Aug 2026"

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
    """The slides in the order of the client deck, then the appendix.

    The story: what we traded and how (1-2), close-only orders hold at the
    close (3-5), pre-traded orders gain by starting early but the close then
    prints against them, mostly on buys (6-11), the Hong Kong and India
    deep dives (12-17), then what to change (18).
    """
    S = []
    ds_all = f"Data Set : {DATA_SET}"
    ds_co = f"Data Set : {DATA_SET}   Filter : {FILTER_CLOSE_ONLY}"
    ds_pre = f"Data Set : {DATA_SET}   Filter : {FILTER_PRE_TRADED}"
    ds_auc = f"Data Set : {DATA_SET}   Filter : {FILTER_AUCTION}"
    ds_hk = f"Data Set : {DATA_SET}   Filter : {FILTER_HK}"
    ds_hk_co = f"Data Set : {DATA_SET}   Filter : {FILTER_HK}, {FILTER_CLOSE_ONLY}"
    ds_hk_pre = f"Data Set : {DATA_SET}   Filter : {FILTER_HK}, {FILTER_PRE_TRADED}"
    ds_in = f"Data Set : {DATA_SET}   Filter : {FILTER_INDIA}"

    # --- 1. overview ------------------------------------------------------------
    prof = run.rows("34_market_profile")
    size = run.table("52_adv_profile")
    b, src = [], []
    if prof:
        tot = sum(num(r, "notional (USDm)") for _, r in prof)
        ranked = sorted(prof, key=lambda kr: -num(kr[1], "notional (USDm)"))
        top5 = 100 * sum(num(r, "notional (USDm)") for _, r in ranked[:5]) / tot
        b.append(f"Executed US${tot / 1000:.1f}bn across {len(prof)} APAC markets.")
        b.append(f"{ranked[0][0]} and {ranked[1][0]} led. The top five markets "
                 f"were {top5:.0f}% of the value.")
        src.append("`34_market_profile`: total " + f"{tot:,.1f}m; "
                   + ", ".join(f"{k} {num(r, '% of notional'):.1f}%" for k, r in ranked[:5]))
    if size is not None and "0-1%" in size.index:
        u1 = num(size.loc["0-1%"], "% of orders")
        b.append(f"{u1:.0f}% of orders were below 1% ADV.")
        src.append(f"`52_adv_profile` [0-1%]: {u1:.1f}% of orders")
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
        b.append(f"{num(co, '% of notional'):.0f}% of the value was close auction "
                 f"only: {int(num(co, 'orders')):,} orders at "
                 f"{num(co, '%ADV (notional-weighted)'):.1f}% ADV.")
        b.append(f"The other {num(pre, '% of notional'):.0f}% started before the close. "
                 f"Those orders were bigger, at {num(pre, '%ADV (notional-weighted)'):.1f}% ADV.")
        b.append(f"They still did {num(pre, '% of notional in the close'):.0f}% of their "
                 f"value in the close auction.")
        src.append("`51_flow_split`: " + "; ".join(
            f"{k} {num(r, 'notional (USDm)'):,.0f}m ({num(r, '% of notional'):.1f}%), "
            f"{int(num(r, 'orders')):,} orders, fill {num(r, 'fill ratio %'):.2f}%, "
            f"{num(r, '%ADV (notional-weighted)'):.2f}% ADV, "
            f"{num(r, '% of notional in the close'):.0f}% in the close"
            for k, r in flow.iterrows()))
    S.append(slide("Pre-traded vs Close only", b, ["23_flow_split.png"], src,
                   "Close-only: 99.5% or more of the order printed in the auction. "
                   "Pre-traded: part of it traded before. The rest of the deck takes "
                   "the two apart.\n\nFrom here on most charts are in spreads: the "
                   "result divided by the average spread, the [x bps] under each bar. "
                   "Multiply the two to get bps.", ds_all))

    # --- 3. close only reversion by market -------------------------------------------
    # A close-only order prints at the close by design, so the only test of
    # the price it got is whether that price held overnight.
    rv = sized(run.rows("71_co_reversion_mkt_spreads"))
    allc = run.row("71b_co_reversion_side_spr", "All")
    co_held = allc is not None and not (holds(allc) and num(allc, "spreads") < 0)
    b, src = [], []
    if allc is not None:
        v = num(allc, "spreads")
        if v >= 0:
            b.append(f"The close held overnight: {v:+.2f} spreads by the next open.")
        else:
            b.append(f"The close gave back {sp(v)} spreads by the next open"
                     + (", within noise." if co_held else ". That holds."))
        src.append(cite("71b_co_reversion_side_spr", "All", allc))
    if rv:
        k, r = min(rv, key=lambda kr: num(kr[1], "spreads"))
        if num(r, "spreads") < 0:
            b.append(f"{k} gave back the most, {sp(num(r, 'spreads'))} spreads"
                     + (". That one holds." if holds(r) else ", but within noise."))
        src += [cite("71_co_reversion_mkt_spreads", k, r) for k, r in rv]
    if co_held:
        b.append("No sign that our close-only orders pushed the close.")
    S.append(slide("Close Only Orders Reversion, by Market", b,
                   ["22_reversion_close_only.png"], src,
                   "Close-only orders print at the close by design, so execution vs close "
                   "is zero and says nothing. Reversion is the test: if our order had "
                   "pushed the close, the price would come back by the next open. "
                   "Negative = it came back against us. Philippines and other small "
                   "markets swing on a handful of orders.", ds_co))

    # --- 4. close only reversion by side -----------------------------------------------
    rb = run.row("71b_co_reversion_side_spr", "Buy")
    rs = run.row("71b_co_reversion_side_spr", "Sell")
    b, src = [], []
    if rb is not None and rs is not None:
        vb, vs = num(rb, "spreads"), num(rs, "spreads")
        b.append(f"Buys {'gained' if vb >= 0 else 'gave back'} {sp(vb)} spreads by the "
                 f"next open. Sells {'gained' if vs >= 0 else 'gave back'} {sp(vs)}.")
        if (vb > 0) != (vs > 0):
            b.append("Opposite signs mean the market moved overnight, not our orders.")
        src += [cite("71b_co_reversion_side_spr", "Buy", rb),
                cite("71b_co_reversion_side_spr", "Sell", rs)]
    ms = sized(run.rows("72_co_reversion_mkt_side_spr", 2))
    bad = sorted([(k, r) for k, r in ms if holds(r) and num(r, "spreads") < 0],
                 key=lambda kr: num(kr[1], "spreads"))
    if bad:
        (m, s_), r = bad[0]
        b.append(f"{m} {s_.lower()}s stand out: {sp(num(r, 'spreads'))} spreads back, "
                 f"about {bps(num(r, 'wtd mean bps'))}.")
        src.append(cite("72_co_reversion_mkt_side_spr", f"{m}, {s_}", r))
    S.append(slide("Close Only Orders Reversion, by Side", b,
                   ["25_close_only_reversion_market_side.png"], src,
                   "If the market rises overnight, buys look good and sells look bad, "
                   "whatever the algo did. When the two sides move in opposite "
                   "directions, that is the market. Our own impact would show as a "
                   "give-back on both sides.", ds_co))

    # --- 5. close only participation ----------------------------------------------------
    pr = sized(run.rows("81_close_pr_mkt_close_only"), 10)
    b, src = [], []
    if pr:
        lo = min(pr, key=lambda kr: num(kr[1], "close PR % (mean)"))
        hi = max(pr, key=lambda kr: num(kr[1], "close PR % (mean)"))
        b.append(f"Our close-only orders were {num(lo[1], 'close PR % (mean)'):.1f}% "
                 f"({lo[0]}) to {num(hi[1], 'close PR % (mean)'):.1f}% ({hi[0]}) "
                 f"of the auction.")
        if num(hi[1], "close PR % (mean)") < 10 and co_held:
            b.append("At that size they are unlikely to move the close.")
        src += [f"`81_close_pr_mkt_close_only` [{k}]: mean {num(r, 'close PR % (mean)'):.2f}%, "
                f"{int(num(r, 'orders')):,} orders" for k, r in pr]
    S.append(slide("Close Only Orders Participation, by Market", b,
                   ["35_close_pr_close_only.png"], src,
                   "Participation = our shares in the close over the official close auction "
                   "volume, averaged per order. India (PR): non-CAS names use the order's "
                   "participation rate.", ds_co))

    # --- 6. pre-traded by side --------------------------------------------------------
    first = run.row("73a_pre_firstexec_side_spr", "All")
    finish = run.row("73b_pre_close_side_spr", "All")
    rev = run.row("73c_pre_reversion_side_spr", "All")
    early_paid = first is not None and num(first, "spreads") > 0 and holds(first)
    b, src = [], []
    if first is not None:
        v = num(first, "spreads")
        b.append(f"Starting early paid: the first fill beat the close by {sp(v)} spreads."
                 if early_paid else
                 f"The first fill was {sp(v)} spreads {'better' if v > 0 else 'worse'} "
                 f"than the close" + (", within noise." if not holds(first) else "."))
        src.append(cite("73a_pre_firstexec_side_spr", "All", first))
    if finish is not None:
        v = num(finish, "spreads")
        b.append(f"The whole order finished {sp(v)} spreads "
                 f"{'better' if v > 0 else 'worse'} than the close.")
        src.append(cite("73b_pre_close_side_spr", "All", finish))
    rb = run.row("73c_pre_reversion_side_spr", "Buy")
    rs = run.row("73c_pre_reversion_side_spr", "Sell")
    if rb is not None and rs is not None and (num(rb, "spreads") > 0) != (num(rs, "spreads") > 0):
        b.append(f"By the next open buys {num(rb, 'spreads'):+.2f}, sells "
                 f"{num(rs, 'spreads'):+.2f}: market drift again.")
        src += [cite("73c_pre_reversion_side_spr", "Buy", rb),
                cite("73c_pre_reversion_side_spr", "Sell", rs)]
    elif rev is not None:
        b.append(f"Overnight the price moved {num(rev, 'spreads'):+.2f} spreads"
                 + (", within noise." if not holds(rev) else "."))
    if rev is not None:
        src.append(cite("73c_pre_reversion_side_spr", "All", rev))
    S.append(slide("Pre-traded Analysis, by Side", b, ["26_pretraded_by_side.png"], src,
                   "First fill vs close: + means the first fill was better for the order "
                   "than the close, so starting early helped. Execution vs close: the "
                   "whole order. Next open: reversion. Positive is a saving.", ds_pre))

    # --- 7. pre-traded first fills by side and market ------------------------------------
    fe2 = sized(run.rows("74_pre_firstexec_mkt_side_spr", 2))
    b, src = [], []
    if fe2:
        def count(side):
            rows = [r for (m, s_), r in fe2 if s_ == side]
            return sum(num(r, "spreads") > 0 for r in rows), len(rows)
        (pb, nb), (ps, ns) = count("Buy"), count("Sell")
        b.append(f"First fills beat the close in {pb}/{nb} markets for buys, "
                 f"{ps}/{ns} for sells.")
        best = {s_: max([(k, r) for k, r in fe2 if k[1] == s_],
                        key=lambda kr: num(kr[1], "spreads"), default=None)
                for s_ in ("Buy", "Sell")}
        if best["Buy"] and best["Sell"]:
            (mb, _), r1 = best["Buy"]
            (ms_, _), r2 = best["Sell"]
            if mb == ms_:
                b.append(f"{mb} gained most: {num(r1, 'spreads'):+.2f} spreads on buys, "
                         f"{num(r2, 'spreads'):+.2f} on sells.")
            else:
                b.append(f"Best: {mb} buys {num(r1, 'spreads'):+.2f}, "
                         f"{ms_} sells {num(r2, 'spreads'):+.2f} spreads.")
            src += [cite("74_pre_firstexec_mkt_side_spr", f"{mb}, Buy", r1),
                    cite("74_pre_firstexec_mkt_side_spr", f"{ms_}, Sell", r2)]
        worst = min(fe2, key=lambda kr: num(kr[1], "spreads"))
        if num(worst[1], "spreads") < 0:
            (m, s_), r = worst
            b.append(f"{m} {s_.lower()}s were the exception at {num(r, 'spreads'):+.2f}"
                     + (", a clear cost." if holds(r) else ", within noise."))
            src.append(cite("74_pre_firstexec_mkt_side_spr", f"{m}, {s_}", r))
    S.append(slide("Pre-traded Analysis, by Side/Market", b,
                   ["28_pretraded_first_exec_market_side.png"], src,
                   "First fill against the close, by market, buys and sells on one scale. "
                   "Counts use markets with at least 30 orders on that side.", ds_pre))

    # --- 8. pre-traded reversion by side and market --------------------------------------
    rv2 = sized(run.rows("76_pre_reversion_mkt_side_spr", 2))
    b, src = [], []
    bad = sorted([(k, r) for k, r in rv2 if holds(r) and num(r, "spreads") < 0],
                 key=lambda kr: num(kr[1], "spreads"))
    if bad:
        (m, s_), r = bad[0]
        again = " As in H1." if (m, s_) == ("South Korea", "Sell") else ""
        b.append(f"{m} {s_.lower()}s gave back {sp(num(r, 'spreads'))} spreads by the "
                 f"next open.{again}")
        src.append(cite("76_pre_reversion_mkt_side_spr", f"{m}, {s_}", r))
    else:
        b.append("No market and side gave back a clear amount overnight.")
    if rv2:
        up = sum(num(r, "spreads") > 0 for (m, s_), r in rv2 if s_ == "Buy")
        nbuy = sum(1 for (m, s_), _ in rv2 if s_ == "Buy")
        dn = sum(num(r, "spreads") < 0 for (m, s_), r in rv2 if s_ == "Sell")
        nsell = sum(1 for (m, s_), _ in rv2 if s_ == "Sell")
        if nbuy and nsell and up / nbuy > 0.6 and dn / nsell > 0.6:
            b.append("Elsewhere buys mostly gained and sells mostly gave back: "
                     "market drift.")
        src.append(f"`76_pre_reversion_mkt_side_spr`: buys positive in {up}/{nbuy}, "
                   f"sells negative in {dn}/{nsell}")
    S.append(slide("Pre-traded Reversion, by Side/Market", b,
                   ["30_pretraded_reversion_market_side.png"], src,
                   "Negative = the price came back against us by the next open. In H1 "
                   "Korea sells stood out here. Buys up and sells down across most "
                   "markets is the market rising overnight, not our impact.", ds_pre))

    # --- 9. pre-traded participation ------------------------------------------------------
    pp = sized(run.rows("82_close_pr_mkt_pretraded"), 10)
    b, src = [], []
    if pp:
        hi3 = sorted(pp, key=lambda kr: -num(kr[1], "close PR % (mean)"))[:3]
        top_co = max((num(r, "close PR % (mean)") for _, r in pr), default=math.nan)
        b.append("Pre-traded orders took a much bigger share of the auction"
                 + (f", up to {num(hi3[0][1], 'close PR % (mean)'):.0f}%." if hi3 else "."))
        b.append("Highest in " + ", ".join(
            f"{k} ({num(r, 'close PR % (mean)'):.1f}%)" for k, r in hi3) + ".")
        if math.isfinite(top_co):
            b.append(f"Close-only orders never went above {top_co:.0f}%.")
        src += [f"`82_close_pr_mkt_pretraded` [{k}]: mean {num(r, 'close PR % (mean)'):.2f}%, "
                f"{int(num(r, 'orders')):,} orders" for k, r in pp]
    S.append(slide("Pre-traded Orders | Close participation by market", b,
                   ["36_close_pr_pretraded.png"], src,
                   "Participation in the close for the part of pre-traded orders that "
                   "reached the auction, averaged per order. These are the bigger "
                   "orders, so a bigger share of the auction is expected.", ds_pre))

    # --- 10. pre-traded auction move by market -------------------------------------------
    # last continuous price vs close, + = the last trade before the auction
    # was better for the order than the close: the close printed against it.
    lc = sized(run.rows("68b_pre_lastcont_mkt_spreads"))
    b, src = [], []
    if lc:
        pos = [(k, r) for k, r in lc if num(r, "spreads") > 0]
        b.append(f"In {len(pos)}/{len(lc)} markets the close printed worse than the "
                 f"last trade before it.")
        clear = [k for k, r in pos if holds(r)]
        if 0 < len(clear) <= 3:
            b.append("The gap is clear in " + ", ".join(clear[:-1])
                     + (" and " if len(clear) > 1 else "") + clear[-1] + ".")
        elif clear:
            b.append(f"The gap is clear in {len(clear)} of them.")
        lo = min(lc, key=lambda kr: num(kr[1], "spreads"))
        hi = max(lc, key=lambda kr: num(kr[1], "spreads"))
        b.append(f"It ran from {num(lo[1], 'spreads'):+.2f} spreads ({lo[0]}) to "
                 f"{num(hi[1], 'spreads'):+.2f} ({hi[0]}).")
        src += [cite("68b_pre_lastcont_mkt_spreads", k, r) for k, r in lc]
    S.append(slide("Pre-traded orders | Auction Move by market", b,
                   ["10c_last_cont_market.png"], src,
                   "Last continuous price vs close. Positive = the last trade before the "
                   "auction was better for the order than the close, so the auction "
                   "itself moved the price against it. This is the cost of finishing in "
                   "the auction rather than in continuous trading.", ds_pre))

    # --- 11. pre-traded auction move by side and market -----------------------------------
    lc2 = sized(run.rows("74b_pre_lastcont_mkt_side_spr", 2))
    b, src = [], []
    if lc2:
        def pos_count(side):
            rows = [r for (m, s_), r in lc2 if s_ == side]
            return sum(num(r, "spreads") > 0 for r in rows), len(rows)
        (pb, nb), (ps, ns) = pos_count("Buy"), pos_count("Sell")
        b.append(f"Buys paid it in {pb}/{nb} markets, sells in {ps}/{ns}.")
        if nb and ns and pb / nb > ps / ns + 0.2:
            b.append("The close tends to print higher, which costs buys and helps sells.")
        worst_b = max([(k, r) for k, r in lc2 if k[1] == "Buy"],
                      key=lambda kr: num(kr[1], "spreads"), default=None)
        if worst_b and holds(worst_b[1]):
            (m, _), r = worst_b
            b.append(f"{m} buys paid most: {sp(num(r, 'spreads'))} spreads, "
                     f"about {bps(num(r, 'wtd mean bps'))}.")
            src.append(cite("74b_pre_lastcont_mkt_side_spr", f"{m}, Buy", r))
        src.append(f"`74b_pre_lastcont_mkt_side_spr`: buys positive {pb}/{nb}, "
                   f"sells positive {ps}/{ns}")
    S.append(slide("Pre-traded orders | Auction Move by Side/Market", b,
                   ["28b_pretraded_last_cont_market_side.png"], src,
                   "Same measure as the slide before, buys and sells apart. Positive = "
                   "the close printed against the order. If the close drifts up into "
                   "the auction, buys pay and sells gain whatever the algo does.", ds_pre))

    # --- 12-15. Hong Kong: cash against swap ------------------------------------------------
    hk_fe = dict(run.rows("120_hk1_first_exec"))
    hk_ex = dict(run.rows("120_hk1_vs_close"))
    hk_pf = dict(run.rows("120_hk0_profile"))
    b, src = [], []
    if {"Cash", "Swap"} <= set(hk_ex):
        c, s_ = hk_ex["Cash"], hk_ex["Swap"]
        gap = abs(num(c, "spreads") - num(s_, "spreads"))
        overlap = (num(c, "CI low") <= num(s_, "CI high")
                   and num(s_, "CI low") <= num(c, "CI high"))
        b.append(f"Cash and swap traded alike: {gap:.2f} spreads apart against the close."
                 if overlap else
                 f"Cash and swap differ by {gap:.2f} spreads against the close.")
        src += [cite("120_hk1_vs_close", g, r) for g, r in hk_ex.items()]
    if {"Cash", "Swap"} <= set(hk_fe):
        fc, fs = num(hk_fe["Cash"], "spreads"), num(hk_fe["Swap"], "spreads")
        both = fc > 0 and fs > 0
        b.append((f"First fills beat the close for both: {fc:+.2f} cash, {fs:+.2f} swap."
                  if both else
                  f"First fill vs close: {fc:+.2f} cash, {fs:+.2f} swap."))
        src += [cite("120_hk1_first_exec", g, r) for g, r in hk_fe.items()]
    if {"Cash", "Swap"} <= set(hk_pf):
        pc, ps = hk_pf["Cash"], hk_pf["Swap"]
        b.append(f"Cash: {num(pc, 'notional (USDm)'):,.0f}m over "
                 f"{int(num(pc, 'orders')):,} orders. Swap: "
                 f"{num(ps, 'notional (USDm)'):,.0f}m over {int(num(ps, 'orders')):,}.")
        src += [f"`120_hk0_profile` [{g}]: {int(num(r, 'orders')):,} orders, "
                f"{num(r, 'notional (USDm)'):,.1f}m, "
                f"{num(r, '% pre-traded (orders)'):.0f}% pre-traded"
                for g, r in hk_pf.items()]
    S.append(slide("Hong Kong Focus: Cash vs Swap", b, ["90_hk_measures.png"], src,
                   "Every Hong Kong order in the review, split by client account. Three "
                   "prices against the close: the first fill, the last continuous trade, "
                   "and the whole order. Positive = better for the order than the close.",
                   ds_hk))

    co_rev = dict(run.rows("120_hk3co_reversion"))
    b, src = [], []
    if co_rev:
        b.append("Close-only orders print at the close. The test is the next open.")
        parts = []
        for g in ("Cash", "Swap"):
            if g in co_rev:
                r = co_rev[g]
                parts.append(f"{g} {num(r, 'spreads'):+.2f}"
                             + ("" if holds(r) else " (noise)"))
                src.append(cite("120_hk3co_reversion", g, r))
        b.append("Next open vs close: " + ", ".join(parts) + " spreads.")
        gave = [(g, r) for g, r in co_rev.items() if holds(r) and num(r, "spreads") < 0]
        if not gave:
            b.append("Neither account shows our orders pushing the close.")
        else:
            g, r = gave[0]
            b.append(f"{g} gave back {sp(num(r, 'spreads'))} spreads"
                     + (": small, but clear." if num(r, "spreads") > -1 else ", a clear amount."))
    S.append(slide("Hong Kong Close Only: Cash vs Swap", b, ["90b_hk_co_measures.png"],
                   src, "Execution vs close is zero by design for close-only orders. "
                   "Negative next open = the price came back against us.", ds_hk_co))

    pre_fe = dict(run.rows("120_hk3pre_first_exec"))
    pre_ex = dict(run.rows("120_hk3pre_vs_close"))
    b, src = [], []
    if {"Cash", "Swap"} <= set(pre_fe):
        fc, fs = num(pre_fe["Cash"], "spreads"), num(pre_fe["Swap"], "spreads")
        b.append(f"Early starts paid: first fills beat the close by {fc:.2f} cash, "
                 f"{fs:.2f} swap." if fc > 0 and fs > 0 else
                 f"First fill vs close: {fc:+.2f} cash, {fs:+.2f} swap.")
        src += [cite("120_hk3pre_first_exec", g, r) for g, r in pre_fe.items()]
    if {"Cash", "Swap"} <= set(pre_ex):
        ec, es = num(pre_ex["Cash"], "spreads"), num(pre_ex["Swap"], "spreads")
        b.append(f"The whole order finished {ec:+.2f} (cash) and {es:+.2f} (swap) "
                 f"against the close.")
        src += [cite("120_hk3pre_vs_close", g, r) for g, r in pre_ex.items()]
    S.append(slide("Hong Kong Pre-traded: Cash vs Swap", b, ["90c_hk_pre_measures.png"],
                   src, "Pre-traded Hong Kong orders. Same three prices as the Hong Kong "
                   "overview. Positive = better for the order than the close.", ds_hk_pre))

    hq = run.rows("120_hk2_first_exec_adv_q", 2)
    hk_large_only = False
    b, src = [], []
    if hq:
        bands = sorted({k[0] for k, _ in hq}, key=_band_lo)
        top = bands[-1]
        small = [(k, r) for k, r in hq if k[0] != top]
        big = {k[1]: r for k, r in hq if k[0] == top}
        edge = top.split("-")[0]
        mx = max((abs(num(r, "spreads")) for _, r in small), default=math.nan)
        if {"Cash", "Swap"} <= set(big):
            c, s_ = num(big["Cash"], "spreads"), num(big["Swap"], "spreads")
            b.append(f"Above {edge}% ADV, first fills beat the close by {c:.2f} (cash) "
                     f"and {s_:.2f} (swap)." if c > 0 and s_ > 0 else
                     f"Above {edge}% ADV, first fill vs close: {c:+.2f} cash, "
                     f"{s_:+.2f} swap.")
        if small:
            b.append(f"Below it they sat within {mx:.2f} spreads of the close."
                     if mx < 0.25 else
                     f"Below it results were mixed, up to {mx:.2f} spreads either way.")
        if {"Cash", "Swap"} <= set(big):
            hk_large_only = (bool(small) and mx < 0.25
                             and min(num(r, "spreads") for r in big.values()) > 0.5)
            if hk_large_only:
                b.append("Starting early only pays on the larger Hong Kong orders.")
        src += [cite("120_hk2_first_exec_adv_q", f"{k[0]}, {k[1]}", r) for k, r in hq]
    S.append(slide("Hong Kong: First Execution by Order Size", b,
                   ["92a_hk_adv_q_first_exec.png"], src,
                   "Order size quartiles cut on Hong Kong's own orders, so cash and swap "
                   "share the bands. Close-only orders sit at zero here by design, so a "
                   "band with mostly close-only orders reads near zero.", ds_hk))

    # --- 16-17. India: CAS against non-CAS ----------------------------------------------------
    in_fe = dict(run.rows("130_in1_first_exec"))
    in_lc = dict(run.rows("130_in1_last_cont"))
    in_ex = dict(run.rows("130_in1_vs_close"))
    b, src = [], []
    if in_ex:
        b.append("From 3 August CAS names close in an auction. Others close on a VWAP.")
    if {"CAS", "Non-CAS"} <= set(in_ex):
        c, n = in_ex["CAS"], in_ex["Non-CAS"]
        b.append((f"CAS orders matched the close ({num(c, 'spreads'):+.2f})."
                  if not holds(c) else
                  f"CAS orders finished {num(c, 'spreads'):+.2f} spreads from the close.")
                 + (f" Non-CAS paid {sp(num(n, 'spreads'))} spreads tracking the VWAP."
                    if num(n, "spreads") < 0 else
                    f" Non-CAS beat the VWAP by {sp(num(n, 'spreads'))} spreads."))
        src += [cite("130_in1_vs_close", g, r) for g, r in in_ex.items()]
    if "CAS" in in_lc:
        r = in_lc["CAS"]
        b.append(f"In CAS names the close printed {sp(num(r, 'spreads'))} spreads from "
                 f"the last trade" + (", against us." if num(r, "spreads") > 0 else "."))
        src.append(cite("130_in1_last_cont", "CAS", r))
    src += [cite("130_in1_first_exec", g, r) for g, r in in_fe.items()]
    S.append(slide("India Focus: CAS vs Non-CAS", b, ["95_in_measures.png"], src,
                   "India orders from 3 Aug 2026. CAS names have an AUC or ACB code in "
                   "qatt that day; the others close on the VWAP of the last half hour, "
                   "so for them 'the close' is that VWAP. CAS spreads are tight "
                   "(about 2 bps), so a few bps read as many spreads. Under three months "
                   "of data: a first look. India reports no close share, so CAS orders "
                   "cannot be split into close-only and pre-traded.", ds_in))

    iq = run.rows("130_in2_first_exec_adv_q", 2)
    b, src = [], []
    if "CAS" in in_fe:
        r = in_fe["CAS"]
        b.append(f"CAS first fills beat the close by {num(r, 'spreads'):.2f} spreads, "
                 f"about {abs(num(r, 'wtd mean bps')):.0f} bps."
                 if num(r, "spreads") > 0 else
                 f"In CAS names the first fill was {sp(num(r, 'spreads'))} spreads "
                 f"worse than the close.")
    cas_q = sorted([(k, r) for k, r in iq if k[1] == "CAS"], key=lambda kr: _band_lo(kr[0][0]))
    if cas_q:
        (band, _), r = cas_q[-1]
        b.append(f"Most of it is the largest orders: {num(r, 'spreads'):+.2f} spreads "
                 f"above {band.split('-')[0]}% ADV" + ("." if holds(r) else ", within noise."))
        rest = cas_q[:-1]
        if rest and all(not holds(x) for _, x in rest):
            b.append("Smaller CAS orders show no clear gain from starting early.")
        src += [cite("130_in2_first_exec_adv_q", f"{k[0]}, {k[1]}", x) for k, x in iq]
    S.append(slide("India CAS: First Execution by Order Size", b,
                   ["97a_in_adv_q_first_exec.png"], src,
                   "Order size quartiles cut on India's orders from 3 Aug, so CAS and "
                   "non-CAS share the bands. Positive = the first fill was better than "
                   "the close.", ds_in))

    # --- 18. takeaways -----------------------------------------------------------------------
    b = []
    if co_held and size is not None:
        b.append("Suggest: keep orders under 1% ADV close-only. The close holds for them.")
    if pp and lc:
        hi_pr = [k for k, r in sorted(pp, key=lambda kr: -num(kr[1], "close PR % (mean)"))[:3]]
        clear_lc = {k for k, r in lc if holds(r) and num(r, "spreads") > 0}
        hit = [k for k in hi_pr if k in clear_lc]
        if hit:
            b.append(f"Suggest: in {' and '.join(hit)}, cap our share of the "
                     f"closing auction.")
    if hq and hk_large_only:
        bands = sorted({k[0] for k, _ in hq}, key=_band_lo)
        b.append(f"Suggest: in Hong Kong, start early only above "
                 f"{bands[-1].split('-')[0]}% ADV.")
    S.append(slide(f"{PERIOD} MOC: Takeaways", b[:3], [], [],
                   "DRAFT - agree each line with the desk. Each Suggest: line rests on a "
                   "slide above: close-only reversion (3-5), pre-traded participation and "
                   "auction move (9-11), Hong Kong by order size (15). India CAS is under "
                   "three months old: watch, do not act yet."))

    # --- appendix -------------------------------------------------------------------------
    for title, charts, ds in [
        ("Appendix: Execution vs Close, by Market", ["15_market_vs_close.png"], ds_all),
        ("Appendix: Where the Cost Came From", ["40_decomposition_market.png"], ds_auc),
        ("Appendix: Auction Move and Dislocation, by Market",
         ["41_auction_move_market.png", "49_dislocation_share_market.png"], ds_auc),
        ("Appendix: Pre-close Drift", ["42_pre_close_drift_market.png"], ds_auc),
        ("Appendix: Close vs Last 30 Min VWAP", ["44_close_vs_vwap_market.png"], ds_auc),
        ("Appendix: Auction Move, by Close Participation",
         ["45a_closepr_auction_move.png", "51_impact_curve.png"], ds_auc),
        ("Appendix: Buy vs Sell", ["48_side_gap_market.png", "46_closepr_by_side.png"],
         ds_auc),
        ("Appendix: Volatility",
         ["56_vol_dislocation.png", "57_vol_close_vs_pretraded.png"], ds_auc),
        ("Appendix: Hong Kong, Last Trade vs Close by Size",
         ["92b_hk_adv_q_last_cont.png"], ds_hk),
        ("Appendix: Hong Kong, Execution vs Close by Size",
         ["92c_hk_adv_q_vs_close.png"], ds_hk),
        ("Appendix: Hong Kong, First Execution by Spread",
         ["91a_hk_spread_q_first_exec.png"], ds_hk),
        ("Appendix: Hong Kong, Last Trade vs Close by Spread",
         ["91b_hk_spread_q_last_cont.png"], ds_hk),
        ("Appendix: Hong Kong, Execution vs Close by Spread",
         ["91c_hk_spread_q_vs_close.png"], ds_hk),
        ("Appendix: India Non-CAS, Close Only", ["95b_in_co_measures.png"], ds_in),
        ("Appendix: India, Last Trade vs Close by Size",
         ["97b_in_adv_q_last_cont.png"], ds_in),
        ("Appendix: India, Execution vs Close by Size",
         ["97c_in_adv_q_vs_close.png"], ds_in),
        ("Appendix: India, First Execution by Spread",
         ["96a_in_spread_q_first_exec.png"], ds_in),
        ("Appendix: India, Last Trade vs Close by Spread",
         ["96b_in_spread_q_last_cont.png"], ds_in),
        ("Appendix: India, Execution vs Close by Spread",
         ["96c_in_spread_q_vs_close.png"], ds_in),
        ("Appendix: India Close, Before and After CAS", ["55_india_cas.png"], ds_all),
        ("Appendix: Spread Relative Performance", ["16_spread_relative.png"], ds_all),
        ("Appendix: Execution vs Close, by ADV% and Spread",
         ["17_close_by_adv.png", "18_close_by_spread.png"], ds_auc),
    ]:
        S.append(slide(title, [], charts, [], "Backup. Show only if asked.", ds, True))
    return S


def _band_lo(label) -> float:
    """The lower edge of a quartile label such as '0.32-28.11%'."""
    try:
        return float(str(label).split("-")[0].rstrip("%"))
    except ValueError:
        return math.inf


# ===========================================================================
# BUILD
# ===========================================================================

def build(run: Run, out: Path, template: Path | None) -> list:
    # No agenda slide: the client deck opens on the overview.
    slides = story(run)
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


SKILL = Path(__file__).resolve().parent.parent / ".claude" / "skills" / "tca-review" / "scripts"


def check_words(slides: list) -> list[str]:
    """The skill's plain-language rules, run on the deck as built."""
    sys.path.insert(0, str(SKILL))
    try:
        from check_readability import check_deck
    except ImportError:
        return ["(readability check not found - skipped)"]
    return check_deck({"slides": [{"type": "content", "title": s["title"],
                                   "bullets": s["bullets"]} for s in slides
                                  if not s["appendix"]]})


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
    problems = check_words(slides)
    if problems:
        print("plain-language check:")
        for x in problems:
            print("  - " + x)
    return 1 if (problems and not problems[0].startswith("(")) else 0


if __name__ == "__main__":
    raise SystemExit(main())
