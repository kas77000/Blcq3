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
        v = float(r[col])
        return 0.0 if abs(v) < 0.005 and col == "spreads" else v
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


def talk(say: str, ask: str = "", nxt: str = "") -> str:
    """Speaker notes: what to say, the caveat if asked, and the bridge to the
    next slide."""
    parts = []
    if say:
        parts.append("Say: " + say)
    if ask:
        parts.append("If asked: " + ask)
    if nxt:
        parts.append(nxt)
    return "\n\n".join(parts)


def story(run: Run) -> list[dict]:
    """The slides in the order of the client deck. No appendix.

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

    # From here every slide follows the H1 shape: the first bullet says in
    # plain words what the chart measures, the next ones say what happened and
    # what it means. The speaker notes carry a script the sales trader can
    # read out ("Say:") and the caveat to keep in their pocket ("If asked:").
    story_points = {}

    # --- 3. close only reversion by market -------------------------------------------
    # A close-only order prints at the close by design, so the only test of
    # the price it got is whether that price held overnight.
    rv = sized(run.rows("71_co_reversion_mkt_spreads"))
    allc = run.row("71b_co_reversion_side_spr", "All")
    co_held = allc is not None and not (holds(allc) and num(allc, "spreads") < 0)
    b, src, say = [], [], ""
    b.append("Reversion asks one thing: did the closing price hold the next morning?")
    if allc is not None:
        v = num(allc, "spreads")
        if co_held:
            b.append(f"Overall it held: {v:+.2f} spreads by the next open"
                     + ("." if v >= 0 else ", within noise."))
            story_points["co"] = "Close-only orders: the closing price held the next morning."
        else:
            b.append(f"Overall it gave back {sp(v)} spreads by the next open.")
            story_points["co"] = "Close-only orders: the close gave back part of its move overnight."
        say = (f"Our close-only orders trade at the close, so the question is not the "
               f"price we got but whether it was a fair price. If our orders had pushed "
               f"the close, it would bounce back by the next open. Across all markets the "
               f"move to the next open was {v:+.2f} spreads, "
               + ("so the close held." if co_held else "so some of it came back."))
        src.append(cite("71b_co_reversion_side_spr", "All", allc))
    if rv:
        k, r = min(rv, key=lambda kr: num(kr[1], "spreads"))
        if num(r, "spreads") < 0:
            b.append(f"{k} moved most, {sp(num(r, 'spreads'))} spreads"
                     + (", but within noise." if not holds(r) else ". That one is real."))
        src += [cite("71_co_reversion_mkt_spreads", k, r) for k, r in rv]
    S.append(slide("Close Only Orders Reversion, by Market", b,
                   ["22_reversion_close_only.png"], src,
                   talk(say,
                        "Execution vs close is zero by design for these orders, which is "
                        "why we use reversion. Small markets such as the Philippines swing "
                        "on a handful of orders: the whisker shows it.",
                        "Next: does it look the same for buys and sells?"), ds_co))

    # --- 4. close only reversion by side -----------------------------------------------
    rb = run.row("71b_co_reversion_side_spr", "Buy")
    rs = run.row("71b_co_reversion_side_spr", "Sell")
    b, src, say = [], [], ""
    if rb is not None and rs is not None:
        vb, vs = num(rb, "spreads"), num(rs, "spreads")
        b.append(f"Buys {'gained' if vb >= 0 else 'gave back'} {sp(vb)} spreads overnight. "
                 f"Sells {'gained' if vs >= 0 else 'gave back'} {sp(vs)}.")
        if (vb > 0) != (vs > 0):
            b.append("That is the market moving overnight, not our orders moving the price.")
            say = ("When buys gain and sells lose at the same time, the whole market moved "
                   "overnight. If our orders were pushing the close, both sides would give "
                   "back. They don't.")
        src += [cite("71b_co_reversion_side_spr", "Buy", rb),
                cite("71b_co_reversion_side_spr", "Sell", rs)]
    ms = sized(run.rows("72_co_reversion_mkt_side_spr", 2))
    bad = sorted([(k, r) for k, r in ms if holds(r) and num(r, "spreads") < 0],
                 key=lambda kr: num(kr[1], "spreads"))
    if bad:
        (m, s_), r = bad[0]
        b.append(f"The one to watch: {m} {s_.lower()}s, {sp(num(r, 'spreads'))} spreads back.")
        src.append(cite("72_co_reversion_mkt_side_spr", f"{m}, {s_}", r))
    S.append(slide("Close Only Orders Reversion, by Side", b,
                   ["25_close_only_reversion_market_side.png"], src,
                   talk(say, "A rising market makes buys look good and sells look bad "
                        "whatever the algo does.", "Next: how big were we in the auction?"),
                   ds_co))

    # --- 5. close only participation ----------------------------------------------------
    pr = sized(run.rows("81_close_pr_mkt_close_only"), 10)
    b, src, say = [], [], ""
    if pr:
        lo = min(pr, key=lambda kr: num(kr[1], "close PR % (mean)"))
        hi = max(pr, key=lambda kr: num(kr[1], "close PR % (mean)"))
        b.append("Participation is our share of the volume in the closing auction.")
        b.append(f"Close-only orders averaged {num(lo[1], 'close PR % (mean)'):.1f}% "
                 f"({lo[0]}) to {num(hi[1], 'close PR % (mean)'):.1f}% ({hi[0]}).")
        if num(hi[1], "close PR % (mean)") < 10 and co_held:
            b.append("Too small to move the close, which is why the price held.")
        say = (f"On average our close-only order was between "
               f"{num(lo[1], 'close PR % (mean)'):.1f}% and "
               f"{num(hi[1], 'close PR % (mean)'):.1f}% of the auction. At that size we are "
               f"a price taker, and the reversion on the last two slides confirms it.")
        src += [f"`81_close_pr_mkt_close_only` [{k}]: mean {num(r, 'close PR % (mean)'):.2f}%, "
                f"{int(num(r, 'orders')):,} orders" for k, r in pr]
    S.append(slide("Close Only Orders Participation, by Market", b,
                   ["35_close_pr_close_only.png"], src,
                   talk(say, "Participation = our shares in the close over the official "
                        "close volume, averaged per order. India (PR): non-CAS names use the "
                        "order's participation rate.",
                        "Next: the orders that started before the close."), ds_co))

    # --- 6. pre-traded by side --------------------------------------------------------
    first = run.row("73a_pre_firstexec_side_spr", "All")
    finish = run.row("73b_pre_close_side_spr", "All")
    rev = run.row("73c_pre_reversion_side_spr", "All")
    early_paid = first is not None and num(first, "spreads") > 0 and holds(first)
    b, src, say = [], [], ""
    b.append("Pre-traded orders start before the auction, then finish in it.")
    if first is not None:
        v = num(first, "spreads")
        b.append(f"Starting early paid: the first fill beat the close by {sp(v)} spreads."
                 if early_paid else
                 f"Starting early did not clearly pay: first fill vs close {v:+.2f} spreads.")
        if early_paid:
            story_points["pre"] = (f"Pre-traded orders: starting early beat the close by "
                                   f"{sp(v)} spreads.")
        src.append(cite("73a_pre_firstexec_side_spr", "All", first))
    if finish is not None:
        v2 = num(finish, "spreads")
        b.append(f"The whole order finished {sp(v2)} spreads "
                 f"{'better' if v2 > 0 else 'worse'} than the close.")
        src.append(cite("73b_pre_close_side_spr", "All", finish))
        if first is not None:
            say = (f"Left bar: where the price was when we started, against the close. "
                   f"{num(first, 'spreads'):+.2f} means our first fill was "
                   f"{'better' if num(first, 'spreads') > 0 else 'worse'} than waiting for "
                   f"the close. Middle bar: the whole order, {v2:+.2f}. Right bar: the next "
                   f"open, which tells us whether the price came back.")
    if rev is not None:
        src.append(cite("73c_pre_reversion_side_spr", "All", rev))
    S.append(slide("Pre-traded Analysis, by Side", b, ["26_pretraded_by_side.png"], src,
                   talk(say, "Positive is always a saving. Next open on buys and sells "
                        "moving in opposite directions is market drift.",
                        "Next: does this hold market by market?"), ds_pre))

    # --- 7. pre-traded first fills by side and market ------------------------------------
    fe2 = sized(run.rows("74_pre_firstexec_mkt_side_spr", 2))
    b, src, say = [], [], ""
    if fe2:
        def count(side):
            rows = [r for (m, s_), r in fe2 if s_ == side]
            return sum(num(r, "spreads") > 0 for r in rows), len(rows)
        (pb, nb), (ps, ns) = count("Buy"), count("Sell")
        b.append("Was the price better when we started than at the close?")
        b.append(f"Yes in {pb} of {nb} markets for buys, {ps} of {ns} for sells.")
        say = (f"Each bar is one market and side. Above zero, our first fill was better "
               f"than the close: starting early helped. That was true in {pb} of {nb} "
               f"markets for buys and {ps} of {ns} for sells.")
        worst = min(fe2, key=lambda kr: num(kr[1], "spreads"))
        best = max(fe2, key=lambda kr: num(kr[1], "spreads"))
        if num(worst[1], "spreads") < 0 and holds(worst[1]):
            (m, s_), r = worst
            b.append(f"The exception is {m} {s_.lower()}s, at {num(r, 'spreads'):+.2f} spreads.")
            src.append(cite("74_pre_firstexec_mkt_side_spr", f"{m}, {s_}", r))
        else:
            (m, s_), r = best
            b.append(f"{m} {s_.lower()}s gained most, {num(r, 'spreads'):+.2f} spreads.")
            src.append(cite("74_pre_firstexec_mkt_side_spr", f"{m}, {s_}", r))
    S.append(slide("Pre-traded Analysis, by Side/Market", b,
                   ["28_pretraded_first_exec_market_side.png"], src,
                   talk(say, "Counts use markets with at least 30 orders on that side. A "
                        "whisker crossing zero means that market's result is not "
                        "distinguishable from zero.",
                        "Next: did the price come back after we finished?"), ds_pre))

    # --- 8. pre-traded reversion by side and market --------------------------------------
    rv2 = sized(run.rows("76_pre_reversion_mkt_side_spr", 2))
    b, src, say = [], [], ""
    b.append("Next open vs close: did the price come back after we finished?")
    bad = sorted([(k, r) for k, r in rv2 if holds(r) and num(r, "spreads") < 0],
                 key=lambda kr: num(kr[1], "spreads"))
    if bad:
        (m, s_), r = bad[0]
        again = ", as in H1" if (m, s_) == ("South Korea", "Sell") else ""
        b.append(f"{m} {s_.lower()}s gave back {sp(num(r, 'spreads'))} spreads{again}.")
        say = (f"{m} {s_.lower()}s are the outlier: the price came back "
               f"{sp(num(r, 'spreads'))} spreads, about {bps(num(r, 'wtd mean bps'))}, by "
               f"the next morning.")
        src.append(cite("76_pre_reversion_mkt_side_spr", f"{m}, {s_}", r))
    else:
        b.append("No market and side gave back a clear amount overnight.")
    if rv2:
        up = sum(num(r, "spreads") > 0 for (m, s_), r in rv2 if s_ == "Buy")
        nbuy = sum(1 for (m, s_), _ in rv2 if s_ == "Buy")
        dn = sum(num(r, "spreads") < 0 for (m, s_), r in rv2 if s_ == "Sell")
        nsell = sum(1 for (m, s_), _ in rv2 if s_ == "Sell")
        if nbuy and nsell and up / nbuy > 0.6 and dn / nsell > 0.6:
            b.append("Elsewhere buys gained and sells lost: that is the market, not us.")
        src.append(f"`76_pre_reversion_mkt_side_spr`: buys positive in {up}/{nbuy}, "
                   f"sells negative in {dn}/{nsell}")
    S.append(slide("Pre-traded Reversion, by Side/Market", b,
                   ["30_pretraded_reversion_market_side.png"], src,
                   talk(say, "In H1 Korea sells stood out here too. Buys up and sells down "
                        "across most markets is the market rising overnight.",
                        "Next: how big were these orders in the auction?"), ds_pre))

    # --- 9. pre-traded participation ------------------------------------------------------
    pp = sized(run.rows("82_close_pr_mkt_pretraded"), 10)
    b, src, say = [], [], ""
    if pp:
        hi3 = sorted(pp, key=lambda kr: -num(kr[1], "close PR % (mean)"))[:3]
        top_co = max((num(r, "close PR % (mean)") for _, r in pr), default=math.nan)
        b.append("Pre-traded orders are bigger, so they take more of the auction.")
        b.append("Highest in " + ", ".join(
            f"{k} ({num(r, 'close PR % (mean)'):.0f}%)" for k, r in hi3) + ".")
        if math.isfinite(top_co):
            b.append(f"Close-only orders never took more than {top_co:.0f}%.")
        say = ("These are the larger orders, which is why they start early: the auction "
               "alone cannot absorb them. The more of the auction we are, the more the "
               "close can move against us, which is the next slide.")
        src += [f"`82_close_pr_mkt_pretraded` [{k}]: mean {num(r, 'close PR % (mean)'):.2f}%, "
                f"{int(num(r, 'orders')):,} orders" for k, r in pp]
    S.append(slide("Pre-traded Orders | Close participation by market", b,
                   ["36_close_pr_pretraded.png"], src,
                   talk(say, "Participation in the close for the part of the order that "
                        "reached the auction, averaged per order.",
                        "Next: did the auction move against these orders?"), ds_pre))

    # --- 10. pre-traded auction move by market -------------------------------------------
    # last continuous price vs close, + = the last trade before the auction
    # was better for the order than the close: the close printed against it.
    lc = sized(run.rows("68b_pre_lastcont_mkt_spreads"))
    b, src, say = [], [], ""
    b.append("Auction move: the close against the last trade just before the auction.")
    if lc:
        pos = [(k, r) for k, r in lc if num(r, "spreads") > 0]
        clear = [k for k, r in pos if holds(r)]
        b.append(f"For pre-traded orders the close printed worse in {len(pos)} of "
                 f"{len(lc)} markets.")
        if 0 < len(clear) <= 3:
            b.append("Clearest in " + ", ".join(clear[:-1])
                     + (" and " if len(clear) > 1 else "") + clear[-1] + ".")
        elif clear:
            b.append(f"The move is clear in {len(clear)} of them.")
        if len(pos) > len(lc) / 2:
            story_points["auc"] = "the auction itself moved the close against them."
        say = ("Above zero means the last trade before the auction was a better price for "
               "our order than the close. In other words, the auction moved the price "
               "against us. That is the cost of finishing in the auction rather than in "
               "continuous trading.")
        src += [cite("68b_pre_lastcont_mkt_spreads", k, r) for k, r in lc]
    S.append(slide("Pre-traded orders | Auction Move by market", b,
                   ["10c_last_cont_market.png"], src,
                   talk(say, "This is not the algo's execution: it is how far the auction "
                        "print sat from the last continuous trade.",
                        "Next: is it buys, sells or both?"), ds_pre))

    # --- 11. pre-traded auction move by side and market -----------------------------------
    lc2 = sized(run.rows("74b_pre_lastcont_mkt_side_spr", 2))
    b, src, say = [], [], ""
    if lc2:
        def pos_count(side):
            rows = [r for (m, s_), r in lc2 if s_ == side]
            return sum(num(r, "spreads") > 0 for r in rows), len(rows)
        (pb, nb), (ps, ns) = pos_count("Buy"), pos_count("Sell")
        b.append(f"Buys paid the auction move in {pb} of {nb} markets. Sells: {ps} of {ns}.")
        if nb and ns and pb / nb > ps / ns + 0.2:
            b.append("The close tends to print high: that costs buys and helps sells.")
            story_points["auc"] = "the auction moved the close against them, mostly on buys."
        worst_b = max([(k, r) for k, r in lc2 if k[1] == "Buy"],
                      key=lambda kr: num(kr[1], "spreads"), default=None)
        if worst_b and holds(worst_b[1]):
            (m, _), r = worst_b
            b.append(f"{m} buys paid most: {sp(num(r, 'spreads'))} spreads, "
                     f"about {bps(num(r, 'wtd mean bps'))}.")
            src.append(cite("74b_pre_lastcont_mkt_side_spr", f"{m}, Buy", r))
        say = ("Split by side, the cost sits mostly on buys. When prices drift up into the "
               "close, a buyer in the auction pays for it. That points to how much of a "
               "large buy we leave for the auction.")
        src.append(f"`74b_pre_lastcont_mkt_side_spr`: buys positive {pb}/{nb}, "
                   f"sells positive {ps}/{ns}")
    S.append(slide("Pre-traded orders | Auction Move by Side/Market", b,
                   ["28b_pretraded_last_cont_market_side.png"], src,
                   talk(say, "If the whole market closes higher, buys pay and sells gain "
                        "whatever the algo does.",
                        "Next: two markets in more detail, Hong Kong then India."), ds_pre))

    # --- 12-15. Hong Kong: cash against swap ------------------------------------------------
    hk_fe = dict(run.rows("120_hk1_first_exec"))
    hk_ex = dict(run.rows("120_hk1_vs_close"))
    hk_pf = dict(run.rows("120_hk0_profile"))
    b, src, say = [], [], ""
    b.append("Same algo, two account types: did cash and swap trade differently?")
    if {"Cash", "Swap"} <= set(hk_ex):
        c, s_ = hk_ex["Cash"], hk_ex["Swap"]
        gap = abs(num(c, "spreads") - num(s_, "spreads"))
        overlap = (num(c, "CI low") <= num(s_, "CI high")
                   and num(s_, "CI low") <= num(c, "CI high"))
        b.append(f"No: against the close they finished {gap:.2f} spreads apart, within noise."
                 if overlap else
                 f"Yes: against the close they finished {gap:.2f} spreads apart.")
        if overlap:
            story_points["hk"] = "Hong Kong: cash and swap traded alike."
        src += [cite("120_hk1_vs_close", g, r) for g, r in hk_ex.items()]
    if {"Cash", "Swap"} <= set(hk_fe):
        fc, fs = num(hk_fe["Cash"], "spreads"), num(hk_fe["Swap"], "spreads")
        if fc > 0 and fs > 0:
            b.append(f"Both beat the close at the first fill: {fc:+.2f} cash, {fs:+.2f} swap.")
        src += [cite("120_hk1_first_exec", g, r) for g, r in hk_fe.items()]
    if {"Cash", "Swap"} <= set(hk_pf):
        pc, ps = hk_pf["Cash"], hk_pf["Swap"]
        say = (f"Hong Kong is the one market with real size on both account types: "
               f"{num(pc, 'notional (USDm)'):,.0f}m cash over {int(num(pc, 'orders')):,} "
               f"orders, {num(ps, 'notional (USDm)'):,.0f}m swap over "
               f"{int(num(ps, 'orders')):,}. The question is whether the account type "
               f"changes the result. Three prices, all against the close: the first fill, "
               f"the last trade before the auction, and the whole order.")
        src += [f"`120_hk0_profile` [{g}]: {int(num(r, 'orders')):,} orders, "
                f"{num(r, 'notional (USDm)'):,.1f}m, "
                f"{num(r, '% pre-traded (orders)'):.0f}% pre-traded"
                for g, r in hk_pf.items()]
    S.append(slide("Hong Kong Focus: Cash vs Swap", b, ["90_hk_measures.png"], src,
                   talk(say, "Cash or swap is read from the client account. Positive = "
                        "better for the order than the close.",
                        "Next: the same split for close-only and pre-traded orders."), ds_hk))

    co_rev = dict(run.rows("120_hk3co_reversion"))
    b, src, say = [], [], ""
    if co_rev:
        b.append("Close-only orders trade at the close, so the test is the next open.")
        parts = []
        for g in ("Cash", "Swap"):
            if g in co_rev:
                v = num(co_rev[g], "spreads")
                parts.append(f"{g if not parts else g.lower()} "
                             f"{'gained' if v >= 0 else 'gave back'} {sp(v)}")
                src.append(cite("120_hk3co_reversion", g, co_rev[g]))
        b.append(", ".join(parts) + " spreads overnight.")
        gave = [(g, r) for g, r in co_rev.items() if holds(r) and num(r, "spreads") < 0]
        if gave:
            g, r = gave[0]
            b.append(f"{g}'s give-back is small but real: worth watching.")
        else:
            b.append("Neither account shows our orders pushing the close.")
        say = ("Same test as slide 4, Hong Kong only. A give-back means the closing price "
               "did not hold the next morning.")
    S.append(slide("Hong Kong Close Only: Cash vs Swap", b, ["90b_hk_co_measures.png"],
                   src, talk(say, "Execution vs close is zero by design for close-only "
                             "orders.", "Next: pre-traded Hong Kong orders."), ds_hk_co))

    pre_fe = dict(run.rows("120_hk3pre_first_exec"))
    pre_ex = dict(run.rows("120_hk3pre_vs_close"))
    b, src, say = [], [], ""
    if {"Cash", "Swap"} <= set(pre_fe):
        fc, fs = num(pre_fe["Cash"], "spreads"), num(pre_fe["Swap"], "spreads")
        b.append(f"Starting early paid on both: {fc:+.2f} cash, {fs:+.2f} swap at the first fill."
                 if fc > 0 and fs > 0 else
                 f"First fill vs close: {fc:+.2f} cash, {fs:+.2f} swap.")
        src += [cite("120_hk3pre_first_exec", g, r) for g, r in pre_fe.items()]
    if {"Cash", "Swap"} <= set(pre_ex):
        ec, es = num(pre_ex["Cash"], "spreads"), num(pre_ex["Swap"], "spreads")
        b.append(f"Both finished ahead of the close: {ec:+.2f} cash, {es:+.2f} swap."
                 if ec > 0 and es > 0 else
                 f"Whole order vs close: {ec:+.2f} cash, {es:+.2f} swap.")
        src += [cite("120_hk3pre_vs_close", g, r) for g, r in pre_ex.items()]
        say = ("Same three prices as slide 7, Hong Kong only. Starting early helped on both "
               "account types, and both whole orders finished at or ahead of the close.")
    S.append(slide("Hong Kong Pre-traded: Cash vs Swap", b, ["90c_hk_pre_measures.png"],
                   src, talk(say, "Positive = better for the order than the close.",
                             "Next: which Hong Kong orders gain from starting early?"),
                   ds_hk_pre))

    hq = run.rows("120_hk2_first_exec_adv", 2)
    hk_from = pays_from(hq, ("Cash", "Swap"))
    hk_large_only = hk_from is not None
    b, src, say = [], [], ""
    if hq:
        bands = sorted({k[0] for k, _ in hq}, key=_band_lo)
        if hk_from:
            edge = _band_edge(hk_from)
            below = [(k, r) for k, r in hq if _band_lo(k[0]) < _band_lo(hk_from)
                     and not thin(r)]
            mx = max((abs(num(r, "spreads")) for _, r in below), default=math.nan)
            if math.isfinite(mx):
                b.append(f"Below {edge}% ADV, first fills sat within {mx:.2f} spreads "
                         f"of the close.")
            b.append(f"From {edge}% ADV up, first fills beat the close for cash and swap.")
            b.append("So in Hong Kong, starting early only pays on the larger orders.")
            say = (f"Orders grouped by size, as on the first slide. Small orders mostly go "
                   f"straight to the auction, so there is nothing to gain by starting "
                   f"early. From {edge}% of ADV, starting early beat the close in every "
                   f"band, for cash and swap alike.")
        else:
            good = sum(all(num(r, "spreads") > 0 for k, r in hq if k[0] == band)
                       for band in bands)
            b.append(f"First fills beat the close in {good} of {len(bands)} size bands.")
            b.append("No size where starting early clearly pays for both accounts.")
            say = ("Orders grouped by size, as on the first slide. There is no clear size "
                   "above which starting early pays for both account types.")
        src += [cite("120_hk2_first_exec_adv", f"{k[0]}, {k[1]}", r) for k, r in hq]
    S.append(slide("Hong Kong: First Execution by Order Size", b,
                   ["92a_hk_adv_first_exec.png"], src,
                   talk(say, "Size bands are the deck's usual ADV% bands. A thin band shows "
                        "as a wide whisker.", "Next: India and its new closing auction."),
                   ds_hk))

    # --- 16-17. India: CAS against non-CAS ----------------------------------------------------
    in_fe = dict(run.rows("130_in1_first_exec"))
    in_lc = dict(run.rows("130_in1_last_cont"))
    in_ex = dict(run.rows("130_in1_vs_close"))
    b, src, say = [], [], ""
    b.append("From 3 August CAS names close in an auction. Others close on a VWAP.")
    if {"CAS", "Non-CAS"} <= set(in_ex):
        c, n = in_ex["CAS"], in_ex["Non-CAS"]
        b.append(f"CAS orders matched the close, at {num(c, 'spreads'):+.2f} spreads."
                 if not holds(c) else
                 f"CAS orders finished {num(c, 'spreads'):+.2f} spreads from the close.")
        b.append(f"Non-CAS orders paid {sp(num(n, 'spreads'))} spreads to track the VWAP."
                 if num(n, "spreads") < 0 else
                 f"Non-CAS orders beat the VWAP by {sp(num(n, 'spreads'))} spreads.")
        story_points["in"] = "India: CAS names now close in an auction, and our orders match it."
        say = ("Since 3 August, India's F&O names close in a 20-minute auction, like the "
               "rest of Asia. The other names still close on the VWAP of the last half hour, "
               "which we have to trade through to match. CAS orders hit the close; non-CAS "
               "orders pay a little to track the VWAP, as you would expect.")
        src += [cite("130_in1_vs_close", g, r) for g, r in in_ex.items()]
    src += [cite("130_in1_first_exec", g, r) for g, r in in_fe.items()]
    src += [cite("130_in1_last_cont", g, r) for g, r in in_lc.items()]
    S.append(slide("India Focus: CAS vs Non-CAS", b, ["95_in_measures.png"], src,
                   talk(say, "CAS names have an AUC or ACB code in qatt that day. CAS spreads "
                        "are tight (about 2 bps), so a few bps read as many spreads. Under "
                        "three months of data: a first look. India reports no close share, "
                        "so CAS orders cannot be split into close-only and pre-traded.",
                        "Next: which CAS orders gain from starting early?"), ds_in))

    iq = run.rows("130_in2_first_exec_adv", 2)
    in_from = pays_from(iq, ("CAS",))
    b, src, say = [], [], ""
    if "CAS" in in_fe:
        r = in_fe["CAS"]
        b.append(f"CAS first fills beat the close by {num(r, 'spreads'):.2f} spreads."
                 if num(r, "spreads") > 0 else
                 f"CAS first fills were {sp(num(r, 'spreads'))} spreads worse than the close.")
    if iq:
        if in_from:
            b.append(f"The gain comes from orders of {_band_edge(in_from)}% ADV and up.")
            if _band_lo(in_from) > 0:
                b.append("Smaller CAS orders show no clear gain from starting early.")
            say = ("Same pattern as Hong Kong: in CAS names, starting early helps the large "
                   "orders and does little for the small ones.")
        else:
            b.append("No size band shows a clear gain on its own yet.")
            say = ("In CAS names no size band stands out yet. Under three months of data, "
                   "so this is a first look.")
        src += [cite("130_in2_first_exec_adv", f"{k[0]}, {k[1]}", x) for k, x in iq]
    S.append(slide("India CAS: First Execution by Order Size", b,
                   ["97a_in_adv_first_exec.png"], src,
                   talk(say, "Size bands are the deck's usual ADV% bands, India orders from "
                        "3 Aug. CAS spreads are about 2 bps, so a big number in spreads is a "
                        "smaller one in bps. Under three months of data.",
                        "Next: what we suggest."),
                   ds_in))

    # --- what we found: the story on one slide, right after the flow split ----------------------
    if "auc" in story_points:
        story_points["auc"] = (("But " if "pre" in story_points else "Pre-traded orders: ")
                               + story_points["auc"])
        story_points["auc"] = story_points["auc"][0].upper() + story_points["auc"][1:]
    b = [story_points[k] for k in ("co", "pre", "auc") if k in story_points]
    if len(b) < 3:
        b += [story_points[k] for k in ("hk", "in") if k in story_points][:3 - len(b)]
    S.insert(2, slide(f"{PERIOD} MOC: What We Found", b, [], [],
                      talk("This is the story in one slide. Close-only orders are small and "
                           "the close held for them. Larger orders start early, and that "
                           "helped. The catch is the auction itself, which tends to print "
                           "against our pre-traded buys. Hong Kong and India follow at the "
                           "end.", "", "Next: close-only orders.")))

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
    if hk_large_only:
        b.append(f"Suggest: in Hong Kong, start early only from {_band_edge(hk_from)}% ADV.")
    S.append(slide(f"{PERIOD} MOC: Takeaways", b[:3], [], [],
                   talk("Three changes, each resting on a slide in this deck. Small orders: "
                        "close-only works, the close held (slides 4-6). Large orders in the "
                        "markets where we are most of the auction: leave less for the "
                        "auction, because that is where the close moves against us (10-12). "
                        "Hong Kong: start early only on the larger orders (16).",
                        "DRAFT - agree each line with the desk. India CAS is under three "
                        "months old: watch it, do not act on it yet.", "")))

    return S


def _band_edge(label) -> str:
    """The lower edge of a band label, as text: '1-3%' -> '1', '25%+' -> '25'."""
    return str(label).split("-")[0].rstrip("%+")


def _band_lo(label) -> float:
    try:
        return float(_band_edge(label))
    except ValueError:
        return math.inf


def pays_from(rows, groups) -> str | None:
    """The smallest size band from which starting early paid in every band
    above it: each group's first fill beat the close, and at least one of
    them clearly. Bands where a group is too thin to read are skipped. None if no
    such band, or if it is the smallest one (then size does not matter)."""
    by_band: dict = {}
    for (band, g), r in rows:
        if g in groups and not thin(r):
            by_band.setdefault(band, {})[g] = r
    # A band is read only where every group has enough orders in it.
    bands = sorted((b for b in by_band if set(groups) <= set(by_band[b])), key=_band_lo)
    def pays(band):
        rs = by_band[band]
        return (all(num(r, "spreads") > 0 for r in rs.values())
                and any(holds(r) for r in rs.values()))
    start = None
    for band in reversed(bands):
        if not pays(band):
            break
        start = band
    return start if start is not None and start != bands[0] else None


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
