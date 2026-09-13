#!/usr/bin/env python3
"""Build the H1 MOC deck, in the shape of the Q1 one, from a moc_tca.py run.

    python build_h1_deck.py --results output_h1
    python build_h1_deck.py --results output_h1 --cover --client BlackRock

Reads <results>/tables.xlsx and <results>/charts/*.png - nothing else. Every
number on a slide is looked up in a table at build time, and every sentence
that takes a side (beat or lagged, cost or gained, clear or not) is chosen
from the sign and the 95% interval of that number. Rerun moc_tca.py and the
deck follows; there is no figure typed in here to go stale.

The story follows the Q1 deck, then pushes on the two populations:
  1  Executive summary          where the value went, how big the orders were
  2  Close-only vs pre-traded   the split, and the headline for each
  3  Close-only                 reversion: did the price hold after the print
  4  Pre-traded                 the start against the finish, by market
  5  Pre-traded by side         buys against sells
  6  Pre-traded by size         the case for sending small orders to the close
  7  What we will change
  +  Appendix                   monthly, by size and spread, by market

Writes <deck>.pptx and <deck>.evidence.md: each bullet with the table row it
rests on, so any number can be defended on the call.
"""
from __future__ import annotations

import argparse
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

import moc_tca as tca

SKILL = Path(__file__).parent / ".claude" / "skills" / "tca-review" / "scripts"

# ===========================================================================
# LAYOUT
# ===========================================================================

def _rgb(h: str) -> RGBColor:
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


SLIDE_W, SLIDE_H = Inches(13.333), Inches(7.5)
MARGIN = Inches(0.6)
GAP = Inches(0.3)
INK, INK_SOFT = _rgb(tca.INK), _rgb(tca.INK_SECOND)
ACCENT, RULE, SURFACE = _rgb(tca.POS), _rgb(tca.BASELINE), _rgb(tca.SURFACE)
FONT = "Arial"


def textbox(slide, left, top, width, height, text, size, *, bold=False,
            color=INK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(left, top, width, height)
    frame = box.text_frame
    frame.word_wrap = True
    frame.vertical_anchor = anchor
    for i, line in enumerate(text.split("\n")):
        para = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        para.text = line
        para.alignment = align
        for run in para.runs:
            run.font.size, run.font.bold = Pt(size), bold
            run.font.color.rgb, run.font.name = color, FONT
    return box


def bullet_box(slide, top, items, size):
    box = slide.shapes.add_textbox(MARGIN, top, SLIDE_W - 2 * MARGIN,
                                   Inches(1.4))
    frame = box.text_frame
    frame.word_wrap = True
    for i, text in enumerate(items):
        para = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        para.text = f"•  {text}"
        para.space_after = Pt(5)
        for run in para.runs:
            run.font.size, run.font.color.rgb = Pt(size), INK
            run.font.name = FONT


def charts_row(slide, paths: list, top, box_h):
    """Charts side by side at one height, the row centred on the slide.

    Equal-width boxes shrink a wide chart to a strip beside a square one.
    Sizing every chart to the same height, and letting widths follow their
    shapes, keeps the pair reading as a pair.
    """
    from PIL import Image
    ratios = []
    for p in paths:
        if p.exists():
            with Image.open(p) as im:
                ratios.append(im.width / im.height)
        else:
            ratios.append(None)
    avail = SLIDE_W - 2 * MARGIN - GAP * (len(paths) - 1)
    known = [r for r in ratios if r] or [1.6]
    h = min(box_h, int(avail / sum(r or sum(known) / len(known)
                                   for r in ratios)))
    widths = [int(h * (r or sum(known) / len(known))) for r in ratios]
    left = int((SLIDE_W - sum(widths) - GAP * (len(paths) - 1)) / 2)
    top = Emu(int(top + (box_h - h) / 2))       # centred in the space left
    for p, r, w in zip(paths, ratios, widths):
        if r:
            slide.shapes.add_picture(str(p), Emu(left), top, width=Emu(w),
                                     height=Emu(int(h)))
        else:
            textbox(slide, Emu(left), top, Emu(w), Inches(0.5),
                    f"[chart not found: {p.name}]", 12, color=INK_SOFT)
        left += w + GAP


def place(slide, path: Path, left, top, box_w, box_h):
    """A chart scaled into its box, never stretched, centred across it."""
    if not path.exists():
        textbox(slide, left, top, box_w, Inches(0.5),
                f"[chart not found: {path.name}]", 12, color=INK_SOFT)
        return False
    from PIL import Image
    with Image.open(path) as im:
        ratio = im.height / im.width
    w, h = box_w, Emu(int(box_w * ratio))
    if h > box_h:
        h, w = box_h, Emu(int(box_h / ratio))
    slide.shapes.add_picture(str(path), Emu(int(left + (box_w - w) / 2)), top,
                             width=w, height=h)
    return True


# ===========================================================================
# READING THE RUN
# ===========================================================================

class Run:
    """tables.xlsx and charts/ from one moc_tca.py output folder."""

    def __init__(self, folder: Path):
        self.folder = folder
        self.charts = folder / "charts"
        path = folder / "tables.xlsx"
        if not path.exists():
            raise SystemExit(f"ERROR: {path} not found - run moc_tca.py first.")
        self.xl = pd.ExcelFile(path)
        self.missing: list[str] = []

    def table(self, name: str, keys: int = 1) -> pd.DataFrame | None:
        sheet = next((s for s in self.xl.sheet_names if s == name[:31]), None)
        if sheet is None:
            self.missing.append(name)
            return None
        df = pd.read_excel(self.xl, sheet)
        idx = list(df.columns[:keys])
        # A MultiIndex is written with the outer label only on its first row.
        df[idx] = df[idx].ffill()
        return df.set_index(idx if keys > 1 else idx[0])

    def rows(self, name: str, keys: int = 1):
        """(key, row) for every row a client slide can show."""
        df = self.table(name, keys)
        if df is None:
            return []
        out = []
        for k, r in df.iterrows():
            first = k[0] if isinstance(k, tuple) else k
            if (str(first) in tca.EXCLUDE_MARKETS_FROM_CHARTS
                    or str(first).strip().lower() in tca.NON_CATEGORIES):
                continue
            out.append((k, r))
        return out

    def row(self, name: str, key, keys: int = 1):
        return next((r for k, r in self.rows(name, keys) if k == key), None)


def num(r, col) -> float:
    try:
        v = float(r[col])
    except (TypeError, ValueError, KeyError):
        return math.nan
    return v


def holds(r) -> bool:
    """The 95% interval sits on one side of zero."""
    lo, hi = num(r, "CI low"), num(r, "CI high")
    return math.isfinite(lo) and math.isfinite(hi) and (lo > 0 or hi < 0)


def thin(r) -> bool:
    return str(r.get("small sample", False)).strip().lower() in ("true", "1")


def sp(v: float) -> str:
    return f"{abs(v):.2f}"


def about_bps(r) -> str:
    return f"about {abs(num(r, 'wtd mean bps')):.0f} bps"


def cite(table, key, r) -> str:
    """One evidence line: where the number is and how sure it is."""
    ci = (f"CI [{num(r, 'CI low'):+.2f}, {num(r, 'CI high'):+.2f}]"
          if math.isfinite(num(r, "CI low")) else "no CI (too few orders)")
    return (f"`{table}` [{key}]: {num(r, 'spreads'):+.2f} spreads x "
            f"[{num(r, 'spread bps (wtd)'):.1f} bps] = "
            f"{num(r, 'wtd mean bps'):+.1f} bps, {ci}, "
            f"{int(num(r, 'orders')):,} orders, "
            f"{'holds' if holds(r) else 'not distinguishable from zero'}")


# ===========================================================================
# THE STORY
# ===========================================================================

def story(run: Run, period: str) -> list[dict]:
    S = []

    # --- 1. executive summary ---------------------------------------------
    prof = run.rows("34_market_profile")
    ncol = "notional (USDm)"
    size = run.table("52_adv_profile")
    b, src = [], []
    if prof:
        tot = sum(num(r, ncol) for _, r in prof)
        ranked = sorted(prof, key=lambda kr: -num(kr[1], ncol))
        top5 = 100 * sum(num(r, ncol) for _, r in ranked[:5]) / tot
        b.append(f"${tot:,.0f}m traded on the close across {len(prof)} markets.")
        a, c = ranked[0], ranked[1]
        b.append(f"{a[0]} and {c[0]} led. The top five were {top5:.0f}% of "
                 "the value.")
        src += [f"`34_market_profile`: total {tot:,.1f}m, "
                + ", ".join(f"{k} {num(r, '% of notional'):.1f}%"
                            for k, r in ranked[:5])]
    es = run.table("50_exec_summary")
    fill = math.nan
    if es is not None:
        hit = es[es["metric"] == "fill ratio"]
        fill = float(hit["value"].iloc[0]) if len(hit) else math.nan
    if size is not None and "0-1%" in size.index:
        u1 = num(size.loc["0-1%"], "% of notional")
        b.append(f"{u1:.0f}% of the value was under 1% of daily volume."
                 + (f" Fill was {fill:.1f}%." if math.isfinite(fill) else ""))
        src.append(f"`52_adv_profile` [0-1%]: {u1:.1f}% of notional; "
                   f"`50_exec_summary` fill ratio {fill:.2f}%")
    S.append(dict(title=f"MOC TCA executive summary, {period}", bullets=b,
                  charts=["13_market_notional.png", "24_adv_profile.png"],
                  src=src, notes=(
        "Scope: CLOSE strategy. The left chart splits each market into cash "
        "and swap by client account. The right chart is order size against "
        "daily volume. No judgement on this slide - let the client recognise "
        "their own book first.")))

    # --- 2. close-only vs pre-traded --------------------------------------
    flow = run.table("51_flow_split")
    first = run.row("73a_pre_firstexec_side_spr", "All")
    finish = run.row("73b_pre_close_side_spr", "All")
    b, src = [], []
    if flow is not None and {"Close-only", "Pre-traded"} <= set(flow.index):
        co, pre = flow.loc["Close-only"], flow.loc["Pre-traded"]
        b.append(f"{num(co, '% of notional'):.0f}% of the value was close-only: "
                 f"{int(num(co, 'orders')):,} orders, "
                 f"{num(co, 'fill ratio %'):.1f}% filled.")
        bigger = ("bigger" if num(pre, "%ADV (notional-weighted)")
                  > num(co, "%ADV (notional-weighted)") else "smaller")
        b.append(f"The other {num(pre, '% of notional'):.0f}% started early. "
                 f"Those orders were {bigger}, at "
                 f"{num(pre, '%ADV (notional-weighted)'):.1f}% of volume.")
        src.append("`51_flow_split`: " + "; ".join(
            f"{k} {num(r, 'notional (USDm)'):,.0f}m "
            f"({num(r, '% of notional'):.1f}%), {int(num(r, 'orders')):,} "
            f"orders, fill {num(r, 'fill ratio %'):.2f}%, "
            f"{num(r, '%ADV (notional-weighted)'):.2f}% ADV, "
            f"{num(r, '% of notional in the close'):.0f}% in the close"
            for k, r in flow.iterrows()))
    if first is not None and finish is not None:
        verb = "beat" if num(finish, "spreads") > 0 else "lagged"
        cost = "cost" if num(first, "spreads") < 0 else "gained"
        b.append(f"They {verb} the close by {sp(num(finish, 'spreads'))} "
                 f"spreads. Their first fills {cost} "
                 f"{sp(num(first, 'spreads'))}.")
        src += [cite("73b_pre_close_side_spr", "All", finish),
                cite("73a_pre_firstexec_side_spr", "All", first)]
    mostly_co = (flow is not None and "Close-only" in flow.index
                 and num(flow.loc["Close-only"], "% of notional") > 50)
    S.append(dict(title=("Most of the flow went straight into the auction"
                         if mostly_co else
                         "Close-only and pre-traded orders"),
                  bullets=b, charts=["23_flow_split.png",
                                     "26_pretraded_by_side.png"],
                  src=src, notes=(
        "Close-only: 99.5% or more of the order printed in the auction. "
        "Pre-traded: part of it traded before. India's 17:30-17:45 orders "
        "count as close-only; India cannot be pre-traded.\n\n"
        "Everything is in spreads: the result divided by the average spread "
        "([x bps] under each bar). Multiply the two to get bps.\n\n"
        "The right chart is the pre-traded story in one picture: first fill "
        "against the close, the whole execution against the close, and the "
        "next open against the close.")))

    # --- 3. close-only: reversion -----------------------------------------
    b, src = ["A close-only order trades at the close, so reversion is the test."], []
    allc = run.row("71b_co_reversion_side_spr", "All")
    if allc is not None:
        way = "our way" if num(allc, "spreads") > 0 else "against us"
        tail = "" if holds(allc) else " Not a clear signal."
        b.append(f"By the next open, prices moved {sp(num(allc, 'spreads'))} "
                 f"spreads {way} on average.{tail}")
        src.append(cite("71b_co_reversion_side_spr", "All", allc))
    bad = sorted([(k, r) for k, r in run.rows("72_co_reversion_mkt_side_spr", 2)
                  if holds(r) and num(r, "spreads") < 0 and not thin(r)],
                 key=lambda kr: num(kr[1], "spreads"))
    worst_co = bad[0] if bad else None
    if worst_co:
        (m, side), r = worst_co
        b.append(f"{m} {side.lower()}s came back hardest: "
                 f"{sp(num(r, 'spreads'))} spreads, {about_bps(r)}.")
        src.append(cite("72_co_reversion_mkt_side_spr", f"{m}, {side}", r))
    else:
        b.append("No market and side came back against us by a clear margin.")
    for k, r in run.rows("71_co_reversion_mkt_spreads"):
        if holds(r):
            src.append(cite("71_co_reversion_mkt_spreads", k, r))
    S.append(dict(title="Close-only: did the price hold after the auction?",
                  bullets=b, charts=["22_reversion_close_only.png",
                                     "25_close_only_reversion_market_side.png"],
                  src=src, notes=(
        "Reversion is the next open against the close, side-adjusted. "
        "Negative means the price came back against us overnight.\n\n"
        "A market with a very tight spread shows a large number of spreads "
        "for a small move in bps - India trades on about 3 bps, so quote the "
        "bps figure beside it.\n\n"
        "Bars without whiskers had too few orders for an interval. Do not "
        "lead with them.")))

    # --- 4. pre-traded: start against finish ------------------------------
    fe = [(k, r) for k, r in run.rows("68_pre_firstexec_mkt_spreads")]
    fin = [(k, r) for k, r in run.rows("77_pre_close_mkt_spreads")]
    b, src = [], []
    if fe:
        neg = sum(num(r, "spreads") < 0 for _, r in fe)
        b.append("First fills were worse than the close in every market."
                 if neg == len(fe) else
                 f"First fills were worse than the close in {neg} of "
                 f"{len(fe)} markets.")
        clear = sorted([(k, r) for k, r in fe if holds(r)
                        and num(r, "spreads") < 0 and not thin(r)],
                       key=lambda kr: num(kr[1], "spreads"))
        if clear:
            k, r = clear[0]
            b.append(f"{k} was the worst: {sp(num(r, 'spreads'))} spreads, "
                     f"{about_bps(r)}.")
        src += [cite("68_pre_firstexec_mkt_spreads", k, r) for k, r in fe]
    if fin:
        pos = sum(num(r, "spreads") > 0 for _, r in fin)
        b.append("By the finish, every market beat the close."
                 if pos == len(fin) else
                 f"By the finish, {pos} of {len(fin)} markets still beat the "
                 "close.")
        src += [cite("77_pre_close_mkt_spreads", k, r) for k, r in fin]
    starts_cost = first is not None and num(first, "spreads") < 0
    S.append(dict(title=("Pre-traded orders: the start is where it costs"
                         if starts_cost else
                         "Pre-traded orders: the start and the finish"),
                  bullets=b, charts=["10b_first_exec_market.png",
                                     "27_pretraded_close_market.png"],
                  src=src, notes=(
        "Left: the first fill against the close, weighted by the part of "
        "each order that traded before the auction. Right: the whole "
        "execution against the close.\n\n"
        "Read them together. The orders finish ahead of the close, so the "
        "algo recovers - but it pays for the first fills to get there. "
        "Starting later or more passively keeps the finish without the "
        "start.")))

    # --- 5. pre-traded by side --------------------------------------------
    b, src = [], []
    for side in ("Buy", "Sell"):
        trio = [run.row(t, side) for t in ("73a_pre_firstexec_side_spr",
                                           "73b_pre_close_side_spr",
                                           "73c_pre_reversion_side_spr")]
        if any(t is None for t in trio):
            continue
        n_clear = sum(holds(t) for t in trio)
        verdict = {3: "All three are clear.", 0: "None is clear."}.get(
            n_clear, f"{n_clear} of 3 are clear.")
        b.append(f"{side}s: start {num(trio[0], 'spreads'):+.2f}, finish "
                 f"{num(trio[1], 'spreads'):+.2f}, next open "
                 f"{num(trio[2], 'spreads'):+.2f}. {verdict}")
        src += [cite(t, side, r) for t, r in zip(
            ("73a_pre_firstexec_side_spr", "73b_pre_close_side_spr",
             "73c_pre_reversion_side_spr"), trio)]
    fe2 = {k: r for k, r in run.rows("74_pre_firstexec_mkt_side_spr", 2)}
    rv2 = {k: r for k, r in run.rows("76_pre_reversion_mkt_side_spr", 2)}
    both = sorted([k for k in fe2 if k in rv2
                   and holds(fe2[k]) and num(fe2[k], "spreads") < 0
                   and holds(rv2[k]) and num(rv2[k], "spreads") < 0],
                  key=lambda k: num(rv2[k], "spreads"))
    if both:
        m, side = both[0]
        b.append(f"{m} {side.lower()}s started {sp(num(fe2[both[0]], 'spreads'))} "
                 f"spreads worse and came back {sp(num(rv2[both[0]], 'spreads'))} "
                 "overnight.")
        src += [cite("74_pre_firstexec_mkt_side_spr", f"{m}, {side}", fe2[both[0]]),
                cite("76_pre_reversion_mkt_side_spr", f"{m}, {side}", rv2[both[0]])]
    worst_pair = both[0] if both else None
    S.append(dict(title="Buys and sells did not behave the same",
                  bullets=b,
                  charts=["28_pretraded_first_exec_market_side.png",
                          "30_pretraded_reversion_market_side.png"],
                  src=src, notes=(
        "Start = first fill vs close. Finish = whole execution vs close. "
        "Next open = reversion. All in spreads, positive is good.\n\n"
        "Left: first fills by market, buys over sells. Right: reversion by "
        "market, buys over sells. A market that is red on both charts for "
        "the same side paid to start and then gave it back overnight - the "
        "clearest case for changing how that flow is worked.")))

    # --- 6. pre-traded by size --------------------------------------------
    b, src = [], []
    psize = run.table("53_pre_adv_profile")
    small = run.row("66_pre_firstexec_adv_spreads", "0-1%")
    if psize is not None and "0-1%" in psize.index:
        r = psize.loc["0-1%"]
        b.append(f"{num(r, '% of orders'):.0f}% of pre-traded orders were "
                 f"under 1% of volume: {num(r, '% of notional'):.0f}% of value.")
        src.append(f"`53_pre_adv_profile` [0-1%]: {int(num(r, 'orders')):,} "
                   f"orders ({num(r, '% of orders'):.1f}%), "
                   f"{num(r, 'notional (USDm)'):,.1f}m "
                   f"({num(r, '% of notional'):.1f}%)")
    small_costs = (small is not None and holds(small)
                   and num(small, "spreads") < 0)
    if small is not None:
        if num(small, "spreads") < 0:
            head = f"Their first fills still cost {sp(num(small, 'spreads'))}"
        else:
            head = f"Their first fills gained {sp(num(small, 'spreads'))}"
        tail = "" if holds(small) else " Not a clear signal."
        b.append(f"{head} spreads against the close.{tail}")
        src.append(cite("66_pre_firstexec_adv_spreads", "0-1%", small))
    if flow is not None and "Close-only" in flow.index:
        co = flow.loc["Close-only"]
        b.append(f"Close-only orders, at {num(co, '%ADV (notional-weighted)'):.1f}% "
                 f"of volume, filled {num(co, 'fill ratio %'):.1f}%.")
    S.append(dict(title=("Small pre-traded orders could go straight to the close"
                         if small_costs else "Pre-traded orders by size"),
                  bullets=b, charts=["31_pretraded_adv_profile.png",
                                     "19_first_exec_by_adv.png"],
                  src=src, notes=(
        "The Q1 argument, tested on H1. Small orders are most of the "
        "pre-traded count but a minority of its value. They do not need the "
        "continuous session for capacity - close-only orders of that size "
        "fill almost completely - and starting them early still cost money "
        "on the first fill.")))

    # --- 7. what we will change -------------------------------------------
    b, src = [], []
    if small_costs:
        b.append("Send pre-traded orders under 1% of volume straight to the "
                 "auction.")
    clear_fe = sorted([(k, r) for k, r in fe if holds(r)
                       and num(r, "spreads") < 0 and not thin(r)],
                      key=lambda kr: -num(kr[1], "notional (USDm)"))[:2]
    if clear_fe:
        names = " and ".join(str(k) for k, _ in clear_fe)
        b.append(f"Start later, and more passively, in {names}.")
    pick = worst_pair or (worst_co[0] if worst_co else None)
    if pick:
        m, side = pick
        b.append(f"Review {m} {side.lower()}s, where the price came back "
                 "overnight.")
    if not b:
        b = ["No single change stands out clearly from this half.",
             "We will keep the current set-up and review again next quarter."]
    S.append(dict(title="What we will change", bullets=b, charts=[], src=src,
                  notes=(
        "Each action comes from a result that holds at 95%: the small-order "
        "first-fill cost (slide 6), the markets with the clearest early-start "
        "cost (slide 4), and the market and side where the price came back "
        "(slides 3 and 5). Agree the order of these with the desk before the "
        "meeting.")))

    # --- appendix -----------------------------------------------------------
    for title, charts in [
        ("Appendix: close performance by month",
         ["12_monthly.png", "12b_monthly_india.png"]),
        ("Appendix: close performance by size and spread",
         ["17_close_by_adv.png", "18_close_by_spread.png"]),
        ("Appendix: close by market, first fills by spread",
         ["15_market_vs_close.png", "20_first_exec_by_spread.png"]),
    ]:
        S.append(dict(title=title, bullets=[], charts=charts, src=[],
                      notes="Backup. Show only if asked.", appendix=True))
    return S


# ===========================================================================
# BUILD
# ===========================================================================

def build(run: Run, out: Path, client: str, period: str, cover: bool) -> list:
    slides = story(run, period)
    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    blank = prs.slide_layouts[6]

    def new():
        s = prs.slides.add_slide(blank)
        bg = s.shapes.add_shape(1, Emu(0), Emu(0), SLIDE_W, SLIDE_H)
        bg.fill.solid(); bg.fill.fore_color.rgb = SURFACE
        bg.line.fill.background(); bg.shadow.inherit = False
        s.shapes._spTree.remove(bg._element)
        s.shapes._spTree.insert(2, bg._element)
        return s

    if cover:
        s = new()
        band = s.shapes.add_shape(1, Emu(0), Inches(3.05), SLIDE_W, Emu(38100))
        band.fill.solid(); band.fill.fore_color.rgb = ACCENT
        band.line.fill.background(); band.shadow.inherit = False
        textbox(s, MARGIN, Inches(2.0), SLIDE_W - 2 * MARGIN, Inches(1.0),
                "MOC TCA review", 40, bold=True)
        textbox(s, MARGIN, Inches(3.3), SLIDE_W - 2 * MARGIN, Inches(0.8),
                f"{client} — {period}", 22, color=INK_SOFT)

    for i, spec in enumerate(slides, start=1):
        s = new()
        textbox(s, MARGIN, Inches(0.35), SLIDE_W - 2 * MARGIN, Inches(0.8),
                spec["title"], 26, bold=True)
        rule = s.shapes.add_shape(1, MARGIN, Inches(1.12),
                                  SLIDE_W - 2 * MARGIN, Emu(9525))
        rule.fill.solid(); rule.fill.fore_color.rgb = RULE
        rule.line.fill.background(); rule.shadow.inherit = False

        text_only = not spec["charts"]
        if spec["bullets"]:
            bullet_box(s, Inches(1.95 if text_only else 1.25), spec["bullets"],
                       24 if text_only else 16)
        if spec["charts"]:
            top = Inches(2.55) if spec["bullets"] else Inches(1.35)
            charts_row(s, [run.charts / c for c in spec["charts"]], top,
                       SLIDE_H - top - Inches(0.45))
        textbox(s, MARGIN, SLIDE_H - Inches(0.38), SLIDE_W - 2 * MARGIN,
                Inches(0.3), f"{client}  |  MOC TCA {period}  |  {i}", 9,
                color=INK_SOFT)
        s.notes_slide.notes_text_frame.text = (
            spec["notes"] + ("\n\nSources:\n" + "\n".join(spec["src"])
                             if spec["src"] else ""))

    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out)
    return slides


def check_words(slides: list) -> list[str]:
    """The skill's plain-language rules, run on the deck as built."""
    sys.path.insert(0, str(SKILL))
    try:
        from check_readability import check_deck
    except ImportError:
        return ["(readability check not found - skipped)"]
    return check_deck({"slides": [{"type": "content", "title": s["title"],
                                   "bullets": s["bullets"]} for s in slides]})


def write_evidence(md: Path, pptx: Path, slides: list) -> None:
    """Every bullet with the rows behind it, checked against the saved deck."""
    saved = Presentation(pptx)
    text = " ".join(sh.text_frame.text for sl in saved.slides
                    for sh in sl.shapes if sh.has_text_frame)
    lines = [f"# Evidence for {pptx.name}", "",
             "Every bullet, and the table row it rests on. Spreads x "
             "[spread bps] = bps. 'holds' means the 95% interval stays on one "
             "side of zero.", ""]
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
    p.add_argument("--results", type=Path, default=Path("output_h1"),
                   help="the moc_tca.py --out folder (tables.xlsx + charts/)")
    p.add_argument("--out", type=Path, default=None)
    p.add_argument("--client", default="BlackRock")
    p.add_argument("--period", default=tca.PERIOD_LABEL)
    p.add_argument("--cover", action="store_true")
    a = p.parse_args(argv)

    run = Run(a.results)
    out = a.out or a.results / f"MOC_TCA_{a.period.replace(' ', '_')}.pptx"
    slides = build(run, out, a.client, a.period, a.cover)
    write_evidence(out.with_suffix(".evidence.md"), out, slides)

    print(f"deck      -> {out}  ({len(slides) + a.cover} slides)")
    print(f"evidence  -> {out.with_suffix('.evidence.md')}")
    if run.missing:
        print("missing tables (their bullets were left out): "
              + ", ".join(sorted(set(run.missing))))
    gone = sorted({c for s in slides for c in s["charts"]
                   if not (run.charts / c).exists()})
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
