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
  6  Pre-traded by ADV%         the case for sending small orders to the close
  7  What we will change
  +  Appendix                   monthly, by ADV% and spread, by market

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
MARGIN = Inches(0.4)
GAP = Inches(0.25)
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


CAPTION_H = Inches(0.34)


def charts_row(slide, paths: list, top, box_h, titles: dict | None = None):
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
    titles = titles or {}
    captions = any(titles.get(p.name, {}).get("title") for p in paths)
    if captions:
        box_h = box_h - CAPTION_H
    avail = SLIDE_W - 2 * MARGIN - GAP * (len(paths) - 1)
    known = [r for r in ratios if r] or [1.6]
    h = min(box_h, int(avail / sum(r or sum(known) / len(known)
                                   for r in ratios)))
    widths = [int(h * (r or sum(known) / len(known))) for r in ratios]
    left = int((SLIDE_W - sum(widths) - GAP * (len(paths) - 1)) / 2)
    top = Emu(int(top + (box_h - h) / 2))       # centred in the space left
    for p, r, w in zip(paths, ratios, widths):
        cap = titles.get(p.name, {}).get("title", "")
        if captions:
            textbox(slide, Emu(left), top, Emu(max(w, Inches(3))), CAPTION_H,
                    cap, 14, bold=True, color=INK)
        img_top = Emu(int(top + (CAPTION_H if captions else 0)))
        if r:
            slide.shapes.add_picture(str(p), Emu(left), img_top, width=Emu(w),
                                     height=Emu(int(h)))
        else:
            textbox(slide, Emu(left), img_top, Emu(w), Inches(0.5),
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
        # The bare deck copies when the run wrote them, with their titles.
        bare = folder / tca.DECK_CHARTS_DIR
        self.charts = bare if (bare / "titles.json").exists() else folder / "charts"
        self.titles = {}
        if self.charts == bare:
            import json
            self.titles = json.loads((bare / "titles.json").read_text(
                encoding="utf-8"))
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
    flow = run.table("51_flow_split")
    b, src = [], []
    if prof:
        tot = sum(num(r, ncol) for _, r in prof)
        ranked = sorted(prof, key=lambda kr: -num(kr[1], ncol))
        top5 = 100 * sum(num(r, ncol) for _, r in ranked[:5]) / tot
        b.append(f"${tot:,.0f}m traded on the close across {len(prof)} markets.")
        a_, c_ = ranked[0], ranked[1]
        b.append(f"{a_[0]} and {c_[0]} led. The top five were {top5:.0f}% of "
                 "the value.")
        src.append(f"`34_market_profile`: total {tot:,.1f}m, "
                   + ", ".join(f"{k} {num(r, '% of notional'):.1f}%"
                               for k, r in ranked[:5]))
    es = run.table("50_exec_summary")
    fill = math.nan
    if es is not None:
        hit = es[es["metric"] == "fill ratio"]
        fill = float(hit["value"].iloc[0]) if len(hit) else math.nan
    if size is not None and "0-1%" in size.index:
        u1 = num(size.loc["0-1%"], "% of notional")
        b.append(f"{u1:.0f}% of the value was under 1% ADV."
                 + (f" Fill was {fill:.1f}%." if math.isfinite(fill) else ""))
        src.append(f"`52_adv_profile` [0-1%]: {u1:.1f}% of notional; "
                   f"`50_exec_summary` fill ratio {fill:.2f}%")
    S.append(dict(title=f"MOC TCA executive summary, {period}", bullets=b,
                  charts=["13_market_notional.png", "24_adv_profile.png"],
                  src=src, notes=(
        "Scope: CLOSE strategy. The left chart splits each market into cash "
        "and swap by client account. The right chart is value by ADV%. "
        "No judgement on this slide - let the client recognise their own "
        "book first.")))

    # --- 1b. close-only against pre-traded ----------------------------------
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
                 f"{num(pre, '%ADV (notional-weighted)'):.1f}% ADV.")
        src.append("`51_flow_split`: " + "; ".join(
            f"{k} {num(r, 'notional (USDm)'):,.0f}m "
            f"({num(r, '% of notional'):.1f}%), {int(num(r, 'orders')):,} "
            f"orders, fill {num(r, 'fill ratio %'):.2f}%, "
            f"{num(r, '%ADV (notional-weighted)'):.2f}% ADV, "
            f"{num(r, '% of notional in the close'):.0f}% in the close"
            for k, r in flow.iterrows()))
        # Every order here printed in the close, so this is not "how many
        # reached the auction" - it is how the pre-traded value divides
        # between the auction and the session before it.
        in_close = num(pre, "% of notional in the close")
        b.append(f"Pre-traded orders executed {in_close:.0f}% of their value in "
                 f"the close, {100 - in_close:.0f}% before.")
    mostly_co = (flow is not None and "Close-only" in flow.index
                 and num(flow.loc["Close-only"], "% of notional") > 50)
    S.append(dict(title=("Most of the flow went straight into the auction"
                         if mostly_co else "Close-only and pre-traded orders"),
                  bullets=b, charts=["23_flow_split.png"],
                  src=src, notes=(
        "Close-only: 99.5% or more of the order printed in the auction. "
        "Pre-traded: part of it traded before. India's 17:30-17:45 orders "
        "count as close-only; India cannot be pre-traded.\n\n"
        "Everything after this slide is in spreads: the result divided by the "
        "average spread, the [x bps] under each bar. Multiply the two to get "
        "bps.")))

    # --- 2. close-only: reversion by market -------------------------------
    b, src = ["A close-only order trades at the close, so reversion is the test."], []
    allc = run.row("71b_co_reversion_side_spr", "All")
    if allc is not None:
        way = "our way" if num(allc, "spreads") > 0 else "against us"
        tail = "" if holds(allc) else " Not a clear signal."
        b.append(f"Overnight, prices moved {sp(num(allc, 'spreads'))} "
                 f"spreads {way}.{tail}")
        src.append(cite("71b_co_reversion_side_spr", "All", allc))
    co_mkt = run.rows("71_co_reversion_mkt_spreads")
    bad_m = sorted([(k, r) for k, r in co_mkt if holds(r)
                    and num(r, "spreads") < 0 and not thin(r)],
                   key=lambda kr: num(kr[1], "spreads"))
    good_m = [k for k, r in co_mkt if holds(r) and num(r, "spreads") > 0]
    if bad_m:
        names = " and ".join(str(k) for k, _ in bad_m[:2])
        b.append(f"The price came back clearly only in {names}.")
    elif good_m:
        b.append("No market came back against us by a clear margin.")
    src += [cite("71_co_reversion_mkt_spreads", k, r) for k, r in co_mkt]
    S.append(dict(title="Close-only: did the price hold after the auction?",
                  bullets=b, charts=["22_reversion_close_only.png"],
                  src=src, notes=(
        "Reversion is the next open against the close, side-adjusted. "
        "Negative means the price came back against us overnight.\n\n"
        "A market with a very tight spread shows a large number of spreads "
        "for a small move in bps - India trades on about 3 bps, so quote the "
        "bps figure beside it.\n\n"
        "Bars without whiskers had too few orders for an interval. Do not "
        "lead with them.")))

    # --- 3. close-only: by side -------------------------------------------
    b, src = [], []
    co_side = run.rows("72_co_reversion_mkt_side_spr", 2)
    bad = sorted([(k, r) for k, r in co_side
                  if holds(r) and num(r, "spreads") < 0 and not thin(r)],
                 key=lambda kr: num(kr[1], "spreads"))
    worst_co = bad[0] if bad else None
    if worst_co:
        (m, side), r = worst_co
        b.append(f"{m} {side.lower()}s came back hardest: "
                 f"{sp(num(r, 'spreads'))} spreads, {about_bps(r)}.")
        opp = next((rr for kk, rr in co_side
                    if kk[0] == m and kk[1] != side), None)
        if opp is not None:
            way = "held" if num(opp, "spreads") >= 0 else "also came back"
            b.append(f"{m} {'buys' if side == 'Sell' else 'sells'} {way}: "
                     f"{num(opp, 'spreads'):+.2f} spreads.")
            src.append(cite("72_co_reversion_mkt_side_spr",
                            f"{m}, {'Buy' if side == 'Sell' else 'Sell'}", opp))
        src.append(cite("72_co_reversion_mkt_side_spr", f"{m}, {side}", r))
        if len(bad) > 1:
            b.append(f"{len(bad) - 1} other market and side pairs came back "
                     "by a clear margin.")
        else:
            b.append("No other market and side came back by a clear margin.")
        src += [cite("72_co_reversion_mkt_side_spr", f"{k[0]}, {k[1]}", rr)
                for k, rr in bad[1:]]
    else:
        b.append("Nothing came back against us by a clear margin, on either side.")
    S.append(dict(title="Close-only reversion, buys against sells",
                  bullets=b,
                  charts=["25_close_only_reversion_market_side.png"],
                  src=src, notes=(
        "Same measure as the previous slide, split by side. A side that comes "
        "back while the other holds points at how that side is sent, not at "
        "the market.")))

    # --- 4. pre-traded: start, finish, next open --------------------------
    first = run.row("73a_pre_firstexec_side_spr", "All")
    finish = run.row("73b_pre_close_side_spr", "All")
    b, src = [], []
    if first is not None and finish is not None:
        verb = "beat" if num(finish, "spreads") > 0 else "lagged"
        cost = "cost" if num(first, "spreads") < 0 else "gained"
        b.append(f"Pre-traded orders {verb} the close by "
                 f"{sp(num(finish, 'spreads'))} spreads. First fills {cost} "
                 f"{sp(num(first, 'spreads'))}.")
        src += [cite("73b_pre_close_side_spr", "All", finish),
                cite("73a_pre_firstexec_side_spr", "All", first)]
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
    starts_cost = first is not None and num(first, "spreads") < 0
    S.append(dict(title=("Pre-traded orders: the start is where it costs"
                         if starts_cost else
                         "Pre-traded orders: start, finish and next open"),
                  bullets=b, charts=["26_pretraded_by_side.png"],
                  src=src, notes=(
        "Start = first fill against the close. Finish = the whole execution "
        "against the close. Next open = reversion. Positive is good.\n\n"
        "The orders finish ahead of the close, so the algo recovers - but it "
        "pays for the first fills to get there. Starting later or more "
        "passively keeps the finish without the start.")))

    # --- 5. pre-traded: by market -----------------------------------------
    fe = run.rows("68_pre_firstexec_mkt_spreads")
    fin = run.rows("77_pre_close_mkt_spreads")
    b, src = [], []
    if fe:
        neg = sum(num(r, "spreads") < 0 for _, r in fe)
        b.append("First fills were worse than the close in every market."
                 if neg == len(fe) else
                 "First fills beat the close in every market." if neg == 0 else
                 f"First fills were worse than the close in {neg} of "
                 f"{len(fe)} markets.")
        gains = sorted([(k, r) for k, r in fe if holds(r)
                        and num(r, "spreads") > 0 and not thin(r)],
                       key=lambda kr: -num(kr[1], "spreads"))
        if gains and neg < len(fe) / 2:
            k, r = gains[0]
            b.append(f"Starting early paid most in {k}: "
                     f"{sp(num(r, 'spreads'))} spreads, {about_bps(r)}.")
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
    S.append(dict(title="Pre-traded orders, market by market",
                  bullets=b, charts=["10b_first_exec_market.png",
                                     "27_pretraded_close_market.png"],
                  src=src, notes=(
        "Left: the first fill against the close, weighted by the part of "
        "each order that traded before the auction. Right: the whole "
        "execution against the close.")))

    # --- 6. pre-traded: first fills by market and side --------------------
    b, src = [], []
    fe2 = {k: r for k, r in run.rows("74_pre_firstexec_mkt_side_spr", 2)}
    for side in ("Buy", "Sell"):
        worst = sorted([(k, r) for k, r in fe2.items() if k[1] == side
                        and holds(r) and num(r, "spreads") < 0 and not thin(r)],
                       key=lambda kr: num(kr[1], "spreads"))
        if worst:
            (m, _), r = worst[0]
            b.append(f"{side}s started worst in {m}: "
                     f"{sp(num(r, 'spreads'))} spreads, {about_bps(r)}.")
            src.append(cite("74_pre_firstexec_mkt_side_spr", f"{m}, {side}", r))
    if not b:
        b.append("Neither side started badly by a clear margin in any market.")
    S.append(dict(title="Pre-traded first fills, buys against sells",
                  bullets=b,
                  charts=["28_pretraded_first_exec_market_side.png"],
                  src=src, notes=(
        "First fill against the close, by market, buys on the left and sells "
        "on the right, on one scale.")))

    # --- 7. pre-traded: reversion by market and side ----------------------
    b, src = [], []
    rv2 = {k: r for k, r in run.rows("76_pre_reversion_mkt_side_spr", 2)}
    both = sorted([k for k in fe2 if k in rv2
                   and holds(fe2[k]) and num(fe2[k], "spreads") < 0
                   and holds(rv2[k]) and num(rv2[k], "spreads") < 0],
                  key=lambda k: num(rv2[k], "spreads"))
    worst_pair = both[0] if both else None
    if worst_pair:
        m, side = worst_pair
        b.append(f"{m} {side.lower()}s started "
                 f"{sp(num(fe2[worst_pair], 'spreads'))} spreads worse and came "
                 f"back {sp(num(rv2[worst_pair], 'spreads'))} overnight.")
        src += [cite("74_pre_firstexec_mkt_side_spr", f"{m}, {side}",
                     fe2[worst_pair]),
                cite("76_pre_reversion_mkt_side_spr", f"{m}, {side}",
                     rv2[worst_pair])]
    rv_bad = sorted([(k, r) for k, r in rv2.items() if holds(r)
                     and num(r, "spreads") < 0 and not thin(r)
                     and k != worst_pair],
                    key=lambda kr: num(kr[1], "spreads"))
    b.append(f"{len(rv_bad)} other market and side pairs came back by a clear "
             "margin." if rv_bad else
             "Beyond that, nothing came back by a clear margin.")
    src += [cite("76_pre_reversion_mkt_side_spr", f"{k[0]}, {k[1]}", r)
            for k, r in rv_bad]
    S.append(dict(title="Pre-traded reversion, buys against sells",
                  bullets=b,
                  charts=["30_pretraded_reversion_market_side.png"],
                  src=src, notes=(
        "A market and side that is red here and on the previous slide paid to "
        "start and then gave it back overnight - the clearest case for "
        "changing how that flow is worked.")))

    # --- 7b. pre-traded against PVWAP, by market --------------------------
    b, src = [], []
    pv_all = run.row("73d_pre_pvwap_side_spr", "All")
    if pv_all is not None:
        verb = "beat" if num(pv_all, "spreads") > 0 else "lagged"
        tail = "" if holds(pv_all) else " Not a clear signal."
        b.append(f"Pre-traded orders {verb} PVWAP by "
                 f"{sp(num(pv_all, 'spreads'))} spreads overall.{tail}")
        src.append(cite("73d_pre_pvwap_side_spr", "All", pv_all))
    pv = run.rows("78_pre_pvwap_mkt_spreads")
    if pv:
        ahead = sum(num(r, "spreads") > 0 for _, r in pv)
        b.append("They beat PVWAP in every market." if ahead == len(pv) else
                 f"They beat PVWAP in {ahead} of {len(pv)} markets.")
        lag = sorted([(k, r) for k, r in pv if holds(r)
                      and num(r, "spreads") < 0 and not thin(r)],
                     key=lambda kr: num(kr[1], "spreads"))
        if lag:
            k, r = lag[0]
            b.append(f"{k} lagged it most clearly: {sp(num(r, 'spreads'))} "
                     f"spreads, {about_bps(r)}.")
        src += [cite("78_pre_pvwap_mkt_spreads", k, r) for k, r in pv]
    S.append(dict(title="Pre-traded orders against PVWAP, by market",
                  bullets=b, charts=["32_pretraded_pvwap_market.png"],
                  src=src, notes=(
        "PVWAP is the market VWAP over the time each order was working. The "
        "close says where the order ended up; PVWAP says whether the trading "
        "on the way was good for the window it ran in.\n\n"
        "If an order beats PVWAP but its first fills lag the close, the algo "
        "traded well - it was the timing of the start that cost.")))

    # --- 7c. pre-traded against PVWAP, by market and side -----------------
    b, src = [], []
    pv2 = {k: r for k, r in run.rows("79_pre_pvwap_mkt_side_spr", 2)}
    for side in ("Buy", "Sell"):
        one = run.row("73d_pre_pvwap_side_spr", side)
        if one is not None:
            src.append(cite("73d_pre_pvwap_side_spr", side, one))
        worst = sorted([(k, r) for k, r in pv2.items() if k[1] == side
                        and holds(r) and num(r, "spreads") < 0 and not thin(r)],
                       key=lambda kr: num(kr[1], "spreads"))
        if worst:
            (m, _), r = worst[0]
            b.append(f"{side}s lagged PVWAP most in {m}: "
                     f"{sp(num(r, 'spreads'))} spreads, {about_bps(r)}.")
            src.append(cite("79_pre_pvwap_mkt_side_spr", f"{m}, {side}", r))
    if not b:
        b.append("Against PVWAP, neither side fell behind by a clear margin "
                 "anywhere.")
    S.append(dict(title="Pre-traded against PVWAP, buys against sells",
                  bullets=b, charts=["33_pretraded_pvwap_market_side.png"],
                  src=src, notes=(
        "Same measure as the previous slide, buys on the left and sells on "
        "the right, on one scale.")))

    # --- 8. pre-traded by ADV% --------------------------------------------
    b, src = [], []
    psize = run.table("53_pre_adv_profile")
    small = run.row("66_pre_firstexec_adv_spreads", "0-1%")
    if psize is not None and "0-1%" in psize.index:
        r = psize.loc["0-1%"]
        b.append(f"{num(r, '% of orders'):.0f}% of pre-traded orders were "
                 f"under 1% ADV, but only {num(r, '% of notional'):.0f}% of "
                 "the value.")
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
        b.append(f"Close-only orders, at "
                 f"{num(co, '%ADV (notional-weighted)'):.1f}% ADV, filled "
                 f"{num(co, 'fill ratio %'):.1f}%.")
    S.append(dict(title=("Small pre-traded orders could go straight to the close"
                         if small_costs else "Pre-traded orders by ADV%"),
                  bullets=b, charts=["31_pretraded_adv_profile.png",
                                     "19_first_exec_by_adv.png"],
                  src=src, notes=(
        "The Q1 argument, tested on H1. Small orders are most of the "
        "pre-traded count but a minority of its value. They do not need the "
        "continuous session for capacity - close-only orders of that size "
        "fill almost completely - and starting them early still cost money "
        "on the first fill.")))

    # --- 9. what we will change -------------------------------------------
    b, src = [], []
    if small_costs:
        b.append("Send pre-traded orders under 1% ADV straight to the auction.")
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
        "first-fill cost, the markets with the clearest early-start cost, and "
        "the market and side where the price came back. Agree the order of "
        "these with the desk before the meeting.")))

    # --- appendix -----------------------------------------------------------
    for title, charts in [
        ("Appendix: close performance by month",
         ["12_monthly.png", "12b_monthly_india.png"]),
        ("Appendix: close performance by ADV% and spread",
         ["17_close_by_adv.png", "18_close_by_spread.png"]),
        ("Appendix: value by ADV%, close by market",
         ["24_adv_profile.png", "15_market_vs_close.png"]),
        ("Appendix: first fills by spread",
         ["20_first_exec_by_spread.png"]),
        ("Appendix: pre-traded against the close, buys and sells",
         ["29_pretraded_close_market_side.png"]),
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
        textbox(s, MARGIN, Inches(0.22), SLIDE_W - 2 * MARGIN, Inches(0.7),
                spec["title"], 24, bold=True)
        rule = s.shapes.add_shape(1, MARGIN, Inches(0.88),
                                  SLIDE_W - 2 * MARGIN, Emu(9525))
        rule.fill.solid(); rule.fill.fore_color.rgb = RULE
        rule.line.fill.background(); rule.shadow.inherit = False

        text_only = not spec["charts"]
        n_b = len(spec["bullets"])
        if n_b:
            bullet_box(s, Inches(1.7 if text_only else 0.95), spec["bullets"],
                       24 if text_only else 15)
        if spec["charts"]:
            top = Inches(1.0 + 0.36 * n_b + (0.08 if n_b else 0))
            # The charts' footnotes, once per slide: the same caveat printed
            # under two pictures is read once and ignored the second time.
            foot = []
            for c in spec["charts"]:
                for note in run.titles.get(c, {}).get("notes", []):
                    if note not in foot:
                        foot.append(note)
            rows = sum(max(1, math.ceil(len(x) / 190)) for x in foot)
            foot_h = Inches(0.035 + 0.135 * rows) if foot else Emu(0)
            bottom = SLIDE_H - Inches(0.3)
            charts_row(s, [run.charts / c for c in spec["charts"]], top,
                       bottom - top - foot_h - Inches(0.05), run.titles)
            if foot:
                textbox(s, MARGIN, bottom - foot_h, SLIDE_W - 2 * MARGIN,
                        foot_h, chr(10).join(foot), 8, color=INK_SOFT)
        textbox(s, MARGIN, SLIDE_H - Inches(0.3), SLIDE_W - 2 * MARGIN,
                Inches(0.25), f"{client}  |  MOC TCA {period}  |  {i}", 8,
                color=INK_SOFT)
        chart_notes = [f"{run.titles[c]['title']}: " + " ".join(
                           run.titles[c]["notes"])
                       for c in spec["charts"]
                       if run.titles.get(c, {}).get("notes")]
        s.notes_slide.notes_text_frame.text = (
            spec["notes"]
            + ("\n\nOn the charts:\n" + "\n".join(chart_notes)
               if chart_notes else "")
            + ("\n\nSources:\n" + "\n".join(spec["src"])
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
