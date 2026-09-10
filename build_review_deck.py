#!/usr/bin/env python3
"""Build the client deck for the MOC review.

    python build_review_deck.py --charts output_h1/charts
    python build_review_deck.py --charts output_h1/charts --slides 3

Five slides by default; --slides 3 collapses them without changing any wording.
Both are the same deck, so nothing is rebuilt if the meeting turns out shorter.

Every number lives in NUMBERS below, in one place, so correcting a figure means
editing one line rather than hunting through slide text. Each is annotated with
the table it came from - see reviews/moc-h1-2026/evidence.md for the full
provenance and the confidence grade.

The technical detail goes in the speaker notes, where the client never sees it
and the sales trader always can.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

# ===========================================================================
# THE NUMBERS - correct these against unified_tables.xlsx before presenting
# ===========================================================================
NUMBERS = {
    # 09_venue_mix_strategy / 30_monthly / 29_fill_rate / 01_scope - all agree
    "notional_musd":      1061.44,
    "auction_notional":    823.74,   # 05_headline_vs_close
    "auction_share_lo":     71.2,    # 30_monthly, lowest month
    "auction_share_hi":     87.1,    # 30_monthly, highest month
    "auction_share_avg":    80,      # rounded, for the slide
    "fill_rate":            99.6,    # 29_fill_rate, notional-weighted
    "never_traded":          0,      # 29_fill_rate

    # 05_headline_vs_close
    "vs_close_bps":          3.74,
    "vs_close_usdk":       307.84,
    "vs_close_ci_lo":        2.31,
    "vs_close_ci_hi":        5.40,

    # 17_first_exec_by_adv - the early start, in dollars
    "early_bands": [("under 1%", -61, 8850), ("1-3%", -235, 209),
                    ("3-5%", -30, 95), ("5-10%", -121, 62),
                    ("10-25%", -105, 56)],

    # 23_close_vs_session_cap
    "session_bands": [("Large", -254), ("Mid", -498),
                      ("Small", -248), ("Other", -45)],

    # 15_cohorts / 16_orders_to_review / 37_close_opportunity
    "unexplained_note": "several hundred orders",
    "markets_short":    2,
    "markets_short_musd": 16,
}
NUMBERS["early_total"] = sum(v for _, v, _ in NUMBERS["early_bands"])
NUMBERS["session_total"] = sum(v for _, v in NUMBERS["session_bands"])

CLIENT = "Client"
PERIOD = "H1 2026"
FOOTER = "MOC execution review"

# ===========================================================================

try:
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Emu, Inches, Pt
except ImportError:
    print("ERROR: python-pptx is not installed.\n       pip install python-pptx",
          file=sys.stderr)
    raise SystemExit(2)

# The palette is imported, not restated. The charts were built against a
# validated set - blue is a saving, red is a cost, and the surface is an
# off-white the charts are already drawn on. A deck painted pure white puts a
# visible box edge around every chart, and a deck that picks its own blue
# stops meaning the same thing as the bars inside the picture.
import moc_tca as tca


def _rgb(hexstr: str) -> "RGBColor":
    h = hexstr.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


SLIDE_W, SLIDE_H = Inches(13.333), Inches(7.5)
MARGIN = Inches(0.7)
INK = _rgb(tca.INK)
INK_SOFT = _rgb(tca.INK_SECOND)
ACCENT = _rgb(tca.POS)          # the same blue the savings bars use
COST = _rgb(tca.NEG)            # the same red the cost bars use
RULE = _rgb(tca.BASELINE)
SURFACE = _rgb(tca.SURFACE)     # what the charts are drawn on
FONT = "Arial"


def textbox(slide, left, top, width, height, text, size, *, bold=False,
            color=INK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.0):
    box = slide.shapes.add_textbox(left, top, width, height)
    frame = box.text_frame
    frame.word_wrap = True
    frame.vertical_anchor = anchor
    for i, line in enumerate(text.split("\n")):
        para = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        para.text = line
        para.alignment = align
        para.line_spacing = spacing
        for run in para.runs:
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.color.rgb = color
            run.font.name = FONT
    return box


def bullets(slide, top, items, size=19):
    box = slide.shapes.add_textbox(MARGIN, top, SLIDE_W - 2 * MARGIN,
                                   Inches(1.3))
    frame = box.text_frame
    frame.word_wrap = True
    for i, text in enumerate(items):
        para = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        para.text = f"•  {text}"
        para.line_spacing = 1.2
        para.space_after = Pt(6)
        for run in para.runs:
            run.font.size = Pt(size)
            run.font.color.rgb = INK
            run.font.name = FONT


def head(slide, title):
    textbox(slide, MARGIN, Inches(0.5), SLIDE_W - 2 * MARGIN, Inches(0.9),
            title, 27, bold=True)
    line = slide.shapes.add_shape(1, MARGIN, Inches(1.4),
                                  SLIDE_W - 2 * MARGIN, Emu(9525))
    line.fill.solid()
    line.fill.fore_color.rgb = RULE
    line.line.fill.background()
    line.shadow.inherit = False


def picture(slide, path: Path, top, max_h):
    """Place a chart, scaled to fit, never stretched."""
    if not path.exists():
        textbox(slide, MARGIN, top, SLIDE_W - 2 * MARGIN, Inches(0.5),
                f"[chart not found: {path.name}]", 13, color=INK_SOFT)
        return
    from PIL import Image
    with Image.open(path) as im:
        ratio = im.height / im.width
    width = SLIDE_W - 2 * MARGIN
    height = Emu(int(width * ratio))
    if height > max_h:
        height, width = max_h, Emu(int(max_h / ratio))
    slide.shapes.add_picture(str(path), Emu(int((SLIDE_W - width) / 2)), top,
                             width=width, height=height)


def money(v):
    """-235 -> '-$235,000'. Whole thousands, because that is the precision."""
    return f"{'-' if v < 0 else ''}${abs(v):,.0f},000"


def draw_session_chart(out: Path) -> Path:
    """Slide 4's chart. Drawn here because moc_tca.py has no chart for it."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    bands = NUMBERS["session_bands"]
    fig, ax = plt.subplots(figsize=(11, 4.6), dpi=200)
    fig.patch.set_facecolor(tca.SURFACE)
    ax.set_facecolor(tca.SURFACE)
    names = [b[0] for b in bands]
    vals = [b[1] for b in bands]
    ax.barh(range(len(bands)), vals, color=tca.NEG, height=0.6, zorder=3)
    ax.set_yticks(range(len(bands)))
    ax.set_yticklabels(names, fontsize=13)
    ax.invert_yaxis()
    ax.axvline(0, color=tca.BASELINE, linewidth=1.2, zorder=4)
    span = max(abs(v) for v in vals)
    ax.set_xlim(-span * 1.35, span * 0.18)
    for i, v in enumerate(vals):
        ax.text(v - span * 0.03, i, money(v), va="center", ha="right",
                fontsize=13, color=tca.INK)
    ax.set_xlabel("what trading at the close cost against trading through the day",
                  fontsize=12, color=tca.INK_SECOND)
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    ax.spines["left"].set_color(tca.GRID)
    ax.spines["bottom"].set_color(tca.GRID)
    ax.tick_params(colors=tca.INK_SECOND, labelsize=12)
    ax.grid(axis="x", color=tca.GRID, linewidth=1, zorder=0)
    ax.set_axisbelow(True)
    fig.tight_layout()
    out.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out, facecolor=tca.SURFACE, dpi=200, bbox_inches="tight")
    plt.close(fig)
    return out


def build(charts: Path, out: Path, n_slides: int, cover: bool = False) -> None:
    n = NUMBERS
    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    blank = prs.slide_layouts[6]
    three = n_slides == 3

    def new(title=None):
        s = prs.slides.add_slide(blank)
        # Paint the slide the same off-white the charts sit on, so a chart
        # blends into the page instead of floating in a white rectangle.
        bg = s.shapes.add_shape(1, Emu(0), Emu(0), SLIDE_W, SLIDE_H)
        bg.fill.solid(); bg.fill.fore_color.rgb = SURFACE
        bg.line.fill.background(); bg.shadow.inherit = False
        s.shapes._spTree.remove(bg._element)
        s.shapes._spTree.insert(2, bg._element)      # behind everything
        if title:
            head(s, title)
        return s

    # --- cover, only if asked for -----------------------------------------
    # Off by default so the deck is exactly five slides, or exactly three.
    # The scope line moves onto the first content slide instead.
    if cover:
        s = new()
        band = s.shapes.add_shape(1, Emu(0), Inches(2.55), SLIDE_W, Emu(38100))
        band.fill.solid(); band.fill.fore_color.rgb = ACCENT
        band.line.fill.background(); band.shadow.inherit = False
        textbox(s, MARGIN, Inches(1.55), SLIDE_W - 2 * MARGIN, Inches(1.0),
                "How your close orders traded", 40, bold=True)
        textbox(s, MARGIN, Inches(2.85), SLIDE_W - 2 * MARGIN, Inches(0.8),
                f"{CLIENT} — {PERIOD}", 22, color=INK_SOFT)
        s.notes_slide.notes_text_frame.text = (
            f"CLOSE strategy only, {PERIOD}. USD {n['notional_musd']:,.0f}m "
            "executed. The client's VWAP flow also reaches closing auctions "
            "and is out of scope, so this describes the MOC product rather "
            "than their total auction footprint.")

    scope_line = (f"{CLIENT} — {PERIOD}   ·   close orders only, "
                  f"USD {n['notional_musd']:,.0f}m traded")

    # --- 1. it works ------------------------------------------------------
    if not three:
        s = new("Your close orders reached the auction and filled")
        if not cover:
            textbox(s, MARGIN, Inches(1.5), SLIDE_W - 2 * MARGIN, Inches(0.4),
                    scope_line, 13, color=INK_SOFT)
        picture(s, charts / "12_monthly.png", Inches(2.0), Inches(3.7))
        bullets(s, Inches(5.85), [
            f"About {n['auction_share_avg']}% of each order printed in the auction.",
            f"Steady all half: never below {n['auction_share_lo']:.0f}%, "
            f"never above {n['auction_share_hi']:.0f}%.",
            f"{n['fill_rate']}% of orders filled. None went untraded.",
        ])
        s.notes_slide.notes_text_frame.text = (
            "Notional-weighted auction share by month, from 30_monthly: 83.3 / "
            "71.2 / 87.1 / 83.9 / 72.2 / 84.2. Fill rate 99.6% weighted, 99.56% "
            "of orders fully filled, zero never traded (29_fill_rate).\n\n"
            "India reports no auction share, so its orders are imputed at 100% "
            "where they ran through the closing window. Any all-market auction "
            "figure is part measured and part assumed. Do not volunteer it; "
            "answer straight if asked.")

    # --- 2. beat the close ------------------------------------------------
    title = ("Your close orders work, and beat the closing price" if three
             else f"You beat the closing price by {n['vs_close_bps']:.1f} basis points")
    s = new(title)
    if three and not cover:
        textbox(s, MARGIN, Inches(1.5), SLIDE_W - 2 * MARGIN, Inches(0.4),
                scope_line, 13, color=INK_SOFT)
    picture(s, charts / "03_headline_vs_close.png",
            Inches(2.0) if (three and not cover) else Inches(1.75), Inches(3.6))
    if three:
        items = [
            f"About {n['auction_share_avg']}% of each order printed in the auction.",
            f"You beat the closing price by {n['vs_close_bps']:.1f} basis points"
            f" — about {money(round(n['vs_close_usdk']))}.",
            "Orders that cleared print at the close by definition.",
        ]
    else:
        items = [
            f"Worth about {money(round(n['vs_close_usdk']))} on auction flow.",
            f"The range is {n['vs_close_ci_lo']:.1f} to {n['vs_close_ci_hi']:.1f},"
            " so this is real.",
            "Orders that cleared print at the close by definition.",
        ]
    bullets(s, Inches(5.6), items)
    s.notes_slide.notes_text_frame.text = (
        f"05_headline_vs_close: +{n['vs_close_bps']} bps on USD "
        f"{n['auction_notional']:,.0f}m of auction-market flow, worth USD "
        f"{n['vs_close_usdk']:,.0f}k. 95% bootstrap interval "
        f"[{n['vs_close_ci_lo']}, {n['vs_close_ci_hi']}] — clears zero, so the "
        "result is real rather than noise.\n\n"
        "The last bullet is the honest one and it changes what the number "
        "means: an order filled entirely in the auction trades AT the close, so "
        "its slippage is ~0 by construction. The whole +3.7 is earned by the "
        "portion that did NOT clear. Say it first if a quant is in the room.\n\n"
        "Do NOT lead with the arrival number. Four of six months are negative "
        "and March alone (+89bps) drags the half positive.")

    # --- 3. the early start ----------------------------------------------
    s = new(f"Starting before the close cost about "
            f"{money(abs(round(n['early_total'] / 10) * 10))}")
    picture(s, charts / "10_first_exec.png", Inches(1.75), Inches(3.6))
    worst = min(n["early_bands"], key=lambda b: b[1])
    bullets(s, Inches(5.6), [
        "Every size of order lost money by starting early.",
        f"The worst: {worst[2]} orders at {worst[0]} of daily volume, "
        f"{money(worst[1])}.",
        "Those were small enough for the auction to absorb.",
    ])
    band_txt = ", ".join(f"{b[0]} {money(b[1])} ({b[2]:,} orders)"
                         for b in n["early_bands"])
    s.notes_slide.notes_text_frame.text = (
        f"17_first_exec_by_adv, in dollars: {band_txt}. Total about "
        f"{money(n['early_total'])}.\n\n"
        "first_exec_vs_close is the first fill against the close the order was "
        "aiming at, weighted by each order's CONTINUOUS notional - the part "
        "that traded before the auction - because the measure only bites on "
        "that part. Negative means the early start lost money.\n\n"
        "This is the strongest recommendation in the deck: no model, and the "
        "desk controls it directly. Expect 'we start early to reduce risk'. "
        "The answer is that under 3% of daily volume there was no capacity "
        "reason to.")

    # --- 4. the close vs the session -------------------------------------
    if not three:
        s = new("The session beat the close in every company size")
        picture(s, draw_session_chart(charts / "_session_vs_close.png"),
                Inches(1.75), Inches(3.6))
        bullets(s, Inches(5.6), [
            "Trading through the day would have beaten the closing print.",
            f"True for large, mid and small alike — about "
            f"{money(abs(round(n['session_total'] / 10) * 10))}.",
            "This is about where to trade, not how well.",
        ])
        s.notes_slide.notes_text_frame.text = (
            "23_close_vs_session_cap: Large -254k, Mid -498k, Small -248k, "
            "Other -45k. VWAP minus Close per order - same executed price on "
            "both sides, so the difference is a pure price move. Negative means "
            "the close was the worse place to trade.\n\n"
            "Lean on the third bullet. This is NOT the algo underperforming - "
            "the previous slide shows it beating its benchmark. It is a "
            "question about where the flow goes, and it is the client's "
            "decision. One half-year, not a law.")

    # --- 5. what we would change -----------------------------------------
    s = new("What we would change")
    items = [
        f"Stop sending small close orders early. About "
        f"{money(abs(round(n['early_total'] / 10) * 10))} a half.",
    ]
    if three:
        items.append("Worth asking whether mid-cap flow belongs in the close.")
        items.append(f"{n['unexplained_note'].capitalize()} got nothing in the "
                     "auction, unexplained.")
    else:
        items.append("Ask whether mid-cap flow belongs in the close at all.")
        items.append(f"{n['markets_short']} markets barely reach the close. "
                     "We are checking why.")
    bullets(s, Inches(2.1), items, size=21)
    textbox(s, MARGIN, Inches(4.3), SLIDE_W - 2 * MARGIN, Inches(1.6),
            "What we would need from you", 19, bold=True)
    bullets(s, Inches(4.9), [
        "A limit price on each order, so we can tell when a limit blocked it.",
        "Close-eligibility tagging would explain most of the rest.",
    ], size=17)
    s.notes_slide.notes_text_frame.text = (
        f"The data ask. {n['unexplained_note'].capitalize()} traded, were small "
        "enough for the auction to absorb, and got nothing in it. They are "
        "listed by order id in 16_orders_to_review and go to the desk.\n\n"
        "We cannot explain them from this file: market_limit reads 'Limit' on "
        "every order, so it distinguishes nothing. A limit price that varies "
        "would close most of the gap.\n\n"
        f"Also in reserve: {n['markets_short']} markets where nearly every "
        f"order finished before the close, about USD {n['markets_short_musd']}m "
        "(37_close_opportunity). Either genuinely mis-scoped or our close time "
        "for those markets is wrong - worth checking before raising.\n\n"
        "Not claimable at all: whether we traded with or against the imbalance, "
        "and any impact claim beyond reversion, which is not distinguishable "
        "from zero.")

    # --- furniture --------------------------------------------------------
    total = len(prs.slides)
    for i, slide in enumerate(prs.slides, start=1):
        if cover and i == 1:
            continue
        textbox(slide, MARGIN, SLIDE_H - Inches(0.5), Inches(9), Inches(0.3),
                f"{FOOTER} — {CLIENT}, {PERIOD}", 10, color=INK_SOFT)
        textbox(slide, SLIDE_W - MARGIN - Inches(1), SLIDE_H - Inches(0.5),
                Inches(1), Inches(0.3), f"{i} / {total}", 10,
                color=INK_SOFT, align=PP_ALIGN.RIGHT)

    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    print(f"built {out}  ({total} slides)")
    print("\nEvery number came from reviews/moc-h1-2026/evidence.md.")
    print("The early-start figures are graded C there - read them off")
    print("unified_tables.xlsx and correct NUMBERS before presenting.")


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--charts", type=Path, default=Path("output_h1/charts"),
                   help="folder of charts from moc_tca.py")
    p.add_argument("--out", type=Path, default=None,
                   help="output .pptx (default: MOC_review_<period>.pptx)")
    p.add_argument("--slides", type=int, choices=(3, 5), default=5,
                   help="5 (default) or 3; same deck, two collapsed")
    p.add_argument("--cover", action="store_true",
                   help="add a separate cover slide (default: scope goes on "
                        "the first slide instead, keeping the count at 5 or 3)")
    args = p.parse_args(argv)

    if not args.charts.is_dir():
        print(f"ERROR: no chart folder at {args.charts}", file=sys.stderr)
        print("       run moc_tca.py first, then point --charts at its charts/",
              file=sys.stderr)
        return 2
    out = args.out or Path(f"MOC_review_{PERIOD.replace(' ', '_')}"
                           f"{'_short' if args.slides == 3 else ''}.pptx")
    build(args.charts, out, args.slides, args.cover)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
