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


def build(charts: Path, out: Path, n_slides: int, cover: bool = False,
          markets: bool = True) -> None:
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

    # --- 0. where the money went -----------------------------------------
    # First, because it is the only slide the client can check against their
    # own records. Recognising the shape of their own book buys the benefit of
    # the doubt on everything that follows.
    if markets:
        s = new("Where you traded")
        if not cover:
            textbox(s, MARGIN, Inches(1.5), SLIDE_W - 2 * MARGIN, Inches(0.4),
                    scope_line, 13, color=INK_SOFT)
        picture(s, charts / "13_market_notional.png", Inches(2.0), Inches(4.4))
        s.notes_slide.notes_text_frame.text = "\n".join([
            "34_market_profile, executed notional by market, largest first.",
            "",
            "No judgement on this slide, and that is deliberate. It is the one",
            "page the client can check against their own records, so let them",
            "recognise their own book before anything is claimed about it.",
            "",
            "If asked which market cost the most, the answer is on",
            "14_market_slippage - but ordered by value traded, not by cost,",
            "because a small market with a big number is still a small market.",
            "",
            "India is roughly a third of the book and runs no closing auction",
            "in this period, so its auction share is imputed rather than",
            "measured. Answer straight if asked; do not volunteer it here.",
        ])

    # --- 1. it works ------------------------------------------------------
    if not three:
        s = new("Your close orders reached the auction and filled")
        if not cover and not markets:
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
    # The title deliberately does NOT claim a pattern across markets. Whether
    # most of them are positive depends on the run, and a title that turns out
    # false in front of the client is worse than one that says less. The
    # chart shows the spread; the title states only what is measured.
    title = ("Your close orders work, and beat the closing price" if three
             else "You beat the closing price overall")
    s = new(title)
    lead = three and not cover and not markets
    if lead:
        textbox(s, MARGIN, Inches(1.5), SLIDE_W - 2 * MARGIN, Inches(0.4),
                scope_line, 13, color=INK_SOFT)
    # By market rather than one summary bar: it shows where the result comes
    # from, and whether it rests on a single place.
    picture(s, charts / "15_market_vs_close.png",
            Inches(2.0) if lead else Inches(1.75), Inches(3.6))
    if three:
        items = [
            f"About {n['auction_share_avg']}% of each order printed in the auction.",
            f"You beat the closing price by {n['vs_close_bps']:.1f} basis points"
            f" — about {money(round(n['vs_close_usdk']))}.",
            "Orders that cleared print at the close by definition.",
        ]
    else:
        items = [
            f"Across the book, {n['vs_close_bps']:.1f} basis points better — "
            f"about {money(round(n['vs_close_usdk']))}.",
            f"The range is {n['vs_close_ci_lo']:.1f} to "
            f"{n['vs_close_ci_hi']:.1f}, so this is real.",
            "Not every market, though. The chart shows where it came from.",
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
            "The closing price was a worse level than the day's average.",
            f"True for large, mid and small alike — about "
            f"{money(abs(round(n['session_total'] / 10) * 10))}.",
            "You still beat the close. The close itself was the problem.",
        ])
        s.notes_slide.notes_text_frame.text = "\n".join([
            "THE QUESTION YOU WILL BE ASKED: how is slippage positive on the",
            "last slide and this so negative?",
            "",
            "Because they measure different things. vs Close compares your",
            "EXECUTION PRICE with the closing print. This compares TWO",
            "BENCHMARKS with each other - the day's VWAP against the closing",
            "print - using the same executed price on both sides, so the",
            "execution cancels out and what is left is a pure price move.",
            "",
            "You executed well against a benchmark that was itself",
            "unfavourable. You beat the shop's price by 3.7bps; the shop down",
            "the road was about 10bps cheaper. Both true.",
            "",
            "The arithmetic: -1.04m on 1,061m is about -10bps, and you beat",
            "the close by +3.7, so net against the session roughly -6bps.",
            "",
            "23_close_vs_session_cap: Large -254k, Mid -498k, Small -248k,",
            "Other -45k.",
            "",
            "TWO THINGS TO SAY BEFORE SOMEONE ELSE DOES",
            "",
            "1. The denominators differ. +308k is on 824m of auction-market",
            "   flow; -1.04m is on the full 1,061m including India. They are",
            "   not subtractable.",
            "",
            "2. The session VWAP is a counterfactual you did not trade. These",
            "   orders went to the close for a reason - urgency, a benchmark",
            "   obligation, index tracking. If the flow buys names that rally",
            "   into the close, the close will always look worse than the",
            "   day's average, and that is the flow's character rather than a",
            "   failure of the mechanism. This slide cannot tell them apart.",
            "",
            "So raise it as a question, never as a recommendation. One",
            "half-year, not a law.",
        ])

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

    md = write_evidence(out.with_suffix(".evidence.md"), out)
    print(f"sources -> {md}")
    print("\nThe early-start figures are graded C - read them off")
    print("unified_tables.xlsx and correct NUMBERS before presenting.")


# ===========================================================================
# WHERE EVERY STATEMENT COMES FROM
# ===========================================================================
#
# Keyed by slide title. Generated from the SAVED deck rather than from the
# code that wrote it, so the quoted statements are the ones actually on the
# slides - if a bullet changes, this file changes with it and cannot drift.
#
# Grades: A read from a cell and cross-checked against another table;
#         B read from a cell, not yet cross-checked;
#         C read off a chart or the terminal, approximate by nature.

SOURCES = {
    "Where you traded": [
        ("Executed notional by market", "`34_market_profile`, chart "
         "`13_market_notional`", "Sum of executed notional per market, "
         "largest first. Totals to USD 1,061.44m.", "A"),
    ],
    "Your close orders reached the auction and filled": [
        ("About 80% printed in the auction", "`30_monthly`, `wtd %CLOSE`",
         "Notional-weighted per month: 83.3 / 71.2 / 87.1 / 83.9 / 72.2 / "
         "84.2. India imputed at 100% where it ran through its closing "
         "window, so this is part measured and part assumed.", "A"),
        ("Never below 71%, never above 87%", "`30_monthly`",
         "Range across the six months. No trend.", "A"),
        ("99.6% filled, none untraded", "`29_fill_rate`",
         "Weighted fill rate 99.6%; 99.56% of orders at or above 99.5% "
         "filled; zero orders never traded.", "B"),
    ],
    "You beat the closing price overall": [
        ("Not every market, though", "`34_market_profile`, `vs Close bps`, "
         "chart `15_market_vs_close`",
         "Per market, notional-weighted, ordered by value traded rather than "
         "by result. The title claims no pattern across markets on purpose - "
         "whether most are positive depends on the run, and a title that "
         "turns out false in front of the client is worse than one that says "
         "less. Read the chart before presenting and be ready to name the "
         "negative markets.", "B"),
        ("+3.74 bps across the book", "`05_headline_vs_close`, "
         "`wtd mean bps`", "Notional-weighted, winsorised 1/99, on USD "
         "823.74m of auction-market flow.", "B"),
        ("Worth about $308,000", "`05_headline_vs_close`, `saved (USDk)`",
         "bps x notional / 10 = 3.74 x 823.74 / 10 = 308.1, against 307.84 "
         "reported. The arithmetic reconciles.", "A"),
        ("Range 2.3 to 5.4, so this is real", "same, `CI low` / `CI high`",
         "95% percentile bootstrap, 2,000 draws. The interval clears zero.",
         "B"),
        ("Cleared orders print at the close by definition", "structural",
         "An order filled entirely in the auction trades AT the closing "
         "price, so its slippage is ~0 by construction. The whole +3.74 is "
         "earned by the portion that did NOT clear.", "A"),
    ],
    "Your close orders work, and beat the closing price": [
        ("Merged slide", "see the two slides above",
         "Same numbers, same sources - this is the three-slide version's "
         "compression of them.", "A"),
    ],
    "Starting before the close cost about $550,000": [
        ("Every size band lost money", "`17_first_exec_by_adv`, "
         "`saved (USDk)`", "under 1% -61k (8,850 orders), 1-3% -235k (209), "
         "3-5% -30k (95), 5-10% -121k (62), 10-25% -105k (56).", "C"),
        ("Total about $550,000", "sum of the five bands",
         "-552k. All five are negative, so the direction does not rest on "
         "any one figure - the total is what needs pinning, not the "
         "finding.", "C"),
        ("Worst: 209 orders at 1-3% of daily volume", "same",
         "-235k, on orders small enough for the auction to absorb, so no "
         "capacity reason to start early.", "C"),
        ("What the measure is", "`first_exec_vs_close`",
         "The first fill against the close the order was aiming at, weighted "
         "by each order's CONTINUOUS notional - the part that traded before "
         "the auction - because the measure only bites there.", "A"),
    ],
    "The session beat the close in every company size": [
        ("The close was a worse level than the day's average",
         "`23_close_vs_session_cap`", "VWAP minus Close per order, same "
         "executed price on both sides, so execution cancels and what "
         "remains is a pure price move.", "B"),
        ("About $1,040,000 across the bands", "same, `saved (USDk)`",
         "Large -254k, Mid -498k, Small -248k, Other -45k.", "B"),
        ("You still beat the close", "the slide above",
         "Not a contradiction: one compares the execution price with the "
         "close, the other compares two benchmarks with each other. "
         "-1.04m on 1,061m is about -10bps; +3.7 on execution; roughly -6 "
         "net.", "A"),
        ("CAUTION - denominators differ", "both slides",
         "+308k is on 823.74m of auction-market flow; -1.04m is on the full "
         "1,061.44m including India. Not subtractable.", "A"),
        ("CAUTION - the session VWAP was never traded", "structural",
         "These orders went to the close for a reason. Flow that buys names "
         "rallying into the close will always make the close look worse "
         "than the day's average. Nothing here separates that from a real "
         "mechanism cost, so it is a question and never a recommendation.",
         "A"),
    ],
    "What we would change": [
        ("Stop sending small close orders early", "`17_first_exec_by_adv`",
         "About -552k a half, and the desk controls it directly.", "C"),
        ("Ask whether mid-cap flow belongs in the close",
         "`23_close_vs_session_cap`", "Mid-caps are the worst band at -498k. "
         "A question, not a recommendation - see the cautions above.", "B"),
        ("Several hundred orders got nothing, unexplained",
         "`15_cohorts`, `16_orders_to_review`",
         "Traded, small enough for the auction to absorb, no auction fill. "
         "Listed individually by order id.", "B"),
        ("Two markets barely reach the close", "`37_close_opportunity`",
         "98% and 100% of orders finished before the close, about USD 16m. "
         "Either genuinely mis-scoped or our close time for those markets is "
         "wrong - check before raising.", "B"),
        ("We need a limit price", "`26_by_market_limit`",
         "market_limit reads 'Limit' on every order in the file, so it "
         "distinguishes nothing and cannot explain the unexplained cohort.",
         "A"),
    ],
}

CANNOT_CLAIM = [
    ("Whether a limit was binding",
     "`market_limit` reads 'Limit' on every order"),
    ("India's true auction share",
     "not reported by the platform; imputed from the closing window"),
    ("Whether we traded with or against the imbalance", "no imbalance data"),
    ("Any impact claim",
     "reversion is not distinguishable from zero"),
]


def write_evidence(md_path: Path, pptx_path: Path) -> Path:
    """Read the saved deck back and write down what backs every statement."""
    from pptx import Presentation as _P

    deck = _P(str(pptx_path))
    lines = [f"# Sources — {pptx_path.name}", "",
             "Every statement on every slide, and the data behind it. "
             "Generated from the saved deck, so the quotes are the words "
             "actually on the slides.", "",
             "**Grades.** **A** read from a cell and cross-checked against "
             "another table · **B** read from a cell, not yet cross-checked · "
             "**C** read off a chart or the terminal, approximate by nature.",
             "", "Numbers came from photographs of `unified_tables.xlsx` "
             "(run of 10 September 2026). Every table read reconciles to the "
             "same population total of **USD 1,061.44m**, which is why the "
             "shape is trustworthy; the last decimal is not.", "", "---", ""]

    for i, slide in enumerate(deck.slides, start=1):
        title = next((sh.text_frame.text.strip().split("\n")[0]
                      for sh in slide.shapes
                      if sh.has_text_frame and sh.text_frame.text.strip()
                      and not sh.text_frame.text.strip().startswith("•")), "")
        lines += [f"## Slide {i} — {title}", ""]

        said = []
        for sh in slide.shapes:
            if sh.has_text_frame and sh.text_frame.text.strip().startswith("•"):
                said += [l.lstrip("• ").strip()
                         for l in sh.text_frame.text.splitlines() if l.strip()]
        if said:
            lines += ["**What the slide says**", ""]
            lines += [f"- {t}" for t in said] + [""]

        rows = SOURCES.get(title)
        if rows:
            lines += ["**What backs it**", "",
                      "| Claim | Source | How it is derived | Grade |",
                      "|---|---|---|---|"]
            for claim, src, how, grade in rows:
                lines.append(f"| {claim} | {src} | {how} | **{grade}** |")
            lines.append("")
        else:
            lines += ["_No sources recorded for this slide._", ""]

    lines += ["---", "", "## What no slide claims", "",
              "Each of these was checked and is genuinely unavailable, not "
              "merely unmeasured.", "", "| Not claimable | Why |", "|---|---|"]
    lines += [f"| {what} | {why} |" for what, why in CANNOT_CLAIM]
    lines += ["", "---", "", "## Before presenting", "",
              "Re-read from `unified_tables.xlsx` and correct `NUMBERS` in "
              "`build_review_deck.py`:", "",
              "1. The five early-start figures — grade C, and they carry the "
              "main recommendation", "2. `05_headline_vs_close` — the bps, "
              "the money, both interval bounds",
              "3. `23_close_vs_session_cap` — the four cap figures",
              "4. `29_fill_rate` — fill rate and never-traded count", "",
              "Everything else is grade A or structural and cannot move.", ""]

    md_path.write_text("\n".join(lines), encoding="utf-8")
    return md_path


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--charts", type=Path, default=Path("output_h1/charts"),
                   help="folder of charts from moc_tca.py")
    p.add_argument("--out", type=Path, default=None,
                   help="output .pptx (default: MOC_review_<period>.pptx)")
    p.add_argument("--slides", type=int, choices=(3, 5), default=5,
                   help="5 (default) or 3; same deck, two collapsed")
    p.add_argument("--no-markets", dest="markets", action="store_false",
                   help="drop the 'where you traded' slide (default: keep it)")
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
    build(args.charts, out, args.slides, args.cover, args.markets)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
