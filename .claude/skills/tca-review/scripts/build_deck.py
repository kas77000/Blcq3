#!/usr/bin/env python3
"""Build the client PowerPoint from verified TCA data.

Reads deck.json and data/ and nothing else, so a number that is not in the
signed-off data cannot reach a slide. Refuses to run until meta.json says the
numbers were verified, and fails on any plain-language violation.

    python build_deck.py --review reviews/client-h1-2026
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Emu, Inches, Pt

sys.path.insert(0, str(Path(__file__).parent))
import charts  # noqa: E402
from check_readability import check_deck  # noqa: E402

SLIDE_W, SLIDE_H = Inches(13.333), Inches(7.5)
MARGIN = Inches(0.7)
INK = RGBColor(0x0B, 0x0B, 0x0B)
INK_SOFT = RGBColor(0x52, 0x51, 0x4E)
ACCENT = RGBColor(0x2A, 0x78, 0xD6)
RULE = RGBColor(0xDE, 0xDE, 0xDA)
SURFACE = RGBColor(0xFF, 0xFF, 0xFF)
BAND = RGBColor(0xF4, 0xF4, 0xF1)
DRAFT = RGBColor(0xE3, 0x49, 0x48)
FONT = "Arial"


def textbox(slide, left, top, width, height, text, size, *, bold=False,
            color=INK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.0):
    box = slide.shapes.add_textbox(left, top, width, height)
    frame = box.text_frame
    frame.word_wrap = True
    frame.vertical_anchor = anchor
    lines = text.split("\n") if isinstance(text, str) else list(text)
    for i, line in enumerate(lines):
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


def bullet_list(slide, left, top, width, height, bullets, size=18):
    box = slide.shapes.add_textbox(left, top, width, height)
    frame = box.text_frame
    frame.word_wrap = True
    for i, text in enumerate(bullets):
        para = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        para.text = f"•  {text}"
        para.line_spacing = 1.25
        para.space_after = Pt(10)
        for run in para.runs:
            run.font.size = Pt(size)
            run.font.color.rgb = INK
            run.font.name = FONT
    return box


def rule(slide, top):
    line = slide.shapes.add_shape(1, MARGIN, top, SLIDE_W - 2 * MARGIN, Emu(9525))
    line.fill.solid()
    line.fill.fore_color.rgb = RULE
    line.line.fill.background()
    line.shadow.inherit = False
    return line


def slide_title(slide, text):
    textbox(slide, MARGIN, Inches(0.55), SLIDE_W - 2 * MARGIN, Inches(1.0),
            text, 26, bold=True)
    rule(slide, Inches(1.5))


def furniture(slide, footer, page, total, draft):
    if footer:
        textbox(slide, MARGIN, SLIDE_H - Inches(0.55), Inches(9.0), Inches(0.35),
                footer, 10, color=INK_SOFT)
    textbox(slide, SLIDE_W - MARGIN - Inches(1.0), SLIDE_H - Inches(0.55),
            Inches(1.0), Inches(0.35), f"{page} / {total}", 10,
            color=INK_SOFT, align=PP_ALIGN.RIGHT)
    if draft:
        textbox(slide, MARGIN, Inches(0.14), Inches(8.0), Inches(0.3),
                "DRAFT - NUMBERS NOT VERIFIED", 11, bold=True, color=DRAFT)


def set_notes(slide, notes):
    if notes:
        slide.notes_slide.notes_text_frame.text = notes


def add_title_slide(prs, spec, deck):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    band = slide.shapes.add_shape(1, Emu(0), Inches(2.5), SLIDE_W, Emu(38100))
    band.fill.solid()
    band.fill.fore_color.rgb = ACCENT
    band.line.fill.background()
    band.shadow.inherit = False
    textbox(slide, MARGIN, Inches(1.5), SLIDE_W - 2 * MARGIN, Inches(1.0),
            spec.get("title", "Execution review"), 40, bold=True)
    subtitle = spec.get("subtitle") or f"{deck.get('client', '')} - {deck.get('period', '')}".strip(" -")
    textbox(slide, MARGIN, Inches(2.8), SLIDE_W - 2 * MARGIN, Inches(0.8), subtitle, 22, color=INK_SOFT)
    if spec.get("footnote"):
        textbox(slide, MARGIN, Inches(5.9), SLIDE_W - 2 * MARGIN, Inches(0.8),
                spec["footnote"], 13, color=INK_SOFT)
    return slide


def add_section_slide(prs, spec):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    band = slide.shapes.add_shape(1, Emu(0), Inches(2.9), SLIDE_W, Inches(1.7))
    band.fill.solid()
    band.fill.fore_color.rgb = BAND
    band.line.fill.background()
    band.shadow.inherit = False
    textbox(slide, MARGIN, Inches(3.15), SLIDE_W - 2 * MARGIN, Inches(1.2),
            spec.get("title", ""), 30, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    return slide


def add_bullets_slide(prs, spec):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_title(slide, spec.get("title", ""))
    bullets = spec.get("bullets") or []
    if bullets:
        bullet_list(slide, MARGIN, Inches(2.0), SLIDE_W - 2 * MARGIN, Inches(4.4), bullets, size=20)
    return slide


def add_chart_slide(prs, spec, data_dir, charts_dir, index):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_title(slide, spec.get("title", ""))
    bullets = spec.get("bullets") or []

    name = spec.get("chart_name") or f"{index:02d}_{spec['chart'].get('kind', 'chart')}"
    png = charts.render(spec["chart"], data_dir, charts_dir / f"{name}.png")

    top = Inches(1.85)
    avail_h = SLIDE_H - top - Inches(0.75) - (Inches(1.0) if bullets else Inches(0))
    width = SLIDE_W - 2 * MARGIN
    height = Emu(int(width * charts.FIG_H / charts.FIG_W))
    if height > avail_h:
        height = avail_h
        width = Emu(int(height * charts.FIG_W / charts.FIG_H))
    left = Emu(int((SLIDE_W - width) / 2))
    slide.shapes.add_picture(str(png), left, top, width=width, height=height)

    if bullets:
        bullet_list(slide, MARGIN, top + height + Inches(0.15),
                    SLIDE_W - 2 * MARGIN, Inches(0.9), bullets, size=16)
    return slide


def add_table_slide(prs, spec, data_dir):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_title(slide, spec.get("title", ""))

    cfg = spec["table"]
    df = pd.read_csv(data_dir / cfg["data"], na_values=["NA", ""])
    for col, value in (cfg.get("filter") or {}).items():
        df = df[df[col].astype(str) == str(value)]
    if cfg.get("sort"):
        key = cfg["sort"]
        df = df.sort_values(key.lstrip("-"), ascending=not key.startswith("-"))
    if cfg.get("limit"):
        df = df.head(int(cfg["limit"]))

    columns = cfg.get("columns") or list(df.columns)
    headers = cfg.get("headers") or columns
    formats = cfg.get("formats") or [None] * len(columns)

    rows, cols = len(df) + 1, len(columns)
    height = min(Inches(4.9), Inches(0.42) * rows)
    shape = slide.shapes.add_table(rows, cols, MARGIN, Inches(1.9),
                                   SLIDE_W - 2 * MARGIN, height)
    table = shape.table

    for c, header in enumerate(headers):
        cell = table.cell(0, c)
        cell.text = str(header)
        cell.fill.solid()
        cell.fill.fore_color.rgb = BAND
        for para in cell.text_frame.paragraphs:
            para.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.RIGHT
            for run in para.runs:
                run.font.size = Pt(14)
                run.font.bold = True
                run.font.color.rgb = INK
                run.font.name = FONT

    for r, (_, row) in enumerate(df.iterrows(), start=1):
        for c, col in enumerate(columns):
            value = row.get(col)
            fmt = formats[c] if c < len(formats) else None
            if pd.isna(value):
                text = "-"
            elif fmt:
                try:
                    text = fmt.format(float(value))
                except (ValueError, TypeError):
                    text = str(value)
            else:
                text = str(value)
            cell = table.cell(r, c)
            cell.text = text
            cell.fill.solid()
            cell.fill.fore_color.rgb = SURFACE
            for para in cell.text_frame.paragraphs:
                para.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.RIGHT
                for run in para.runs:
                    run.font.size = Pt(14)
                    run.font.color.rgb = INK
                    run.font.name = FONT
    return slide


BUILDERS = {"title", "section", "bullets", "chart", "table"}


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--review", required=True)
    ap.add_argument("--out", help="output .pptx (default: <review>/<client>-<period>.pptx)")
    ap.add_argument("--allow-unverified", action="store_true",
                    help="build anyway and stamp every slide DRAFT. For your own eyes only.")
    args = ap.parse_args()

    root = Path(args.review)
    data_dir, charts_dir = root / "data", root / "charts"
    deck_path = root / "deck.json"

    for path in (data_dir, deck_path):
        if not path.exists():
            print(f"ERROR: missing {path}", file=sys.stderr)
            return 2

    meta = json.loads((data_dir / "meta.json").read_text(encoding="utf-8"))
    if not meta.get("verified"):
        if not args.allow_unverified:
            print("REFUSED: data/meta.json does not say the numbers were verified.\n", file=sys.stderr)
            print("  Re-read every source image, write verification.md, get sign-off,", file=sys.stderr)
            print('  then set "verified": true. Or pass --allow-unverified for a', file=sys.stderr)
            print("  draft stamped on every slide.", file=sys.stderr)
            return 2
        print("WARNING: building unverified. Every slide is stamped DRAFT.\n")

    deck = json.loads(deck_path.read_text(encoding="utf-8"))

    problems = check_deck(deck)
    if problems:
        print(f"REFUSED: {len(problems)} plain-language violation(s)\n", file=sys.stderr)
        for p in problems:
            print(f"  {p}", file=sys.stderr)
        print("\nRewrite the wording. The client has to understand it on one reading.",
              file=sys.stderr)
        return 1

    if meta.get("sign_convention") == "positive_is_cost":
        print("NOTE: source convention is positive_is_cost. Confirm every sign in")
        print("      deck.json was flipped during analysis, and that the deck says so.\n")

    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H

    slides = deck["slides"]
    total = len(slides)
    draft = not meta.get("verified")
    footer = deck.get("footer", "")

    for i, spec in enumerate(slides, start=1):
        kind = spec.get("type")
        if kind not in BUILDERS:
            print(f"ERROR: slide {i} has unknown type {kind!r}", file=sys.stderr)
            return 2
        if kind == "title":
            slide = add_title_slide(prs, spec, deck)
        elif kind == "section":
            slide = add_section_slide(prs, spec)
        elif kind == "bullets":
            slide = add_bullets_slide(prs, spec)
        elif kind == "chart":
            slide = add_chart_slide(prs, spec, data_dir, charts_dir, i)
        else:
            slide = add_table_slide(prs, spec, data_dir)

        set_notes(slide, spec.get("notes"))
        if kind != "title":
            furniture(slide, footer, i, total, draft)
        elif draft:
            furniture(slide, "", i, total, draft)

    stem = f"{meta.get('client', 'client')}-{meta.get('period', '')}".strip("- ")
    stem = "".join(ch if ch.isalnum() or ch in " -_" else "" for ch in stem).strip().replace(" ", "-")
    out = Path(args.out) if args.out else root / f"{stem or 'execution-review'}.pptx"
    prs.save(str(out))

    print(f"built {out}  ({total} slides)")
    print(f"charts in {charts_dir}")
    if draft:
        print("\nThis is a DRAFT. Do not send it to a client.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
