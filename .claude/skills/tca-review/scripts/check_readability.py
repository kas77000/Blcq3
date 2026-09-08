#!/usr/bin/env python3
"""Plain-language check for a TCA client deck.

The rules in references/plain-language.md, made mechanical. Speaker notes are
exempt - they carry the technical backup for the sales trader.

    python check_readability.py --deck reviews/client-h1-2026/deck.json

Exit code 1 means at least one slide breaks a rule. Fix the wording. Do not
loosen the check.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

MAX_TITLE_WORDS = 10
MAX_BULLETS = 3
MAX_BULLET_WORDS = 14
MAX_SENTENCE_WORDS = 20

BANNED = [
    "implementation shortfall", "notional-weighted", "notional weighted",
    "value-weighted", "winsorised", "winsorized", "bootstrap",
    "confidence interval", "statistically significant", "heteroscedastic",
    "frontier", "cohort", "decomposition", "degenerate", "endogenous",
    "orthogonal", "alpha decay", "adverse selection", "participation-weighted",
    "basis point dispersion", "interquartile", "percentile",
]

# Phrasings that make a deck read as machine-written. A client can tell, and once
# they can, they stop trusting the numbers too.
TELLS = [
    "delve", "leverage", "leveraging", "seamless", "seamlessly", "robust",
    "landscape", "underscores", "underscoring", "it is worth noting",
    "it's worth noting", "moreover", "furthermore", "in conclusion",
    "a testament to", "plays a key role", "plays a crucial role",
    "when it comes to", "in today's", "fast-paced", "ever-evolving",
    "unlock", "harness", "elevate", "streamline", "holistic",
    "actionable insights", "deep dive", "at the end of the day",
    "navigate the complexities", "tapestry", "realm", "pivotal",
    "not only", "comprehensive overview", "key takeaway", "best-in-class",
    "state-of-the-art", "significant improvement", "optimal outcomes",
]
MAX_SAME_OPENING = 3  # slides whose first bullet may start with the same word

_WORD = re.compile(r"[A-Za-z0-9$%£€.,+\-/']+")
_SENTENCE = re.compile(r"[.!?]+\s+|[.!?]+$")


def words(text: str) -> list[str]:
    return _WORD.findall(text or "")


def sentences(text: str) -> list[str]:
    parts = [p.strip() for p in _SENTENCE.split(text or "") if p and p.strip()]
    return parts or ([text.strip()] if text and text.strip() else [])


def banned_in(text: str, allowed: set[str]) -> list[str]:
    low = (text or "").lower()
    return [t for t in BANNED if t in low and t not in allowed]


def tells_in(text: str) -> list[str]:
    low = (text or "").lower()
    return [t for t in TELLS if t in low]


def check_slide(index: int, slide: dict) -> list[str]:
    where = f"slide {index + 1} ({slide.get('type', '?')})"
    problems: list[str] = []
    allowed = {d.lower() for d in slide.get("defines", [])}

    title = slide.get("title", "")
    if title:
        n = len(words(title))
        if n > MAX_TITLE_WORDS:
            problems.append(f"{where}: title is {n} words, limit {MAX_TITLE_WORDS} - {title!r}")
        if ";" in title:
            problems.append(f"{where}: semicolon in the title")
        for term in banned_in(title, allowed):
            problems.append(f"{where}: title uses {term!r} - define it on the slide or drop it")
        for term in tells_in(title):
            problems.append(f"{where}: title uses {term!r} - reads as boilerplate, say it your own way")
        if title.count("—") + title.count(" - ") > 1:
            problems.append(f"{where}: more than one dash in the title - use a full stop")

    bullets = slide.get("bullets") or []
    if len(bullets) > MAX_BULLETS:
        problems.append(f"{where}: {len(bullets)} bullets, limit {MAX_BULLETS}")
    for bullet in bullets:
        n = len(words(bullet))
        if n > MAX_BULLET_WORDS:
            problems.append(f"{where}: bullet is {n} words, limit {MAX_BULLET_WORDS} - {bullet!r}")
        if ";" in bullet:
            problems.append(f"{where}: semicolon in a bullet - split it into two sentences")
        for sentence in sentences(bullet):
            sn = len(words(sentence))
            if sn > MAX_SENTENCE_WORDS:
                problems.append(
                    f"{where}: sentence is {sn} words, limit {MAX_SENTENCE_WORDS} - {sentence!r}"
                )
        for term in banned_in(bullet, allowed):
            problems.append(f"{where}: bullet uses {term!r} - define it on the slide or drop it")
        for term in tells_in(bullet):
            problems.append(f"{where}: bullet uses {term!r} - reads as boilerplate, say it your own way")

    subtitle = slide.get("subtitle", "")
    if subtitle:
        for sentence in sentences(subtitle):
            sn = len(words(sentence))
            if sn > MAX_SENTENCE_WORDS:
                problems.append(f"{where}: subtitle sentence is {sn} words, limit {MAX_SENTENCE_WORDS}")
        for term in banned_in(subtitle, allowed):
            problems.append(f"{where}: subtitle uses {term!r}")

    if slide.get("chart") and slide.get("table"):
        problems.append(f"{where}: carries both a chart and a table - use one")

    for term in slide.get("defines", []):
        if term.lower() not in BANNED:
            problems.append(f"{where}: defines {term!r}, which is not on the ban list - remove it")

    return problems


def check_deck(deck: dict) -> list[str]:
    problems: list[str] = []
    slides = deck.get("slides") or []
    if not slides:
        return ["deck has no slides"]
    for i, slide in enumerate(slides):
        problems.extend(check_slide(i, slide))
    problems.extend(check_rhythm(slides))
    return problems


def check_rhythm(slides: list[dict]) -> list[str]:
    """Human writing varies. A deck where every slide opens the same way does not."""
    problems: list[str] = []

    openings: dict[str, int] = {}
    for slide in slides:
        bullets = slide.get("bullets") or []
        if bullets:
            first = (words(bullets[0])[:1] or [""])[0].lower()
            if first:
                openings[first] = openings.get(first, 0) + 1
    for word, count in sorted(openings.items(), key=lambda kv: -kv[1]):
        if count > MAX_SAME_OPENING:
            problems.append(
                f"deck: {count} slides open a bullet with {word!r} - vary the phrasing"
            )

    lengths = [len(words(b)) for s in slides for b in (s.get("bullets") or [])]
    if len(lengths) >= 6:
        spread = max(lengths) - min(lengths)
        if spread <= 2:
            problems.append(
                "deck: every bullet is the same length - real writing varies, "
                "let some run short"
            )

    titles = [s.get("title", "").lower() for s in slides if s.get("title")]
    if len(titles) != len(set(titles)):
        problems.append("deck: two slides share a title - each one states its own finding")

    return problems


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--deck", required=True)
    args = ap.parse_args()

    path = Path(args.deck)
    if not path.exists():
        print(f"ERROR: no deck at {path}", file=sys.stderr)
        return 2

    deck = json.loads(path.read_text(encoding="utf-8"))
    problems = check_deck(deck)

    slides = len(deck.get("slides") or [])
    if not problems:
        print(f"readability: {slides} slides, no violations")
        return 0

    print(f"readability: {len(problems)} violation(s) across {slides} slides\n")
    for p in problems:
        print(f"  {p}")
    print("\nRewrite the wording. The client has to understand it on one reading.")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
