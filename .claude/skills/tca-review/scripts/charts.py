#!/usr/bin/env python3
"""Deck-sized chart rendering for a TCA client review.

Charts are drawn at slide dimensions with presentation-sized type. Colours come
from the validated reference palette: blue for savings, red for cost, grey at
zero, and a fixed categorical order so a strategy keeps its colour across the
whole deck.

Used by build_deck.py. Also runnable to preview one chart:

    python charts.py --data reviews/x/data --spec '{"kind":"bar",...}' --out c.png
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import pandas as pd  # noqa: E402

# Validated reference palette, light surface.
CATEGORICAL = [
    "#2a78d6", "#eb6834", "#1baf7a", "#eda100",
    "#e87ba4", "#008300", "#4a3aa7", "#e34948",
]
SAVINGS = "#2a78d6"
COST = "#e34948"
NEUTRAL = "#8a8a85"
GRID = "#dededa"
INK = "#0b0b0b"
INK_SOFT = "#52514e"
SURFACE = "#ffffff"

FIG_W, FIG_H, DPI = 11.0, 5.6, 200
BASE_PT = 15

plt.rcParams.update(
    {
        "figure.facecolor": SURFACE,
        "axes.facecolor": SURFACE,
        "font.size": BASE_PT,
        "axes.labelsize": BASE_PT,
        "xtick.labelsize": BASE_PT,
        "ytick.labelsize": BASE_PT,
        "text.color": INK,
        "axes.labelcolor": INK_SOFT,
        "xtick.color": INK_SOFT,
        "ytick.color": INK_SOFT,
        "axes.edgecolor": GRID,
        "axes.linewidth": 1.0,
        "grid.color": GRID,
        "grid.linewidth": 1.0,
        "figure.autolayout": False,
    }
)

_COLOR_CACHE: dict[str, dict[str, str]] = {}

# Column names the client should never see on a chart.
PRETTY = {
    "pct_close": "in the closing auction",
    "pct_open": "in the opening auction",
    "pct_post": "posted passively",
    "pct_take": "took liquidity",
    "pct_dark": "dark",
    "fill_rate": "filled",
    "participation_actual": "achieved",
    "participation_target": "asked for",
    "pct_adv": "size against a normal day",
    "spread_capture_pct": "spread earned",
    "reversion_bps": "price came back",
    "duration_min": "minutes working",
    "arrival": "when the order reached us",
    "pvwap": "market average while working",
    "vwap": "market average for the day",
    "close": "closing auction price",
    "open": "opening auction price",
    "nextopen": "next day's open",
}


def pretty(name, spec: dict | None = None) -> str:
    """Plain label for a raw column value. `series_labels` in the spec wins."""
    key = str(name)
    if spec:
        override = (spec.get("series_labels") or {}).get(key)
        if override:
            return override
    return PRETTY.get(key, key)


def strategy_colors(data_dir: Path) -> dict[str, str]:
    """Stable colour per strategy: largest by value takes slot 1, and so on."""
    key = str(data_dir.resolve())
    if key in _COLOR_CACHE:
        return _COLOR_CACHE[key]
    mapping: dict[str, str] = {}
    scope = data_dir / "scope.csv"
    if scope.exists():
        df = pd.read_csv(scope)
        if "notional_musd" in df.columns:
            df = df.sort_values("notional_musd", ascending=False)
        for i, name in enumerate(df["strategy"].astype(str)):
            mapping[name] = CATEGORICAL[i % len(CATEGORICAL)]
    _COLOR_CACHE[key] = mapping
    return mapping


def _series_colors(names, data_dir: Path) -> list[str]:
    mapping = strategy_colors(data_dir)
    out = []
    for i, name in enumerate(names):
        out.append(mapping.get(str(name), CATEGORICAL[i % len(CATEGORICAL)]))
    return out


def load_frame(spec: dict, data_dir: Path) -> pd.DataFrame:
    df = pd.read_csv(data_dir / spec["data"], na_values=["NA", ""])
    for col, value in (spec.get("filter") or {}).items():
        if isinstance(value, list):
            df = df[df[col].astype(str).isin([str(v) for v in value])]
        else:
            df = df[df[col].astype(str) == str(value)]
    sort = spec.get("sort")
    if sort:
        desc = sort.startswith("-")
        df = df.sort_values(sort.lstrip("-"), ascending=not desc)
    if spec.get("limit"):
        df = df.head(int(spec["limit"]))
    return df.reset_index(drop=True)


def _finish(fig, ax, spec: dict, horizontal: bool) -> None:
    ax.set_xlabel(spec.get("xlabel", ""))
    ax.set_ylabel(spec.get("ylabel", ""))
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    note = spec.get("axis_note")
    if note:
        fig.text(0.5, 0.015, note, ha="center", va="bottom", fontsize=BASE_PT - 2, color=INK_SOFT)
    top = 0.93 if getattr(ax, "legend_above", False) else 0.99
    fig.tight_layout(rect=(0, 0.05 if note else 0.01, 1, top))


def _fmt(value, spec: dict) -> str:
    try:
        return (spec.get("value_fmt") or "{:+.1f}").format(value)
    except (ValueError, TypeError):
        return str(value)


def chart_bar(df, spec, ax, fig, data_dir):
    """Horizontal bars, one per category. Blue above zero, red below."""
    labels = df[spec["x"]].astype(str).tolist()
    values = pd.to_numeric(df[spec["y"]], errors="coerce").fillna(0).tolist()
    if spec.get("diverging", True):
        colors = [SAVINGS if v >= 0 else COST for v in values]
    else:
        colors = _series_colors(labels, data_dir)

    pos = range(len(labels))
    ax.barh(list(pos), values, color=colors, height=0.62, zorder=3)
    ax.set_yticks(list(pos))
    ax.set_yticklabels([pretty(l, spec) for l in labels])
    ax.invert_yaxis()
    ax.axvline(0, color=NEUTRAL, linewidth=1.4, zorder=4)
    ax.grid(axis="x", zorder=0)
    ax.set_axisbelow(True)

    span = max((abs(v) for v in values), default=1.0) or 1.0
    pad = span * 0.04
    for i, v in enumerate(values):
        ax.text(
            v + (pad if v >= 0 else -pad), i, _fmt(v, spec),
            va="center", ha="left" if v >= 0 else "right",
            fontsize=BASE_PT, color=INK,
        )
    ax.set_xlim(min(0, min(values, default=0)) - span * 0.28,
                max(0, max(values, default=0)) + span * 0.28)
    _finish(fig, ax, spec, horizontal=True)


def _pivot(df, spec):
    """Categories down the index, series across the columns, both in deck order."""
    pivot = df.pivot_table(index=spec["x"], columns=spec["series"], values=spec["y"],
                           aggfunc="first")
    if spec.get("series_order"):
        pivot = pivot[[c for c in spec["series_order"] if c in pivot.columns]]
    if spec.get("x_order"):
        keep = [c for c in spec["x_order"] if c in pivot.index]
        pivot = pivot.loc[keep + [c for c in pivot.index if c not in keep]]
    return pivot


def chart_grouped_bar(df, spec, ax, fig, data_dir):
    pivot = _pivot(df, spec)
    cats = pivot.index.astype(str).tolist()
    series = pivot.columns.astype(str).tolist()
    width = 0.8 / max(len(series), 1)
    for i, name in enumerate(series):
        offs = [j - 0.4 + width * (i + 0.5) for j in range(len(cats))]
        vals = pd.to_numeric(pivot[pivot.columns[i]], errors="coerce").fillna(0).tolist()
        ax.bar(offs, vals, width=width * 0.92, label=pretty(name, spec),
               color=CATEGORICAL[i % len(CATEGORICAL)], zorder=3)
        for x, v in zip(offs, vals):
            ax.text(x, v, _fmt(v, spec), ha="center",
                    va="bottom" if v >= 0 else "top", fontsize=BASE_PT - 3, color=INK)
    ax.set_xticks(range(len(cats)))
    ax.set_xticklabels(cats)
    ax.axhline(0, color=NEUTRAL, linewidth=1.4, zorder=4)
    ax.grid(axis="y", zorder=0)
    ax.set_axisbelow(True)
    ax.legend(frameon=False, loc="best", fontsize=BASE_PT - 2)
    _finish(fig, ax, spec, horizontal=False)


def chart_stacked_bar(df, spec, ax, fig, data_dir):
    pivot = _pivot(df, spec).fillna(0)
    cats = pivot.index.astype(str).tolist()
    pos = list(range(len(cats)))
    left = [0.0] * len(cats)
    for i, col in enumerate(pivot.columns):
        vals = pd.to_numeric(pivot[col], errors="coerce").fillna(0).tolist()
        ax.barh(pos, vals, left=left, height=0.62, label=pretty(col, spec),
                color=CATEGORICAL[i % len(CATEGORICAL)], zorder=3,
                edgecolor=SURFACE, linewidth=1.6)
        for j, (l, v) in enumerate(zip(left, vals)):
            if v >= 8:  # only label a segment wide enough to hold the text
                ax.text(l + v / 2, j, f"{v:.0f}", ha="center", va="center",
                        fontsize=BASE_PT - 3, color=SURFACE)
        left = [l + v for l, v in zip(left, vals)]
    ax.set_yticks(pos)
    ax.set_yticklabels(cats)
    ax.invert_yaxis()
    ax.set_xlim(0, max(left + [100]))
    ax.grid(axis="x", zorder=0)
    ax.set_axisbelow(True)
    # Plain-word legend labels are long, so wrap them rather than run off the edge.
    labels = [pretty(c, spec) for c in pivot.columns]
    ncol = 5 if sum(len(l) for l in labels) <= 55 else 3
    ax.legend(frameon=False, ncol=min(len(labels), ncol),
              loc="upper center", bbox_to_anchor=(0.5, 1.13), fontsize=BASE_PT - 3)
    ax.legend_above = True
    _finish(fig, ax, spec, horizontal=True)


def chart_waterfall(df, spec, ax, fig, data_dir):
    """Components adding to a total. The total bar is drawn from zero."""
    labels = df[spec["x"]].astype(str).tolist()
    values = pd.to_numeric(df[spec["y"]], errors="coerce").fillna(0).tolist()
    total_label = spec.get("total_label")

    # Lay the bars out first so the axis can be sized before anything is drawn -
    # otherwise a label on a bar that reaches the axis edge collides with the ticks.
    bars, running = [], 0.0
    for label, value in zip(labels, values):
        is_total = total_label is not None and label == total_label
        bottom = 0.0 if is_total else running
        bars.append((label, value, bottom, bottom + value, is_total))
        if not is_total:
            running = bottom + value

    edges = [0.0] + [b[2] for b in bars] + [b[3] for b in bars]
    lo, hi = min(edges), max(edges)
    span = (hi - lo) or 1.0
    ax.set_ylim(lo - span * 0.18, hi + span * 0.18)
    pad = span * 0.03

    for i, (label, value, bottom, top, is_total) in enumerate(bars):
        color = NEUTRAL if is_total else (SAVINGS if value >= 0 else COST)
        ax.bar(i, value, bottom=bottom, width=0.6, color=color, zorder=3)
        ax.text(i, top + (pad if value >= 0 else -pad), _fmt(value, spec), ha="center",
                va="bottom" if value >= 0 else "top", fontsize=BASE_PT, color=INK)
        if not is_total and i < len(bars) - 1:
            ax.plot([i + 0.3, i + 0.7], [top, top], color=NEUTRAL,
                    linewidth=1.0, linestyle=(0, (3, 3)), zorder=2)

    ax.set_xticks(range(len(labels)))
    ax.set_xticklabels([pretty(l, spec) for l in labels])
    ax.axhline(0, color=NEUTRAL, linewidth=1.4, zorder=4)
    ax.grid(axis="y", zorder=0)
    ax.set_axisbelow(True)
    _finish(fig, ax, spec, horizontal=False)


def chart_scatter(df, spec, ax, fig, data_dir):
    x = pd.to_numeric(df[spec["x"]], errors="coerce")
    y = pd.to_numeric(df[spec["y"]], errors="coerce")
    series_col = spec.get("series")
    if series_col and series_col in df.columns:
        names = list(dict.fromkeys(df[series_col].astype(str)))[:3]  # 3 slots validate all-pairs
        for i, name in enumerate(names):
            mask = df[series_col].astype(str) == name
            ax.scatter(x[mask], y[mask], s=90, label=name, alpha=0.85,
                       color=CATEGORICAL[i % len(CATEGORICAL)],
                       edgecolor=SURFACE, linewidth=1.4, zorder=3)
        ax.legend(frameon=False, fontsize=BASE_PT - 2)
    else:
        ax.scatter(x, y, s=90, color=CATEGORICAL[0], alpha=0.85,
                   edgecolor=SURFACE, linewidth=1.4, zorder=3)

    line = spec.get("line")
    if line:
        ldf = load_frame(line, data_dir) if line.get("data") else df
        lx = pd.to_numeric(ldf[line["x"]], errors="coerce")
        ly = pd.to_numeric(ldf[line["y"]], errors="coerce")
        order = lx.argsort()
        ax.plot(lx.iloc[order], ly.iloc[order], color=INK_SOFT, linewidth=2.0,
                linestyle=(0, (5, 3)), zorder=4, label=line.get("label", ""))
        if line.get("label"):
            ax.legend(frameon=False, fontsize=BASE_PT - 2)

    if spec.get("hline") is not None:
        ax.axhline(float(spec["hline"]), color=NEUTRAL, linewidth=1.4, zorder=2)
    ax.grid(zorder=0)
    ax.set_axisbelow(True)
    _finish(fig, ax, spec, horizontal=False)


def chart_line(df, spec, ax, fig, data_dir):
    xcol, ycol = spec["x"], spec["y"]
    series_col = spec.get("series")
    if series_col and series_col in df.columns:
        names = list(dict.fromkeys(df[series_col].astype(str)))
        colors = _series_colors(names, data_dir)
        for name, color in zip(names, colors):
            sub = df[df[series_col].astype(str) == name]
            xs = sub[xcol].astype(str).tolist()
            ys = pd.to_numeric(sub[ycol], errors="coerce").tolist()
            ax.plot(xs, ys, marker="o", markersize=8, linewidth=2.0, color=color, zorder=3)
            if xs:
                ax.annotate(name, (xs[-1], ys[-1]), textcoords="offset points",
                            xytext=(8, 0), va="center", fontsize=BASE_PT - 2, color=color)
    else:
        xs = df[xcol].astype(str).tolist()
        ys = pd.to_numeric(df[ycol], errors="coerce").tolist()
        ax.plot(xs, ys, marker="o", markersize=8, linewidth=2.0, color=CATEGORICAL[0], zorder=3)
    ax.axhline(0, color=NEUTRAL, linewidth=1.4, zorder=2)
    ax.grid(axis="y", zorder=0)
    ax.set_axisbelow(True)
    ax.margins(x=0.12)
    _finish(fig, ax, spec, horizontal=False)


KINDS = {
    "bar": chart_bar,
    "grouped_bar": chart_grouped_bar,
    "stacked_bar": chart_stacked_bar,
    "waterfall": chart_waterfall,
    "scatter": chart_scatter,
    "line": chart_line,
}


def render(spec: dict, data_dir: Path, out_path: Path) -> Path:
    kind = spec.get("kind")
    if kind not in KINDS:
        raise ValueError(f"unknown chart kind {kind!r}; expected one of {sorted(KINDS)}")
    df = load_frame(spec, Path(data_dir))
    if df.empty:
        raise ValueError(f"chart {kind!r} on {spec.get('data')} has no rows after filtering")

    fig, ax = plt.subplots(figsize=(FIG_W, FIG_H), dpi=DPI)
    try:
        KINDS[kind](df, spec, ax, fig, Path(data_dir))
        out_path = Path(out_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        fig.savefig(out_path, dpi=DPI, facecolor=SURFACE)
    finally:
        plt.close(fig)
    return out_path


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--data", required=True)
    ap.add_argument("--spec", required=True, help="chart spec as JSON, or a path to a JSON file")
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    spec_path = Path(args.spec)
    spec = json.loads(spec_path.read_text(encoding="utf-8")) if spec_path.exists() else json.loads(args.spec)
    path = render(spec, Path(args.data), Path(args.out))
    print(f"wrote {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
