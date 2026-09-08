# deck.json

The slide manifest. `build_deck.py` reads this plus `data/`, and nothing else.

```json
{
  "client": "Example Asset Management",
  "period": "H1 2026",
  "footer": "Execution review - prepared for Example Asset Management",
  "slides": [ ... ]
}
```

Every slide accepts `title` and `notes`. **Notes are exempt from the readability
check** - write them technically, with the method, the sample size, whether the
difference is real, and the source image.

## Slide types

**title**
```json
{"type": "title", "title": "Execution review", "subtitle": "H1 2026",
 "notes": "..."}
```

**section** - a divider before each strategy section.
```json
{"type": "section", "title": "How your VWAP orders did"}
```

**bullets**
```json
{"type": "bullets", "title": "Waiting for the close cost more than the trading",
 "bullets": ["...", "...", "..."],
 "defines": ["reversion"],
 "notes": "..."}
```

`defines` lists banned terms this slide explains in its own text. Anything listed
there passes the check on this slide only.

**table**
```json
{"type": "table", "title": "What we looked at",
 "table": {"data": "scope.csv",
           "columns": ["strategy", "orders", "notional_musd", "pct_notional"],
           "headers": ["Strategy", "Orders", "Value ($m)", "Share"],
           "formats": [null, "{:,.0f}", "{:,.1f}", "{:.1f}%"],
           "sort": "-notional_musd", "limit": 12},
 "notes": "..."}
```

**chart**
```json
{"type": "chart", "title": "Your VWAP orders tracked the market closely",
 "chart": {"kind": "bar", "data": "benchmarks.csv",
           "filter": {"benchmark": "pvwap"},
           "x": "strategy", "y": "value_bps",
           "diverging": true, "value_fmt": "{:+.1f}",
           "ylabel": "basis points against the market average"},
 "bullets": ["Optional, up to two lines beside the chart"],
 "notes": "..."}
```

## Chart kinds

| kind | Required keys | Use for |
|---|---|---|
| `bar` | `x`, `y` | One value per category. Blue for savings, red for cost when `diverging` |
| `grouped_bar` | `x`, `y`, `series` | Same measure across two or three groups |
| `stacked_bar` | `x`, `y`, `series` | Composition that sums to 100, such as venue mix |
| `waterfall` | `x`, `y`, optional `total_label` | Cost decomposition - waiting plus execution |
| `scatter` | `x`, `y`, optional `series`, optional `line` | Size against outcome, with a reference line |
| `line` | `x`, `y`, optional `series` | Trend over the period |

Common optional keys: `filter` (exact-match column filters), `sort`, `limit`,
`ylabel`, `xlabel`, `value_fmt`, `diverging`, `axis_note`, `hline`.

`axis_note` prints under the axis. On any cost measure use
`"left = cost | right = savings"` (bar) or the vertical equivalent, so the reader
never has to remember the sign.

## Rules the build enforces

- Title 10 words or fewer.
- Three bullets maximum, 14 words each, 20 words per sentence, no semicolons.
- Banned terms from `plain-language.md` fail unless listed in `defines`.
- A chart slide and a table slide cannot carry both.
- `meta.json` must have `"verified": true`.

## A working example

`assets/deck.example.json` is a complete twelve-slide deck exercising every
slide type and all six chart kinds. Copy it and replace the content.
