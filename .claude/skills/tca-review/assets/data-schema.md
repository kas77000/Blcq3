# Data schema

Everything transcribed lands here. `build_deck.py` reads nothing else, so a
number that is not in these files cannot reach a slide.

```
reviews/<client>-<period>/
  images/            the .HEIC files as supplied
  images/png/        converted, written by intake.py
  data/
    meta.json        required
    scope.csv        required
    benchmarks.csv   required
    sources.csv      required
    execution.csv    optional
    cohorts.csv      optional
    monthly.csv      optional
  verification.md    the sign-off sheet
  findings.md
  deck.json
  charts/
  <client>-<period>.pptx
```

Two columns appear on every CSV:

- `source_image` - the PNG the number was read from. Required on every row.
- `approx` - `Y` if the value was read off a chart rather than a table, else `N`.
  Anything marked `Y` is written as "about" wherever it appears.

Unreadable value → `NA`. Never a guess.

## meta.json

```json
{
  "client": "Example Asset Management",
  "period": "H1 2026",
  "date_from": "2026-01-02",
  "date_to": "2026-06-30",
  "currency": "USD",
  "notional_unit": "USD millions",
  "sign_convention": "positive_is_savings",
  "prepared_by": "Execution Consulting",
  "verified": false,
  "verified_on": null,
  "caveats": [
    "Korea and Taiwan close times not verified against the desk's sessions",
    "India reported separately - NSE has no single-price closing auction"
  ]
}
```

`sign_convention` is `positive_is_savings` or `positive_is_cost`. The build
converts everything to positive-is-savings internally and says so on the deck.

## scope.csv

One row per strategy.

| Column | Meaning |
|---|---|
| `strategy` | As printed in the source. Do not rename |
| `family` | From `strategy-playbooks.md`; blank if unknown |
| `orders` | Order count |
| `notional_musd` | Value traded, in the unit named in `meta.json` |
| `pct_notional` | Share of total value |
| `fill_rate_pct` | Mean fill rate, blank if not shown |
| `markets` | Semicolon-separated, or blank |

## benchmarks.csv

Long format - one row per strategy and benchmark. This carries the whole of the
performance story and every chart of it.

| Column | Meaning |
|---|---|
| `strategy` | Matches `scope.csv` |
| `benchmark` | `arrival`, `pvwap`, `vwap`, `twap`, `close`, `open`, `nextopen`, `decision` |
| `value_bps` | Signed, in the convention named in `meta.json` |
| `n` | Orders behind this number. Required |
| `ci_lo`, `ci_hi` | Range if the source shows one, else blank |
| `weighting` | `notional`, `equal`, or blank if not stated |

## execution.csv

Long format - one row per strategy and metric. Everything that is not a
benchmark slippage.

| Column | Meaning |
|---|---|
| `strategy` | Matches `scope.csv`, or `ALL` |
| `metric` | `pct_close`, `pct_open`, `pct_dark`, `pct_post`, `pct_take`, `fill_rate`, `participation_actual`, `participation_target`, `duration_min`, `pct_adv`, `spread_bps`, `reversion_bps`, `spread_capture_pct` |
| `value` | Numeric |
| `unit` | `pct`, `bps`, `min`, `x` |
| `n` | Orders behind it |

The five venue shares sum to 100 where all five are present - `verify.py` checks
it.

## cohorts.csv

For the miss taxonomy, or any other grouping of orders.

`cohort, orders, notional_musd, pct_notional, note`

## monthly.csv

`period, strategy, benchmark, value_bps, n` - `period` as `2026-01`.

## sources.csv

The audit trail. One row per table transcribed, plus one row per `NA`.

`table, field, source_image, note`
