# MOC / close-algo TCA

H1 review of every close algo traded. **Tables and charts, no deck** — the
narrative comes out of the numbers, not the other way round.

**The whole period ran on the previous trading platform.** This client has not
migrated yet, so H1 is a single regime end to end — no cutover inside the
window, nothing to split, no date filter needed. `DATE_FROM` / `DATE_TO` stay
`None`.

Two consequences:

- Nothing here describes the current platform's close path, and no field name
  from it is assumed. Where a close-eligibility flag is mentioned, it means
  whatever the previous system called it.
- This is the **pre-migration baseline**, and it is worth more as that than as a
  remediation list. Every number is the "before" measurement against which the
  new platform will be judged when this client moves — so the conclusions to
  draw out are the ones that say what should change at cutover, and what the
  client should expect to see move.

One file, `moc_tca.py`. Copy it to the machine holding the order extract and run
it there. No network, no config file, nothing to install beyond the libraries.

```bash
python moc_tca.py --data orders.csv --probe    # audit the file and STOP - run this first
python moc_tca.py --data orders.csv            # tables + charts
python moc_tca.py --sample                     # synthetic file, end to end
python moc_tca.py --self-test                  # analytics only, no data file
```

Requires `pandas numpy matplotlib openpyxl`. Without matplotlib the charts are
skipped with a warning and the tables still build.

Output lands in `output/`:

| File | What it is |
|---|---|
| `tables.xlsx` | every table, one sheet each |
| `charts/*.png` | every exhibit standalone at 200dpi |
| `run_log.txt` | sanity report, close-algo selection, all tables, findings |

**What to check on the target machine:** [docs/WHAT-TO-CHECK.md](docs/WHAT-TO-CHECK.md)

**Which columns are needed:** [docs/COLUMNS-CHECKLIST.md](docs/COLUMNS-CHECKLIST.md)
(tick-list) and [docs/COLUMNS-REQUIRED.md](docs/COLUMNS-REQUIRED.md) (why each
one, and what dies without it).

## Run this first

```bash
python moc_tca.py --data orders.csv --probe
```

`--probe` reads the file and stops. It prints every header with dtype and a
sample value, which columns mapped, the field distributions, the venue-mix sum
check, the side-adjustment check, the markets found, and **every `Strategy`
value with its order count, notional, mean `%CLOSE` and mean `vs Close`**.

Read that table, pin `CLOSE_STRATEGIES` at the top of the script, rerun without
`--probe`. Left empty, the close algos are *derived* (by name pattern or by mean
`%CLOSE`) and the run says loudly that it did so.

## What makes this a MOC review rather than a generic one

`vs Close` is degenerate for a close algo: an order that clears in the auction
executes *at* the close, so its close slippage is ~0 by construction. Judging
close algos on it produces a chart of near-zero bars that says nothing.

Two things fix that.

**1. `%CLOSE` measures the auction/continuous split directly.** With
`%OPEN + %CLOSE + %POST + %TAKE + %DARK = 100`, whether the order reached the
auction is a measurement, not an inference. The degeneracy then becomes useful:
the auction portion contributes ~0 to `vs Close`, so the whole of that slippage
is carried by the portion that *missed*. Dividing by the missed share recovers
what that portion achieved against the close it did not get —
`implied continuous vs Close = Close / (1 - %CLOSE/100)` — and times notional,
that is the cost of leakage in currency.

**2. The benchmarks share one executed price, so their differences are pure
price moves.** That gives the decomposition with no extra columns:

```
vs Arrival  =  (vs Arrival - vs Close)  +  (vs Close)
                cost of waiting            execution vs the auction print
                for the close
                -- the schedule            -- the algo
```

and two more: `Vwap - Close` (was the close a better place to trade than the
session — the venue-choice question), and `NextOpen - Close` (reversion;
negative = the price moved back against us, i.e. temporary impact we paid).

## Size is a confound, and it is controlled for

A large order legitimately begins before the close because the auction cannot
absorb it. Continuous execution is **expected**, not a defect. So the run fits a
**capacity frontier** — the median achieved `%CLOSE` within each
(market, `%Adv`) cell — and an order is only called short when it sits below the
frontier *for its own size*. That is the difference between a defensible finding
and a wrong one.

## The miss taxonomy

Without close-eligibility tagging, a missed auction cannot be attributed to a
mechanism inside the algo. `%CLOSE` still measures the outcome, so the fact is
solid; only the cause is unavailable. Causes are therefore assigned by
exclusion, in precedence order:

| Cohort | Test |
|---|---|
| Never traded | `FR < 1` |
| Cleared the auction | `%CLOSE >= 90` |
| Size explains it | `%Adv >= 5` |
| Limit did not cross | no auction fill, and a limit order |
| **No auction fill — unexplained** | traded, small, no limit, still nothing |
| Partial — below the frontier | partial fill, >15pp under the frontier |
| Partial — in line with peers | partial fill, at or above the frontier |

The unexplained cohort is the deliverable: those orders are listed individually
by id, date and symbol in `16_orders_to_review`, with the statement that the
cause is not in this file. If the cohort is large, that is itself the argument
for getting tagging into the extract.

## Conventions

- **Positive bps = savings** everywhere. The sanity report re-derives it from the
  data and shouts on a mismatch. Every axis carries `← cost | savings →`.
- **`$Mln` is executed notional; `#Shares` is the order quantity.** Different
  bases — `$Mln / #Shares` is *not* an average price and is never computed.
  Executed shares = `#Shares × FR/100`.
- Every aggregate is **notional-weighted**. Medians, hit rates and box
  distributions are per-order by design and labelled as such.
- Means are **winsorised at 1/99**; the tails are pulled in, never dropped.
- CIs are **95% percentile bootstrap**, 2,000 draws — slippage is fat-tailed and
  a normal-theory interval would be too narrow. Under n=8 there is no CI and the
  row is flagged small-sample. **A CI crossing zero means not distinguishable
  from zero**, and the findings block says so per line.
- **India is never pooled with the rest.** NSE has no single-price closing
  auction — it closes on a VWAP of the last half hour — so its "MOC" is a
  different product and is reported separately.
- Market close times are in HKT. Hong Kong, Japan and Australia are verified
  against the desk's own session windows; the rest are derived from published
  exchange hours and are **flagged UNVERIFIED wherever they affect a number**.
  Australia shifts an hour against HKT under AEDT and is handled per date.
- Unrecognised `Cap` values are **kept and shown**, never dropped, and named in
  the run log.

## What this analysis cannot show

Stated on every run, so the gaps are explicit rather than discovered late:

- Whether an order was **tagged** for the close. Without that flag, an order that
  worked out in continuous cannot be told apart from one never meant for the
  auction.
- **Our share of the closing auction** — capacity is expressed as `%Adv` rather
  than as a share of the auction itself. Closing-auction volume by sym/date would
  fix this and may be cheap to get.
- Whether we traded **into or against the published imbalance**.
- Any impact claim beyond what `NextOpen - Close` supports.

## Charts

Palette validated against the computable checks — OKLCH lightness band, chroma
floor, protan/deutan ΔE on adjacent pairs, normal-vision floor, WCAG contrast vs
the surface. The categorical set passes with zero failures; three hues sit in the
sub-3:1 contrast relief band, so those charts carry direct labels rather than
relying on colour alone. The ordinal ramp is evenly spaced in OKLab at ΔL≈0.087.
One axis per panel — two measures of different scale get two panels, never two
y-scales.
