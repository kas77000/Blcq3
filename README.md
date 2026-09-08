# MOC / close-algo TCA

Close-algo review covering orders **to 4 September 2026**. **Tables and charts,
no deck** — the narrative comes out of the numbers, not the other way round.

**The trading platform did not change inside the window.** This client has not
migrated, so there is no platform cutover to split on and no date filter is
needed — `DATE_FROM` / `DATE_TO` stay `None`.

**One market did change inside the window.** India's Closing Auction Session
went live on 3 August 2026, so India carries three regimes and is never pooled.
See the India note under Conventions.

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
| Auction only | `%CLOSE >= 99.5` |
| Cleared the auction | `%CLOSE >= 90` |
| Size explains it | `%Adv >= 5` |
| Limit did not cross | no auction fill, and a limit order |
| **No auction fill — unexplained** | traded, small, no limit, still nothing |
| Partial — below the frontier | partial fill, >15pp under the frontier |
| Partial — in line with peers | partial fill, at or above the frontier |

`%CLOSE` at 100 means every executed share printed in the auction, so the order
behaved as **auction only**. That is an outcome rather than a permission: an
order free to trade continuously that happened to fill entirely in the auction
looks identical, and an order that never filled cannot be classified at all. It
is enough to split the cleared population into orders that never left the
auction and orders that needed continuous help to get there. It is not enough
to explain a miss, and the run log says so.

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
- **India is never pooled, and it carries three regimes.** It closed on a VWAP
  of the last half hour until the Closing Auction Session went live on
  **3 August 2026** — a 20-minute call auction, 15:15 to 15:35, referenced to
  the 15:00–15:15 VWAP, and **only for stocks in the derivatives segment**.

  Before that date there is no auction to reach: auction share means nothing,
  and `vs Close` is a genuine tracking result rather than a degenerate one —
  you cannot print at a VWAP, you have to work the last half hour to track it.

  After it, the extract carries no segment flag, so an F&O name and an ordinary
  one cannot be told apart. It cannot be inferred from `%CLOSE` either: an F&O
  name whose order *missed* the auction looks exactly like a name that never
  had one, and inferring would drop India's misses out of the auction
  population — the very orders the review exists to find. So post-CAS India is
  its own regime, pooled with neither side and asserted about nothing, and it
  appears in `32_close_regimes`.

  To close this properly: get the F&O eligibility list, drop `India` from
  `AUCTION_SEGMENT_UNKNOWN`, and filter on the symbol.
- Market close times are in HKT. Hong Kong, Japan and Australia are verified
  against the desk's own session windows; the rest are derived from published
  exchange hours and are **flagged UNVERIFIED wherever they affect a number**.
  Australia shifts an hour against HKT under AEDT and is handled per date.
- Unrecognised `Cap` values are **kept and shown**, never dropped, and named in
  the run log.
- The export is **already side-adjusted** — plus is good, minus is bad, on both
  sides. `SIDE_ADJUSTED = True` records that, and the by-side means are then
  printed as a result to explain rather than as a data error.

## Scope, and what leaves the study

`STRATEGY_SCOPE` pins the review to **VWAP and CLOSE**. CLOSE puts 53% of its
value through the auction and VWAP only 7.6%, but VWAP is four times the book,
so it carries 37% of every dollar this client sends to a closing auction.
Between them the two are 98.8% of it. Keeping CLOSE alone would understate the
close footprint by more than a third and hide the algo-selection question,
which is usually worth more than algo performance. Every strategy dropped is
named in the run log with its order count and value.

The miss taxonomy runs on `MOC_STRATEGIES` only. A VWAP order was never aiming
at the auction, so calling its low auction share an unexplained miss would be
nonsense. VWAP keeps its benchmark, venue and decomposition tables and stays
out of clearance, capacity and the cohorts. The capacity frontier is computed
**within strategy** — pooling a 53% algo with a 7% one would drag the reference
line down until nothing looked short.

Exclusions work at the **value** level, not the order level. An infinity in
`NextOpen` says nothing about that order's auction share, size or notional, so
the cell is cleared and the order stays in every table its other columns can
support. Dropping whole orders for one bad cell would bias the rest, because
orders with broken cells are not a random sample. Only orders that cannot
contribute anywhere leave — no notional, no quantity, no readable side — and
each exit is counted with its value. **Nothing is ever removed for being
large**: winsorising handles the tails, and deleting the extremes would delete
the orders the review exists to find.

## Basis points for money, spreads for comparison

Slippage in bps answers "what did it cost". It does not compare across names or
markets: a wide-spread mid-cap costs more bps than a large-cap for reasons that
have nothing to do with the algo, so ranking on bps ranks the names. Dividing
by the spread gives "how many spreads did we pay", which does compare.

`eIS/Sprd`, `ePvwap/Sprd` and `Pvwap/Sprd` are read straight from the export
where it carries them and derived from `slip / spread` where it does not, so
either shape of file behaves the same. Tables `06a`–`06c` put the two side by
side, by strategy, by market and by size band.

The ratio is taken **between the two averages**, never as the average of
per-order ratios — one name with a 0.5bp spread would otherwise produce a ratio
in the hundreds and dominate the mean. The median per-order ratio sits beside
it for what the typical order paid.

## What this analysis cannot show

Stated on every run, so the gaps are explicit rather than discovered late:

- Whether an order was **tagged** for the close. `%CLOSE = 100` identifies
  orders that *behaved* as auction-only, but it is an outcome, not a
  permission, and it says nothing about an order that never filled. So an
  order that worked out in continuous still cannot be told apart from one
  never meant for the auction.
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
