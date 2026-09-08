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

## Two windows off one extract

`--from` / `--to` / `--label` override the constants, so a second window needs
no code edit and no second copy of the script. Give each one its own `--out`:

```bash
python moc_tca.py --data orders.csv --to 2026-06-30        --out output_h1   --label "H1 2026"
python moc_tca.py --data orders.csv        --out output_full --label "Jan to 4 Sep 2026"
```

H1 is the clean baseline: one platform, one market structure, India on a VWAP
close throughout. The full window adds July to 4 September and crosses the
India auction change on 3 August, which the run log says out loud.

**If the question is "has anything moved", run the two windows DISJOINT**, not
nested:

```bash
python moc_tca.py --data orders.csv --from 2026-07-01        --out output_h2 --label "Jul to 4 Sep 2026"
```

H1 against the full period compares a window with a window that contains it,
so two thirds of the orders are on both sides and any change is diluted by its
own baseline. H1 against Jul-Sep compares two separate populations, which is
the comparison that can actually move.

Each run stamps `run_window.txt` in its output directory, and a run that lands
on a directory holding a different window says so before overwriting it.

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
- **Short sells keep their own label.** The house code is `SSH` and it is 13%
  of the book. A short sell is still a sell, so the arithmetic is unchanged,
  but locate requirements and short-sale rules can change how an order
  executes — so buys, sells and short sells are reported as three groups
  (`33_by_side`) rather than two. A side code the script does not know is
  never silently dropped: it is named, and past half a percent of the book the
  run stops and asks for it.

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

The spread-normalised columns are **optional**. Whatever the export carries is
read straight from it; everything else is computed as `slippage / spread`, so a
file with all of them, some of them or none behaves identically. A missing one
is never reported as a gap. The run log names which were read and which were
computed, so the two are not confused for each other.

Where a file has both a plain and an `e`-prefixed variant, the plain one wins,
because it pairs with the same benchmark as the bps column beside it — taking
`ePvwap/Sprd` next to a `Pvwap` slippage would put two different benchmarks in
one row. Tables `06a`–`06c` put bps and spreads side by side, by strategy, by
market and by size band.

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

## Markets

`34_market_profile` puts each market's size and cost in one place, ordered by
value traded: orders, share of orders, notional, share of notional, the
weighted spread, auction share, slippage against arrival, PVWAP and the close
in both bps and spreads, fill rate and median size. Two charts come off it —
`13_market_notional` for where the value went, `14_market_slippage` for what
it cost.

Both are **ordered by share of value, never by cost**. Ordering by cost puts a
sixty-order market at the top of the slide, and each bar carries its own share
of the book so nobody has to guess which ones can actually move the number.

`35_market_by_strategy` splits the same rows by strategy, because a market
effect can be a mix effect: if one market is nearly all VWAP and another nearly
all CLOSE, comparing the markets compares the strategies as much as the venues.
Check this before attributing anything to a market.

## Charts

Palette validated against the computable checks — OKLCH lightness band, chroma
floor, protan/deutan ΔE on adjacent pairs, normal-vision floor, WCAG contrast vs
the surface. The categorical set passes with zero failures; three hues sit in the
sub-3:1 contrast relief band, so those charts carry direct labels rather than
relying on colour alone. The ordinal ramp is evenly spaced in OKLab at ΔL≈0.087.
One axis per panel — two measures of different scale get two panels, never two
y-scales.
