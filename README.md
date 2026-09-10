# MOC / close-algo TCA

**H1 2026** review of the MOC product. **Tables and charts, no deck** — the
narrative comes out of the numbers, not the other way round.

`DATE_FROM` / `DATE_TO` are pinned to **2026-01-01 .. 2026-06-30**. The extract
runs past June, so the filter is doing real work, and `--from` / `--to`
override it for a second window without touching the file.

**The trading platform did not change inside the window**, and neither did any
market. This client has not migrated, so there is no platform cutover to split
on. India's Closing Auction Session went live on 3 August 2026, which is after
the period ends, so India is on a VWAP close throughout and the CAS split never
fires — the run log confirms that on every run rather than leaving it to be
assumed. The machinery stays in place because the extract itself extends past
the date; see the India note under Conventions.

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
| `unified_tables.xlsx` | the same tables stacked down **one** sheet, titled and spaced, for capturing a whole run in a few screenshots instead of thirty tab clicks |
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

  **Known and unresolved:** India reports `%CLOSE` as 0 on every order, CLOSE
  orders included, in both regimes. So it is not only the segment that cannot
  be identified — auction share is not measured for that market at all. India
  is about a third of the book, and while it sits inside the reviewed
  population its zero drags the overall weighted auction share down. Read the
  auction-market figures, not the all-market ones, until this is settled.
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

## The AWS extract

A second source, one parquet per date/country/client, holding the same orders
with more columns on them. Drop them in `data/aws/` and the run picks them up;
with the folder empty it runs on the order file alone and says so.

Every parquet is concatenated — files with a different column set are unioned
and named in the log rather than dropped — deduplicated on `aggrTgtId`, and
**left-joined onto the order file**. The order file stays the population: an
order with no match is kept, and only the columns that came from AWS are empty
on it.

**Name collisions are resolved on the normalised name, not the exact one.**
`resolve_columns` is case- and separator-blind, so an AWS `sym` sitting beside
the order file's `Sym` is not a harmless near-miss — both answer to the same
logical field, and whichever came first in column order would win. Every
collision, exact or not, keeps the order file's version and renames the AWS one
to `<name>_aws`. So nothing that already works can change underneath a run, and
the alternative is there by name when it is wanted.

## India, without an auction share

India runs no closing auction in this period and the platform reports `%CLOSE`
as 0 on every one of its orders, so auction share cannot say whether an order
was aimed at the close. **When it started can.** `fstart_time` comes from the
AWS extract, and India orders that began outside `INDIA_CLOSE_WINDOW_HKT` are
not close orders whatever strategy label they carry, so they leave the study.
Only India is touched; every other market keeps all of its orders.

The window is `17:30–17:45` HKT. Worth checking against the desk: NSE closes
15:30 IST — 18:00 HKT — and the pre-CAS closing VWAP runs over the last half
hour, so the full window is 17:30–18:00 HKT and this is the first half of it.
Widening it is a one-value change.

Without `fstart_time` nothing is filtered: the run says India cannot be
windowed and leaves those orders in, flagged, rather than filtering on a column
it does not have.

**The India orders that survive are then given `%CLOSE = 100`.** They ran
through the VWAP that *is* the close, so their close share is 100 by the
mechanism rather than by measurement — and leaving it at 0 would drag every
all-market auction-share figure down with a number that measures nothing.

It is **imputed, not observed**, and nothing hides that: the original value
stays in `pct_close_measured`, the rows are flagged in `pct_close_imputed`, the
run log says how many were set, and the findings block names the count and
tells you to say so on the slide. Any auction-share figure that includes India
is part measured and part assumed. Set `INDIA_CLOSE_PROXY = False` to leave the
zeros alone.

Because it is set before any venue field is derived, `close_notional`,
`pct_continuous`, `close_bucket` and the rest all follow from the same number
rather than from a stale one.

India still stays **out of the auction-market tables** — clearance, capacity,
the cohorts. Its close is a half-hour VWAP, not a single print, so "did it
clear the auction" is not a question that can be asked there, and 100% imputed
rows would answer it with an assumption.

## Seeing the merged data for yourself

`merge_aws.py` writes the two intermediate files out as CSV:

```bash
python merge_aws.py --orders orders.csv --aws data/aws
```

- `aws_all.csv` — every parquet in one frame, deduplicated on `aggrTgtId`
- `merged.csv` — the order file with the AWS columns joined on

It **imports the join from `moc_tca.py`** rather than writing it again, so what
lands in `merged.csv` is exactly what the analysis sees: same collision rules,
same direction, same suffix. Nothing about the script can drift away from the
run. `--aws-only` writes the concatenated frame and stops.

## The pipeline, in order

1. Read `orders.csv` — this is the population, and it stays the population.
2. Concatenate every parquet in `data/aws/` into one frame, dedupe on
   `aggrTgtId`, and left-join it on. Name collisions resolve on the normalised
   name and the order file wins, unless its column is empty.
3. **Keep only the orders that had a closing auction to reach**
   (`REQUIRE_CAS_ELIGIBLE`, measured by `marketCloseSize > 0`). Markets in
   `NO_CLOSING_AUCTION` are exempt — India runs no auction in this period, so
   its zero is about the market rather than the order, and applying the test
   would delete a third of the book on a technicality. India keeps its own
   window filter, its own imputed close share, and its own section.
4. Window India, scope the strategies, filter the period, clear unusable
   values, drop orders that cannot contribute anywhere.
5. Everything else — every check, table and chart — runs on what survives.

Step 3 is an **opportunity** test, not an outcome test. A day with no auction
puts a zero in the denominator of every auction-share figure for a reason that
has nothing to do with the order or the algo. It never removes an order that
reached for an auction and missed — those are the finding, and they stay in
the cohorts. The run prints what it dropped per market.

## Three levers, not just a status

A status report says how the close product did. These say what to change.

**`38_auction_share` — our share of the closing auction.** `%Adv` measures an
order against a normal day; the auction is not a normal day, it is one print,
and how much of it we were is the constraint that actually binds. Built from
`fillCloseSize / marketCloseSize` in the AWS extract, by market and size band.
An order that took a large share of the auction and missed the rest hit
capacity. An order that took a sliver and still missed did not — and those are
opposite fixes.

**`39_lateness` — how late the order arrived.** `fstart_time` against each
market's own close, banded from "0–2 min" out to "more than 2 hours", with
auction share and the miss rate for each band. The desk controls when an order
is sent, so if auction share falls away inside the last few minutes that is a
lever — and nobody can pull it without knowing where the cliff is. A band
labelled *after the close* is a warning, not a finding: it means the close time
in `MARKET_CLOSE_HKT` is wrong for that market.

**"Limit could not cross" — a real cause in the taxonomy.** A buy limit under
the close, or a sell limit over it, could not have traded in the auction. Built
from the order's limit against the closing price, and tested *before* the
residual, so those orders stop being filed as unexplained. `market_limit` could
never answer this because it reads "Limit" on every order.

**Eligibility comes from `fillCloseSize`, not from a flag.** A positive fill
*is* the order having been in the auction — that day, in that name — which is
an outcome rather than a permission, and no flag can beat it. `marketCloseSize`
adds the other half: where the auction itself had no size there was nothing to
miss, so those orders get their own cohort, **"No auction that day"**, tested
ahead of every other cause. Filing them as failures would send the desk hunting
a cause that cannot exist.

The two also **check each other**. `%CLOSE` and `fillCloseSize` answer the same
question — was this order in the auction — from different systems, so the run
reports how often they agree and names the disagreements. Where they differ,
trust `fillCloseSize`: it is a quantity actually printed in the auction, while
`%CLOSE` is a share of executed quantity that this export has already been
caught leaving empty.

**`auctionOnly` is deliberately not used.** It reports the market's mechanism
as much as the order's permission — India's close orders come back
`ContinuousOnly` because India has no auction to be eligible for — so reading
it as permission would rule out a third of the book for a reason that was
never about the order.

## Opportunity is not outcome

`37_close_opportunity` splits every market into orders that were still live
into the closing window and orders that had finished before it. An order that
ended at 11am never had a chance to reach the auction, so judging it as a close
order says nothing — that is the honest analogue of the India window.

**This is not the same as `%CLOSE = 0`, and the difference decides whether the
review has a finding in it.** A zero on an order that *was* live into the close
is a real outcome, and those orders are the most valuable population in the
run: they could have cleared and did not. Dropping them would leave only the
orders that worked, push auction share toward 100% by construction, and have
the deck conclude that everything clears because everything that did not was
removed.

So it flags and drops nothing by default. `DROP_NO_CLOSE_OPPORTUNITY = True`
removes them, after reading the table.

**The table checks itself.** An order that finished before the close cannot
also have printed in the auction. A market showing both is not telling you
about its orders — it is telling you its close time in `MARKET_CLOSE_HKT` is
wrong, and it names the ones flagged UNVERIFIED. Do not filter on this column
for a market until its close time is confirmed against the desk's own sessions.

## Scope, and what leaves the study

`STRATEGY_SCOPE` pins the review to **CLOSE**. This is a review of the MOC
product, and on this platform the CLOSE label already covers it end to end —
MOC and IIS both report under it. Every strategy dropped is named in `01_scope`
and in the run log with its order count and its value.

**State the consequence on the deck rather than leaving it implicit.** The
client's VWAP flow also reaches closing auctions — on this book it is around
70% of the value traded — so this describes the **MOC product**, not the
client's total auction footprint. Those are different questions and only one
of them is answered here.

With a single strategy in scope there is nothing to compare it against, so
`35_market_by_strategy` and `36_algo_choice` are skipped and the run log says
why. Both come back on their own if a second strategy is added to the scope.

The capacity frontier is computed **within strategy**, so it stays correct
either way.

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
Check it before attributing anything to a market. With one strategy in scope
it would only repeat `34`, so it is skipped.

## Money, and the algo-choice question

Every notional-weighted table carries a `saved (USDk)` column beside its bps:
`bps x notional / 10`, positive still meaning savings, so a negative figure
reads as a cost without a second convention to remember. The early-start
tables are weighted by **continuous** notional, so their money is that portion
of the order rather than the whole of it.

`36_algo_choice` compares two strategies on the same market and the same size
band, and `15_algo_choice` charts it. Every cell needs at least 30 orders on
**both** sides, so with `STRATEGY_SCOPE` on a single strategy the table is
skipped entirely. What follows applies when a second one is in scope.

**It is not a controlled comparison, and must never be presented as one.**
Orders are not assigned to a strategy at random. A trader choosing CLOSE for
one order and VWAP for another is acting on urgency, on a view, on an
instruction the extract does not carry — and those reasons drive cost too.
Holding market and size constant removes the two biggest confounds and leaves
the rest standing. So read a row as *"orders like these, on this algo, cost
this much"* — a question for the desk — and never as *"the other algo would
have saved that"*. The table, the chart footnote and the run log all say so.

Colour on that chart is the **algo**, not the direction, so it deliberately
avoids the blue/red diverging pair every other chart uses — otherwise a blue
bar at −37bps would read as good news.

Each row also carries the same cells measured against **PVWAP**, and
`gap that is execution %` — how much of the arrival gap survives once every
order is scored against its own window. That column is the test that matters.
Against arrival an order carries every basis point the market moved while it
worked; against PVWAP it does not. A gap that survives is the algos working
differently. A gap that collapses is the two sets of orders having faced
different markets — a routing question, not a quality one. Without it a clean
sweep for one algo reads as "always use that one", which is the exact
conclusion the data cannot support.

## Charts

Palette validated against the computable checks — OKLCH lightness band, chroma
floor, protan/deutan ΔE on adjacent pairs, normal-vision floor, WCAG contrast vs
the surface. The categorical set passes with zero failures; three hues sit in the
sub-3:1 contrast relief band, so those charts carry direct labels rather than
relying on colour alone. The ordinal ramp is evenly spaced in OKLab at ΔL≈0.087.
One axis per panel — two measures of different scale get two panels, never two
y-scales.
