# MOC TCA — the column list

H1 review, close algos. **One source: the order extract.** No kdb.

Three facts you gave that shape everything below:

- **`$Mln` is EXECUTED notional. `#Shares` is the ORDER quantity.** Different
  bases. `$Mln / #Shares` is **not** the average price and is never computed.
  Executed shares = `#Shares x FR/100`.
- **`%OPEN + %CLOSE + %POST + %TAKE + %DARK = 100`.** The auction/continuous
  split is measured, not inferred.
- **`first_exec_vs_close` is bps: the close price against the first execution
  price.** Not a time. It says whether starting early paid — see 4B.

---

## 1. Required

| Column | Use |
|---|---|
| `Strategy` | Selects the close algos. `--probe` prints every observed value with count, notional, mean `%CLOSE` and mean `Close`; you tick the close ones into a one-line config. |
| `Date` | H1 scope, monthly trend. |
| `Sym` | Market from the suffix. India split out — **NSE has no single-price closing auction**, it closes on the last-30-min VWAP, so its "MOC" is a different product and is never pooled with HK/JP/AU/KS/TT. |
| `Side` | Buy/sell split, and the check that slippages are side-adjusted. |
| `$Mln` | Executed notional — the weight on every aggregate, and the currency P&L. |
| `#Shares` | Order qty. With `FR`, gives executed qty and the unfilled residual. |
| `FR` | Fill rate. A missed close cannot be retried — the day is gone, so `FR` carries more weight here than in any other TCA. |

## 2. Venue mix — the core

| Column | Use |
|---|---|
| `%CLOSE` | Share of executed qty in the closing auction. **The most important column in the file.** |
| `%TAKE` | Continuous, aggressive. |
| `%POST` | Continuous, passive. |
| `%DARK` | Continuous, dark. |
| `%OPEN` | Opening auction. On a close algo this should be ~0 — **if it is not, that is a finding**, reported separately rather than folded into "continuous". |

`pct_continuous = %TAKE + %POST + %DARK`. The five are checked to sum to 100 per
order; rows that fail are listed in the run log, never silently normalised.

## 3. Benchmarks

| Column | Use |
|---|---|
| `Close` | Slippage vs the closing price. With `%CLOSE` this becomes the money number: the auction portion is ~0 by construction, so **implied continuous execution vs the close = `Close / (1 - %CLOSE/100)`** — how much worse the portion that missed the auction did than the auction it missed. Times notional = the cost of leakage in dollars. |
| `IS` | Arrival slippage — the headline, and the only non-degenerate performance number for a close algo. |
| `Pvwap` | Interval VWAP — separates market drift over the order's life from execution inside it. |
| `Vwap` | Day VWAP. `Vwap - Close` is the side-adjusted move from the session average to the close: **"should this flow have gone MOC at all"**, with no extra column. |
| `NextOpen` | Confirmed a bps slippage. `NextOpen - Close` is the side-adjusted move from the closing price to the next open — **reversion**. Negative = the price moved back against where we traded = temporary impact we paid. The only route to an impact claim. |
| `Open` | Low value here. Include if free. |

## 4. Capacity and behaviour

### 4A. The confound you raised

Large orders legitimately begin before the close, because the auction cannot
absorb them. Continuous execution is therefore **expected**, not a defect, and
the analysis must control for size before calling anything a miss.

| Column | Use |
|---|---|
| `%Adv` | **The control variable for the whole review.** Method: fit achieved `%CLOSE` against `%Adv` per market to get a capacity frontier — how much of an order this size can realistically clear in the auction — then flag only orders **below the frontier for their own size**. That separates "had to leg it" from "leaked when it did not need to". |
| `Adv` | Denominator behind `%Adv`. |
| `fPR_cont` | Continuous participation rate. Asks whether the pre-close portion was worked too hard — a high `fPR_cont` on flow that could have waited is a costed, specific recommendation. |

### 4B. `first_exec_vs_close` — close price vs first execution price, in bps

This is the column that answers **"was starting early right?"** at order level,
which no other field can.

- **Positive** (first fill better than the close): legging in early *paid*. The
  order was right to start when it did — and if `%CLOSE` was high anyway, it may
  have been right to start *earlier still*.
- **Negative** (first fill worse than the close): the early start *cost*. If
  `%Adv` was also low, the order had no capacity reason to start early — it gave
  up money for nothing.

Weighted by the continuous share of each order, this aggregates into a single
defensible statement about whether the close algos begin working too soon —
which is exactly the guidance the client can act on.

### 4C. Segmentation and controls

| Column | Use |
|---|---|
| `Cap` | large / mid / small / micro / other, matched case-insensitively; **unrecognised values are kept and shown, never dropped** — the run log names any value not in the expected set. Auction depth differs by an order of magnitude across these, so this is the main segmentation. |
| `marketLimit` | Market or Limit. **A limit that cannot cross the auction price is a mechanical miss**, and `%CLOSE = 0` on a Limit order is its signature. Likely the most actionable finding available. |
| `Start(HK)` / `End(HK)` | Working interval. `End` against the market's continuous end gives `close_gap_min` — separates "finished early" from "worked to the bell and still missed". |
| `Sprd` | Difficulty control; also prices the leakage. |
| `Vol` | Difficulty control. |
| `PR` | Overall participation. |
| `arrivalTime` | Pre-Open / First30Mins / Day / Last30Mins — how much runway the order was given. |
| `sector` | Segmentation. |
| `client` / `Trader` | Sub-account or desk breakdown, if behaviour differs by desk. |
| `aggrTgtId` | Order id — dedup, and naming specific orders in the exception list. |

---

## 5. What losing kdb costs, and what replaces it

Gone: close-eligibility tagging (was the order marked for the close, under
whatever the previous platform called that flag) and the MOC-sent funnel
(size reserved, size sent, venue rejects). So **we can no longer attribute a
missed auction to a mechanism inside the algo.**

What survives is stronger than the earlier proxy, because `%CLOSE` measures the
outcome directly. `%CLOSE = 0` on a close-algo order is a *fact*, not an
inference. Only the cause is unavailable — so causes are assigned by exclusion,
in a taxonomy built from the extract alone:

| Cohort | Test | Reading |
|---|---|---|
| Never traded | `FR = 0` | Not an auction problem. |
| Blocked by a limit | `%CLOSE = 0`, `FR > 0`, `marketLimit = Limit` | Mechanical: the limit did not cross the auction price. |
| Size-driven | `%CLOSE` low, `%Adv` high | Correct behaviour. Explicitly cleared, not counted against the algo. |
| **Unexplained** | `%CLOSE = 0`, `FR > 0`, `%Adv` low, `marketLimit = Market` | **The deliverable.** Orders that had every reason to clear the auction and did not, with no benign explanation in the data. Listed individually by `aggrTgtId`, `Date`, `Sym` for the desk to check against the logs. |

That last cohort is the honest version of the finding: the analysis names the
orders and states plainly that the cause is not in this file. If it turns out to
be large, that is itself the argument for getting the tagging data.

---

## 6. Out of scope, and stated as such

| Wanted | Effect |
|---|---|
| Auction market volume per sym/date | Capacity stays expressed as `%Adv` rather than as our share of the auction. `%Adv` is a fair proxy; auction depth is the real constraint and differs most in small and micro caps. |
| Published imbalance | Cannot say whether we traded into or against the imbalance. |
| Auction price level vs last continuous price | Cannot say whether the auction printed away from the book. |
| Close-eligibility tagging | Cannot separate "the algo missed the close" from "the order was never marked for it". See the unexplained cohort above. |

No impact claim will be made beyond what `NextOpen - Close` supports, and no
cause will be asserted for the unexplained cohort.
