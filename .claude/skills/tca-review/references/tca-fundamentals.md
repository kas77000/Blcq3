# TCA fundamentals

What you need to hold in your head before judging anyone's execution.

## The sign convention comes first

Before reading a single number, establish which way the sign runs, and write it
into `meta.json`. Two conventions are in the wild:

- **Positive = savings** (you beat the benchmark). Common on buy-side desks.
- **Positive = cost** (you paid away). Common in broker reports.

Get this backwards and every conclusion inverts. Two ways to settle it:

1. A cost-like benchmark that should be negative on average. Spread-crossing
   costs money, so aggregate arrival slippage is usually slightly adverse.
2. Ask. It is one message and it removes the risk entirely.

State the convention on the deck's "how to read this" slide in plain words:
"plus means you did better than the benchmark".

## The benchmarks and the question each one answers

| Benchmark | The question it answers | Careful of |
|---|---|---|
| **Arrival** | What did the whole order cost from the moment we had it? | Contains market drift the algo did not cause |
| **PVWAP** (interval / participation VWAP) | Did we beat the market **over the window we actually traded**? | The fairest algo score. Needs correct start and end stamps |
| **VWAP** (full session) | Did we beat the whole day? | Unfair to any order that did not run the whole day |
| **TWAP** | Did we beat an even spread over the window? | Only meaningful for time-based schedules |
| **Close** | How did we do against the closing auction print? | ~0 by construction for anything that cleared the auction |
| **Open** | How did we do against the opening print? | Same degeneracy as Close for opening algos |
| **Next open** | Did the price come back after we finished? | Reversion, not cost. Overnight news is noise here |
| **Decision / IS** | What did the idea cost from decision to done? | Includes delay the desk may not control |

**One benchmark is never enough.** A single number invites the algo to be tuned
to that number at the expense of the others. Show at least the order's own-window
benchmark plus arrival, and say what each one means.

## The decomposition that needs no extra data

When benchmarks share one executed price, their differences are pure price moves.
That gives an attribution with nothing more than the columns you already have:

```
vs Arrival  =  (vs Arrival - vs Benchmark)  +  (vs Benchmark)
                   the cost of the schedule      the algo's own work
                   - when we chose to trade      - how it traded
```

For a close algo, the benchmark is the close and the first term is the cost of
waiting for the auction. For a VWAP, the benchmark is PVWAP and the first term is
what the timing of the order cost against the day.

This is the single most useful slide in most reviews: it separates **what the
client decided** (when to send it) from **what the algo did** (how it worked it).
Those two have different owners and different fixes.

Two more differences worth showing when the columns exist:

- `VWAP - Close` - was the close a better place to trade than the session?
- `NextOpen - Close` - did the price move back after we were done? A move back
  means part of what we paid was temporary.

## Size is a confound and must be controlled

A big order costs more. That is physics, not failure. Any comparison that does
not hold size constant is measuring order size, not execution quality.

Control by comparing within size bands, expressed as **% of average daily
volume**. The practical bands:

| Band | %ADV | What it means |
|---|---|---|
| Small | < 1% | The market absorbs it without noticing |
| Medium | 1-5% | Normal working size |
| Large | 5-15% | Needs a schedule; costs are expected |
| Very large | > 15% | Multi-day or high impact; judge separately |

Volatility and spread are the next two confounds. If a strategy's orders sit in
wider-spread names, part of its cost is the names, not the algo. Say so.

## Is the difference real?

Slippage is fat-tailed and skewed. Three rules keep you honest:

1. **Weight by value, and say so.** A notional-weighted average answers "what did
   this cost us in money"; an unweighted one answers "what did the typical order
   do". They differ, often a lot. Show one, name it, be consistent.
2. **Pull the tails in, do not cut them off.** Winsorise at 1/99. Dropping
   outliers hides the exact orders most worth discussing.
3. **State the sample size next to every number.** Rough guide, and label it as a
   guide rather than a test:

| n | How to talk about it |
|---|---|
| < 8 | Show the orders individually. No average. |
| 8-30 | "Indicative". Never "better" or "worse". |
| 30-100 | A difference of a few bps is suggestive, not proven |
| > 100 | An average is meaningful; still quote a range |

If confidence intervals exist in the source, use them, and translate: an interval
that crosses zero means **"we cannot tell this apart from zero"**. Say that in
those words, not "not statistically significant".

## Fill rate is part of the score, never a footnote

The easiest way to look good on slippage is to not trade. A passive algo that
fills 60% and shows +3bps has not beaten anything - it has left 40% of the order
undone, and the client still needs those shares, probably at a worse price.

**Any slippage figure for a passive, dark or opportunistic strategy is shown with
its fill rate on the same slide.** Where unfilled shares can be priced against a
later mark, price them and add the cost back.

## Reversion, and what it is allowed to prove

Compare the price shortly after the last fill, and again at the next open, with
the price we paid.

- Price stays where we pushed it → the move was probably information.
- Price comes back → we paid temporary impact. Trading slower or smaller may
  recover part of that.

Reversion is evidence about impact. It is not a cost number and must not be added
to one. One overnight gap on a small sample proves nothing.

## Common pitfalls

| Pitfall | What goes wrong |
|---|---|
| One benchmark for everything | Hides the trade-offs and invites tuning to the number |
| Full-session VWAP on short orders | Scores the market's drift, not the algo |
| Comparing algos without size control | You measure order size, not skill |
| Averaging across markets | One market's structure dominates the average |
| Dropping outliers | Deletes the orders that most need discussing |
| Ignoring unfilled shares | Rewards not trading |
| Ranking on a 12-order sample | Noise presented as a finding |
| Mixing order-level and share-level bases | `notional / shares` is not an average price |
| Pooling a market with a different auction | See `moc-close-algos.md` on India/NSE |
