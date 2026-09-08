# Auction algos - MOC, LOC, MOO

Why a close algo needs its own treatment, and what to measure instead.

## Distance from the close is a dead number

An order that clears in the closing auction executes **at** the closing price. Its
slippage against the close is zero by construction. Charting it produces a row of
near-zero bars that look like excellent execution and carry no information.

So do not lead with it. Two things replace it.

### 1. Measure how much of the order reached the auction

If the data carries the split of executed quantity across open, close, post,
take and dark - they sum to 100 - then whether the order reached the auction is
a **measurement**, not an inference. That single number is the honest score for a
close algo.

The degeneracy then becomes useful. The part that cleared the auction contributes
nothing to the slippage against the close, so all of that slippage is carried by
the part that missed. Dividing by the missed share recovers what the missing part
achieved against the close it did not get:

```
what the missed part achieved = (vs Close) / (1 - auction share)
```

Times the value of the missed portion, that is the cost of missing the auction,
in money. That is the number the client cares about, and it is the one to lead
with.

### 2. Separate waiting from executing

```
vs Arrival  =  (vs Arrival - vs Close)  +  (vs Close)
                the cost of waiting        what execution
                for the close              added or saved
                - the client's decision    - the algo's work
```

The first term prices the decision to hold the order until the close. The second
is what the algo did. Different owners, different fixes - and clients rarely see
these separated, so it is usually the most valuable slide in the deck.

## Size decides what was achievable

A large order legitimately starts before the close because the auction cannot
absorb it. Trading in continuous is **expected** for those, not a defect.

So an order is only short of the auction if it sits below **what orders of the
same size in the same market normally achieve**. Build that reference line from
the data itself - the median auction share within each market and size band - and
compare each order against its own band. That is the difference between a finding
that survives the client meeting and one that does not.

In the deck, never say "frontier". Say: "compared with orders of the same size".

## Why an order missed the auction

Without a flag saying the order was tagged for the close, you cannot attribute a
miss to a mechanism inside the algo. The auction share still measures the
outcome, so the fact is solid; only the cause is unavailable. Say exactly that.

Assign causes by elimination, in this order, and stop at the first that fits:

| Group | Test | What it means |
|---|---|---|
| Never traded | Nothing filled | Not an auction question |
| Cleared | Auction share high | Working as intended |
| Size explains it | Large versus daily volume | Expected, not a defect |
| Limit did not cross | No auction fill and a limit order | The limit, not the algo |
| **Unexplained** | Traded, small, no limit, still no auction fill | **The deliverable** |
| Partial, below its peers | Partial fill, well under its size band | Worth investigating |
| Partial, in line | Partial fill, at or above its size band | Normal |

The unexplained group is what the review is for. List those orders individually
by id, date and symbol, and state plainly that the cause is not in the data. If
the group is large, that is the argument for getting close-eligibility tagging
into the extract - a concrete, cheap ask.

## Market structure - do not pool

Closing mechanisms differ, and pooling them averages away the thing being
measured.

- **Single-price closing auction** (Hong Kong, Japan, Australia, Korea, Taiwan,
  most of Europe, US): one print, and the analysis above applies.
- **India (NSE)**: no single-price closing auction - the close is a volume
  weighted average of the last half hour. Its "MOC" is a different product.
  **Report it separately, always.** Pooling it corrupts every close number.

Check the same for any market before including it. When a market's close time or
mechanism has not been verified against the desk's own session data, mark it
unverified wherever it affects a number.

## Two things worth asking the desk for

Both cheap, both would materially improve the next review:

1. **Close-eligibility tagging** - turns "we cannot say why it missed" into a
   cause.
2. **Closing auction volume by symbol and date** - lets capacity be expressed as
   a share of the auction itself rather than as a share of the day, which is what
   actually constrains a close order.
