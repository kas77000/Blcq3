# Strategy playbooks

One page per algo family. **Judge each algo against the job it was given.**

House names map onto families like this. Extend the table when a new name shows
up; do not guess a family from the name alone.

| House name | Family |
|---|---|
| VWAP | Volume schedule |
| PVWAP | Volume schedule, own window |
| TMX | Time schedule (TWAP) |
| PART | Participation (POV) |
| MOC, LOC | Closing auction |
| MOO | Opening auction |
| IIS | Implementation shortfall |
| NINJA | Liquidity seeking |
| SHADOW | Passive / dark |

**Unknown algo:** stop and ask the desk what it is trying to do, then add a
playbook here. Do not review an algo whose objective you are guessing at. If the
client needs it in the deck anyway, show its volume and basic slippage in a
"not classified" section and make no performance claim.

---

## Volume schedule - VWAP, PVWAP

**The job:** match the volume-weighted average price over a window.

| | |
|---|---|
| Score on | vs PVWAP over the order's own window |
| Show beside it | vs Arrival, and the order's duration |
| Good looks like | Slippage near zero **and** low spread of outcomes |
| Classic wrong call | Scoring it on arrival - that measures the market's drift |

Consistency matters as much as the average. A VWAP averaging 0bps with a wide
spread of outcomes is worse than one at -1bps that lands there every time,
because the client can plan around the second one.

Check the volume curve assumption. A VWAP that misses on days with a heavy close
is following a curve that does not match the market it is trading.

**Say it like:** "Your VWAP orders finished within a basis point of the market
average over the same window, and did it consistently."

---

## Time schedule - TMX (TWAP)

**The job:** spread the order evenly over time.

| | |
|---|---|
| Score on | vs TWAP over the window |
| Show beside it | vs PVWAP, to show what the even spread cost |
| Good looks like | Even fills through the window, small deviation from TWAP |
| Classic wrong call | Scoring it on VWAP - a time algo cannot track a volume benchmark, and will look wrong on any day with a volume skew |

The honest question for a TWAP is not "did it beat VWAP" but "was an even spread
the right choice here". If the name's volume is concentrated at the close, an
even spread is a decision that costs money and the client should know the size of
it.

---

## Participation - PART (POV)

**The job:** stay a fixed share of the volume that trades.

| | |
|---|---|
| Score on | Realised participation against target, then vs PVWAP |
| Show beside it | Fill rate and duration |
| Good looks like | Realised rate close to target, low deviation, order completes |
| Classic wrong call | Blaming it for taking a long time - duration is an output, not an input. It finishes when the volume shows up |

Two real failure modes: it consistently undershoots the target rate (it is not
getting the volume it was asked for, so it drifts long and takes on risk), or it
overshoots in thin names (it becomes the volume and pays for it). Both show up as
realised-versus-target, which is why that comes before slippage.

If the client sets a rate too high for the name, the fix is the rate, not the
algo. Compare requested rate against %ADV to see it.

---

## Implementation shortfall - IIS

**The job:** minimise total cost against the price when the order arrived,
balancing impact against the risk of waiting.

| | |
|---|---|
| Score on | vs Arrival, controlled for size and volatility |
| Show beside it | Reversion, fill rate, duration |
| Good looks like | Beats a size- and volatility-matched cost expectation |
| Classic wrong call | Penalising it for costing more when it was told to be urgent. Urgency is bought with impact - that is the trade the client asked for |

Where an urgency setting exists, group by it. Fast orders costing more than slow
ones is the algo working, not failing. It is only a finding if the extra cost is
larger than the risk it removed.

---

## Closing auction - MOC, LOC

**The job:** trade at the closing auction price.

| | |
|---|---|
| Score on | How much of the order reached the auction |
| Show beside it | vs Close, cost of waiting (arrival minus close), and size |
| Good looks like | The order clears in the auction; what did not clear did not cost much |
| Classic wrong call | Scoring it on distance from the close - anything that cleared the auction traded **at** the close, so that number is ~0 by construction and says nothing |

Full detail in `moc-close-algos.md`. The short version: the measurement that
matters is the share that reached the auction, and the size of the order sets
what share was achievable.

---

## Opening auction - MOO

**The job:** trade at the opening auction price.

Same structure as MOC, with two differences worth stating:

- Overnight risk is the whole story. Sending at the open is a decision about
  overnight exposure, and the review should price it: how far did the open sit
  from the previous close on these names?
- Opening auctions are thinner than closing auctions in most markets, so the
  size that clears is smaller. Do not carry a close-auction expectation over.

---

## Liquidity seeking - NINJA

**The job:** find liquidity opportunistically across lit and dark, without
signalling.

| | |
|---|---|
| Score on | vs Arrival **and** fill rate, together, never separately |
| Show beside it | Reversion, dark share, duration |
| Good looks like | Good arrival slippage with a high fill rate and low reversion |
| Classic wrong call | Praising slippage that came from not filling. Selective execution flatters every price-based benchmark |

Read the pair as a grid:

| | High fill | Low fill |
|---|---|---|
| **Good slippage** | Genuinely good. Say so. | Selection, not skill. Price the unfilled shares. |
| **Poor slippage** | It paid up to complete - correct if the order was urgent | The bad case. Investigate the names and the sizes. |

High reversion with good slippage means it took liquidity aggressively and paid
temporary impact. Worth raising even when the headline looks fine.

---

## Passive / dark - SHADOW

**The job:** work passively with minimal footprint.

| | |
|---|---|
| Score on | Spread capture and passive share, then fill rate |
| Show beside it | vs Arrival, reversion, and what was left undone |
| Good looks like | High passive fill, positive spread capture, low reversion, order still completes |
| Classic wrong call | The same trap as NINJA. Unfilled shares are a real cost that no price benchmark shows |

The specific question for a passive algo: **when it did not fill, what happened
to the price?** If the price ran away from the unfilled portion, patience cost
real money, and the client should see that number next to the spread it captured.

---

## Cross-strategy rules

1. **Never rank two families on one benchmark.** Each is scored against its own
   target. State this on the slide so nobody builds the ranking themselves.
2. **Cross-algo comparison is like-for-like only** - same size band, same market,
   same volatility band. Say which controls were applied.
3. **Algo choice is usually the bigger finding than algo performance.** A
   well-executed VWAP on an order that needed the close is a loss the VWAP number
   will never show. Check, for each order: was the strategy right for its size,
   its urgency and its name?
4. **A strategy under 30 orders is described, not scored.**
