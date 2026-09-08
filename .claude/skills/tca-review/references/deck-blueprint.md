# Deck blueprint

Structure of the client deck. 15-20 slides depending on how many strategies the
client traded. **Every slide states one thing.**

Slide types available in `deck.json`: `title`, `bullets`, `chart`, `table`,
`section`. Contract in `assets/deck-manifest.md`.

## Opening - 3 slides

**1. Title.** Client, period, what is covered, who prepared it.

**2. How to read this.** The only slide that teaches anything. Cover, in one line
each: which way the sign runs ("plus means you did better"), what a basis point
is worth in money, and the two or three benchmarks used. Nothing else.

**3. What we looked at.** Orders, value traded, markets, strategies used, dates.
A table. This is also where the scope limits go: what is excluded and why.

## Headline - 2 slides

**4. The overall result.** One number per strategy family, against that family's
own benchmark, with the benchmark named on the row. State on the slide that the
families are not comparable to each other.

**5. Where the cost came from.** The waiting-versus-executing split. A waterfall.
This separates the client's decision from the algo's work, and it is usually the
slide the meeting turns on.

## One section per strategy - 2 slides each

Only for strategies with enough orders to score. Under 30 orders, the strategy
goes in the appendix as a description.

**A. How it did against the job it was given.** The primary benchmark, the number,
the sample size, and for anything passive or opportunistic the fill rate next to
it. One chart.

**B. The one thing worth noting.** The diagnostic that matters for that family -
consistency for VWAP, realised rate against target for PART, auction share for
MOC, fill rate against slippage for NINJA and SHADOW. One chart, one conclusion.

## Cross-cutting - 3 slides

**C1. Was the right algo used.** Strategy choice against order size and urgency.
Usually the most valuable slide in the deck, because algo selection costs more
than algo performance.

**C2. Markets and venues.** Where the flow went and what that cost. Keep markets
with different closing mechanisms apart.

**C3. Over the period.** Monthly trend on one or two measures. Says whether
anything is moving.

## Closing - 3 slides

**D1. What we suggest.** Three items maximum, each with the effect you would
expect and roughly what it is worth. Ordered by value.

**D2. What we could not see.** The data that was missing and what it would let
the next review answer. Turns a gap into a concrete ask.

**D3. Appendix.** Definitions, sample sizes per strategy, method, and the
strategies too small to score.

## Rules for every slide

- One message. If a slide needs "and", it is two slides.
- The title is the message, not the category.
- Chart or table, not both.
- Three bullets maximum, and often zero - a good chart with a good title needs
  no bullets.
- Speaker notes carry the method, the sample size, whether the difference is
  real, and which source image the number came from. The sales trader reads these
  to prepare.
- No slide requires having read another slide.

## Charts

`scripts/charts.py` provides: `bar`, `grouped_bar`, `stacked_bar`, `waterfall`,
`scatter`, `line`. All render at slide size with presentation-sized type.

- Cost and savings use the diverging pair - blue for savings, red for cost, grey
  at zero. Every axis carries the words "cost" and "savings" so nobody has to
  remember the sign.
- Strategies use the categorical palette in fixed slot order, the same colour for
  the same strategy on every slide in the deck.
- Direct labels on the bars. Never a legend when a direct label will do.
- One measure per chart. Two measures of different scale are two charts.
