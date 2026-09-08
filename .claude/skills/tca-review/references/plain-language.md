# Plain language

The rules for everything the client sees. `scripts/check_readability.py` enforces
the mechanical ones and fails the build. The rest is judgement.

## The reader

A client who trades algos every day and does not speak TCA. They know what a VWAP
is. They do not know what winsorising is and do not need to.

The deck is presented by a sales trader who must answer any question on the call
without a quant beside them. So the slide stays simple and the **speaker notes
carry the depth**. Simple costs the specialist nothing - it just moves the detail
to where the client does not have to read it.

## Enforced rules

| Rule | Limit |
|---|---|
| Slide title | 10 words, states the finding |
| Bullets per slide | 3 |
| Words per bullet | 14 |
| Sentence length | 20 words |
| Semicolons on a slide | none |
| Banned terms | fail unless the slide defines the term |

Speaker notes are exempt from all of it. Write them properly and technically.

## Titles state the finding

A title is the sentence the client would repeat afterwards. Not a category.

| No | Yes |
|---|---|
| Auction clearance analysis | Most of your close orders reached the auction |
| Benchmark performance summary | Your VWAP orders tracked the market closely |
| Participation rate distribution | PART traded slower than you asked for |
| Cost decomposition | Waiting for the close cost more than the trading did |

## Word swaps

| Instead of | Say |
|---|---|
| vs Arrival | the price when the order reached us |
| vs PVWAP | the market average while your order was working |
| vs VWAP | the market average for the whole day |
| vs Close | the closing auction price |
| auction share, %CLOSE | how much of the order traded in the closing auction |
| capacity frontier | what orders of that size normally achieve |
| reversion | whether the price came back afterwards |
| notional-weighted | bigger orders count for more |
| winsorised | extreme trades pulled in so one order cannot skew the average |
| the CI crosses zero | too few orders to call this a real difference |
| implementation shortfall | the cost from when you sent it to when it was done |
| cost of delay | what it cost to wait |
| adverse selection | we traded just before the price moved against us |
| spread capture | how often we earned the spread instead of paying it |
| %ADV | the order's size against a normal day's volume |
| impact | how much our own trading moved the price |
| opportunity cost | what the shares we did not get would have cost later |
| execution shortfall of -4.2bps | cost 4.2 basis points, about $4,200 per $10 million |

## Every finding has three beats

In this order, one sentence each:

1. **What happened.** "About a third of your MOC orders did not reach the auction."
2. **What it cost or saved.** "That cost roughly $180,000 over the half year."
3. **What to do.** "Send the smaller ones later and they should clear."

If you cannot write beat 3, it is an observation, not a finding. Put it in the
appendix.

## Numbers

- Round to what the client can act on. 4.2bps, not 4.23bps. $180k, not $178,412.
- Give bps a money translation the first time: "4 basis points, about $4,000 per
  $10 million traded".
- Say the sign in words once, on the "how to read this" slide, and never again.
- Never put a number on a slide without saying how many orders are behind it, in
  the notes if not on the slide.
- Under 30 orders, describe it. Do not score it.

## Tone

Write to a professional counterparty, not about them. "Your MOC orders" not "the
client's MOC orders". No blame, no salesmanship. Where the desk could have done
better, say so plainly - it is the reason the review is credible.

## Write like a person, because a person is presenting it

The deck has to read as though a trader on the desk wrote it after looking at the
numbers. The moment it reads as generated, the client stops trusting the numbers
as well as the prose. `check_readability.py` catches the obvious tells; the rest
is on you.

**Vary the rhythm.** Real writing has short sentences next to long ones. Four
slides where every bullet is eleven words and starts with "Your" reads as a
template. Let one be three words. Let one be a fragment.

**Say the thing, then stop.** No wind-up, no summary of what you are about to
say, no restating the title in the first bullet.

**Commit to a view.** "This is worth changing" beats "this may potentially
represent an area for consideration". A review with no opinion in it is not worth
presenting - hedging reads as generated, and it also reads as evasive.

**Be specific where a machine would be general.** "PART ran at 11% when you asked
for 15%" is human. "Participation metrics indicate suboptimal alignment" is not.
Name the algo, the market, the month, the number.

**Speak from the desk.** "We looked at the ones that missed and most were under
$2m" - you did the work, so say so. First person plural is fine and reads as
someone who was there.

**Contractions are fine.** "It didn't clear" is how people talk.

**Allow an ordinary observation.** "Nothing much to say about the VWAP flow - it
did what it was supposed to." Not every slide needs to be a finding, and saying
so plainly is more credible than manufacturing significance.

**Never write these**, and they will fail the build: delve, leverage, seamless,
robust, landscape, underscores, moreover, furthermore, it is worth noting, a
testament to, when it comes to, unlock, harness, elevate, streamline, holistic,
actionable insights, deep dive, key takeaway, best-in-class, not only.

One test before the deck goes out: read a slide aloud. If you would not say it in
those words to the client on the phone, rewrite it.

## Ban list

These fail the check unless the slide declares it defines the term:

implementation shortfall, notional-weighted, notional weighted, value-weighted,
winsorised, winsorized, bootstrap, confidence interval, statistically
significant, heteroscedastic, frontier, cohort, decomposition, degenerate,
endogenous, orthogonal, alpha decay, adverse selection, participation-weighted,
basis point dispersion, interquartile, percentile

Some of these are fine **once defined on the slide**, which is the point: define
it or drop it.
