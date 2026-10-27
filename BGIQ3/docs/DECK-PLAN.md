# YTD 2026 MOC TCA deck: plan

Period: 1 January to 23 October 2026. The deck goes to BlackRock and is presented by the sales trader on the 3 November call.

It keeps the shape of the H1 deck, so the client sees the same story with the period extended. New slides answer the sales trader's review email, and one slide checks whether the H1 suggestions were followed and what they changed.

## How each slide is built

Same as H1:

- A title, one to three short bullets, then one large chart.
- Each bullet states a result with its number. No bullet just describes the chart.
- Results are in spreads, with the spread in bps under each market ("[12.5 bps]"), so markets compare fairly.
- 95% confidence whiskers on every bar. Where a whisker crosses zero, the bullet does not call the result a gain or a loss.
- Blue above zero is a saving, red below zero is a cost. **Positive = saving, negative = cost** everywhere. The email asked for the opposite and we are not following it.
- A footer line names the data set and the filter, e.g. `Data Set : YTD 2026 (1 Jan - 23 Oct) Filter : CloseFillRate > 99.5`.
- Small samples (under 30 orders in a bar) are either flagged on the slide or left off it. H1 showed Philippines at +28 spreads on a handful of orders; that does not happen again.
- The technical detail (method, t-stats, regression output) goes in the speaker notes or the appendix, not on the slide.

Wording: plain sentences, the way the sales trader would say it on the call. It must not read as generated.

## Slide order

| # | Slide | Chart(s) | What the bullets say |
|---|---|---|---|
| 1 | YTD MOC TCA Overview | `13_market_notional`, `24_adv_profile` | Total executed, number of markets, top markets and their share, % of orders under 1% ADV, overall fill rate and why the residual was left. Same as H1. |
| 2 | Pre-traded vs Close only | `23_flow_split` | Split of value, orders, fill, %ADV and % executed in the close for each. Same as H1. |
| 3 | Where the cost came from | `40_decomposition_market` | **New.** vs Arrival split into the stock's trend before the close, the auction move and our execution. Answers the email's headline question: was the cost the market or the auction? |
| 4 | The auction move, by market | `41_auction_move_market`, `49_dislocation_share_market` | **New.** How far the close printed from the last continuous trade, and the share of orders in names that moved more than 1, 2, 3 spreads. |
| 5 | Before the auction | `42_pre_close_drift_market` | **New.** The move over the last 60 and 15 minutes, then in the auction. Was the cost already there before the auction started? |
| 6 | Is the close an outlier? | `44_close_vs_vwap_market` (+ `43` in backup) | **New.** Close against the last 30 minutes' VWAP, markets ranked. The email asked for 15; say on the slide it is 30. |
| 7 | Close Only Participation, by Market | `35_close_pr_close_only` | Same as H1. |
| 8 | Close Only Reversion, by Market / by Side | `22_reversion_close_only`, `25_close_only_reversion_market_side` | Same as H1. Say "the price did not come back", not "no impact". |
| 9 | Does cost rise with participation? | `51_impact_curve`, `45a_closepr_auction_move` | **New.** Auction move by ClosePR bucket (<5, 5-10, 10-20, 20-30, 30%+) and where it steepens. Covers every market the way the H1 Korea table covered one. |
| 10 | Pre-traded Analysis, by Side | `26_pretraded_by_side` | Same as H1. Fix the H1 slip: the first bullet is first execution (+2.45 in H1), not execution vs close. |
| 11 | Pre-traded Analysis, by Side/Market | `28_pretraded_first_exec_market_side` | Same as H1. |
| 12 | Pre-traded Reversion, by Side/Market | `30_pretraded_reversion_market_side` | Same as H1. |
| 13 | Pre-traded Close Participation, by Market | `36_close_pr_pretraded` | Same as H1. |
| 14 | Buys against sells | `48_side_gap_market` | **New.** Where buys and sells really differ, with significance. Note that a rising market makes buys look worse. |
| 15 | Volatility | `56_vol_dislocation`, `57_vol_close_vs_pretraded` | **New.** Do high-vol names dislocate more at the close, and does pre-trading calmer names pay off? |
| 16 | India and the CAS | `55_india_cas` | **New.** India before 3 August, then CAS names and non-CAS names after it. Small sample, said on the slide. |
| 17 | H1 suggestions: what happened | tables from a Jan-Jun run and a Jul-Oct run | **New.** Did pre-traded orders under 1% ADV go close-only? Did Korea pre-traded reversion improve? Run `moc_tca.py` twice with `--from/--to` and compare. |
| 18 | YTD MOC 2026: Takeaways | none | "Suggest:" lines, as in H1. Each one points back to the slide that supports it. |

## Appendix

- `50_decomposition_market_adjusted`: the split after removing the index move. Only once `INDEX_SYMS` is set. Sector moves are not removed; say so.
- `52_cost_per_vol_market`, `18_close_by_spread`, `17_close_by_adv`: cost normalised by vol, spread and size.
- `53_concentration_market`: where a few names drive a market's figure.
- `54_cap_dislocation`: closing dislocation by market cap.
- Regression tables (`101`-`105` in `tables.xlsx`): drivers of cost, one-sd effects, ClosePR segment slopes.
- `47_opportunity_cost_market`: cost of the unfilled part.

## Before any slide is written

1. Run `moc_tca.py` on the YTD extract with qatt set up. Read the run log warnings first, especially the qatt clock check and the Vol unit line.
2. Look at the numbers before building anything. If a chart says nothing, it comes out of the deck.
3. Transcribe the numbers for each slide from `tables.xlsx`, not from the chart.
