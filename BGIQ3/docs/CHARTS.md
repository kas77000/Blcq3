# Charts produced by moc_tca.py

Every run writes the charts to `output/charts/<folder>/` and a copy without titles and with larger type to `output/charts_deck/<folder>/`. The folders follow the sales trader's review email (points 1 to 13). Charts that answer none of those points sit in named folders at the end.

## Conventions on every chart

- **Sign:** positive = saving, negative = cost, side-adjusted. A buy that paid less than the benchmark, or a sell that got more, is positive.
- **Units:** most charts are in spreads. The result is divided by the average spread, and that spread is printed under each bar as "[x bps]". Multiply the two to get bps.
- **Whiskers:** 95% bootstrap confidence intervals. A whisker crossing zero means the result can't be told apart from zero.
- **Weighting:** averages are weighted by notional unless the chart says otherwise.
- **Populations:**
  - close-only = 99.5% or more of the order printed in the auction
  - pre-traded = part of it traded before the auction
- **Unknown markets:** markets that can't be attributed are left off the charts. They stay in the tables.

## Terms used below

- **Auction move:** the official close against the last trade before the continuous session ended. It is the price change the auction itself produced.
- **Pre-close trend:** arrival price to that last continuous trade, i.e. what the stock did before the auction.
- **Reversion:** close to the next day's open. Negative means the price came back against where we traded.
- **ClosePR (close participation):** our shares executed in the close over the official close auction volume.

The auction move, the trend, the drift and the 15-minute VWAP need intraday prices from qatt. Without qatt set up, those charts are not produced.

---

## 01_core_cost_metrics (point 1)

| Chart | What it shows | How to read it |
|---|---|---|
| `15_market_vs_close` | Execution against the official close, by market, all orders | The tracking-accuracy chart. Markets ordered by value traded, not by result. |
| `27_pretraded_close_market` | The same for pre-traded orders only | Close-only orders sit at zero by construction, so this is where any vs Close difference lives. |
| `32_pretraded_pvwap_market` | Pre-traded orders against PVWAP (the market VWAP while the order was working) | Judges how well the pre-close part traded, apart from when the order started. |
| `47_opportunity_cost_market` | Cost of the unfilled part: unfilled share × the move from the close to the next open | Approximate, because the price the residual was later completed at is not in the extract. Small when fill rates are close to 100%. |

## 02_buy_vs_sell (point 2)

| Chart | What it shows | How to read it |
|---|---|---|
| `25_close_only_reversion_market_side` | Close-only reversion by market, Buy panel and Sell panel | Same as the H1 slide. |
| `26_pretraded_by_side` | Pre-traded orders: first execution, execution vs close, next open vs close; All / Buy / Sell | Same as the H1 slide. |
| `28_pretraded_first_exec_market_side` | Pre-traded first execution vs close, by market, buys and sells | Was starting before the close the right call, per market and side? |
| `29_pretraded_close_market_side` | Pre-traded execution vs close, by market, buys and sells | |
| `30_pretraded_reversion_market_side` | Pre-traded reversion, by market, buys and sells | Same as the H1 slide (Korea sells stood out in H1). |
| `33_pretraded_pvwap_market_side` | Pre-traded vs PVWAP, by market, buys and sells | |
| `46_closepr_by_side` | Auction move by ClosePR bucket, Buy panel and Sell panel | Does asymmetry grow with participation? |
| `48_side_gap_market` | Buy minus sell on the auction move, by market | Positive = buys did better. Whiskers are ±1.96 standard errors clustered by date. In a rising market buys look worse for reasons that are not execution. |

## 03_auction_move (point 3)

| Chart | What it shows | How to read it |
|---|---|---|
| `41_auction_move_market` | Auction move by market, side-adjusted | Negative = the close moved against our orders. Answers "which markets have the largest closing dislocations". |
| `49_dislocation_share_market` | % of orders in names whose close moved more than 1, 2 and 3 spreads from the last trade, either direction | Answers "what % of orders were in names with close moves above x spreads". |
| `56_vol_dislocation` | Median distance of the close from the last trade and from the last-15m VWAP, by daily-vol quartile (from `Vol`) | Do high-vol names dislocate more at the close? |
| `57_vol_close_vs_pretraded` | vs Arrival for close-only against pre-traded, by daily-vol quartile | Does completing risk before the auction pay off in calmer names? This is the email's "low dislocation names allow completing earlier". |

## 04_auction_vs_trend (point 4)

| Chart | What it shows | How to read it |
|---|---|---|
| `40_decomposition_market` | vs Arrival split into three bars per market: pre-close trend, auction move, execution vs close | The three add up to vs Arrival exactly. Shows whether the cost was the stock trending all day or the auction itself. The email's "+120 bps arrival vs +5 bps auction" example (−120 and −5 in our sign). |
| `40b_decomposition_side` | The same split, buys and sells | |
| `40c_decomposition_strategy` | The same split, by strategy | |

## 05_pre_close_drift (point 5)

| Chart | What it shows | How to read it |
|---|---|---|
| `42_pre_close_drift_market` | The stock's move over the last 60, 30 and 15 minutes before the continuous close, then in the auction, by market | Negative = the price was already moving against the order. If the drift is large and the auction small, the cost came before the auction started. Drift toward our side is not proof of leakage: published imbalances move prices too. |
| `42b_pre_close_drift_side` | The same, buys and sells | |

## 06_close_vs_last_15m_vwap (point 6)

| Chart | What it shows | How to read it |
|---|---|---|
| `43_close_vs_vwap15_distribution` | Distribution of the close against the last 15 minutes' VWAP, one count per name per day | Market measure, not side-adjusted. Positive = the close printed above the VWAP. |
| `44_close_vs_vwap15_market` | Markets ranked by the average distance (either direction) | A higher market has a close more out of line with the trading just before it. |

## 07_participation (point 7)

| Chart | What it shows | How to read it |
|---|---|---|
| `34_close_pr_all` | Average ClosePR per order, by market, all orders | Simple mean per order, not weighted by value. India uses its PR on window orders. |
| `35_close_pr_close_only` | The same, close-only orders | Same as the H1 slide. |
| `36_close_pr_pretraded` | The same, pre-traded orders | Same as the H1 slide. |
| `45a_closepr_auction_move` | Auction move by ClosePR bucket: <5, 5-10, 10-20, 20-30, 30%+ | The participation question. Execution vs close cannot answer it: an auction fill prints at the close, so it is zero by construction. |
| `45b_closepr_reversion` | Reversion by ClosePR bucket | Did bigger participation leave a move that came back overnight? |

## 08_regression (point 8)

| Chart | What it shows | How to read it |
|---|---|---|
| `51_impact_curve` | Auction move against ClosePR: bucket averages with 95% CIs, plus a fitted line that bends at 5, 10, 20 and 30% with spread and vol held constant | The footnote names the steepest segment, i.e. where cost speeds up. Full regression output (drivers, one-sd effects, how much ClosePR and %ADV move together) is in `tables.xlsx` sheets 101-105. |

## 09_normalisation (point 9)

| Chart | What it shows | How to read it |
|---|---|---|
| `16_spread_relative` | Results by market in spreads beside bps | Why bps alone mislead across markets. |
| `17_close_by_adv` | vs Close by ADV% band | Cost against order size. |
| `18_close_by_spread` | vs Close by spread quartile | Cost against how wide the stock trades. |
| `52_cost_per_vol_market` | vs Arrival, vs Close, trend and auction move in units of daily vol, by market | Daily vol comes from `Vol`, read as annualised % (`VOL_UNIT`). The run log prints the median daily vol to check. |

## 10_signal_vs_noise (point 10)

| Chart | What it shows | How to read it |
|---|---|---|
| `50_decomposition_market_adjusted` | The three-part split after taking off the index move × beta | Empty until `INDEX_SYMS` is set. Beta comes from the extract where it has one, else 1. Sector moves are not removed. |

## 11_statistical_robustness (point 11)

| Chart | What it shows | How to read it |
|---|---|---|
| `53_concentration_market` | Share of each market's value in its five largest names | Above 50%, the market's average is mostly those names. Every table also carries standard errors clustered by date, t-stats and an equal-weighted mean beside the notional-weighted one. |

## 12_liquidity_buckets (point 12)

| Chart | What it shows | How to read it |
|---|---|---|
| `24_adv_profile` | Value traded and orders by ADV% band | Same as the H1 overview inset. |
| `31_pretraded_adv_profile` | The same, pre-traded orders | |
| `54_cap_dislocation` | Median auction move (either direction) by market cap | Do small caps dislocate more at the close? |

## 13_market_deep_dives (point 13)

| Chart | What it shows | How to read it |
|---|---|---|
| `13_market_notional` | Value traded by market, Cash and Swap | Same as the H1 overview. |
| `12_monthly` | vs Close by month, excluding India | |
| `12b_monthly_india` | The same, India only | India changes regime on 3 August. |
| `55_india_cas` | India before the CAS, then CAS names and non-CAS names after it | CAS names are those with an AUC or ACB code in qatt that day. Needs qatt. Small sample. |

## first_execution

| Chart | What it shows | How to read it |
|---|---|---|
| `10b_first_exec_market` | First execution vs close, by market, pre-traded orders | Was starting before the close the right call? |
| `19_first_exec_by_adv` | The same, by ADV% band | |
| `20_first_exec_by_spread` | The same, by spread quartile | |

## reversion

| Chart | What it shows | How to read it |
|---|---|---|
| `11_reversion` | Reversion by strategy | Did the auction print come back by the next open? |
| `21_reversion_pretraded` | Reversion by market, pre-traded orders | |
| `22_reversion_close_only` | Reversion by market, close-only orders | Same as the H1 slide. |

## flow_and_algo_choice

| Chart | What it shows | How to read it |
|---|---|---|
| `23_flow_split` | Close-only against pre-traded: value, orders, fill, %ADV, % executed in the close | Same as the H1 slide. |
| `15_algo_choice` | Strategies compared inside the same market and size band | Only produced when the scope has two or more close strategies. |
