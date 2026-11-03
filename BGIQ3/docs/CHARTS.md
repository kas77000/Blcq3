# Charts produced by moc_tca.py

`output/deck_charts/` holds only the charts the YTD deck uses, numbered in slide order, with the newer charts beside the slide they belong with.

Every run writes the charts to `output/charts/<folder>/` and a copy without titles and with larger type to `output/charts_deck/<folder>/`. The folders follow the sales trader's review email (points 1 to 13). Charts that answer none of those points sit in named folders at the end.

## Conventions on every chart

- **Sign:** positive = saving, negative = cost, side-adjusted. A buy that paid less than the benchmark, or a sell that got more, is positive.
- **Units:** most charts are in spreads. The result is divided by the average spread, and that spread is printed under each bar as "[x bps]". Multiply the two to get bps.
- **Whiskers:** 95% bootstrap confidence intervals. A whisker crossing zero means the result can't be told apart from zero.
- **Weighting:** averages are weighted by notional unless the chart says otherwise.
- **Populations:**
  - auction-only = 99.5% or more of the order printed in the auction
  - pre-traded = part of it traded before the auction
- **Unknown markets:** markets that can't be attributed are left off the charts. They stay in the tables.

## Terms used below

- **Auction move:** the official close against the last trade before the continuous session ended. It is the price change the auction itself produced.
- **Pre-close trend:** arrival price to that last continuous trade, i.e. what the stock did before the auction.
- **Reversion:** close to the next day's open. Negative means the price came back against where we traded.
- **ClosePR (close participation):** our shares executed in the close over the official close auction volume.

The auction move, the trend, the drift and the 30-minute VWAP come from the extract's `last_cont_price`, `price60bef_close`, `price_15bef_close` and `last30min_vwap_cont`. qatt only fills a price the extract is missing. The email asked for the last 15 min VWAP; the extract carries 30.

---

## 01_core_cost_metrics (point 1)

| Chart | What it shows | How to read it |
|---|---|---|
| `15_market_vs_close` | Execution against the official close, by market, all orders | The tracking-accuracy chart. Markets ordered by value traded, not by result. |
| `27_pretraded_close_market` | The same for pre-traded orders only | Auction-only orders sit at zero by construction, so this is where any vs Close difference lives. |
| `32_pretraded_pvwap_market` | Pre-traded orders against PVWAP (the market VWAP while the order was working) | Judges how well the pre-close part traded, apart from when the order started. |
| `37_close_only_arrival_market` | Auction-only orders against the arrival price, by market, in spreads | The full cost from the order's arrival to the close print. |
| `38_pretraded_arrival_market` | Pre-traded orders against the arrival price, by market, in spreads | The same for orders that traded before the auction. |
| `47_opportunity_cost_market` | Cost of the unfilled part: unfilled share × the move from the close to the next open | Approximate, because the price the residual was later completed at is not in the extract. Small when fill rates are close to 100%. |

## 02_buy_vs_sell (point 2)

| Chart | What it shows | How to read it |
|---|---|---|
| `25_close_only_reversion_market_side` | Auction-only reversion by market, Buy panel and Sell panel | Same as the H1 slide. |
| `25b_close_only_reversion_participation_side` | The same with close participation for that side as a green line across the top on the right axis, Buy and Sell side by side | Takes the place of 25 in `deck_charts`. One scale per measure across both panels. No title. |
| `26_pretraded_by_side` | Pre-traded orders: first execution vs close, vs PVWAP, execution vs close, next open vs close; All / Buy / Sell | In the desk's order. |
| `26c_pretraded_by_order_type` | Pre-traded orders: first execution vs close, vs PVWAP, execution vs close, next open vs close; Discretionary panel, Stipulated panel | Same measures and weights as 26, split by order type instead of side (tables `73e`-`73h`). Orders with no order type are left out; order count in each panel title. |
| `26e_pretraded_by_order_subtype` | Pre-traded: first execution vs close, vs PVWAP, execution vs close, next open vs close - Model against % of Order, order counts in the panel titles | Tables 73i-73l. Close Auction Only and unknown sub-types left out. |
| `26d_pretraded_close_order_type_market` | Pre-traded execution vs close by market; Discretionary panel, Stipulated panel; buys and sells together | Table `75d_pre_close_mkt_type_spr`. Unknown order types left out. |
| `26f_pretraded_close_subtype_market` | Pre-traded execution vs close by market, Model against % of Order, buys and sells together | Table 75e. Same chart as 26d, split by sub-type. Close Auction Only and unknown sub-types left out. |
| `22d_auction_only_by_order_type` | Auction-only orders: last continuous vs close, vs PVWAP, execution vs close, next open vs close; Discretionary panel, Stipulated panel | Execution vs close is zero by design. Tables `72e`-`72h`. Unknown order types left out. |
| `28c_pretraded_measures_market_side` | Pre-traded orders, the same four measures by market: one row per measure, Buy left, Sell right | Buy and Sell share a scale within a row; rows do not. No longer in the deck. |
| `28f_pretraded_close_reversion_market` | Pre-traded orders by market, buys and sells together: execution vs close (left) and next open vs close (right), each with a strip above showing close and continuous participation % | Participation is a simple mean per order (table `82b_pr_pair_mkt_pretraded`). On slide 8. |
| `28g_pretraded_close_market_pr` | 28f's left half on its own, larger: pre-traded execution vs close by market, with the close/continuous participation strip and its legend | In deck_charts in place of 28f. |
| `28h_pretraded_reversion_market_pr` | 28f's right half on its own, larger: pre-traded next open vs close by market, with the participation strip and its legend | In deck_charts in place of 28f. |
| `28e_pretraded_measures_market` | Pre-traded orders, the same four measures by market, buys and sells together: one panel per measure, 2x2 | Same markets in the same order in every panel; each panel has its own scale. Execution vs close in bold. |
| `28d_pretraded_close_best_worst` | Pre-traded orders: the 5 best and 5 worst single orders on execution vs close, in spreads | Each bar is one order over its own spread, so tight-spread names reach large numbers; table `75c_pre_close_best_worst` has the bps beside it. |
| `28_pretraded_first_exec_market_side` | Pre-traded first execution vs close, by market, buys and sells | Was starting before the close the right call, per market and side? |
| `28b_pretraded_last_cont_market_side` | Pre-traded last continuous price vs close, by market, buys and sells (+ = the auction moved our way, - = against us) | How far did the close move from where the session ended, per market and side? |
| `29_pretraded_close_market_side` | Pre-traded execution vs close, by market, buys and sells | |
| `30_pretraded_reversion_market_side` | Pre-traded reversion, by market, buys and sells | Same as the H1 slide (Korea sells stood out in H1). |
| `33_pretraded_pvwap_market_side` | Pre-traded vs PVWAP, by market, buys and sells | |
| `37b_close_only_arrival_market_side` | Auction-only arrival slippage, by market, buys and sells | |
| `38b_pretraded_arrival_market_side` | Pre-traded arrival slippage, by market, buys and sells | |
| `39b_close_only_last_cont_market_side` | Auction-only last continuous price vs close, by market, buys and sells (auction names only) | Same sign as 28b. |
| `46_closepr_by_side` | Auction move by ClosePR bucket, Buy panel and Sell panel | Does asymmetry grow with participation? |
| `48_side_gap_market` | Buy minus sell on the auction move, by market | Positive = buys did better. Whiskers are ±1.96 standard errors clustered by date. In a rising market buys look worse for reasons that are not execution. |

## 03_auction_move (point 3)

| Chart | What it shows | How to read it |
|---|---|---|
| `41_auction_move_market` | Auction move by market, side-adjusted | Negative = the close moved against our orders. Answers "which markets have the largest closing dislocations". |
| `49_dislocation_share_market` | % of orders in names whose close moved more than 1, 2 and 3 spreads from the last trade, either direction | Answers "what % of orders were in names with close moves above x spreads". |
| `56_vol_dislocation` | Median distance of the close from the last trade and from the last 30 min VWAP, by daily-vol quartile (from `Vol`) | Do high-vol names dislocate more at the close? |
| `57_vol_close_vs_pretraded` | vs Arrival for auction-only against pre-traded, by daily-vol quartile | Does completing risk before the auction pay off in calmer names? This is the email's "low dislocation names allow completing earlier". |

## 04_auction_vs_trend (point 4)

| Chart | What it shows | How to read it |
|---|---|---|
| `40_decomposition_market` | vs Arrival split into three bars per market: pre-close trend, auction move, execution vs close | The three add up to vs Arrival exactly. Shows whether the cost was the stock trending all day or the auction itself. The email's "+120 bps arrival vs +5 bps auction" example (−120 and −5 in our sign). |
| `40b_decomposition_side` | The same split, buys and sells | |
| `40c_decomposition_strategy` | The same split, by strategy | |

## 05_pre_close_drift (point 5)

| Chart | What it shows | How to read it |
|---|---|---|
| `42_pre_close_drift_market` | The stock's move over the last 60 and 15 minutes before the continuous close (`price60bef_close`, `price_15bef_close` to `last_cont_price`), then in the auction, by market | Negative = the price was already moving against the order. If the drift is large and the auction small, the cost came before the auction started. Drift toward our side is not proof of leakage: published imbalances move prices too. |
| `42b_pre_close_drift_side` | The same, buys and sells | |

## 06_close_vs_last_vwap (point 6)

| Chart | What it shows | How to read it |
|---|---|---|
| `43_close_vs_vwap_distribution` | Distribution of the close against the VWAP of the last 30 minutes of the continuous session (`last30min_vwap_cont`), one count per name per day | Market measure, not side-adjusted. Positive = the close printed above the VWAP. |
| `44_close_vs_vwap_market` | Markets ranked by the average distance (either direction) | A higher market has a close more out of line with the trading just before it. |

## 07_participation (point 7)

| Chart | What it shows | How to read it |
|---|---|---|
| `34_close_pr_all` | Average ClosePR per order, by market, all orders | Simple mean per order, not weighted by value. India uses its PR on window orders. |
| `35_close_pr_close_only` | The same, auction-only orders | Same as the H1 slide. |
| `35b_close_pr_close_only_side` | The same, buys and sells side by side | One scale, markets in the order of 35. |
| `36_close_pr_pretraded` | The same, pre-traded orders | Same as the H1 slide. |
| `36b_pretraded_pr_close_cont` | Pre-traded orders: close participation (ClosePR) beside continuous participation (fPR_cont), by market | Two bars, not stacked: they are shares of different volumes (the auction's, and the market's while the order traded before it). |
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
| `52_cost_per_vol_market` | vs Arrival, vs Close, trend and auction move in units of daily vol, by market | Daily vol is `Vol`, Parkinson daily volatility in % (`VOL_UNIT = "daily_pct"`). The run log prints the median daily vol to check. |

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
| `24b_adv_market` | Order size as % of ADV by market, notional-weighted, order count under each market | Table 52b (also has the median). All orders with a %ADV. |
| `31_pretraded_adv_profile` | The same, pre-traded orders | |
| `54_cap_dislocation` | Median auction move (either direction) by market cap | Do small caps dislocate more at the close? |

## 13_market_deep_dives (point 13)

| Chart | What it shows | How to read it |
|---|---|---|
| `13_market_notional` | Value traded by market, Cash and Swap | Same as the H1 overview. |
| `13b_market_notional_cash` / `13c_market_notional_swap` | Pie charts: value traded by market, cash alone and swap alone, legend on the right | Largest slice first, each percentage a share of that account type's total. Markets under 3% go into Others, named in the footnote. |
| `12_monthly` | vs Close by month, excluding India | |
| `12b_monthly_india` | The same, India only | India changes regime on 3 August. |
| `55_india_cas` | India before the CAS, then CAS names and non-CAS names after it | CAS names are those with an AUC or ACB code in qatt that day. Needs qatt. Small sample. |

## first_execution

| Chart | What it shows | How to read it |
|---|---|---|
| `10b_first_exec_market` | First execution vs close, by market, pre-traded orders | Was starting before the close the right call? |
| `10c_last_cont_market` | Last continuous price vs close, by market, pre-traded orders; same sign as 10b, the mirror of the auction move (41). CAS names use qatt's last trade before the auction | Would the last continuous trade have beaten the close? |
| `19_first_exec_by_adv` | The same, by ADV% band | |
| `20_first_exec_by_spread` | The same, by spread quartile | |
| `39_close_only_last_cont_market` | Last continuous price vs close, by market, auction-only orders on auction names | The auction-only twin of 10c. India on a VWAP close is left out: it has no auction. |

## reversion

| Chart | What it shows | How to read it |
|---|---|---|
| `11_reversion` | Reversion by strategy | Did the auction print come back by the next open? |
| `21_reversion_pretraded` | Reversion by market, pre-traded orders | |
| `22_reversion_close_only` | Reversion by market, auction-only orders | Same as the H1 slide. |
| `22b_close_only_reversion_participation` | Auction-only reversion by market on the left axis, with close participation as a green line across the top on the right axis | Takes the place of 22 and 35 in `deck_charts`. The line joins markets to read across; it is not a trend. |
| `22c_auction_only_order_type_market` | Auction-only orders: discretionary vs stipulated, % of each market's value | One 100% bar per market, value traded on top. Beside 22b on slide 5. |

## flow_and_algo_choice

| Chart | What it shows | How to read it |
|---|---|---|
| `23_flow_split` | Auction-only against pre-traded: value, orders, fill, %ADV, % executed in the close | Same as the H1 slide. |
| `23b_flow_split_order_type` | The same two bars, each stacked by order type (from `bgi_order_types.csv`) | Each segment's % is its share of that bar's value. Unknown = no row in the order-types file. |
| `23c_order_type_market` | All orders reaching the close: order sub-type (Model, % of Order, Close Auction Only), % of each market's value | Sub-type from the OrderSubType column of bgi_order_types.csv (table 51f). |
| `26b_pretraded_order_type_market` | Pre-traded orders: discretionary vs stipulated, % of each market's value | Same chart as 22c and 23c, pre-traded orders only. |
| `15_algo_choice` | Strategies compared inside the same market and size band | Only produced when the scope has two or more close strategies. |

## hk focus

Hong Kong, every order in the review, split Cash / Swap by client account (`CASH_CLIENTS`, `SWAP_CLIENTS`). An HK account in neither list shows as a third panel, Other, and the run log names it. Spread quartiles are cut on Hong Kong's own orders, so Cash and Swap sit in the same bands; size uses the fixed ADV% bands. Everything is notional-weighted. Tables `120_hk*` in `tables.xlsx`, with `120_hk0_profile` giving orders, value, % pre-traded, spread and median %ADV per group.

| Chart | What it shows | How to read it |
|---|---|---|
| `90_hk_measures` | First execution vs close, last continuous price vs close, execution vs close; Cash panel, Swap panel | + is always good for the order: better than the close, and for last continuous vs close, the auction moved our way. |
| `90b_hk_co_measures` | Auction-only orders: last continuous price vs close, vs PVWAP, execution vs close, next open vs close (reversion); Cash panel, Swap panel | Execution vs close sits at zero by design; reversion is the measure that moves. |
| `90c_hk_pre_measures` | Pre-traded orders: first execution vs close, last continuous price vs close, vs PVWAP, execution vs close; Cash panel, Swap panel | Orders with nothing printed in the close are in neither 90b nor 90c. |
| `90d_hk_flow_order_type` | Hong Kong, auction-only vs pre-traded, each bar split discretionary / stipulated; cash and swap together | Same chart as 23b, Hong Kong only (table `120_hk7_flow_order_type`). Orders with no order type are left out. |
| `91a`/`91b`/`91c_hk_spread_q_*` | Each of the three by spread quartile, Cash and Swap side by side | Does the gap between Cash and Swap hold in names that trade alike? |
| `92a`/`92b`/`92c_hk_adv_*` | The same by ADV% band (0-1, 1-3, 3-5, 5-10, 10-25, 25%+) | The same question for order size. Fixed bands, not quartiles: most orders are tiny, so quartiles would put every large order in one band. |
| `92d_hk_pre_adv_q_first_exec` | Pre-traded only: first execution vs close by ADV% quartile, Cash and Swap | Quartiles cut on the pre-traded orders, so the four bands cover the sizes that trade before the auction. |

## India focus

India orders from 3 August 2026 (CAS go-live), split CAS / Non-CAS. CAS = the name had an AUC or ACB code in qatt that day. Orders with no qatt answer are left out and counted in the run log. Pre-CAS India is left out so both groups cover the same months. Non-CAS names close on the VWAP, so their "close" is that VWAP. Tables `130_in*`.

| Chart | What it shows | How to read it |
|---|---|---|
| `95_in_measures` | The three measures; CAS panel, Non-CAS panel | |
| `95b_in_co_measures` | Auction-only: last continuous price vs close, execution vs close, reversion; CAS panel, Non-CAS panel | Non-CAS window orders count as auction-only (their close share is imputed at 100). |
| `95f_in_flow_measures` | India from the CAS date, 2x2: CAS and Non-CAS by auction only and pre-traded; first execution, last continuous, execution vs close, next open vs close | Flow read from the price (first fill at the close = auction only). Order count in each panel title. Tables `130_in7_flow_*`. In the deck in place of 95d. |
| `95h_in_subtype_flow_measures` | Same four panels as 95f (CAS / Non-CAS x auction only / pre-traded), flow from the order sub-type: Close Auction Only = auction only, % of Order = pre-traded; Model orders left out | Tables 130_in9_subflow_*. |
| `95g_in_pre_cas_measures` | India pre-traded orders, CAS vs Non-CAS: first execution vs close, vs PVWAP, execution vs close, next open vs close, order counts in the panel titles | Tables 130_in8_pre_*. Same layout as 26c. Every order behind 95d/95f/95g is in india_orders.xlsx (one sheet per group and flow). |
| `95i_in_pre_subtype_measures` | India CAS pre-traded orders (Non-CAS left out), Model vs % of Order: first execution vs close, vs PVWAP, execution vs close, next open vs close | Tables 130_in10_pre_sub_*. India's version of 26e. |
| `95c_in_pre_measures` | Pre-traded: first execution, last continuous price, execution, each vs close | CAS names only in practice: a Non-CAS order cannot be pre-traded. |
| `96a`/`96b`/`96c_in_spread_q_*` | By spread quartile, CAS and Non-CAS side by side | |
| `97a`/`97b`/`97c_in_adv_*` | By ADV% band | |
| `95d_in_cas_flow_measures` | India CAS only: auction-only and pre-traded side by side. First execution, last continuous price, execution and next open, each vs close | Split by price: auction only = first fill at the close. Order counts and value shares in the footnote. |
| `95e_in_cas_flow_order_type` | India CAS: auction only and pre-traded, each bar stacked by order type, with mean close participation (ClosePR) and the % of executed quantity in the close under the bar | Each segment's % is its share of that bar's value. |
| `97d_in_pre_adv_q_first_exec` | Pre-traded only: first execution vs close by ADV% quartile, CAS and Non-CAS | India reports no close share, so pre-traded here means the first fill was not at the close. Non-CAS names trade through the closing VWAP, so nearly all of them count. |
| `97e_in_cas_pre_pr_adv` | India CAS, pre-traded: ClosePR beside fPR_cont, by ADV% band | Same reading as 36b. |

## australia focus

Pre-traded Australia orders, by the time from the order's first start to the close (AEDT handled). Tables `140_au*`. Orders that appear to start after the close are left out and counted in the run log.

| Chart | What it shows | How to read it |
|---|---|---|
| `98a_au_pre_start_profile` | Share of pre-traded value by start band, Australia against the other auction markets | Do Australia's orders start later than the rest of the book? |
| `98b_au_pre_perf_by_start` | First execution, execution and arrival slippage, each vs its benchmark, by start band | Does the result get worse for late starters? Size moves with the start time, so read it with median %ADV in `140_au0_start_profile`. |
| `98c_au_pre_pr_by_start` | ClosePR beside continuous participation, by start band | Do late starters trade harder? Two different volumes, so never added. |

