# MOC TCA — column checklist

One source: the order extract. No kdb. Mark each row and send back.

Header names are matched **case-insensitively, ignoring spaces, dots, underscores
and brackets**, so `Dark %`, `dark%` and `%DARK` all resolve. Alternative
spellings in brackets are accepted automatically.

## Required — the run stops without these

- [ ] `Strategy`
- [ ] `Date`
- [ ] `Sym`  (market comes from the suffix)
- [ ] `Side`
- [ ] `$Mln`  — **executed** notional
- [ ] `#Shares`  — **order** quantity, in **thousands**
- [ ] `FR`  — fill rate
- [ ] `aggrTgtId`  — order id

## Venue mix — the core of the analysis

Must sum to 100 per order.

- [ ] `%CLOSE`
- [ ] `%TAKE`
- [ ] `%POST`
- [ ] `%DARK`
- [ ] `%OPEN`

## Benchmarks — all in bps, positive = savings

- [ ] `Close`
- [ ] `IS`
- [ ] `Pvwap`
- [ ] `Vwap`
- [ ] `NextOpen`
- [ ] `Open`  *(optional)*

## Capacity and behaviour

- [ ] `%Adv`
- [ ] `Adv`
- [ ] `fPR_cont`  — continuous participation rate
- [ ] `first_exec_vs_close`  — bps, close price vs first execution price

## Segmentation and controls

- [ ] `Cap`  — large / mid / small / micro / other, and anything else
- [ ] `marketLimit`  — Market / Limit
- [ ] `Start(HK)`
- [ ] `End(HK)`
- [ ] `Sprd`
- [ ] `Vol`
- [ ] `PR`
- [ ] `arrivalTime`
- [ ] `sector`
- [ ] `client`
- [ ] `Trader`

## Not available, and therefore out of scope

Listed so the gaps are explicit rather than discovered late.

- [ ] Closing-auction market volume per sym/date — would replace `%Adv` as the
      capacity measure with our actual share of the auction
- [ ] Published auction imbalance
- [ ] Auction price level vs the last continuous price
- [ ] Close-eligibility tagging — whatever the previous platform called the flag
      that marked an order for the close

Full reasoning for every column: **[COLUMNS-REQUIRED.md](COLUMNS-REQUIRED.md)**.
