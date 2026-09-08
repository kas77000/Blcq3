# What to check on the target machine

Ordered by how much each answer changes the analysis. **Step 1 answers about
fifteen of these at once** — if you only do one thing, do that.

---

## Step 1 — run the probe and send me the log

```bash
python moc_tca.py --data <the extract> --probe
```

It reads the file and stops. Nothing is written except `output/run_log.txt`.
**Send me that file.** It contains no order-level data — headers, counts,
percentiles and group means only — so it is safe to pass around.

That one output settles:

- every header in the file, its dtype and a sample value
- which columns mapped and which did not
- **every `Strategy` value with order count, notional, mean `%CLOSE` and mean
  `vs Close`** — this is what pins the close-algo list
- whether the five venue columns sum to 100
- whether the bps columns are side-adjusted
- whether the `%` columns are percentages or fractions
- which markets are present, and whether any symbol suffix failed to map
- any `Cap` / `marketLimit` / `arrivalTime` value outside the expected set

If the script will not run there, this gets me most of the same thing:

```python
import pandas as pd
df = pd.read_csv("<the extract>")          # or read_excel
print(list(df.columns))
print(df.dtypes)
print(df.head(3).to_string())
g = df.groupby("Strategy")
print(pd.DataFrame({"orders": g.size(), "mln": g["$Mln"].sum(),
                    "pctClose": g["%CLOSE"].mean(),
                    "close_bps": g["Close"].mean()}).to_string())
```

---

## Step 2 — six questions the probe cannot answer

These are about what the columns *mean*. Each one silently changes a headline
number if it is assumed wrong, and none of them shows up as an error.

### 1. What exactly does `%CLOSE` count? — **the biggest risk on the list**

- Is it the **closing auction only**, or all auction volume, or everything
  traded in a closing *period* (last N minutes) whether or not it printed in
  the auction?
- Is it a share of **quantity** or of **notional**? The script assumes quantity.

If `%CLOSE` is a time window rather than the auction itself, then "cleared the
auction" is not what it measures and the whole taxonomy needs relabelling. The
presence of a separate `%OPEN` suggests it really is the auction, but it is
worth one sentence of confirmation.

### 2. Is `NextOpen` adjusted for corporate actions?

The reversion exhibit is `NextOpen - Close`. An ex-dividend date between the
close and the next open puts the dividend straight into that number as fake
reversion, and on a dividend-heavy stretch that is not a small effect. If the
benchmark set is unadjusted, reversion gets reported per market with a warning,
or dropped.

### 3. Is `first_exec_vs_close` side-adjusted?

The probe now checks this alongside the other bps columns, so the log will
answer it — but confirm it independently if you can. If it is raw rather than
side-adjusted, every sell is inverted and the "was starting early right?"
conclusion flips sign.

### 4. Is `$Mln` gross or net of commission and fees?

It is the weight on every aggregate and the basis of every currency figure. Net
of fees is fine, it just has to be stated, because the cost-of-leakage number is
then after-fee and is not comparable to a gross figure from elsewhere.

### 5. One row per order, or can `aggrTgtId` repeat?

If an amend or a replacement writes a second row under the same id, both the
order count and the notional are inflated. The probe prints duplicate ids if
there are any; a yes/no from the data owner is a useful cross-check.

### 6. Are `Start(HK)` / `End(HK)` in HKT for every market, or local time?

The names say HKT. If Japan and Australia rows are actually stamped local, the
`close_gap_min` measure is wrong by the offset for those markets. This affects
one secondary exhibit only, so it is last.

Also, for India specifically: **what does `%CLOSE` even mean there?** NSE has no
single-price closing auction — it closes on a VWAP of the last half hour. India
is already reported separately for exactly this reason, but knowing what the
platform put in that column for Indian orders decides whether the number is
usable at all.

---

## Step 3 — what else exists in the extract

We have what we need for a strong review. These would make it stronger, in
order of value. A yes/no on each is enough.

| Wanted | What it unlocks |
|---|---|
| **Closing-auction market volume**, per symbol and date | Turns capacity from a `%Adv` proxy into **our actual share of the auction** — the real constraint, and the difference is largest in small and micro caps, which is exactly where the misses concentrate. The desk's own market-stats query already splits auction from continuous volume by the clock, so this may be close to free. **Highest value on this list.** |
| **Close-eligibility flag** — whatever the old platform called it | The only way to separate "the algo missed the close" from "the order was never marked for it". Would convert the unexplained cohort from a residual into a measurement. |
| **Limit price** | Confirms the limit-did-not-cross finding with the actual price instead of inferring it from `marketLimit`. |
| **Cancel / amend timestamp** | Separates "the algo missed the close" from "the client pulled the order before it". A real and common confound in the unexplained cohort. |
| **Published auction imbalance** | Whether we traded into or against the imbalance. The most actionable close-specific behaviour there is, and currently entirely absent. |

---

## Step 4 — a sanity read on the data itself

Quick eyeball, either from the probe log or directly. Any "yes" here is worth
telling me about before I interpret anything.

- Do any **close-algo orders have `%OPEN` > 0**? They should not — a close order
  touching the opening auction is either mis-bucketed or a genuinely odd order.
- Is `%CLOSE` **bimodal** (piled at 0 and near 100), or spread across the middle?
  Bimodal means orders mostly either make the auction or miss it entirely, which
  points at a mechanical gate rather than a capacity limit — a different story
  and a different recommendation.
- Any rows with `FR > 100`, `%CLOSE > 100`, or the five venue columns not summing
  to 100? The probe reports all three; they mean a column is not what we think.
- Does the **order count per month** hold roughly steady across H1, or is there a
  step? A step would mean something changed in scope, coverage or extraction
  part-way through.
