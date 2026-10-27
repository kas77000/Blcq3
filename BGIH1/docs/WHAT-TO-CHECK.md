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

## Step 2 — what the columns mean

### Confirmed with the data owner — no longer open

| | |
|---|---|
| `$Mln` | EXECUTED notional, `sum(cumqty * avgprice * fx_last) / 1e6`. **Millions of USD**, so the multiplier back to USD is 1e6. `fx_last` is inside it, so the column is already USD-converted. |
| `%CLOSE` | The % of order quantity executed through CLOSE. The venue shares are shares of **executed** quantity — the run tests this explicitly by checking whether they sum to 100 or to `FR`, and says which. |
| Sign | **All data is side-adjusted: `+` is good, `-` is bad.** Matches `POSITIVE_IS_SAVING = True` and `SIDE_ALREADY_ADJUSTED = True`. |
| `aggrTgtId` | Cannot repeat. One row per order, so no dedup and no double-counted notional. |
| `Start(HK)` / `End(HK)` | HKT for every market, so `close_gap_min` needs no per-market offset. |

Because a wrong notional scale moves every currency figure by a power of ten
and nothing downstream would notice, the run derives the **implied USD share
price** (executed notional / executed shares) and prints it per market, warning
if the median leaves a plausible band. On the synthetic file it lands around
USD 20 a share across all six markets, which is what a correct scale looks like.

### Still open — one item

**Is `NextOpen` adjusted for corporate actions?** Your read is that it should be,
which is probably right, but it is worth confirming because it is the one thing
that would quietly corrupt the reversion exhibit. Reversion is
`NextOpen - Close`; an ex-dividend date between the two puts the dividend
straight into that number as impact that never happened. On a dividend-heavy
stretch — and H1 covers the main APAC dividend season — that is not a small
effect, and it biases in one direction rather than averaging out.

If it turns out unadjusted, the fix is cheap: report reversion per market with
the caveat, or drop the exhibit. Nothing else in the pack depends on it.

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
