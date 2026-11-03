"""Build bgi_order_types.csv by deriving OrderType from the hedge file
and merging it onto the order file on CUST_SG_SILK_COMMON_REFERENCE."""
import pandas as pd

HEDGE_FILE = "BLACROC_Orders_0126_1026_hedge.csv"
ORDER_FILE = "BLACROC_Orders_0126_1026_swap.csv"
OUTPUT_FILE = "bgi_order_types.csv"

# The source files are UTF-16 encoded.
ENCODING = "utf-16"


def is_populated(series: pd.Series) -> pd.Series:
    """Return a boolean mask of non-empty (populated) values."""
    return series.notna() & (series.astype(str).str.strip() != "")


def compute_order_type(df: pd.DataFrame) -> pd.Series:
    """Stipulated if CUST_XP_DATA_54 is populated OR
    CUST_XP_DATA_68 == "CLOSEONLY"; otherwise Discretionary."""
    stipulated = is_populated(df["CUST_XP_DATA_54"]) | (
        df["CUST_XP_DATA_68"].astype(str).str.strip() == "CLOSEONLY"
    )
    return stipulated.map({True: "Stipulated", False: "Discretionary"})


def compute_order_subtype(df: pd.DataFrame) -> pd.Series:
    """Sub-category, alongside OrderType:
    CloseAuctionOnly if CUST_XP_DATA_68 == "CLOSEONLY";
    else PctOfOrder if CUST_XP_DATA_54 == "PCTORDERSIZE";
    otherwise Model."""
    close_only = df["CUST_XP_DATA_68"].astype(str).str.strip() == "CLOSEONLY"
    pct_order = df["CUST_XP_DATA_54"].astype(str).str.strip() == "PCTORDERSIZE"
    sub = pd.Series("Model", index=df.index)
    sub[pct_order] = "PctOfOrder"
    sub[close_only] = "CloseAuctionOnly"
    return sub


def drop_buy_sell(df: pd.DataFrame) -> pd.Series:
    """Mask of rows where CL_BUY_SELL == "S" and BUY_SELL_QUALIFIER == "S"."""
    return (df["CL_BUY_SELL"].astype(str).str.strip() == "S") & (
        df["BUY_SELL_QUALIFIER"].astype(str).str.strip() == "S"
    )


def main() -> None:
    hedge = pd.read_csv(HEDGE_FILE, encoding=ENCODING, dtype=str)
    order = pd.read_csv(ORDER_FILE, encoding=ENCODING, dtype=str)

    hedge["OrderType"] = compute_order_type(hedge)
    hedge["OrderSubType"] = compute_order_subtype(hedge)

    # A reference can (rarely) match more than one hedge row. Keep the
    # latest one by CL_ENTERED_DATETIME so the merge is 1:1 and uses the
    # most recent OrderType.
    hedge["_entered"] = pd.to_datetime(
        hedge["CL_ENTERED_DATETIME"].str.replace(r"\s+", " ", regex=True),
        format="%b %d %Y %I:%M%p",
        errors="coerce",
    )
    hedge_lookup = (
        hedge.sort_values("_entered")
        .drop_duplicates("CUST_SG_SILK_COMMON_REFERENCE", keep="last")
    )[["CUST_SG_SILK_COMMON_REFERENCE", "OrderType", "OrderSubType"]]

    merged = order.merge(hedge_lookup, on="CUST_SG_SILK_COMMON_REFERENCE", how="left")

    # For order rows whose CLIENT_VIEW_CODE contains "XXX" or "DXA",
    # apply the hedge rule directly on the order row's own columns.
    direct_check = merged["CLIENT_VIEW_CODE"].astype(str).str.contains(
        "XXX|DXA", na=False
    )
    merged.loc[direct_check, "OrderType"] = compute_order_type(merged[direct_check])
    merged.loc[direct_check, "OrderSubType"] = compute_order_subtype(
        merged[direct_check]
    )

    # Order rows with no matching hedge reference have no OrderType from the
    # merge; by the rule they evaluate to Discretionary / Model, so default them.
    merged["OrderType"] = merged["OrderType"].fillna("Discretionary")
    merged["OrderSubType"] = merged["OrderSubType"].fillna("Model")

    # Drop rows where CL_BUY_SELL == "S" and BUY_SELL_QUALIFIER == "S".
    merged = merged[~drop_buy_sell(merged)]

    order_result = merged[["CL_ORDER_ID", "OrderType", "OrderSubType"]]

    # Also include the hedge-file elements with their own OrderType,
    # dropping rows where CL_BUY_SELL == "S" and BUY_SELL_QUALIFIER == "S".
    hedge_result = hedge[~drop_buy_sell(hedge)][
        ["CL_ORDER_ID", "OrderType", "OrderSubType"]
    ]

    result = pd.concat([order_result, hedge_result], ignore_index=True).rename(
        columns={"CL_ORDER_ID": "OrderId"}
    )

    result.to_csv(OUTPUT_FILE, index=False)
    print(f"Wrote {len(result)} rows to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
