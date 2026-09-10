#!/usr/bin/env python3
"""Concatenate the AWS parquet extract and left-join it onto the order file.

Two outputs, both CSV:

    aws_all.csv   every parquet in one frame, deduplicated on aggrTgtId
    merged.csv    the order file with the AWS columns joined on

    python merge_aws.py --orders orders.csv --aws data/aws
    python merge_aws.py --orders orders.csv --aws data/aws --out-dir out

The join rules come from moc_tca.py rather than being written again here, so
the merged file is exactly what the analysis sees - the order file stays the
population, an unmatched order is kept with its AWS columns empty, and a name
that exists in both keeps the order file's version with the AWS one suffixed.
Nothing about this script can drift away from the run.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

import moc_tca as tca


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--orders", type=Path, required=True,
                   help="the order extract (.csv or .xlsx)")
    p.add_argument("--aws", type=Path, default=Path(tca.AWS_DIR),
                   help=f"folder of parquet files (default: {tca.AWS_DIR})")
    p.add_argument("--out-dir", type=Path, default=Path("."),
                   help="where to write the two CSVs (default: here)")
    p.add_argument("--aws-only", action="store_true",
                   help="write aws_all.csv and stop, without joining")
    args = p.parse_args(argv)

    if not args.aws.is_dir():
        print(f"ERROR: no such folder: {args.aws}", file=sys.stderr)
        return 2
    args.out_dir.mkdir(parents=True, exist_ok=True)

    print(f"reading parquet from {args.aws}")
    aws = tca.read_aws(args.aws)
    if aws.empty:
        print(f"ERROR: no parquet files found in {args.aws}", file=sys.stderr)
        return 2

    key = tca.AWS_JOIN_KEY
    if key in aws.columns:
        dupes = int(aws[key].duplicated().sum())
        if dupes:
            print(f"  {dupes:,} duplicate {key} values - keeping the first of each")
            aws = aws.drop_duplicates(subset=[key], keep="first")

    aws_path = args.out_dir / "aws_all.csv"
    aws.to_csv(aws_path, index=False)
    print(f"  wrote {aws_path}  ({len(aws):,} rows x {len(aws.columns)} columns)")
    if args.aws_only:
        return 0

    if not args.orders.exists():
        print(f"ERROR: no such file: {args.orders}", file=sys.stderr)
        return 2
    orders = tca.read_file(args.orders)
    print(f"read {args.orders}  ({len(orders):,} rows x {len(orders.columns)} columns)")

    for name, frame in (("order file", orders), ("AWS extract", aws)):
        if key not in frame.columns:
            print(f"ERROR: {key} is not in the {name}; cannot join.", file=sys.stderr)
            return 2

    # Exactly the join the analysis performs - same collision rules, same
    # direction, same suffix.
    overlap = sorted(c for c in aws.columns
                     if c != key and tca._norm_name(c) in
                     {tca._norm_name(o) for o in orders.columns})
    if overlap:
        print(f"  {len(overlap)} AWS column(s) share a name with the order file;"
              f" suffixed with {tca.AWS_SUFFIX}:")
        print("    " + ", ".join(overlap[:12])
              + (f" ... and {len(overlap) - 12} more" if len(overlap) > 12 else ""))

    merged = tca._merge_frames(orders, aws)
    matched = int(merged[key].isin(set(aws[key])).sum())
    print(f"  {matched:,} of {len(orders):,} orders matched "
          f"({100.0 * matched / max(len(orders), 1):.1f}%)")
    if matched < len(orders):
        print(f"  {len(orders) - matched:,} did not match. They are KEPT - the "
              f"order file is the population - with their AWS columns empty.")

    out = args.out_dir / "merged.csv"
    merged.to_csv(out, index=False)
    print(f"  wrote {out}  ({len(merged):,} rows x {len(merged.columns)} columns)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
