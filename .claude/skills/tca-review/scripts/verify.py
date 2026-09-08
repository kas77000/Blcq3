#!/usr/bin/env python3
"""Arithmetic checks on transcribed TCA data.

Catches what arithmetic can catch. It cannot catch a digit transposed inside a
single cell, so the human re-read of every source image is still required.

    python verify.py --review reviews/client-h1-2026
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import pandas as pd

BENCHMARKS = {"arrival", "pvwap", "vwap", "twap", "close", "open", "nextopen", "decision"}
VENUE = ["pct_open", "pct_close", "pct_post", "pct_take", "pct_dark"]
PCT_METRICS = {
    "pct_close", "pct_open", "pct_post", "pct_take", "pct_dark", "fill_rate",
    "participation_actual", "participation_target", "pct_adv", "spread_capture_pct",
}


class Report:
    def __init__(self) -> None:
        self.rows: list[tuple[str, str, str]] = []

    def add(self, status: str, check: str, detail: str = "") -> None:
        self.rows.append((status, check, detail))

    def ok(self, c, d=""):
        self.add("PASS", c, d)

    def warn(self, c, d=""):
        self.add("WARN", c, d)

    def fail(self, c, d=""):
        self.add("FAIL", c, d)

    @property
    def failed(self) -> int:
        return sum(1 for s, _, _ in self.rows if s == "FAIL")

    @property
    def warned(self) -> int:
        return sum(1 for s, _, _ in self.rows if s == "WARN")

    def render(self) -> str:
        width = max((len(c) for _, c, _ in self.rows), default=10)
        out = []
        for status, check, detail in self.rows:
            line = f"  {status:<4}  {check:<{width}}"
            if detail:
                line += f"  {detail}"
            out.append(line)
        return "\n".join(out)


def load(data: Path, name: str, rep: Report, required: bool):
    path = data / name
    if not path.exists():
        (rep.fail if required else rep.warn)(f"{name} present", "missing")
        return None
    try:
        df = pd.read_csv(path, keep_default_na=True, na_values=["NA", ""])
    except Exception as exc:
        rep.fail(f"{name} readable", str(exc))
        return None
    rep.ok(f"{name} present", f"{len(df)} rows")
    return df


def check_sources(df: pd.DataFrame, name: str, rep: Report) -> None:
    if "source_image" not in df.columns:
        rep.fail(f"{name} source_image column", "column absent")
        return
    missing = int(df["source_image"].isna().sum())
    if missing:
        rep.fail(f"{name} every row has a source", f"{missing} row(s) without one")
    else:
        rep.ok(f"{name} every row has a source")


def close_to(value: float, target: float, tol: float) -> bool:
    return abs(value - target) <= tol


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--review", required=True)
    args = ap.parse_args()

    root = Path(args.review)
    data = root / "data"
    if not data.is_dir():
        print(f"ERROR: no data folder at {data}", file=sys.stderr)
        return 2

    rep = Report()

    # ---- meta ----------------------------------------------------------
    meta = {}
    meta_path = data / "meta.json"
    if not meta_path.exists():
        rep.fail("meta.json present", "missing")
    else:
        meta = json.loads(meta_path.read_text(encoding="utf-8"))
        rep.ok("meta.json present")
        for key in ("client", "period", "currency", "notional_unit", "sign_convention"):
            if meta.get(key):
                rep.ok(f"meta.{key} set", str(meta[key]))
            else:
                rep.fail(f"meta.{key} set", "empty")
        conv = meta.get("sign_convention")
        if conv not in {"positive_is_savings", "positive_is_cost"}:
            rep.fail("meta.sign_convention valid", f"got {conv!r}")

    scope = load(data, "scope.csv", rep, True)
    bench = load(data, "benchmarks.csv", rep, True)
    load(data, "sources.csv", rep, True)
    execu = load(data, "execution.csv", rep, False)
    monthly = load(data, "monthly.csv", rep, False)
    cohorts = load(data, "cohorts.csv", rep, False)

    strategies = set()

    # ---- scope ---------------------------------------------------------
    if scope is not None and len(scope):
        check_sources(scope, "scope.csv", rep)
        strategies = set(scope["strategy"].dropna().astype(str))
        dupes = scope["strategy"][scope["strategy"].duplicated()].tolist()
        if dupes:
            rep.fail("scope.csv one row per strategy", f"repeated: {dupes}")
        else:
            rep.ok("scope.csv one row per strategy", f"{len(strategies)} strategies")

        if "pct_notional" in scope.columns:
            total = float(pd.to_numeric(scope["pct_notional"], errors="coerce").sum())
            if close_to(total, 100.0, 1.0):
                rep.ok("scope share of value sums to 100", f"{total:.1f}")
            else:
                rep.fail("scope share of value sums to 100", f"sums to {total:.1f}")

        if {"notional_musd", "pct_notional"} <= set(scope.columns):
            n = pd.to_numeric(scope["notional_musd"], errors="coerce")
            p = pd.to_numeric(scope["pct_notional"], errors="coerce")
            if n.sum() > 0 and p.notna().any():
                implied = n / n.sum() * 100.0
                worst = float((implied - p).abs().max())
                if worst <= 1.5:
                    rep.ok("value column agrees with its share column", f"worst gap {worst:.2f}pp")
                else:
                    rep.fail(
                        "value column agrees with its share column",
                        f"worst gap {worst:.2f}pp - one of the two is mistyped",
                    )

        if "orders" in scope.columns:
            o = pd.to_numeric(scope["orders"], errors="coerce")
            if (o.fillna(1) <= 0).any():
                rep.fail("scope order counts positive")
            else:
                rep.ok("scope order counts positive", f"{int(o.sum()):,} orders")

        if "fill_rate_pct" in scope.columns:
            fr = pd.to_numeric(scope["fill_rate_pct"], errors="coerce").dropna()
            if len(fr) and ((fr < 0) | (fr > 100)).any():
                rep.fail("fill rates within 0-100")
            elif len(fr):
                rep.ok("fill rates within 0-100")

    # ---- benchmarks ----------------------------------------------------
    if bench is not None and len(bench):
        check_sources(bench, "benchmarks.csv", rep)
        unknown = sorted(set(bench["benchmark"].dropna().astype(str)) - BENCHMARKS)
        if unknown:
            rep.warn("benchmark names recognised", f"unrecognised: {unknown}")
        else:
            rep.ok("benchmark names recognised")

        if strategies:
            orphan = sorted(set(bench["strategy"].dropna().astype(str)) - strategies - {"ALL"})
            if orphan:
                rep.fail("every benchmark strategy is in scope.csv", f"not in scope: {orphan}")
            else:
                rep.ok("every benchmark strategy is in scope.csv")

        if "n" in bench.columns:
            missing = int(pd.to_numeric(bench["n"], errors="coerce").isna().sum())
            if missing:
                rep.fail("every benchmark row states its sample size", f"{missing} row(s) blank")
            else:
                rep.ok("every benchmark row states its sample size")
            small = bench[pd.to_numeric(bench["n"], errors="coerce") < 30]
            if len(small):
                names = ", ".join(f"{r.strategy}/{r.benchmark}(n={r.n})" for r in small.itertuples())
                rep.warn("samples large enough to score", f"describe only: {names}")

        v = pd.to_numeric(bench["value_bps"], errors="coerce")
        wild = bench[v.abs() > 500]
        if len(wild):
            rep.fail("slippage values plausible", f"{len(wild)} row(s) beyond +/-500bps - check the unit")
        else:
            rep.ok("slippage values plausible")

        if {"ci_lo", "ci_hi"} <= set(bench.columns):
            lo = pd.to_numeric(bench["ci_lo"], errors="coerce")
            hi = pd.to_numeric(bench["ci_hi"], errors="coerce")
            both = lo.notna() & hi.notna()
            if both.any():
                bad = int(((lo > hi) & both).sum())
                outside = int((both & ((v < lo) | (v > hi))).sum())
                if bad or outside:
                    rep.fail("ranges bracket their value", f"{bad} inverted, {outside} outside")
                else:
                    rep.ok("ranges bracket their value")

        arr = bench[bench["benchmark"] == "arrival"]
        if len(arr):
            mean_arr = float(pd.to_numeric(arr["value_bps"], errors="coerce").mean())
            conv = meta.get("sign_convention")
            expected_adverse = mean_arr < 0 if conv == "positive_is_savings" else mean_arr > 0
            detail = f"mean arrival {mean_arr:+.2f}bps under {conv}"
            if expected_adverse:
                rep.ok("sign convention consistent with arrival cost", detail)
            else:
                rep.warn(
                    "sign convention consistent with arrival cost",
                    detail + " - arrival is usually adverse; confirm the convention",
                )

    # ---- execution -----------------------------------------------------
    if execu is not None and len(execu):
        check_sources(execu, "execution.csv", rep)
        pct = execu[execu["metric"].isin(PCT_METRICS)]
        val = pd.to_numeric(pct["value"], errors="coerce")
        bad = pct[(val < 0) | (val > 100)]
        if len(bad):
            rep.fail("percentage metrics within 0-100", f"{len(bad)} row(s) outside")
        elif len(pct):
            rep.ok("percentage metrics within 0-100")

        pivot = execu.pivot_table(index="strategy", columns="metric", values="value", aggfunc="first")
        have = [c for c in VENUE if c in pivot.columns]
        if len(have) == len(VENUE):
            sums = pivot[VENUE].sum(axis=1, min_count=len(VENUE)).dropna()
            offenders = sums[(sums - 100).abs() > 1.0]
            if len(offenders):
                detail = ", ".join(f"{i}={s:.1f}" for i, s in offenders.items())
                rep.fail("venue shares sum to 100", detail)
            elif len(sums):
                rep.ok("venue shares sum to 100", f"{len(sums)} strategies")
        elif have:
            rep.warn("venue shares sum to 100", f"only {len(have)} of 5 splits transcribed")

        if {"participation_actual", "participation_target"} <= set(pivot.columns):
            gap = (pivot["participation_actual"] - pivot["participation_target"]).abs().dropna()
            if len(gap):
                rep.ok("participation actual vs target present", f"worst gap {gap.max():.1f}pp")

    # ---- monthly / cohorts --------------------------------------------
    if monthly is not None and len(monthly):
        check_sources(monthly, "monthly.csv", rep)
        if strategies:
            orphan = sorted(set(monthly["strategy"].dropna().astype(str)) - strategies - {"ALL"})
            if orphan:
                rep.fail("monthly strategies are in scope.csv", f"not in scope: {orphan}")
            else:
                rep.ok("monthly strategies are in scope.csv")

    if cohorts is not None and len(cohorts):
        check_sources(cohorts, "cohorts.csv", rep)
        if "pct_notional" in cohorts.columns:
            total = float(pd.to_numeric(cohorts["pct_notional"], errors="coerce").sum())
            if close_to(total, 100.0, 1.5):
                rep.ok("cohort shares sum to 100", f"{total:.1f}")
            else:
                rep.warn("cohort shares sum to 100", f"sums to {total:.1f} - overlapping groups?")

    # ---- NA inventory --------------------------------------------------
    na_total = 0
    for name, df in (("scope.csv", scope), ("benchmarks.csv", bench),
                     ("execution.csv", execu), ("monthly.csv", monthly),
                     ("cohorts.csv", cohorts)):
        if df is None:
            continue
        count = int(df.isna().sum().sum())
        na_total += count
        if count:
            cols = [c for c in df.columns if df[c].isna().any()]
            rep.warn(f"{name} unreadable cells", f"{count} NA in {cols}")

    approx = 0
    for df in (scope, bench, execu, monthly, cohorts):
        if df is not None and "approx" in df.columns:
            approx += int((df["approx"].astype(str).str.upper() == "Y").sum())
    if approx:
        rep.warn("values read off charts", f"{approx} marked approx - write them as 'about'")

    # ---- output --------------------------------------------------------
    print(f"\nVERIFY  {root}\n")
    print(rep.render())
    print()
    print(f"  {len(rep.rows)} checks: {rep.failed} FAIL, {rep.warned} WARN")
    if na_total:
        print(f"  {na_total} unreadable cell(s) - name every one in verification.md")
    print()
    if rep.failed:
        print("Arithmetic checks failed. Fix the transcription before the human pass.")
    else:
        print("Arithmetic is consistent. That is necessary, not sufficient:")
        print("re-read every source image fresh, THEN compare with the CSV.")
    print("Write verification.md, get the user sign-off, then set")
    print('"verified": true in meta.json.')
    return 1 if rep.failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
