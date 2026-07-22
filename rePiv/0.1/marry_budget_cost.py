#!/usr/bin/env python3
"""
marry_budget_cost.py -- the differential (variance) analytics layer.

WHAT IT DOES
    Reads the two STAGED, already-ingested tables (cost + budget), verifies they
    share the schema and grain the two pipelines are contracted to emit, marries
    them at the fixed comparison key, and writes a compiled Budget-vs-Cost
    variance pivot -- plus a "seam check" listing any key present on only one
    side (your all-surplus / all-deficit signal).

WHY IT READS STAGED DATA, NOT SOURCE FILES
    This is the ANALYTICS side of the two-part puzzle. Each pipeline (cost via
    pivot_extract, budget via its own scripts against its OWN config) owns
    ingestion and stages tidy output. This layer consumes that staged output and
    never re-ingests -- so it stays decoupled from either config. (If you ever
    DO want it to ingest raw source files instead, that is a one-line swap: see
    read_staged().)

LAYOUT (resolved relative to THIS file, so it runs from any directory)
    <folder>/
        marry_budget_cost.py        <- this script
        staged/                     <- drop the two *_ingested.xlsx here
            <cost>_ingested.xlsx      (sheet 'ingested_data', meta_report_type=Cost)
            <budget>_ingested.xlsx    (sheet 'ingested_data', meta_report_type=Budget)
        outputs/                    <- variance pivot lands here (auto-created)

RUN
    python marry_budget_cost.py                       # auto-detect the two staged files
    python marry_budget_cost.py  cost.xlsx  budget.xlsx   # or name them explicitly

REQUIREMENTS
    Python 3.9+ with:  pandas  openpyxl
"""
import os
import sys
import glob

import pandas as pd


# --- Everything resolves relative to THIS file, never the current directory. -
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STAGED_DIR = os.path.join(BASE_DIR, "staged")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")

# THE COMPARISON KEY -- the fixed hierarchical level both pipelines agree on.
# Variant is now included, so this is the full declared hierarchy.
KEY = ["DEPT", "Service", "Variant", "DEPT DESCRIPTION", "Project"]
PERIOD, VALUE, RTYPE = "FiscalYear", "Value", "meta_report_type"

# The contract: every staged table MUST carry these for the marriage to be valid.
REQUIRED = set(KEY + [PERIOD, VALUE, RTYPE])


def read_staged(path):
    """Load one staged ingested table (the 'ingested_data' sheet) and assert it
    honours the shared schema contract.

    To ingest a RAW source pivot here instead of a staged table, swap the read
    for:  import pivot_extract as px; df, _ = px.ingest([path], config=...)
    """
    df = pd.read_excel(path, sheet_name="ingested_data")
    missing = REQUIRED - set(df.columns)
    if missing:
        raise ValueError(
            f"{os.path.basename(path)} is missing required column(s) "
            f"{sorted(missing)} -- the cost and budget pipelines must stage the "
            f"same schema for the marriage to be valid."
        )
    return df


def real_periods(df):
    """Drop the Total-Cost / Total-Budget row-total pseudo-period. Those labels
    differ between the streams, so leaving them in would both fail to align AND
    double-count each side's money."""
    return df[~df[PERIOD].astype(str).str.startswith("Total")]


def seam_check(cost, budg):
    """Make the cost/budget seam OBSERVABLE: return the comparison keys that
    appear on ONE side only. With full coexistence this is empty; when it isn't,
    each row is an all-surplus (budget-only) or all-deficit (cost-only) line --
    signal, not an error."""
    ck = cost[KEY].drop_duplicates()
    bk = budg[KEY].drop_duplicates()

    def only_in(a, b, label):
        merged = a.merge(b, on=KEY, how="left", indicator=True)
        out = merged[merged["_merge"] == "left_only"].drop(columns="_merge")
        out["present_in"] = label
        return out

    return pd.concat(
        [only_in(ck, bk, "cost only (all-deficit line)"),
         only_in(bk, ck, "budget only (all-surplus line)")],
        ignore_index=True,
    )


def marry(cost, budg):
    """Stack the two streams and pivot to a compiled Budget/Cost/Variance table
    at the comparison key (summed across all real fiscal years)."""
    married = pd.concat(
        [real_periods(cost).assign(Measure="Cost"),
         real_periods(budg).assign(Measure="Budget")],
        ignore_index=True,
    )
    piv = pd.pivot_table(
        married, index=KEY, columns="Measure", values=VALUE,
        aggfunc="sum", fill_value=0, margins=True, margins_name="GRAND TOTAL",
    )
    piv = piv[["Budget", "Cost"]]
    piv["Variance"] = piv["Budget"] - piv["Cost"]
    piv["Var %"] = (piv["Variance"] / piv["Cost"].replace(0, pd.NA) * 100).round(1)
    piv["Status"] = piv["Variance"].apply(
        lambda v: "surplus" if v > 0 else ("deficit" if v < 0 else "on-budget"))
    return piv


def _fy_order(fiscal_years):
    """Chronological order of FY labels by their numeric part (FY25 < FY40)."""
    return sorted(fiscal_years,
                  key=lambda s: int("".join(ch for ch in str(s) if ch.isdigit())))


def marry_stacked(cost, budg):
    """Measures as ROWS. Budget / Cost / Variance become an inner row level
    right after the comparison key (the smallest categorization grain), so for
    each key you read a 3-row block VERTICALLY, scanning the fiscal-year columns
    across, then a compiled TOTAL column.

    Index : DEPT / Service / Variant / DEPT DESCRIPTION / Project / Measure
    Columns : FY25 ... FY40, TOTAL
    """
    c = pd.pivot_table(real_periods(cost), index=KEY, columns=PERIOD,
                       values=VALUE, aggfunc="sum", fill_value=0)
    b = pd.pivot_table(real_periods(budg), index=KEY, columns=PERIOD,
                       values=VALUE, aggfunc="sum", fill_value=0)

    fys = _fy_order(set(b.columns) | set(c.columns))
    b = b.reindex(columns=fys, fill_value=0)
    c = c.reindex(columns=fys, fill_value=0)
    v = b - c                                  # Variance = Budget - Cost, per FY
    for frame in (b, c, v):
        frame["TOTAL"] = frame[fys].sum(axis=1)

    # Stack the three measures under each key. reorder_levels puts Measure last
    # (right after Project); sort_index orders keys, then Budget < Cost < Variance.
    stacked = pd.concat({"Budget": b, "Cost": c, "Variance": v}, names=["Measure"])
    stacked = stacked.reorder_levels(KEY + ["Measure"]).sort_index()
    return stacked[fys + ["TOTAL"]]


def _identify(frames):
    """Pick out the single Cost stream and single Budget stream by their
    self-describing meta_report_type."""
    def one(kind):
        hits = [d for d in frames.values() if set(d[RTYPE].dropna().unique()) == {kind}]
        return hits[0] if len(hits) == 1 else None
    return one("Cost"), one("Budget")


def main(argv):
    os.makedirs(STAGED_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Two explicit paths, or auto-detect the pair from ./staged.
    if len(argv) > 2:
        cost = read_staged(os.path.abspath(argv[1]))
        budg = read_staged(os.path.abspath(argv[2]))
    else:
        files = sorted(glob.glob(os.path.join(STAGED_DIR, "*.xlsx")))
        if len(files) < 2:
            sys.exit(f"[stop] need two staged *_ingested.xlsx files in {STAGED_DIR}")
        cost, budg = _identify({p: read_staged(p) for p in files})
        if cost is None or budg is None:
            sys.exit("[stop] could not identify exactly one Cost and one Budget "
                     "stream by meta_report_type in ./staged.")

    seam = seam_check(cost, budg)
    piv = marry(cost, budg)
    piv_stacked = marry_stacked(cost, budg)

    # split the margin row from the body for reporting
    is_gt = piv.index.get_level_values(0) == "GRAND TOTAL"
    body, grand = piv[~is_gt], piv[is_gt].iloc[0]

    stamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
    out = os.path.join(OUTPUT_DIR, f"variance_pivot_{stamp}.xlsx")
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        piv_stacked.to_excel(writer, sheet_name="variance_stacked")
        piv.to_excel(writer, sheet_name="variance_compiled")
        seam.to_excel(writer, sheet_name="seam_check", index=False)

    print("comparison key :", " / ".join(KEY))
    print("key lines      :", len(body), "  (x3 measure rows =",
          len(piv_stacked), "stacked rows)")
    print("fiscal years   :", len(piv_stacked.columns) - 1, "(+ TOTAL column)")
    print("seam (one-side):", len(seam), "(expected 0 given full coexistence)")
    print(f"grand total    : Budget {grand.Budget:,.0f} | Cost {grand.Cost:,.0f} | "
          f"Variance {grand.Variance:,.0f} ({grand['Var %']}%)")
    print("written        :", out)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
