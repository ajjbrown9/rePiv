#!/usr/bin/env python3
"""
run_local.py -- local test harness for the pivot_extract ingestion pipeline.

WHAT IT DOES
    Reads every .xlsx in ./inputs, runs the full ingestion + all four gates,
    and writes the tidy result, the audit report, and a focused review list to
    a timestamped folder under ./outputs. It prints the gate VERDICT (would this
    batch load, or would it block?) without ever losing the artifacts, so you
    can open them and see exactly why.

LAYOUT (all resolved relative to THIS file, so it works from any directory)
    <folder>/
        run_local.py            <- this script
        pivot_extract.py        <- the library
        configuration.json      <- the declaration (SME-owned)
        inputs/                 <- drop your source .xlsx exports here
        outputs/                <- results land here (auto-created)

RUN
    python run_local.py
    python run_local.py  path/to/other_config.json      # optional config override
    python run_local.py  path/to/config.json  path/to/inputs_dir   # both overridable

REQUIREMENTS
    Python 3.9+ with:  pandas  openpyxl  numpy
    (install once:  python -m pip install pandas openpyxl numpy)
"""
import os
import sys
import glob
import datetime as dt

import pandas as pd

import pivot_extract as px


# --- Resolve every path relative to THIS file, never the current directory. --
# os.path.abspath(__file__) is the script's own location; everything hangs off
# its directory, so `python run_local.py` behaves the same no matter where you
# launch it from.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_CONFIG = os.path.join(BASE_DIR, "configuration.json")
DEFAULT_INPUTS = os.path.join(BASE_DIR, "inputs")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")


def discover_inputs(input_dir):
    """Every .xlsx in input_dir, ignoring Excel's temporary ~$ lock files."""
    pattern = os.path.join(input_dir, "*.xlsx")
    return sorted(
        p for p in glob.glob(pattern)
        if not os.path.basename(p).startswith("~$")
    )


def main(argv):
    # Optional overrides:  argv[1] = config path,  argv[2] = inputs directory.
    config_path = os.path.abspath(argv[1]) if len(argv) > 1 else DEFAULT_CONFIG
    input_dir = os.path.abspath(argv[2]) if len(argv) > 2 else DEFAULT_INPUTS

    # Make sure the working folders exist (first run creates them).
    os.makedirs(input_dir, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if not os.path.isfile(config_path):
        sys.exit(f"[stop] no configuration file at: {config_path}")

    paths = discover_inputs(input_dir)
    if not paths:
        sys.exit(f"[stop] no .xlsx files found in: {input_dir}\n"
                 f"        drop your source exports there and re-run.")

    print(f"config : {config_path}")
    print(f"inputs : {len(paths)} file(s) in {input_dir}")
    for p in paths:
        print("   -", os.path.basename(p))

    # Run the pipeline with the gate OFF so we always get artifacts to inspect,
    # then ask the gate for its verdict ourselves. (In production you'd leave
    # gate=True so a bad batch raises before anything is written downstream.)
    data, report = px.ingest(paths, config=config_path, gate=False)
    ok, high = px.load_gate(report, raise_on_fail=False)
    verdict = ("WOULD LOAD  (no blocking issues)" if ok
               else f"WOULD BLOCK ({len(high)} HIGH issue(s))")

    # Timestamped run folder so successive runs don't clobber each other.
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    run_dir = os.path.join(OUTPUT_DIR, f"ingest_{stamp}")
    os.makedirs(run_dir, exist_ok=True)

    # Name the output after the source when the batch is a single file, so the
    # cost and budget runs produce distinct, self-describing artifacts that can
    # be dropped straight into the analytics layer's ./staged folder.
    if len(paths) == 1:
        stem = os.path.splitext(os.path.basename(paths[0]))[0]
        data_name = f"{stem}_ingested.xlsx"
    else:
        data_name = "batch_ingested.xlsx"
    data_path = os.path.join(run_dir, data_name)
    with pd.ExcelWriter(data_path, engine="openpyxl") as writer:
        (data if not data.empty else pd.DataFrame()).to_excel(
            writer, sheet_name="ingested_data", index=False)
        report.to_excel(writer, sheet_name="audit_report", index=False)

    review_path = os.path.join(run_dir, "review_high.csv")
    high.to_csv(review_path, index=False)

    # --- summary ------------------------------------------------------------
    print()
    print("=" * 64)
    print("VERDICT      :", verdict)
    print("rows ingested:", len(data))
    print("HIGH issues  :", len(high))
    if len(high):
        for method, n in high.method.value_counts().items():
            print(f"   - {method}: {n}")
    print("outputs      :")
    print("   data   ->", data_path)
    print("   review ->", review_path)
    print("=" * 64)

    # Non-zero exit code when the batch would block -- handy for CI/scripts.
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
