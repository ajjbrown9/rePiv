# Budget vs Cost Pipeline — Operating Manual

Beginning-to-end instructions for the whole pipeline: ingesting Excel report
exports, marrying cost against budget, and producing the variance deliverables.

---

## Contents

1. [What this pipeline does](#1-what-this-pipeline-does)
2. [Environment setup](#2-environment-setup)
3. [Folder layout](#3-folder-layout)
4. [Stage 1 — Ingestion](#4-stage-1--ingestion)
5. [Stage 2 — Staging](#5-stage-2--staging)
6. [Stage 3 — The marriage (variance analytics)](#6-stage-3--the-marriage-variance-analytics)
7. [Stage 4 — Reporting (PDF)](#7-stage-4--reporting-pdf)
8. [The full run, start to finish](#8-the-full-run-start-to-finish)
9. [When a load is blocked](#9-when-a-load-is-blocked)
10. [Changing things (the config)](#10-changing-things-the-config)
11. [Reference outputs and sample files](#11-reference-outputs-and-sample-files)
12. [Developer tools](#12-developer-tools)
13. [Known limits — read before trusting a number](#13-known-limits--read-before-trusting-a-number)

---

## 1. What this pipeline does

The pipeline has two halves that meet at a comparison layer.

**Ingestion** takes messy Excel pivot/report exports and turns them into a tidy,
validated, long-format table — one row per (dimensional key × fiscal period).
It refuses to load anything it cannot vouch for. Four independent gates run on
every file:

| Gate | Question it asks | Catches |
|---|---|---|
| **Arithmetic** | Do the numbers tie? | Leaf sums ≠ the sheet's own total row |
| **Structural** | Is the shape right? | A declared hierarchy level missing/unmapped |
| **Dimensional** | Is every row addressable? | Null keys that vanish from every `GROUP BY` |
| **Identity** | Is every key a declared member? | `"A "` vs `"A"`, `"b"` vs `"B"`, new members |

Any HIGH finding blocks the load. That is deliberate: a file that parses
beautifully and loads *wrong numbers* is the failure this exists to prevent.

**Analytics** takes the two staged outputs — cost and budget, same schema, same
grain — aggregates both to the fixed comparison key, and nets them into
surplus/deficit variance. Output is an Excel pivot and a canned PDF.

```
  cost export.xlsx  ─┐                                  ┌─ variance_pivot.xlsx
                     ├─► [ ingestion + 4 gates ] ─► staged ─┤
budget export.xlsx  ─┘                                  └─ variance_report.pdf
```

**The governing principle:** all business and format rules live in
`configuration.json`, not in the code. Point the module at a different config
and it ingests a different business, with no code change.

---

## 2. Environment setup

**Python 3.9 or newer.** Tested on 3.12.

```bash
python -m pip install -r requirements.txt
```

That installs `pandas`, `openpyxl`, `numpy` (ingestion + analytics),
`reportlab` (PDF stage — pure Python, no system libraries), and `pytest`
(dev regression suite only).

> **Python 3.7 note.** The code itself contains no 3.8+ syntax and would run,
> but 3.7 forces pip back to `pandas==1.3.5` / `numpy==1.21.6`, which is *not*
> what this was tested against. Subtle differences in nullable-string dtypes and
> `groupby` behaviour between pandas 1.3 and 3.x are exactly the kind of thing
> that produces right-looking-but-wrong numbers. If 3.7 is unavoidable, run the
> `dev/` regression suite first and treat green-or-not as the gate. Python 3.7
> has also been end-of-life since June 2023.

---

## 3. Folder layout

Every script resolves its paths with `os.path` relative to **its own location**,
so you can run them from any working directory.

```
budget_cost_pipeline/
├── MANUAL.md                     ← this file
├── CODE_GUIDE.md                 ← how to navigate pivot_extract.py internals
├── INPUT_REQUIREMENTS.md         ← hand this to SMEs who submit source files
├── requirements.txt
│
├── pivot_extract.py              ← the ingestion library (never run directly)
├── run_local.py                  ← STAGE 1  ingestion runner
├── marry_budget_cost.py          ← STAGE 3  variance analytics
├── variance_to_pdf.py            ← STAGE 4  PDF, indented outline form
├── variance_to_pdf_columnar.py   ← STAGE 4  PDF, columnar form (alternative)
├── configuration.json            ← THE DECLARATION (5-level, current)
│
├── inputs/
│   ├── cost/                     ← drop cost exports here
│   └── budget/                   ← drop budget exports here
├── staged/                       ← ingested outputs land here for the marriage
├── outputs/                      ← all generated results (timestamped)
│
├── reference_outputs/            ← known-good results from the tested run
├── samples/                      ← other sample source files + legacy config
└── dev/                          ← regression suite, fixtures, refactor branch
```

**Why `inputs/cost` and `inputs/budget` are separate:** the ingestion runner
processes every file in a directory as *one batch*. Cost and budget must be
staged as two distinct tables so the analytics layer can tell them apart, so
they are ingested in two separate runs.

---

## 4. Stage 1 — Ingestion

Run once for cost, once for budget:

```bash
python run_local.py configuration.json inputs/cost
python run_local.py configuration.json inputs/budget
```

Both arguments are optional (defaults: `./configuration.json` and `./inputs`),
but for this pipeline you want the explicit input directories above.

**What you get:**

```
================================================================
VERDICT      : WOULD LOAD  (no blocking issues)
rows ingested: 8500
HIGH issues  : 0
outputs      :
   data   -> outputs/ingest_<timestamp>/<source>_ingested.xlsx
   review -> outputs/ingest_<timestamp>/review_high.csv
================================================================
```

- The output workbook has two sheets: **`ingested_data`** (the tidy rows) and
  **`audit_report`** (every heal, fill, match and check — the full audit trail).
- **`review_high.csv`** is the focused list of blocking issues. Empty on a clean run.
- **Exit code** is `0` when the batch would load and `1` when it would block, so
  the runner can be scripted or dropped into CI.

The runner deliberately writes its artifacts **even when the batch blocks**, so
you can open the audit and see exactly why. (In a production pipeline you would
call `px.ingest(...)` directly with `gate=True`, which raises before anything
reaches the warehouse.)

**Output columns:** `source_file`, `source_sha256`, `ingested_at`, the `meta_*`
metadata fields, the declared hierarchy levels, `nulled_levels`, `sub_level`,
`is_parent`, `is_leaf`, `parent_label`, `path`, `FiscalYear`, `Value`, `row_key`.

---

## 5. Stage 2 — Staging

Copy the two ingested workbooks into `staged/`:

```bash
cp outputs/ingest_*/Extraction_Sample_7_2_500_ingested.xlsx        staged/
cp outputs/ingest_*/Extraction_Sample_7_2_500_BUDGET_ingested.xlsx staged/
```

(Windows: use `copy`, or drag them across.)

The analytics layer identifies which file is which by reading the
**`meta_report_type`** column — `Cost` vs `Budget` — not by filename. So the
names don't matter, but exactly one of each must be present.

> This is the seam between the two halves. In production, the budget pipeline
> writes its own staged output here directly; the analytics layer never
> re-ingests, which keeps it decoupled from either side's configuration.

---

## 6. Stage 3 — The marriage (variance analytics)

```bash
python marry_budget_cost.py
```

Or name the two files explicitly:

```bash
python marry_budget_cost.py staged/cost_ingested.xlsx staged/budget_ingested.xlsx
```

**What it does:**

1. **Validates the schema contract** — both tables must carry the five key
   columns plus `FiscalYear`, `Value`, `meta_report_type`, or it refuses. The
   two pipelines must agree on shape for the marriage to mean anything.
2. **Drops the row-total pseudo-period** (see [Known limits](#13-known-limits--read-before-trusting-a-number)).
3. **Marries and pivots** at the comparison key:
   `DEPT / Service / Variant / DEPT DESCRIPTION / Project`.
4. **Runs a seam check** for keys present on only one side.

**Output** — `outputs/variance_pivot_<timestamp>.xlsx`, three sheets:

| Sheet | What it is |
|---|---|
| `variance_stacked` | Measures as **rows**: Budget / Cost / Variance stacked under each key, fiscal years across, TOTAL column. The OLAP-style view. |
| `variance_compiled` | One row per key, Budget / Cost / Variance / Var % / Status compiled across all years, with a grand total. |
| `seam_check` | Keys present on one side only — an all-surplus (budget-only) or all-deficit (cost-only) line. Empty when coverage is complete. |

**Console summary:**

```
comparison key : DEPT / Service / Variant / DEPT DESCRIPTION / Project
key lines      : 368   (x3 measure rows = 1104 stacked rows)
seam (one-side): 0 (expected 0 given full coexistence)
grand total    : Budget 19,776,314 | Cost 20,111,285 | Variance -334,971 (-1.7%)
```

**To change the comparison level**, edit `KEY` at the top of
`marry_budget_cost.py`. Levels you remove are aggregated over.

---

## 7. Stage 4 — Reporting (PDF)

Two layouts. Both are landscape, sectioned by DEPT, with column headers
repeating on every page and variance coloured green (surplus) / red (deficit),
negatives in accounting parentheses.

```bash
python variance_to_pdf.py            # indented outline form
python variance_to_pdf_columnar.py   # columnar form
```

| Layout | Row labels | Pages* | Best for |
|---|---|---|---|
| **Indented** (`variance_to_pdf.py`) | One `Category` column, each level stepped right | ~55 | Narrow label block, larger font, classic pivot outline |
| **Columnar** (`variance_to_pdf_columnar.py`) | Five separate label columns | ~34 | Fewer pages; every level readable as its own column |

\* for the 368-key sample at FY25–FY32.

**The reporting window** is set by `FY_KEEP` at the top of either script:

```python
FY_KEEP = [f"FY{y}" for y in range(25, 33)]   # FY25..FY32 inclusive (8 years)
```

Change the range and the `TOTAL` column recomputes over whatever window you set.

Both scripts import `marry_budget_cost` for data prep, so they must stay in the
same folder as it.

---

## 8. The full run, start to finish

```bash
# 0. one time
python -m pip install -r requirements.txt

# 1. ingest both sides (four gates run here)
python run_local.py configuration.json inputs/cost
python run_local.py configuration.json inputs/budget

# 2. stage the two ingested workbooks
cp outputs/ingest_*/*_ingested.xlsx staged/

# 3. marry them into the variance pivot
python marry_budget_cost.py

# 4. render the PDF
python variance_to_pdf.py

# results in ./outputs
```

Expected on the shipped sample data: both ingests report **WOULD LOAD, 0 HIGH,
8,500 rows**; the marriage reports **368 key lines, seam 0, variance
−334,971 (−1.7%)**.

---

## 9. When a load is blocked

A block is the system working. Open `review_high.csv` (or the `audit_report`
sheet) and read the `method` and `note` columns. The common verdicts:

| `method` | Meaning | Usual fix |
|---|---|---|
| `member_undeclared` | A key is not in the level's vocabulary | Add the member to `configuration.json` (a business decision), or correct the source |
| `hierarchy_missing` | A declared level isn't present after aliasing | Add the source column's spelling to that level's `aliases` |
| `hierarchy_unexpected` | A column isn't a declared level | Declare it, or remove it from the source |
| `dimension_null` | Rows have no value for a declared level | Fill the gap in the source |
| `reconciliation_mismatch` | Leaf sums ≠ the sheet's total row | The parse or the source total is wrong — investigate before loading |
| `metadata_missing` / `report_name_mismatch` | The `Report:` block is absent or ≠ the sheet name | Fix the metadata block or the tab name |
| `file_error` | The file could not be parsed at all | See the note; often a structural mismatch |

The identity gate hands you the answer rather than applying it: an undeclared
`'b'` reports *"closest known member is 'B' (1.0) — approve or correct."*
**Fuzzy matching suggests; it never renames a member on your behalf.**

Priorities: `OK` = informational (heals, fills — auditable, non-blocking),
`OVERRIDDEN` = a declared exception someone approved, `HIGH` = blocks.

Hand `INPUT_REQUIREMENTS.md` to whoever submits source files — it is the
layman-readable version of what the gates require.

---

## 10. Changing things (the config)

**`configuration.json` is the artifact. The code only obeys it.**

| You want to… | Change |
|---|---|
| Add a department / service / project member | that level's `vocabulary` |
| Accept a new source column spelling | that level's `aliases` |
| Add or reorder a hierarchy level | the `hierarchy` list (order **is** the business level) |
| Adjust fuzzy strictness / recon tolerance | `matching` / `reconciliation` blocks |
| Support a localized workbook | `total_labels`, `aggregation_prefixes`, `blank_member_literal` |
| Let one report legitimately lack a level | `overrides.allow_missing_levels` |

The current hierarchy is five levels: **DEPT › Service › Variant ›
DEPT DESCRIPTION › Project**.

Two config-level safety features worth knowing:

- **Self-validation at load.** A contradictory declaration fails at startup, not
  halfway through a batch — e.g. a `closed` level with no vocabulary, or a
  normalization license that would merge two declared members.
- **Governance modes** per level: `closed` (hand-written vocabulary; small,
  stable levels), `registry` (list builds itself in `members.json`, new members
  block until bulk-approved; granular levels), `observed` (no membership check;
  near-duplicate and cardinality-drift detection only).

If you find yourself editing Python to change a *business* rule, it belongs in
the config instead.

> **Cost/budget config sync.** The two pipelines currently run against separate
> config files kept in step by hand. While that's true, a divergence at the
> comparison level is invisible to both pipelines — each file still validates,
> each reconciles, and the marriage silently nets against the wrong bucket. If
> the interim runs long, a read-only check comparing the two configs'
> comparison-level members is cheap and doesn't depend on merging them.

---

## 11. Reference outputs and sample files

**`reference_outputs/`** — known-good results from the tested run. Compare
against these after a change to confirm nothing drifted.

| File | |
|---|---|
| `Extraction_Sample_7_2_500_ingested.xlsx` | Staged cost (8,500 rows, 0 HIGH) |
| `Extraction_Sample_7_2_500_BUDGET_ingested.xlsx` | Staged budget (8,500 rows, 0 HIGH) |
| `Budget_vs_Cost_variance_stacked.xlsx` | Measures-as-rows pivot |
| `Budget_vs_Cost_variance_by_FY.xlsx` | Fiscal years across, Budget/Cost/Variance per year |
| `Budget_vs_Cost_variance_pivot_5key.xlsx` | Compiled pivot, full 5-level key |
| `Budget_vs_Cost_variance_pivot.xlsx` | Compiled pivot, 4-level key (no Variant) |
| `Budget_vs_Cost_variance_report_indented.pdf` | PDF, indented outline |
| `Budget_vs_Cost_variance_report.pdf` | PDF, columnar |

**`inputs/`** — the core sample source files used in the final iteration:

- `inputs/cost/Extraction_Sample_7_2_500.xlsx` — 500 rows, five hierarchy
  levels, FY25–FY40, `Report Type: Cost`. All members drawn from the declared
  config vocabularies.
- `inputs/budget/Extraction_Sample_7_2_500_BUDGET.xlsx` — the same 500
  dimensional keys, `Report Type: Budget`, with a per-DEPT bias and per-cell
  jitter applied so the variance is directional rather than noise (~15% of cells
  left exactly on budget).

**`samples/`** — other files from development: `Extraction_Sample_6.xlsx` (the
original three-level sample), `Extraction_Sample_7_2.xlsx` (the 20-row template
that established the current structure), and
`configuration_3level_legacy.json` (the three-level config those older samples
were built against).

> **Reading the pivots back into pandas.** They carry a real MultiIndex, so Excel
> renders them as normal grouped pivots, but a plain `read_excel` shows repeated
> labels as `NaN`. Reattach the index:
> `pd.read_excel(f, sheet_name="variance_stacked", index_col=[0,1,2,3,4,5])`
> (and `header=[0,1]` for the by-FY sheet).

---

## 12. Developer tools

Everything in `dev/` is for changing the code, never for production runs.

- **`test_pivot_extract.py`** — the regression suite. Tests assert the *gate
  verdict*, not that parsing "worked." The headline test,
  `test_dirty_member_values_are_blocked`, feeds a file whose arithmetic ties
  exactly and which has no nulls — and demands a block anyway, because it turns
  three departments into six. If it ever passes, the identity gate has regressed.
- **`fixtures/`** — the ten deliberate mutation files (real Excel indent, blanked
  repeat labels, accounting negatives, a department named "Total Rewards",
  text-stored numbers, compact form, numeric-coded dimensions, subtotal rows,
  deep nesting, dirty member values).
- **`make_fixtures.py`** — regenerates those ten from `Extraction_Sample_6.xlsx`,
  so they're reproducible rather than snapshots.
- **`pivot_extract.staged.py`** + **`.diff`** — a reviewed refactor branch
  (shared header/classification helpers, memoized sheet reads cutting per-parse
  disk reads from 6 to 4, `PROVENANCE_COLUMNS` used as the single source of
  truth). Proven to produce byte-identical output on the sample data. **Not
  adopted** — swap it in deliberately, with the suite watching.
- **`pivot_extract_README.md`** — the library's own design README.

**To run the suite** you must adjust two paths first: `UPLOADS` at the top of
`test_pivot_extract.py` (it points at an absolute path), and run from a folder
where the fixtures sit at the expected relative locations. The three
`Extraction_Sample_1/2/3.xlsx` legacy tests will fail regardless — those SME
originals are not part of this package.

---

## 13. Known limits — read before trusting a number

**1. The row-total column rides along as an extra period.** Source files carry a
`Total Cost` / `Total Budget` column. The parser cannot distinguish a row-total
from a fiscal period, so it is melted in alongside FY25–FY40 — and it
*reconciles*, so no gate objects. **Summing `Value` across all periods
double-counts.** The analytics layer filters it out (`real_periods()`), and you
should too in any downstream query. The clean fix is to drop that column before
ingest. Note the labels differ between the two streams, so it also fails to
align across the marriage.

**2. Misattribution is invisible.** A row assigned to the wrong *but valid*
department loads cleanly — sums tie, keys are declared, no nulls, **all four
gates pass**. This is inherent to reconciling against a single grand total. Do
not treat "an analyst will spot it downstream" as a control. The real fix is a
second independent statement of the numbers — per-DEPT control totals from the
source system — after which the arithmetic gate simply runs at a finer grain.

**3. Header detection needs a non-numeric header row.** Periods labelled as bare
integers (`2025`) make the header row read as data and defeat detection
entirely — the parse fails before any gate runs. Text labels (`FY25`,
`Sum of FY26`) are required. This is the single most common reason a new
template won't ingest.

**4. Compact-form pivots are rejected, not parsed.** All levels collapsed into
one indented column is detected and blocked; supporting it is new work.

**5. The vocabulary must be owned by an SME.** The current one was inferred from
sample data. Until a human owns it, the identity gate is enforcing a fiction
with great rigour.

**6. Recon tolerance is absolute (0.01), not relative.** At large scale, float
summation error alone can exceed a penny. Expect false failures, and expect
someone to "temporarily" loosen it.

**7. `row_key` uniqueness assumes real pivot grain.** The key is
`hash(source + dimensions + period)` so a reload merges onto itself instead of
doubling the facts. It is unique only when each dimensional coordinate appears
once — true of an aggregated pivot, but not of randomly generated test data
where combinations repeat.
