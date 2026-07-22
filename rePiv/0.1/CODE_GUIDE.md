# `pivot_extract.py` — Navigation Guide

A map for reading the script: the 30,000-foot purpose of each section, the call
path a file actually travels, and the one critical thing to understand at each
major step. This is a guide to the *code*; the operational README covers how to
*run* it.

---

## 1. The mental model (read this first)

Two kinds of error arrive in a spreadsheet, and they need opposite treatment:

- **Decode errors** — the cell's *form* is misread. `(948)` is a negative; a
  blank means "same as above"; `"8,057"` is a number stored as text. These are
  *loud*: they break arithmetic, and a totals check catches them.
- **Identity errors** — the cell decodes perfectly but is the *wrong key*.
  `"A "` ≠ `"A"`; `"b"` ≠ `"B"`; a non-breaking space makes two identical-looking
  strings unequal. These are *silent*: every sum still ties, no nulls appear,
  and the rows look right in any report. They split one department into two and
  no arithmetic check can see it.

The whole design follows from that split:

> **Heal what is unambiguous. Fail on what is a decision.
> Nothing enters the warehouse as a key until it has been declared as one.**

The parser handles decode errors; a set of **four gates** handles the rest. Any
HIGH finding from any gate blocks the load. The gates are:

| Gate | Question | Lives in section |
|---|---|---|
| **Arithmetic** | Do the leaf rows sum to the source's own total row? | Reconciliation |
| **Structural** | Are exactly the declared hierarchy levels present? | Auto-dispatch (`_apply_hierarchy`) |
| **Dimensional** | Is every row addressable at every level (no null keys)? | Dimension-completeness gate |
| **Identity** | Is every member a *declared* member? | Member-identity gate |

And one principle that pervades everything: **the config is the artifact, the
module only obeys it.** All business/format knowledge lives in
`configuration.json`; point the module at a different config and it ingests a
different business with no code change.

---

## 2. The end-to-end flow

A single file travels this path. Two public entry points, `parse()` (one file)
and `ingest()` (a batch that wraps `parse`):

```
ingest(paths, config)                         [batch, fail-closed by default]
  └─ for each file:
     parse(path, config)                      [the front door / router]
        ├─ _expand_config            load + self-validate the declaration
        ├─ _sheet_is_hierarchical    peek: indented outline column?
        │
        ├─ extract_pivot  ──OR──  extract_hierarchical_pivot
        │     read sheet → find header → resolve headers → classify value/dim
        │     → drop total rows → coerce numbers
        │     → validate_members     [IDENTITY gate: heal + check membership]
        │     → resolve_blank_dims   [suppressed-repeat vs. missing]
        │     → melt to long form
        │
        └─ _finish()  (post-parse gates + provenance)
              ├─ read_metadata_block + validate_metadata   [Report: == sheet]
              ├─ _harmonize                 uniform schema (flat == hier shape)
              ├─ _apply_hierarchy           [STRUCTURAL gate: levels present?]
              ├─ validate_dimension_completeness  [DIMENSIONAL gate: no nulls]
              └─ read_total_row + total_validation  [ARITHMETIC gate: sums tie]
  └─ add source_file / source_sha256 / ingested_at / row_key   [provenance]
  └─ load_gate(report)   → raises "Load blocked" if ANY row is HIGH
```

The **output** is always two things: a tidy long-format DataFrame *and* a report
(one row per column/member/check). Rows in the report carry a `review_priority`
of `OK` (informational), `OVERRIDDEN` (a pre-approved exception), or `HIGH`
(blocks the load). Filtering the report to HIGH is `review_items()`.

---

## 3. Section-by-section map

The file is organized top-to-bottom as constants → building blocks → gates →
extractors → the router → reconciliation. Each `# ====` banner is a section.

### GLOBALS — every constant, in one place
**Purpose:** all module-level constants: the two workhorse regexes (`TOTAL_RE`
for total/subtotal labels, `NUMERIC_RE` for finance-style numbers incl. `(948)`
and `$1,234`), the report column shape (`REPORT_COLUMNS`), the review
priorities (`REVIEW_OK` / `REVIEW_HIGH`), the governance modes (`MODE_CLOSED` /
`MODE_REGISTRY` / `MODE_OBSERVED`), and the full vocabulary of `method` strings
(`member_undeclared`, `label_repeat_filled`, `reconciliation_mismatch`, …).
**Critical understanding:** every `method` string is a machine-readable verdict
category — the report *is* the product, and these names are its schema. Also
here: `to_number` (accounting negatives are negative — treating `(948)` as
unreadable is *wrong*, not conservative) and `_is_total_row` (a total row is
judged on the **outline column only**, never a whole-row scan — that's what stops
a department named "Total Rewards" from being deleted).

### Metadata block — the `Report:` key/value rows above the pivot
**Purpose:** `read_metadata_block` scans the rows above the header for
`Label:` / value pairs; `validate_metadata` enforces the contract.
**Critical understanding:** the block is located **by label, not by cell
address**, so it survives junk rows being added above. Every *declared* field
always becomes a `meta_*` column (null if absent) so the schema never drifts; an
**undeclared** label is a HIGH flag — that's how a typo like `Reportr:` is caught
instead of silently creating a column. The one hard name rule: the `Report:`
value must equal the sheet name exactly.

### Header-name helpers — mapping source columns to canonical levels
**Purpose:** `_norm` (cosmetic normalization), `_build_alias_index`,
`_resolve_one` (the decision: exact canonical → exact alias → fuzzy → passthrough),
and `_resolve_collisions` (two headers claiming one name → higher-confidence
wins, other reverts + HIGH).
**Critical understanding:** fuzzy matching on **headers** is applied
automatically, because a header is *one decision per file and a wrong one is
loud*. (Contrast with members, below, where fuzzy only ever *suggests*.) Two
thresholds drive it: `accept` (silent) and `review` (applied but flagged HIGH).

### Value healing — dimension MEMBERS, not headers
**Purpose:** the healing tiers. `heal_universal` (NFKC, zero-width chars,
whitespace — *always* on), `heal_licensed` (punctuation / case — only where a
level's config licenses it), `member_key` (the comparison key), and
`validate_license` (the safety proof).
**Critical understanding:** `validate_license` is the subtle one. A license is a
claim ("dashes don't matter for Service"); the function proves the claim is
**self-consistent against that level's own vocabulary** — if the healing would
fold two declared members into one, it's rejected at load. The machine doesn't
decide what's ignorable; it decides whether *your* answer is coherent.

### Configuration — the SME-owned declaration
**Purpose:** `load_configuration` expands `configuration.json` into everything
the parser needs (aliases, ordered levels, vocabularies, licenses, modes,
parents, thresholds, reconciliation settings) and **validates the spec against
itself** before any file is opened.
**Critical understanding:** a contradictory declaration fails at *startup*, not
halfway through a batch (e.g. a `closed` level with no vocabulary, or a license
that merges two members). This is where "the config is the artifact" is enforced.

### Blank dimension cells — suppressed repeat vs. genuinely missing
**Purpose:** `resolve_blank_dimensions` decides what a blank dimension cell
*means* and either fills it or flags it.
**Critical understanding:** Excel suppresses a repeated label **only when the
whole prefix above it is unchanged**, so suppression blanks always form a
**left-contiguous prefix**. `(∅, N, C)` is a suppressed repeat → fill; `(A, ∅, C)`
can't be → it's a real hole → HIGH, never filled. And the fill comes from the
last label **at the same indent level**, *not* the cell above — on an indented
sheet the cell above is a *child*, and a naive fill would stamp `"A subtype 2"`
onto A's siblings while every sum still tied. That is the silent-wrong this
module exists to prevent.

### The member registry (`MODE_REGISTRY`)
**Purpose:** `load_registry` / `save_registry` / `approve_members` /
`pending_members` / `_observe_member` — the machinery for levels whose
membership *builds itself* but still needs a human signature.
**Critical understanding:** `configuration.json` is **policy** (small, stable,
human-owned); `members.json` is **the register** (large, churning,
machine-appended, human-approved, never hand-edited). Approval is **bulk** by
design — the legitimate path through the gate must be faster than the path
around it, or someone widens the vocabulary to a wildcard and the control dies.

### Member-identity gate — THE KEY GATE
**Purpose:** `validate_members` — heals each member, then checks it against the
level's approved set, dispatching on the level's mode (`closed` = hand-written
vocabulary; `registry` = approved register, new members block as pending;
`observed` = no membership check, near-duplicate + cardinality-drift only).
**Critical understanding:** this is the gate that catches what arithmetic can't —
`"A "` vs `"A"`, a lowercase `"b"`, a new department. It **writes the canonical
value back** to the frame but keeps the raw value in the report, so a heal is
auditable but never invisible. Crucially, fuzzy here only **suggests** (names the
closest member in the note) — it never renames, because a member is *thousands of
silent decisions per file* and auto-accepting one would manufacture the exact
orphan the gate exists to catch.

### Dimension-completeness gate
**Purpose:** `validate_dimension_completeness` — any null in a declared level is
HIGH, regardless of how it got there.
**Critical understanding:** a null key is worse than a wrong number — it silently
drops out of every `GROUP BY` downstream and nobody sees it go. Five lines that
make that entire class impossible to pass.

### Sheet-shape detection
**Purpose:** `_looks_numeric` (how numeric is a column?) and `_find_header_row`
(locate the real header by finding the first numeric data run and taking the
label row just above it).
**Critical understanding:** this is the most *fragile* part by nature — it
assumes the header row is **non-numeric**. Bare-integer period headers (e.g.
`2025`) make the header row look like data and defeat detection; text labels like
`FY25` or `Sum of FY26` keep it working. Same routine also splits value vs.
dimension columns by numeric fraction, probing on non-total rows so a subtotal
can't skew the vote.

### Main extraction + batch ingest
**Purpose:** `extract_pivot` (the flat parser) and `ingest` (the batch driver).
**Critical understanding — two things.** (1) **Order matters inside
`extract_pivot`:** heal/canonicalize members *first*, then resolve blanks (so the
repeat-fill copies canonical labels), then melt; and "the body ends where the
numbers end," which cuts the trailing notes block before the member gate can
mistake `*NOTES…` for a member. (2) `ingest` is **fail-closed by default**
(`gate=True`) — a safety property you must remember to invoke is not a safety
property. It also stamps provenance (`source_sha256`, `ingested_at`, `row_key`)
so a re-run **merges onto itself** instead of doubling the facts, and never lets
one bad file stop the batch (that file becomes a `file_error` HIGH row).

### Hierarchical (indented) pivot parsing — the second parse "class"
**Purpose:** `extract_hierarchical_pivot`, `_read_sheet_with_indent`,
`_indent_levels`, `_hierarchy_lineage` — parse an *indented* outline (parent
subtotal rows with children nested beneath) into tidy rows tagged with
`sub_level` / `is_parent` / `is_leaf` / `parent_label` / `path`.
**Critical understanding:** **leading spaces in the outline column ARE the
hierarchy**, so indent must be extracted *before* any `strip()`/heal touches the
labels — otherwise every `sub_level` flattens to 0 and parents get double-counted.
Parents are subtotals of their children; both live in the output, but **only
`is_leaf` rows are summed** for reconciliation.

### Auto-dispatch — one entry point that routes
**Purpose:** `parse` (the front door), plus `_sheet_is_hierarchical` (the router
sniff), `_harmonize` (make flat and hierarchical outputs share one schema),
`build_hierarchy` (expand a business-hierarchy spec), `_apply_hierarchy` (the
**structural gate** — order levels + enforce presence), `hierarchy_audit`, and
`to_level_columns`.
**Critical understanding:** `parse` picks the parser, then runs the post-parse
gates inside its nested `_finish` — metadata, structural (`_apply_hierarchy`),
dimensional, and arithmetic, in that order, on the *final* frame. `_harmonize`
is why a mixed batch stacks cleanly: flat rows get the hierarchy columns as
level-0 leaves so every output has one identical shape.

### Reconciliation — THE ARITHMETIC GATE
**Purpose:** `read_total_row` (extract the source's own total row *before* it's
discarded) and `total_validation` (assert leaf sums tie to it, per period).
**Critical understanding:** the pivot's total row is the source's **self-declared
answer**; this is the check that makes silent-wrong *loud*. Sum **only leaves**
(parents are subtotals — including them double-counts). The total row is anchored
to the outline column, so a stray "total" in a notes block can't become the
baseline. No total row at all is itself reportable — without it the gate can't run.

---

## 4. The two entry points, and reading the output

```python
import pivot_extract as px

# One file:
tidy, report = px.parse("Q3.xlsx", config="configuration.json", return_report=True)

# A batch (production): fail-closed — raises "Load blocked" on any HIGH row.
data, report = px.ingest(["Q3.xlsx", "Q4.xlsx"], config="configuration.json")

# Inspect a failing batch instead of raising:
data, report = px.ingest(paths, config="configuration.json", gate=False)
print(px.review_items(report)[["source_file", "method", "note"]])
```

**Output columns:** `source_file, source_sha256, ingested_at`, the declared
hierarchy levels, `nulled_levels, sub_level, is_parent, is_leaf, parent_label,
path`, the period column, the value column, `row_key`, and any `meta_*` fields.

**Report priorities:** `OK` = informational (heals, fills, bootstrapped members —
auditable, non-blocking); `OVERRIDDEN` = a declared exception someone approved;
`HIGH` = blocks the load.

---

## 5. Where to change things

| You want to… | Change… | Not… |
|---|---|---|
| Add a department / service / member | `configuration.json` vocabulary (or approve in `members.json` for registry levels) | the code |
| Accept a new source column spelling | that level's `aliases` in the config | the code |
| Adjust fuzzy strictness, recon tolerance, scan depth | the `matching` / `reconciliation` / `parsing` blocks in config | the code |
| Support a localized workbook (e.g. German totals) | `total_labels` / `aggregation_prefixes` / `blank_member_literal` in config | six regexes in code |
| Permit a report to legitimately lack a level | `overrides.allow_missing_levels` in config | the code |

If you find yourself editing the code to change a *business* rule, it belongs in
the config instead — that separation is the point.

---

## 6. Running it locally

Use `run_local.py` (included). It resolves every path with `os.path` relative to
its own location, reads `.xlsx` files from `./inputs`, runs the full pipeline +
gates, and writes the tidy data, audit report, and a `review_high.csv` to a
timestamped folder under `./outputs`. It prints the gate verdict and returns a
non-zero exit code when a batch would block (handy for scripting/CI).

```
python run_local.py                          # uses ./configuration.json and ./inputs
python run_local.py  other_config.json       # override the config
python run_local.py  cfg.json  some/inputs   # override both
```
