# pivot_extract

Ingests Excel pivot and flat report exports into a tidy long-format DataFrame, and **refuses to load anything it cannot vouch for**.

The parsing is the easy half. The point of this module is the other half: a spreadsheet that parses cleanly and loads *wrong numbers* is the failure mode that costs you the platform's credibility, and you don't get that back. Everything here is built to make that impossible — or, where it isn't possible, to say so out loud.

---

## Contents

- [The core idea](#the-core-idea)
- [The four gates](#the-four-gates)
- [Quick start](#quick-start)
- [The declaration (`configuration.json`)](#the-declaration-configurationjson)
- [Governance modes](#governance-modes)
- [The register (`members.json`)](#the-register-membersjson)
- [Bootstrap](#bootstrap)
- [Healing vs. deciding](#healing-vs-deciding)
- [Operating it](#operating-it)
- [Known limits](#known-limits-read-this)
- [Roadmap](#roadmap)
- [Testing](#testing)

---

## The core idea

Two kinds of error arrive in a spreadsheet, and they need opposite treatments.

**Decode errors** — the cell's *form* doesn't match its meaning. `(948)` is a negative. A blank means "same as above." A numeric-coded category looks like data. These are loud: they break something visible, and the arithmetic gate catches them.

**Identity errors** — the cell decodes perfectly and is still the *wrong key*. `"A "` is not `"A"`. `"b"` is not `"B"`. A non-breaking space makes two identical-looking strings unequal. These are **silent**: every sum still ties, no nulls appear, and the rows are visually indistinguishable in any report. They quietly split a department into two, and no arithmetic check on earth can see it.

The module handles the first class in the parser and the second class in a **gate**, because they cannot be caught by the same mechanism.

> **Heal what is unambiguous. Fail on what is a decision.
> Nothing enters the warehouse as a key until it has been declared as one.**

---

## The four gates

Every gate catches a class the others are blind to. A HIGH from any of them blocks the load.

| Gate | Question | Catches |
|---|---|---|
| **Arithmetic** | Do the numbers tie? | Leaf sum ≠ the source's own total row, per period. Parents excluded (they're subtotals). |
| **Structural** | Is the shape right? | A declared hierarchy level missing or unmapped. |
| **Dimensional** | Is every row addressable? | Nulls in a declared level. A null key silently drops out of every `GROUP BY` downstream. |
| **Identity** | Is every key a declared member? | `"A "` vs `"A"`. New/undeclared members. Near-duplicates. |

`ingest()` is **fail-closed by default** (`gate=True`). A safety property you have to remember to invoke is not a safety property.

---

## Quick start

```python
import pivot_extract as px

data, report = px.ingest(
    ["Q3_costs.xlsx", "Q3_headcount.xlsx"],
    config="configuration.json",
)
```

That's the whole call site. No column names, no thresholds, no hierarchy — it all lives in the declaration. A HIGH row anywhere raises `ValueError: Load blocked`.

To inspect a failing batch instead of raising:

```python
data, report = px.ingest(paths, config="configuration.json", gate=False)
print(px.review_items(report)[["source_file", "method", "note"]])
```

**Output columns:** `source_file`, `source_sha256`, `ingested_at`, `row_key`, the declared hierarchy levels, `sub_level`, `is_parent`, `is_leaf`, `parent_label`, `path`, the period column, the value column, plus any `meta_*` fields from the report's metadata block.

---

## The declaration (`configuration.json`)

**The config is the artifact. The module is only the thing that obeys it.**

Every business and format decision lives here — hierarchy, aliases, member vocabularies, normalization licenses, metadata contract, thresholds, output column names. Point it at a different config and it ingests a different business, with no code change.

```json
{
  "hierarchy": [
    {
      "canonical": "DEPT",
      "aliases": ["Row labels", "Department"],
      "mode": "closed",
      "vocabulary": ["A", "B", "C"],
      "normalization": { "punctuation": false, "case": false }
    }
  ]
}
```

### Normalization licenses

A license is a claim: *"dashes don't matter for Service."* On load, the module **proves the claim is self-consistent** against your own vocabulary:

> A normalization is safe for a level **iff** applying it to that level's declared members never causes two of them to collide.

Declare `F-1` and `F1` as distinct members *and* license punctuation-stripping, and startup fails. The machine doesn't decide what's ignorable — **it decides whether your answer is coherent.** And because the check re-runs as the vocabulary grows, the day someone approves a genuinely distinct `F-1`, punctuation-stripping stops being licensed for that level automatically.

### US-locale assumption

`total_labels`, `aggregation_prefixes` and `blank_member_literal` are the strings a US-English Excel emits. They're in the config rather than the code, so a localized workbook is a five-minute edit rather than a hunt through six regexes.

---

## Governance modes

**Choose by cardinality.** Getting this wrong is costly in both directions.

| Mode | Use for | New member behaviour |
|---|---|---|
| **`closed`** | Small, stable, high-blast-radius levels (3–20 members; Department). Everyone sees them on a dashboard; a corruption is catastrophic. | **Blocked.** You hand-write the vocabulary. |
| **`registry`** | Granular, churning levels (hundreds/thousands; Service codes, cost centres). | **Blocked until approved — in bulk.** The list builds itself; a human signs off once. |
| **`observed`** | Too granular to gate at all (project codes, free-text). | **Accepted.** No membership check — but near-duplicate detection and cardinality drift still run. |

**Why `closed` is wrong for a granular level:** hand-maintaining thousands of members is a treadmill, and **a treadmill is not a control, because it gets abandoned.** Someone widens the vocabulary to a wildcard "just for this load," and you're back to six departments named `A`.

**Why `observed` is wrong for a small level:** it's the dimension everybody actually looks at. Gate it.

**The near-duplicate check matters *more* as cardinality rises, not less.** With three departments, a human eventually notices `b` and `B` both exist. With 5,000 project codes, **nobody will ever look** — and `PROJ-1234` / `PROJ 1234` will split every rollup that touches them, silently, forever. Automated detection is the *only* control that can work up there.

---

## The register (`members.json`)

Not shipped — **created at runtime**, only when a level runs in `registry` mode.

```json
{
  "levels": {
    "Service": {
      "members": {
        "F1": {
          "approved": true,
          "approved_by": "jane.smith",
          "approved_at": "2026-07-14T02:58:32Z",
          "first_seen":  "2026-07-01T09:12:04Z",
          "parents": { "A": 1, "B": 1, "C": 1 }
        }
      }
    }
  }
}
```

**Config is *policy*. The register is *the register*.** Small/stable/human-owned vs. large/churning/machine-appended. Conflating them is exactly what makes a controlled vocabulary feel unmaintainable.

### Rules

- **Commit it.** Track it in git like a schema migration. A PR diff then reads *"added Service F7, approved by Jane"* — a reviewable record of every dimension change. That's why it's boring flat JSON.
- **Never hand-edit it.** Append via the code, approve via `approve_members()`. The moment someone types a member in by hand, the `first_seen`/`approved_by` provenance is a lie and the register stops being evidence of anything.
- **Approval is bulk.** *"47 new Service codes this load — approve all / selected / reject."* One decision, not 47 edits.

```python
reg = px.load_registry("members.json")
px.pending_members(reg)                                    # the worklist
px.approve_members(reg, "Service", approver="jane.smith")  # one call, all of them
px.save_registry(reg, "members.json")
```

**Approval must be cheaper than the workaround.** If clearing a new member is a week-long ticket, someone routes around the gate — and the way you route around an identity gate is by lying about the member name. **The gate's real design constraint isn't detection accuracy. It's that the legitimate path through it must be faster than the path around it.**

---

## Bootstrap

On the **first** load of a `registry` level, the register is empty — so *every* member is new, and *everything* blocks. A reviewer facing 5,000 HIGH rows will either click approve-all without looking, or disable the gate. **Either way the control is bypassed on day one**, which is worse than no control, because everyone believes it's running.

```json
{ "canonical": "Service", "mode": "registry", "bootstrap": true }
```

For that one run, unknown members are registered and **auto-approved**, stamped `approved_by: "bootstrap"`. Then you turn it off and the gate is live.

### Bootstrap is a deliberate hole in the gate

For exactly one run, the identity check is off. **Whatever is in that seed file becomes truth** — including dirty members already sitting in it. If `f-1` and `F1` are both there, bootstrap approves *both*, permanently.

- **Seed from a source you trust** — ideally the source system's dimension table, not a spreadsheet.
- **It's loud on purpose.** Every bootstrapped member logs a `member_registered` note, and `approved_by: "bootstrap"` persists forever — so you can always ask *"which members did a human actually review, and which just showed up on day one?"* That query is the audit trail.
- **The failure mode is leaving it on.** Permanent bootstrap = `observed` mode with extra bookkeeping: a gate that reports instead of blocking. The config validator catches nonsensical uses (bootstrap on a `closed` level raises), but it **cannot** catch "you forgot," because that's indistinguishable from "you meant it." See [Roadmap](#roadmap).

---

## Healing vs. deciding

**Healing targets how text was *encoded*. Never what it *says*.**

| Always on (encoding noise) | Never merges two members under any vocabulary |
|---|---|
| NFKC normalization | non-breaking space, full-width chars |
| Zero-width char removal | ZWSP/ZWNJ/BOM — invisible, and they make `"A" != "A"` |
| Whitespace strip + collapse | trailing spaces, wrapped-cell line breaks |

| Licensed only (conditionally safe) | Verified against the vocabulary first |
|---|---|
| Punctuation stripping | `F-1` → `F1` — fine *iff* nothing collides |
| Case folding | `b` → `B` — fine *iff* nothing collides |

| Never | Why |
|---|---|
| Leading zeros | `01` ≠ `1` in coded dimensions |
| Diacritics | `José` ≠ `Jose` — different people |
| Internal spaces | destruction, not cleanup |

**Fuzzy matching *suggests*, it never *applies*.** A header is one decision per file and a wrong one is loud. A **member** is thousands of decisions per file and a wrong one is silent. Auto-accepting a fuzzy member match would manufacture the exact orphan the gate exists to catch.

```
'b' → member_undeclared  HIGH  closest declared member is 'B' (1.0) — approve or correct
```

The reviewer gets the answer handed to them. The machine doesn't make the call.

### Ordering constraint (do not refactor)

```
read cells → extract indent → heal → fill blanks → build lineage → melt
```

**Leading spaces in the outline column ARE the hierarchy.** If healing runs before indent extraction, `strip()` flattens every `sub_level` to 0, `is_leaf` becomes true everywhere, and parents get counted alongside their children. Commented in-place with the reason.

### Blank cells: suppressed repeat vs. missing value

Excel suppresses a repeated label **only when the entire prefix above it is unchanged** — if DEPT changes, Service always prints. So blanks from suppression always form a **left-contiguous prefix**.

- `(∅, N, C)` → suppressed repeat → **fill from the last label at the same indent level.** (Not the cell above — on an indented sheet that's a *child*, and a naive `ffill` writes `"A subtype 2"` into A's siblings while every sum still ties.)
- `(A, ∅, C)` → cannot arise from suppression → **genuinely missing → HIGH, never filled.**

No SME declaration needed, no pivot-vs-flat sniffing. The blank tells you what it is by where it sits.

---

## Operating it

### Idempotent reloads

Every row carries `source_sha256` (content hash), `ingested_at`, and a deterministic `row_key`.

**This is the silent-wrong that *defeats* the gates rather than tripping them.** Re-run a batch after fixing one file and, if the ETL does `INSERT` rather than `MERGE`, the other 29 land twice. Every fact doubles. Each file reconciled perfectly *on its own*, so nothing complains.

- **MERGE on `row_key`.** Never `INSERT`.
- Same bytes → same hash → same keys → a reload merges onto itself.
- A resubmitted file with the same name but different content is **visibly** a different file.

### Reading the report

| Priority | Meaning |
|---|---|
| `OK` | Informational. Heals, fills, bootstrapped members. **Auditable, non-blocking.** |
| `OVERRIDDEN` | A declared exception (`allow_missing_levels`). Someone approved this. |
| `HIGH` | **Load blocked.** |

A batch **always completes** — one bad file is reported, never fatal. You get thirty results and a list of the four to look at.

### S3 / distributed deployment

`load_registry()` and `save_registry()` are the **only two I/O points**, and `ingest()` reads once and writes once per batch. A parse never mutates governance state as a side effect — that was deliberate, so storage can move.

**But the register is shared mutable state**, and it uses last-write-wins. Concurrent workers will clobber each other: worker A registers `F7`, worker B registers `F8`, B's write wins, `F7` vanishes. Worse, an *approval* can get clobbered.

Options, cheapest first: **serialize batches** (fine for nightly loads); **S3 conditional writes** (`If-Match` on ETag, with retry); or **move the register to DynamoDB/Postgres**, which is what it actually is — a keyed store of `(level, member) → {approved, first_seen, parents}`.

Note the register is saved **before** the gate raises: a blocked batch has still learned what the new members are, and that list is exactly what the reviewer needs to unblock it.

---

## Known limits (read this)

**1. Misattribution is invisible.** A row assigned to the wrong *but valid* department loads cleanly. Sums tie, keys are declared, no nulls. **All four gates pass.** This is inherent to reconciling against a single grand total — the file is internally consistent with a wrong answer.

Do **not** treat "the analyst will spot it downstream" as a control. That only works if someone is looking, knows what to expect, and the error is big enough to see. **The errors most likely to be caught downstream are the ones that mattered least; the ones that survive are the ones that hid in the noise.** Write the limit into your acceptance criteria in the SME's language, so it's a governance decision someone signed rather than a gap that got quietly reassigned.

The real fix is a **second independent statement of the numbers** — a per-DEPT control total from the source system, not just a grand total. Then the arithmetic gate simply runs at a finer grain.

**2. Compact-form pivots are rejected, not parsed.** All levels collapsed into one indented "Row Labels" column. Detected and blocked; supporting it is new work.

**3. The vocabulary must be owned by an SME.** The current one was inferred from sample data. Until a human owns it, the identity gate is enforcing a fiction with great rigour.

**4. Recon tolerance is absolute (0.01), not relative.** At billions-scale, float summation error alone can exceed a penny. Expect false failures, and expect someone to "temporarily" loosen it to 100.

---

## Roadmap

Ordered by value.

| | Item | Why |
|---|---|---|
| 1 | **Per-DEPT control totals** from the source system | The only thing that closes the misattribution blind spot. Everything else is mitigation. |
| 2 | **Relative recon tolerance** | Absolute 0.01 will produce false failures at scale, and false failures kill trust in the gate. |
| 3 | **Single-use bootstrap** | Record `bootstrapped_at` per level and refuse to bootstrap twice. Five minutes now; a mystery in eight months. |
| 4 | **Register → DynamoDB/Postgres** | Before any concurrent execution. It's a database table pretending to be a file. |
| 5 | **Parent-change flag, once history exists** | Already captured (`parents` per member). A Service that has lived under DEPT B for six months and arrives under DEPT A is the **first signal we have that reaches into the misattribution blind spot.** The file can never tell you this. History can. |
| 6 | **Period-over-period variance report** | Doesn't prove correctness, but makes the invisible visible — turns "someone might notice" into "something surfaces it." |
| 7 | **Multi-sheet workbooks** | `sheet` is a single index today. |
| 8 | **Localized Excel** | Strings are already in config; nothing to build unless it happens. |

---

## Testing

```bash
pytest -v test_pivot_extract.py     # 32 tests
python regression.py                # human-readable pass/block matrix, 14 fixtures
python demo_modes.py                # registry lifecycle end to end
```

Fixtures cover the four SME originals plus ten mutations: real Excel indent, blanked repeat labels, accounting negatives, a department named "Total Rewards", text-stored numbers, compact form, numeric-coded dimensions, subtotal rows, deep nesting with `(blank)`, and dirty member values.

**Tests assert the *gate verdict*, not that parsing "worked."** A file that parses beautifully and loads wrong numbers is the failure this module exists to prevent, so the only assertion worth making is *did it block?*

### The one test that earns its keep

`test_dirty_member_values_are_blocked` deliberately confirms the other three gates are **satisfied** — sums tie *exactly*, zero nulls — and demands a block anyway:

```python
assert leaf_sum(tidy) == SOURCE_TOTAL_FY26   # arithmetic still ties!
assert tidy.DEPT.isna().sum() == 0           # no nulls either!
assert px.MEMBER_UNDECLARED in methods(high) # ...and yet: BLOCKED
```

**That file used to pass.** Three departments in, six out. If it ever passes again, the identity gate has regressed and the warehouse is taking corrupted keys.

### When a new file breaks something

Add it as a fixture, assert the verdict you want, then make it pass. The suite becomes the accumulated memory of every way a file has tried to lie to you.

---

## Files

| File | |
|---|---|
| `pivot_extract.py` | The module. |
| `configuration.json` | **The declaration.** SME-owned, versioned. All business + format decisions. |
| `members.json` | The register. Runtime-created. Commit it; never hand-edit it. |
| `INPUT_REQUIREMENTS.md` | Layman-readable submission guide for SMEs. |
| `test_pivot_extract.py` | 32 tests. |
| `regression.py` | Pass/block matrix across all fixtures. |
| `demo_modes.py` | Registry lifecycle demo. |
