# How to Send Us Your Report File

A short guide for anyone submitting a report for automated loading.

**The good news first:** you cannot break our numbers. If your file doesn't meet these rules, the system **rejects it and tells us why** — it never quietly loads something wrong. So if you're unsure, send it. The worst case is we come back and ask you to fix one thing.

---

## The short version

1. **Excel `.xlsx` file.** Not `.xls`, not CSV, not Google Sheets.
2. **One report per file**, on the **first tab**.
3. The tab name and the **`Report:`** cell must be **spelled exactly the same**.
4. Your table needs a **Total row at the bottom**. We use it to check our own maths.
5. **Don't invent new categories** without telling us first (see the last section).

---

## What your file needs to look like

```
  A                          B                C          D           E
1  Random Header Info         DROPDOWN         Helper text
2
3  Report:                    Q3-COST-REPORT              <- must match the tab name
4  Reporter:                  Smith, Jane
5  Report Type:               Cost
6  Report Team:               Production
7  Report Date:               2026-07-13
8
9  Row labels                 Division         Type       FY26        FY27   <- header row
10 A                          MC               B          2,822       8,719
11     A subtype 1            MC               B          1,411       4,359
12     A subtype 2            MC               B          1,411       4,360
13 B                          N                C          3,670       5,730
14 total                                                  86,273      89,843  <- Total row
15
16 *NOTES: any comments you like down here (but see the warning below)
```

### The information block (rows 3–7 above)

- Each line is a **label ending in a colon**, with its **value in the cell immediately to the right**.
- **`Report:` is required.** Everything else is optional but helpful.
- It must sit **within 12 rows above your table's header row**.
- **`Report:` must exactly match the tab name** — same spelling, same capitalisation. And it must be **31 characters or fewer** (that's an Excel limit on tab names, not ours).
- **Don't add your own labels here.** Any *other* line ending in a colon above the table gets flagged as unexpected. That's deliberate — it's how we catch a typo like `Reportr:` instead of silently ignoring it.

### The table

- **One header row.** Category columns first, then one column per period.
- **Each category gets its own column** — Department in one, Service in another, Variant in another.
- **Sub-levels are indented in the first column** (see rows 11–12 above). Real Excel indenting or just leading spaces — either works, any depth.
- **A parent row's value should equal the sum of its children.**

### The Total row

- **Required.** Without it we can't verify our numbers, and the file is rejected.
- Put the word **`Total`** (or `Grand Total`) in the **first column**, and **leave the other category columns blank**.
- If you have sub-totals per department *and* a grand total, that's fine — put the **grand total last**.

---

## Things we handle for you (don't worry about these)

✅ **Pivot table or plain table** — either is fine.
✅ **Repeated labels turned on or off.** If Excel blanks out repeating department names, we fill them back in correctly.
✅ **Numbers stored as text** — `"8,057"` is fine.
✅ **Negatives in brackets** — `(948)` reads as minus 948, as you'd expect.
✅ **Dollar signs and commas** — `$1,234.50` is fine.
✅ **Extra blank rows** inside or around the table.
✅ **Trailing spaces or odd invisible characters** in your category names — we clean those silently.
✅ **Notes and comments** below the table — as long as they contain no numbers (see below).

---

## Things that will get your file rejected

| ❌ Don't | Why |
|---|---|
| **Put numbers in the notes area below the table** | We'd read them as data. Keep the area under your table text-only. |
| **Collapse all categories into one column** | Some pivot layouts stack Department, Service and Variant into a single indented "Row Labels" column. We can't read that. Use **Tabular layout** — one column per category. |
| **Rename the tab without updating the `Report:` cell** (or vice versa) | They must match exactly. |
| **Leave a category blank in the middle** | If Department is filled in but Service is empty, we won't guess. Fill it in, or tell us that category genuinely doesn't apply. |
| **Use number codes for a category** | If Service is `10, 20, 30` instead of `MC, N, F`, we'll mistake it for data. Use the names. |
| **Change the capitalisation of a category** | `b` is not `B`. We won't assume — we'll ask. |
| **Name a category exactly "Total" or "Subtotal"** | We'd mistake the row for a totals row. |
| **Send a file with uncalculated formulas** | Open it in Excel and save it before sending, so the numbers are actually stored in the cells. |

---

## Before you add a new department, service or variant

**Tell us first.** This is the one that surprises people.

We keep an approved list of every valid Department, Service and Variant. If a value shows up that isn't on the list, **the file is rejected** — even if it's a perfectly legitimate new department.

That's on purpose. A brand-new category is a **business change**, not a spreadsheet detail, and it should be a decision someone signed off on rather than something that quietly appears in a report one Tuesday. The alternative is a warehouse where nobody can say what the valid department list actually is.

**It takes us about five minutes to add one.** Just email us before you send the file, and we'll have it on the list.

---

## One thing we genuinely cannot check

We verify that your numbers **add up correctly**, that every row is **properly categorised**, and that every category is one we **recognise**.

We **cannot** tell whether a cost was put under the **wrong (but valid) department**. If a figure belonging to Department B is recorded under Department A, everything still adds up perfectly and the file loads cleanly.

**Accuracy of what goes where remains with you.** That's the one thing the system can't do for you — please give the categorisation a once-over before sending.

---

## Quick checklist before you hit send

- [ ] Saved as `.xlsx`, report on the **first tab**
- [ ] **`Report:`** cell matches the **tab name** exactly (31 characters or fewer)
- [ ] One header row, **each category in its own column**
- [ ] **Total row at the bottom**, labelled in the first column, other category columns blank
- [ ] **No numbers** in the notes area below the table
- [ ] **No new** departments/services/variants that you haven't cleared with us
- [ ] Opened and saved in Excel (so formulas have real values)

---

*Questions, or a new category to add? Contact the data platform team.*
