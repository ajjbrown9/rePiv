Pre-populated with the two ingested workbooks from the tested run, so the
analytics (marry_budget_cost.py) and reporting (variance_to_pdf.py) stages can
be run immediately without ingesting first.

On a real run, replace these with the fresh outputs of Stage 1:
    cp outputs/ingest_*/*_ingested.xlsx staged/

Exactly one Cost table and one Budget table must be present. The analytics layer
identifies them by the meta_report_type column, not by filename.
