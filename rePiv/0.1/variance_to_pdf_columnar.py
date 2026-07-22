#!/usr/bin/env python3
"""
variance_to_pdf.py -- render the Budget-vs-Cost variance report to a canned PDF.

Reuses the analytics layer (marry_budget_cost.py) to build the stacked pivot,
slices it to the reporting window (FY25-FY32 by default), and renders a
landscape PDF: one section per DEPT, measures (Budget/Cost/Variance) stacked per
key, fiscal years across, variance coloured by surplus/deficit.

Only ONE extra dependency beyond the analytics stack: reportlab (pure Python).
    python -m pip install reportlab

RUN
    python variance_to_pdf.py
Output lands in ./outputs next to this file.
"""
import os
import glob
from datetime import datetime

import pandas as pd

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, LongTable, TableStyle,
                                Paragraph, Spacer)

import marry_budget_cost as mbc   # data prep lives in the analytics layer


# --- reporting window & layout constants ------------------------------------
FY_KEEP = [f"FY{y}" for y in range(25, 33)]        # FY25..FY32 (8 inclusive)
NUMERIC = FY_KEEP + ["TOTAL"]
KEY_COLS = ["Service", "Variant", "DEPT DESCRIPTION", "Project"]  # DEPT = section
HDR = ["Service", "Variant", "DEPT DESC", "Project", "Measure"] + FY_KEEP + ["TOTAL"]

INK = colors.HexColor("#1f2a37")
BAND = colors.HexColor("#334155")
SURPLUS = colors.HexColor("#15803d")
DEFICIT = colors.HexColor("#b91c1c")
FAINT = colors.HexColor("#d7dbe0")
ZEBRA = colors.HexColor("#f5f7fa")

COL_W = [42, 40, 52, 88, 52] + [50] * len(NUMERIC)   # 5 labels + 9 numeric


def money(v):
    """Thousands-separated integer; negatives in accounting parentheses."""
    n = int(round(0 if v is None or pd.isna(v) else v))
    return f"({abs(n):,})" if n < 0 else f"{n:,}"


def load_stacked():
    """Build the stacked (measures-as-rows) pivot and slice to the window."""
    files = sorted(glob.glob(os.path.join(mbc.STAGED_DIR, "*.xlsx")))
    cost, budg = mbc._identify({p: mbc.read_staged(p) for p in files})
    st = mbc.marry_stacked(cost, budg).copy()
    st["TOTAL"] = st[FY_KEEP].sum(axis=1)          # TOTAL over the shown years only
    return st[FY_KEEP + ["TOTAL"]]


def _footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(colors.grey)
    w, _ = landscape(letter)
    canvas.drawString(0.4 * inch, 0.28 * inch,
                      "Budget vs Cost — Variance Report")
    canvas.drawRightString(w - 0.4 * inch, 0.28 * inch, f"Page {doc.page}")
    canvas.restoreState()


def build_pdf(out_path):
    st = load_stacked()
    flat = st.reset_index()

    # headline totals over the window
    tot = st.groupby(level="Measure")["TOTAL"].sum()
    gb, gc = tot.get("Budget", 0), tot.get("Cost", 0)
    gv = tot.get("Variance", 0)
    gpct = (gv / gc * 100) if gc else 0

    styles = getSampleStyleSheet()
    title_s = ParagraphStyle("t", parent=styles["Title"], fontSize=15,
                             textColor=INK, spaceAfter=2)
    meta_s = ParagraphStyle("m", parent=styles["Normal"], fontSize=8,
                            textColor=colors.grey, leading=11)
    dept_s = ParagraphStyle("d", parent=styles["Heading2"], fontSize=11,
                            textColor=INK, spaceBefore=10, spaceAfter=3)

    story = [
        Paragraph("Budget vs Cost — Variance Report", title_s),
        Paragraph(
            f"Window <b>FY25–FY32</b> &nbsp;|&nbsp; key: "
            f"DEPT / Service / Variant / DEPT DESC / Project &nbsp;|&nbsp; "
            f"generated {datetime.now():%Y-%m-%d %H:%M}", meta_s),
        Paragraph(
            f"Totals — Budget <b>{money(gb)}</b> &nbsp; Cost <b>{money(gc)}</b> "
            f"&nbsp; Variance <b>{money(gv)}</b> ({gpct:+.1f}%)", meta_s),
        Spacer(1, 6),
    ]

    base = [
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 6.8),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 7),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 0), (-1, 0), BAND),
        ("ALIGN", (5, 0), (-1, -1), "RIGHT"),      # numeric block right-aligned
        ("ALIGN", (0, 0), (4, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 1.6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.6),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("LINEBELOW", (0, 0), (-1, 0), 0.6, INK),
    ]

    for dept, dsub in flat.groupby("DEPT", sort=False):
        data = [HDR]
        cmds = list(base)
        r = 1
        for _, block in dsub.groupby(KEY_COLS, sort=False):
            block = (block.set_index("Measure")
                          .reindex(["Budget", "Cost", "Variance"]).reset_index())
            for i, (_, row) in enumerate(block.iterrows()):
                labels = ([row["Service"], row["Variant"],
                           row["DEPT DESCRIPTION"], row["Project"]]
                          if i == 0 else ["", "", "", ""])
                data.append(labels + [row["Measure"]]
                            + [money(row[c]) for c in NUMERIC])
                if row["Measure"] == "Variance":
                    col = (SURPLUS if row["TOTAL"] > 0
                           else DEFICIT if row["TOTAL"] < 0 else colors.grey)
                    cmds.append(("TEXTCOLOR", (4, r), (-1, r), col))
                    cmds.append(("FONTNAME", (4, r), (4, r), "Helvetica-Bold"))
                r += 1
            cmds.append(("LINEBELOW", (0, r - 1), (-1, r - 1), 0.25, FAINT))

        table = LongTable(data, colWidths=COL_W, repeatRows=1)
        table.setStyle(TableStyle(cmds))
        story.append(Paragraph(f"DEPT — {dept}", dept_s))
        story.append(table)

    doc = SimpleDocTemplate(
        out_path, pagesize=landscape(letter),
        leftMargin=0.4 * inch, rightMargin=0.4 * inch,
        topMargin=0.45 * inch, bottomMargin=0.5 * inch,
        title="Budget vs Cost Variance Report",
    )
    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    return out_path, len(st)


if __name__ == "__main__":
    os.makedirs(mbc.OUTPUT_DIR, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = os.path.join(mbc.OUTPUT_DIR, f"variance_report_{stamp}.pdf")
    path, nrows = build_pdf(out)
    print("rows rendered:", nrows)
    print("written      :", path)
