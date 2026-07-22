#!/usr/bin/env python3
"""
variance_to_pdf.py -- canned Budget-vs-Cost variance PDF, INDENTED-OUTLINE form.

Collapses the category columns (Service / Variant / DEPT DESC / Project) into a
single indented label column -- the classic compact pivot outline -- so four
label columns become one, freeing horizontal space. DEPT stays a section
heading; Budget/Cost/Variance sit indented deepest under each Project.

One extra dependency beyond the analytics stack: reportlab (pure Python).
    python -m pip install reportlab

RUN
    python variance_to_pdf.py           # output -> ./outputs
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


# --- reporting window & layout ----------------------------------------------
FY_KEEP = [f"FY{y}" for y in range(25, 33)]        # FY25..FY32 (8 inclusive)
NUMERIC = FY_KEEP + ["TOTAL"]
KEY_COLS = ["Service", "Variant", "DEPT DESCRIPTION", "Project"]  # DEPT = section
HDR = ["Category"] + FY_KEEP + ["TOTAL"]

INK = colors.HexColor("#1f2a37")
BAND = colors.HexColor("#334155")
SURPLUS = colors.HexColor("#15803d")
DEFICIT = colors.HexColor("#b91c1c")
FAINT = colors.HexColor("#e2e6ea")
LVL_BG = colors.HexColor("#eef1f5")               # tint for the top category level

LABEL_W = 236
NUM_W = 55
COL_W = [LABEL_W] + [NUM_W] * len(NUMERIC)        # 1 label + 9 numeric
INDENT0 = 3                                        # base left padding (pt)
STEP = 11                                          # indent per hierarchy level


def money(v):
    n = int(round(0 if v is None or pd.isna(v) else v))
    return f"({abs(n):,})" if n < 0 else f"{n:,}"


def load_stacked():
    files = sorted(glob.glob(os.path.join(mbc.STAGED_DIR, "*.xlsx")))
    cost, budg = mbc._identify({p: mbc.read_staged(p) for p in files})
    st = mbc.marry_stacked(cost, budg).copy()
    st["TOTAL"] = st[FY_KEEP].sum(axis=1)          # TOTAL over the shown window
    return st[FY_KEEP + ["TOTAL"]]


def _footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(colors.grey)
    w, _ = landscape(letter)
    canvas.drawString(0.4 * inch, 0.28 * inch, "Budget vs Cost - Variance Report")
    canvas.drawRightString(w - 0.4 * inch, 0.28 * inch, f"Page {doc.page}")
    canvas.restoreState()


def build_pdf(out_path):
    st = load_stacked()
    flat = st.reset_index()

    tot = st.groupby(level="Measure")["TOTAL"].sum()
    gb, gc, gv = tot.get("Budget", 0), tot.get("Cost", 0), tot.get("Variance", 0)
    gpct = (gv / gc * 100) if gc else 0

    styles = getSampleStyleSheet()
    title_s = ParagraphStyle("t", parent=styles["Title"], fontSize=15,
                             textColor=INK, spaceAfter=2)
    meta_s = ParagraphStyle("m", parent=styles["Normal"], fontSize=8,
                            textColor=colors.grey, leading=11)
    dept_s = ParagraphStyle("d", parent=styles["Heading2"], fontSize=11,
                            textColor=INK, spaceBefore=10, spaceAfter=3)

    story = [
        Paragraph("Budget vs Cost - Variance Report", title_s),
        Paragraph("Window <b>FY25-FY32</b> &nbsp;|&nbsp; outline: DEPT &gt; Service "
                  "&gt; Variant &gt; DEPT DESC &gt; Project &gt; Measure "
                  "&nbsp;|&nbsp; generated " + f"{datetime.now():%Y-%m-%d %H:%M}",
                  meta_s),
        Paragraph(f"Totals - Budget <b>{money(gb)}</b> &nbsp; Cost <b>{money(gc)}</b>"
                  f" &nbsp; Variance <b>{money(gv)}</b> ({gpct:+.1f}%)", meta_s),
        Spacer(1, 6),
    ]

    base = [
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 0), (-1, 0), BAND),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
        ("ALIGN", (0, 0), (0, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 1.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("LINEBELOW", (0, 0), (-1, 0), 0.6, INK),
    ]

    for dept, dsub in flat.groupby("DEPT", sort=False):
        data = [HDR]
        cmds = list(base)
        r = 1
        prev = [None, None, None, None]
        for _, block in dsub.groupby(KEY_COLS, sort=False):
            key = [block.iloc[0][c] for c in KEY_COLS]
            start = 0
            while start < 4 and key[start] == prev[start]:
                start += 1
            for lvl in range(start, 4):
                data.append([str(key[lvl])] + [""] * len(NUMERIC))
                cmds.append(("LEFTPADDING", (0, r), (0, r), INDENT0 + lvl * STEP))
                cmds.append(("FONTNAME", (0, r), (0, r), "Helvetica-Bold"))
                cmds.append(("TEXTCOLOR", (0, r), (0, r), INK))
                if lvl == 0:
                    cmds.append(("BACKGROUND", (0, r), (-1, r), LVL_BG))
                r += 1
            prev = key

            block = (block.set_index("Measure")
                          .reindex(["Budget", "Cost", "Variance"]).reset_index())
            for _, row in block.iterrows():
                data.append([row["Measure"]] + [money(row[c]) for c in NUMERIC])
                cmds.append(("LEFTPADDING", (0, r), (0, r), INDENT0 + 4 * STEP))
                if row["Measure"] == "Variance":
                    col = (SURPLUS if row["TOTAL"] > 0
                           else DEFICIT if row["TOTAL"] < 0 else colors.grey)
                    cmds.append(("TEXTCOLOR", (0, r), (-1, r), col))
                    cmds.append(("FONTNAME", (0, r), (0, r), "Helvetica-Bold"))
                r += 1
            cmds.append(("LINEBELOW", (0, r - 1), (-1, r - 1), 0.25, FAINT))

        table = LongTable(data, colWidths=COL_W, repeatRows=1)
        table.setStyle(TableStyle(cmds))
        story.append(Paragraph(f"DEPT - {dept}", dept_s))
        story.append(table)

    doc = SimpleDocTemplate(
        out_path, pagesize=landscape(letter),
        leftMargin=0.4 * inch, rightMargin=0.4 * inch,
        topMargin=0.45 * inch, bottomMargin=0.5 * inch,
        title="Budget vs Cost Variance Report (indented)",
    )
    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    return out_path


if __name__ == "__main__":
    os.makedirs(mbc.OUTPUT_DIR, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = os.path.join(mbc.OUTPUT_DIR, f"variance_report_indented_{stamp}.pdf")
    print("written:", build_pdf(out))
