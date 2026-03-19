"""
Google Sheet → PDF Management Report
Reads project data, analyzes keyword overlap, generates charts + network graph.

Requirements:
    pip install pandas matplotlib networkx reportlab Pillow

Usage:
    python gsheet_to_pdf_report.py
"""

import pandas as pd
import json
import os
import platform
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.font_manager as fm
import networkx as nx
import numpy as np
from collections import Counter
from itertools import combinations
from io import BytesIO
from PIL import Image as PILImage

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_RIGHT

import re

# ═══════════════════════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════════════════════
SHEET_ID = "131Vl8q-bQFqMi7CAnLwy-rkLGBswKmgx9o_YOUSQ7zc"
GID = "0"
OUTPUT_PDF = "project_report.pdf"
OUTPUT_JSON = "projects_full.json"

# ═══════════════════════════════════════════════════════════════
# FONT SETUP — auto-detect by OS
# ═══════════════════════════════════════════════════════════════
if platform.system() == "Windows":
    _FDIR = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    FONT = os.path.join(_FDIR, "arial.ttf")
    FONT_BOLD = os.path.join(_FDIR, "arialbd.ttf")
    FF = "Arial"
elif platform.system() == "Darwin":
    FONT = "/System/Library/Fonts/Supplemental/Arial.ttf"
    FONT_BOLD = "/System/Library/Fonts/Supplemental/Arial Bold.ttf"
    FF = "Arial"
else:
    FONT = "/usr/share/fonts/truetype/freefont/FreeSans.ttf"
    FONT_BOLD = "/usr/share/fonts/truetype/freefont/FreeSansBold.ttf"
    FF = "FreeSans"

pdfmetrics.registerFont(TTFont("Heb", FONT))
pdfmetrics.registerFont(TTFont("HebB", FONT_BOLD))
fm.fontManager.addfont(FONT)
fm.fontManager.addfont(FONT_BOLD)
plt.rcParams['font.family'] = FF

C_PRIMARY = "#1B4F72"
C_ACCENT = "#2E86C1"
C_LIGHT = "#D6EAF8"
C_GREEN = "#27AE60"
C_ORANGE = "#E67E22"
C_RED = "#E74C3C"
C_GRAY = "#95A5A6"
C_PURPLE = "#8E44AD"
C_YELLOW = "#F1C40F"

STATUS_COLORS = {
    "בפרודקשן": C_GREEN,
    "הושלם": C_ACCENT,
    "מושהה": C_ORANGE,
    "גרסה ישנה": C_GRAY,
}


def heb(text):
    """Reverse Hebrew text for correct RTL display in matplotlib/reportlab."""
    s = str(text)
    parts = re.split(r'([A-Za-z0-9_.:/\-]+)', s)
    reversed_parts = []
    for p in parts:
        if re.match(r'^[A-Za-z0-9_.:/\-]+$', p):
            reversed_parts.append(p)
        else:
            reversed_parts.append(p[::-1])
    return ''.join(reversed_parts[::-1])


def fig_to_image(fig, width=170 * mm):
    buf = BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    buf.seek(0)
    w, h = PILImage.open(buf).size
    buf.seek(0)
    return Image(buf, width=width, height=width * (h / w))


# ═══════════════════════════════════════════════════════════════
# 1) READ DATA
# ═══════════════════════════════════════════════════════════════
def read_data():
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"
    df = pd.read_csv(url, dtype=str).fillna("")
    print(f"✔ Read {len(df)} rows from Google Sheet")

    projects = []
    for _, row in df.iterrows():
        projects.append({
            "id": row["מספר סידורי"],
            "name": row["שם פרוייקט"],
            "responsible": row["אחראי פרוייקט"],
            "developer": row["מפתח"],
            "developer_b": row["מפתח ב"],
            "manager": row["מנהל פרוייקט"],
            "manager_b": row["מנהל פרוייקט ב"],
            "start": row["תאריך התחלה"],
            "end": row["תאריך סיום"],
            "status": row["סטטוס"],
            "clients": row["לקוחות"],
            "description": row["תיאור"],
            "keywords": [k.strip().lower() for k in row["מילות מפתח"].split(",") if k.strip()],
        })
    return projects


# ═══════════════════════════════════════════════════════════════
# 2) BUILD SIMILARITY
# ═══════════════════════════════════════════════════════════════
def build_similarity(projects):
    for p in projects:
        similar = []
        for o in projects:
            if o["id"] == p["id"] and o["name"] == p["name"]:
                continue
            shared = set(p["keywords"]) & set(o["keywords"])
            if shared:
                similar.append({
                    "project_id": o["id"],
                    "project_name": o["name"],
                    "shared_keywords": sorted(shared),
                    "score": len(shared),
                })
        similar.sort(key=lambda x: x["score"], reverse=True)
        p["similar_projects"] = similar
        p["total_connections"] = len(similar)
        p["max_similarity_score"] = similar[0]["score"] if similar else 0
    return projects


# ═══════════════════════════════════════════════════════════════
# 3) SAVE JSON
# ═══════════════════════════════════════════════════════════════
def save_json(projects):
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(projects, f, ensure_ascii=False, indent=2)
    print(f"✔ JSON saved: {OUTPUT_JSON}")


# ═══════════════════════════════════════════════════════════════
# 4) GENERATE CHARTS
# ═══════════════════════════════════════════════════════════════
def chart_status_pie(projects):
    counts = Counter(p["status"] for p in projects)
    fig, ax = plt.subplots(figsize=(6, 4))
    labels = [heb(k) for k in counts.keys()]
    clrs = [STATUS_COLORS.get(k, C_GRAY) for k in counts.keys()]
    _, _, autotexts = ax.pie(list(counts.values()), labels=labels, colors=clrs,
                             autopct='%1.0f%%', startangle=90,
                             textprops={'fontsize': 11, 'fontfamily': FF})
    for t in autotexts:
        t.set_fontsize(10); t.set_color('white'); t.set_fontweight('bold')
    ax.set_title(heb("התפלגות סטטוס פרוייקטים"), fontsize=14, fontweight='bold', fontfamily=FF)
    fig.tight_layout()
    return fig


def chart_manager_bar(projects):
    counts = Counter(p["manager"] for p in projects)
    fig, ax = plt.subplots(figsize=(7, 3.5))
    bars = ax.barh([heb(m) for m in counts.keys()], list(counts.values()),
                   color=[C_ACCENT, C_GREEN, C_ORANGE, C_PURPLE][:len(counts)])
    ax.set_xlabel(heb("מספר פרוייקטים"), fontsize=11, fontfamily=FF)
    ax.set_title(heb("עומס לפי מנהל פרוייקט"), fontsize=14, fontweight='bold', fontfamily=FF)
    for bar, c in zip(bars, counts.values()):
        ax.text(bar.get_width() + 0.1, bar.get_y() + bar.get_height() / 2, str(c), va='center', fontsize=11)
    fig.tight_layout()
    return fig


def chart_keywords(projects):
    all_kw = [k for p in projects for k in p["keywords"]]
    freq = Counter(all_kw).most_common(12)
    fig, ax = plt.subplots(figsize=(7, 4))
    colors_list = [C_PRIMARY, C_ACCENT, C_GREEN, C_ORANGE, C_PURPLE, C_RED,
                   C_YELLOW, C_GRAY, C_PRIMARY, C_ACCENT, C_GREEN, C_ORANGE]
    bars = ax.bar(range(len(freq)), [v for _, v in freq], color=colors_list[:len(freq)])
    ax.set_xticks(range(len(freq)))
    ax.set_xticklabels([heb(k) for k, _ in freq], fontsize=9, fontfamily=FF, rotation=30, ha='right')
    ax.set_ylabel(heb("מספר פרוייקטים"), fontsize=11, fontfamily=FF)
    ax.set_title(heb("מילות מפתח נפוצות"), fontsize=14, fontweight='bold', fontfamily=FF)
    for bar, (_, v) in zip(bars, freq):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.1, str(v), ha='center', fontsize=10)
    fig.tight_layout()
    return fig


def chart_clients(projects):
    counts = Counter(p["clients"] for p in projects)
    fig, ax = plt.subplots(figsize=(6, 3.5))
    clrs = [C_ACCENT, C_GREEN, C_ORANGE, C_PURPLE, C_RED, C_YELLOW][:len(counts)]
    ax.barh([heb(k) for k in counts.keys()], list(counts.values()), color=clrs)
    ax.set_xlabel(heb("מספר פרוייקטים"), fontsize=11, fontfamily=FF)
    ax.set_title(heb("פרוייקטים לפי לקוח"), fontsize=14, fontweight='bold', fontfamily=FF)
    for i, v in enumerate(counts.values()):
        ax.text(v + 0.1, i, str(v), va='center', fontsize=11)
    fig.tight_layout()
    return fig


def chart_network(projects):
    G = nx.Graph()
    for p in projects:
        G.add_node(p["id"], label=p["name"], status=p["status"])
    for a, b in combinations(projects, 2):
        shared = set(a["keywords"]) & set(b["keywords"])
        if shared:
            G.add_edge(a["id"], b["id"], weight=len(shared))

    fig, ax = plt.subplots(figsize=(10, 8))
    pos = nx.spring_layout(G, k=2.5, iterations=80, seed=42)

    edges = G.edges(data=True)
    weights = [d["weight"] for _, _, d in edges]
    max_w = max(weights) if weights else 1
    nx.draw_networkx_edges(G, pos, ax=ax,
                           width=[1 + (w / max_w) * 5 for w in weights],
                           edge_color=[C_ORANGE if w >= 2 else "#BDC3C7" for w in weights], alpha=0.6)

    node_colors = [STATUS_COLORS.get(G.nodes[n].get("status", ""), C_GRAY) for n in G.nodes()]
    nx.draw_networkx_nodes(G, pos, ax=ax, node_color=node_colors,
                           node_size=[max(300, G.degree(n) * 200) for n in G.nodes()],
                           edgecolors=C_PRIMARY, linewidths=1.5, alpha=0.9)

    nx.draw_networkx_labels(G, pos, {n: heb(G.nodes[n]["label"]) for n in G.nodes()},
                            ax=ax, font_size=8, font_family=FF)

    strong = {(u, v): str(d["weight"]) for u, v, d in edges if d["weight"] >= 2}
    nx.draw_networkx_edge_labels(G, pos, strong, ax=ax, font_size=8, font_color=C_RED)

    legend = [mpatches.Patch(color=c, label=heb(s)) for s, c in STATUS_COLORS.items()]
    ax.legend(handles=legend, loc='upper left', fontsize=9, prop={'family': FF})
    ax.set_title(heb("רשת קשרים בין פרוייקטים"), fontsize=14, fontweight='bold', fontfamily=FF)
    ax.axis('off')
    fig.tight_layout()
    return fig


# ═══════════════════════════════════════════════════════════════
# 5) BUILD PDF
# ═══════════════════════════════════════════════════════════════
def build_pdf(projects):
    doc = SimpleDocTemplate(OUTPUT_PDF, pagesize=A4,
                            leftMargin=15 * mm, rightMargin=15 * mm,
                            topMargin=20 * mm, bottomMargin=15 * mm)
    W = A4[0] - 30 * mm
    story = []

    s_title = ParagraphStyle("T", fontName="HebB", fontSize=26, alignment=TA_CENTER, textColor=HexColor(C_PRIMARY), spaceAfter=6 * mm)
    s_sub = ParagraphStyle("S", fontName="Heb", fontSize=13, alignment=TA_CENTER, textColor=HexColor(C_ACCENT), spaceAfter=12 * mm)
    s_h1 = ParagraphStyle("H1", fontName="HebB", fontSize=16, alignment=TA_RIGHT, textColor=HexColor(C_PRIMARY), spaceAfter=4 * mm, spaceBefore=8 * mm)
    s_h2 = ParagraphStyle("H2", fontName="HebB", fontSize=13, alignment=TA_RIGHT, textColor=HexColor(C_ACCENT), spaceAfter=3 * mm, spaceBefore=5 * mm)
    s_body = ParagraphStyle("B", fontName="Heb", fontSize=10, alignment=TA_RIGHT, leading=14, spaceAfter=2 * mm)
    s_cell = ParagraphStyle("C", fontName="Heb", fontSize=7.5, alignment=TA_CENTER, leading=10)
    s_cell_r = ParagraphStyle("CR", fontName="Heb", fontSize=7.5, alignment=TA_RIGHT, leading=10)

    # ── Page 1: Title + Summary ───────────────────────────────
    story.append(Spacer(1, 40 * mm))
    story.append(Paragraph(heb("דוח סיכום פרוייקטים"), s_title))
    story.append(Paragraph(heb("ניתוח קשרים, סטטוסים ומילות מפתח"), s_sub))
    story.append(Spacer(1, 10 * mm))

    total = len(projects)
    active = sum(1 for p in projects if p["status"] == "בפרודקשן")
    done = sum(1 for p in projects if p["status"] == "הושלם")
    paused = sum(1 for p in projects if p["status"] == "מושהה")

    boxes = [
        (str(paused), heb("מושהים"), C_ORANGE),
        (str(done), heb("הושלמו"), C_ACCENT),
        (str(active), heb("בפרודקשן"), C_GREEN),
        (str(total), heb("סה״כ"), C_PRIMARY),
    ]
    row1 = [Paragraph(f"<b>{n}</b>", ParagraphStyle("x", fontName="HebB", fontSize=20, alignment=TA_CENTER, textColor=HexColor(c))) for n, _, c in boxes]
    row2 = [Paragraph(lbl, ParagraphStyle("x", fontName="Heb", fontSize=10, alignment=TA_CENTER)) for _, lbl, _ in boxes]

    t = Table([row1, row2], colWidths=[W / 4] * 4)
    t.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        *[('BOX', (i, 0), (i, -1), 1, HexColor(c)) for i, (_, _, c) in enumerate(boxes)],
        ('TOPPADDING', (0, 0), (-1, 0), 8),
        ('BOTTOMPADDING', (0, -1), (-1, -1), 8),
    ]))
    story.append(t)
    story.append(PageBreak())

    # ── Page 2: Status + Manager ──────────────────────────────
    story.append(Paragraph(heb("התפלגות סטטוס פרוייקטים"), s_h1))
    story.append(fig_to_image(chart_status_pie(projects), width=130 * mm))
    story.append(Spacer(1, 5 * mm))
    story.append(Paragraph(heb("עומס עבודה לפי מנהל פרוייקט"), s_h1))
    story.append(fig_to_image(chart_manager_bar(projects), width=150 * mm))
    story.append(PageBreak())

    # ── Page 3: Keywords + Clients ────────────────────────────
    story.append(Paragraph(heb("מילות מפתח נפוצות"), s_h1))
    story.append(fig_to_image(chart_keywords(projects), width=155 * mm))
    story.append(Spacer(1, 5 * mm))
    story.append(Paragraph(heb("פרוייקטים לפי לקוח"), s_h1))
    story.append(fig_to_image(chart_clients(projects), width=140 * mm))
    story.append(PageBreak())

    # ── Page 4: Network ───────────────────────────────────────
    story.append(Paragraph(heb("רשת קשרים בין פרוייקטים"), s_h1))
    story.append(Paragraph(heb("קווים עבים וכתומים = 2+ מילות מפתח משותפות. גודל עיגול = מספר חיבורים."), s_body))
    story.append(fig_to_image(chart_network(projects), width=170 * mm))
    story.append(PageBreak())

    # ── Page 5: Full table ────────────────────────────────────
    story.append(Paragraph(heb("טבלת פרוייקטים"), s_h1))
    hdr = [heb(h) for h in ["קשרים", "סטטוס", "לקוח", "מנהל", "תיאור", "שם פרוייקט", "#"]]
    tdata = [hdr]
    for p in projects:
        mx = f" (max {p['max_similarity_score']})" if p["similar_projects"] else ""
        tdata.append([
            Paragraph(f"{p['total_connections']}{mx}", s_cell),
            Paragraph(heb(p["status"]), s_cell),
            Paragraph(heb(p["clients"]), s_cell),
            Paragraph(heb(p["manager"]), s_cell),
            Paragraph(heb(p["description"][:60]), s_cell_r),
            Paragraph(heb(p["name"]), s_cell),
            Paragraph(str(p["id"]), s_cell),
        ])
    cw = [22 * mm, 22 * mm, 18 * mm, 24 * mm, 52 * mm, 28 * mm, 10 * mm]
    tbl = Table(tdata, colWidths=cw, repeatRows=1)
    tbl.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, 0), 'HebB'), ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('BACKGROUND', (0, 0), (-1, 0), HexColor(C_PRIMARY)), ('TEXTCOLOR', (0, 0), (-1, 0), white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, HexColor("#BDC3C7")),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [white, HexColor(C_LIGHT)]),
        ('TOPPADDING', (0, 0), (-1, -1), 3), ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    story.append(tbl)
    story.append(PageBreak())

    # ── Page 6: Strong connections + isolated ─────────────────
    story.append(Paragraph(heb("קשרים חזקים בין פרוייקטים"), s_h1))
    story.append(Paragraph(heb("זוגות עם 2+ מילות מפתח משותפות:"), s_body))

    pairs = []
    seen = set()
    for p in projects:
        for s in p["similar_projects"]:
            if s["score"] >= 2:
                key = tuple(sorted([p["id"], s["project_id"]]))
                if key not in seen:
                    seen.add(key)
                    pairs.append((p["name"], s["project_name"], s["score"], s["shared_keywords"]))
    pairs.sort(key=lambda x: x[2], reverse=True)

    chdr = [heb(h) for h in ["מילות מפתח משותפות", "ציון", "פרוייקט ב", "פרוייקט א"]]
    cdata = [chdr]
    for a, b, score, shared in pairs:
        cdata.append([
            Paragraph(heb(", ".join(shared)), s_cell_r),
            Paragraph(f"<b>{score}</b>", ParagraphStyle("x", fontName="HebB", fontSize=9, alignment=TA_CENTER,
                                                         textColor=HexColor(C_RED if score >= 3 else C_ORANGE))),
            Paragraph(heb(b), s_cell),
            Paragraph(heb(a), s_cell),
        ])
    ctbl = Table(cdata, colWidths=[55 * mm, 15 * mm, 45 * mm, 45 * mm], repeatRows=1)
    ctbl.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, 0), 'HebB'), ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('BACKGROUND', (0, 0), (-1, 0), HexColor(C_ACCENT)), ('TEXTCOLOR', (0, 0), (-1, 0), white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, HexColor("#BDC3C7")),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [white, HexColor("#EBF5FB")]),
        ('TOPPADDING', (0, 0), (-1, -1), 4), ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(ctbl)
    story.append(Spacer(1, 8 * mm))

    isolated = [p for p in projects if not p["similar_projects"]]
    if isolated:
        story.append(Paragraph(heb("פרוייקטים ללא קשרים (מבודדים)"), s_h2))
        for p in isolated:
            story.append(Paragraph(heb(f"• {p['name']} — {p['description']}"), s_body))

    doc.build(story)
    print(f"✅ PDF saved: {OUTPUT_PDF}")


# ═══════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════
if __name__ == "__main__":
    projects = read_data()
    projects = build_similarity(projects)
    save_json(projects)
    build_pdf(projects)