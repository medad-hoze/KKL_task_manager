"""
Google Sheet → PDF Management Report
Executive summary with time-based analytics and manager dashboard.

Requirements:
    pip install pandas matplotlib networkx reportlab Pillow
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
from collections import Counter, defaultdict
from itertools import combinations
from io import BytesIO
from PIL import Image as PILImage
from datetime import datetime

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white, black
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak,
    Flowable
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
from reportlab.pdfgen import canvas

import re

# ═══════════════════════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════════════════════
SHEET_ID = "131Vl8q-bQFqMi7CAnLwy-rkLGBswKmgx9o_YOUSQ7zc"
GID = "0"
OUTPUT_PDF = "project_report.pdf"
OUTPUT_JSON = "projects_full.json"

# Column name in the Google Sheet for the new "client manager" field
COL_CLIENT_MANAGER = "מנהל לקוחות"

# ═══════════════════════════════════════════════════════════════
# FONT SETUP
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

# Professional Color Palette
C_PRIMARY = "#1A365D"
C_ACCENT  = "#3182CE"
C_LIGHT   = "#EBF8FF"
C_GREEN   = "#38A169"
C_ORANGE  = "#DD6B20"
C_RED     = "#E53E3E"
C_GRAY    = "#718096"
C_PURPLE  = "#805AD5"
C_YELLOW  = "#D69E2E"
C_DARK    = "#2D3748"
C_TEAL    = "#319795"

STATUS_COLORS = {
    "בפרודקשן": C_GREEN,
    "הושלם": C_ACCENT,
    "מושהה": C_ORANGE,
    "גרסה ישנה": C_GRAY,
}
MANAGER_COLORS = [C_ACCENT, C_GREEN, C_ORANGE, C_PURPLE, C_RED,
                  "#319795", "#D69E2E", "#2B6CB0"]


# ───────────────────────────────────────────────────────────────
# HELPERS
# ───────────────────────────────────────────────────────────────
def heb(text):
    MIRROR = str.maketrans('()[]{}', ')(][}{')
    s = str(text)
    parts = re.split(r'([A-Za-z0-9_.:/\-]+)', s)
    out = []
    for p in parts:
        if re.match(r'^[A-Za-z0-9_.:/\-]+$', p):
            out.append(p)
        else:
            out.append(p[::-1].translate(MIRROR))
    return ''.join(out[::-1])


def parse_date(d):
    for fmt in ("%d/%m/%Y", "%d.%m.%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(d.strip(), fmt)
        except (ValueError, AttributeError):
            continue
    return None


def fig_to_image(fig, width=170 * mm, max_height=210 * mm):
    buf = BytesIO()
    fig.savefig(buf, format='png', dpi=200, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close(fig)
    buf.seek(0)
    w, h = PILImage.open(buf).size
    buf.seek(0)
    img_h = width * (h / w)
    if img_h > max_height:
        width = max_height * (w / h)
        img_h = max_height
    return Image(buf, width=width, height=img_h)


# ─── Colored divider line flowable ────────────────────────────
class ColorLine(Flowable):
    def __init__(self, width, color=C_PRIMARY, thickness=1.5):
        Flowable.__init__(self)
        self.width = width
        self.color = color
        self.thickness = thickness
        self.height = thickness + 2 * mm

    def draw(self):
        self.canv.setStrokeColor(HexColor(self.color))
        self.canv.setLineWidth(self.thickness)
        self.canv.line(0, self.thickness / 2, self.width, self.thickness / 2)


# ─── Page numbers + header/footer ─────────────────────────────
class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self._draw_extras(num_pages)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def _draw_extras(self, page_count):
        pg = self._pageNumber
        if pg == 1:
            return
        w, h = A4
        self.setStrokeColor(HexColor(C_PRIMARY))
        self.setLineWidth(1.5)
        self.line(15 * mm, h - 12 * mm, w - 15 * mm, h - 12 * mm)
        self.setStrokeColor(HexColor("#E2E8F0"))
        self.setLineWidth(0.8)
        self.line(15 * mm, 13 * mm, w - 15 * mm, 13 * mm)
        self.setFont("Heb", 8)
        self.setFillColor(HexColor(C_GRAY))
        self.drawCentredString(w / 2, 7 * mm, f"{pg} / {page_count}")
        self.setFont("Heb", 7)
        self.drawString(15 * mm, 7 * mm, heb("דוח ניהולי — סיכום פרוייקטים"))
        self.drawRightString(w - 15 * mm, 7 * mm,
                             datetime.now().strftime("%d/%m/%Y"))


# ═══════════════════════════════════════════════════════════════
# 1) READ DATA
# ═══════════════════════════════════════════════════════════════
def _clean_ein(val):
    """Blank out a cell whose stripped value is exactly 'אין'."""
    return "" if str(val).strip() == "אין" else val


def read_data():
    url = (f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
           f"/export?format=csv&gid={GID}")
    df = pd.read_csv(url, dtype=str).fillna("")
    print(f"✔ Read {len(df)} rows from Google Sheet")

    # ── Replace every cell that is exactly "אין" with "" ──
    df = df.applymap(lambda v: "" if str(v).strip() == "אין" else v)

    # Tolerate the new column being missing (older sheet copies)
    if COL_CLIENT_MANAGER not in df.columns:
        print(f"⚠ Column '{COL_CLIENT_MANAGER}' not found in sheet — "
              f"defaulting to empty.")
        df[COL_CLIENT_MANAGER] = ""

    projects = []
    for _, row in df.iterrows():
        projects.append({
            "id":             row["מספר סידורי"],
            "name":           row["שם פרוייקט"],
            "responsible":    row["אחראי פרוייקט"],
            "developer":      row["מפתח"],
            "developer_b":    row["מפתח ב"],
            "manager":        row["מנהל פרוייקט"],
            "manager_b":      row["מנהל פרוייקט ב"],
            "client_manager": row[COL_CLIENT_MANAGER],
            "start":          row["תאריך התחלה"],
            "end":            row["תאריך סיום"],
            "start_dt":       parse_date(row.get("תאריך התחלה", "")),
            "end_dt":         parse_date(row.get("תאריך סיום", "")),
            "status":         row["סטטוס"],
            "clients":        row["לקוחות"],
            "description":    row["תיאור"],
            "keywords":       [k.strip().lower()
                               for k in row["מילות מפתח"].split(",") if k.strip()],
        })
    return projects


def build_similarity(projects):
    for p in projects:
        similar = []
        for o in projects:
            if o["id"] == p["id"] and o["name"] == p["name"]:
                continue
            shared = set(p["keywords"]) & set(o["keywords"])
            if shared:
                similar.append({"project_id": o["id"],
                                "project_name": o["name"],
                                "shared_keywords": sorted(shared),
                                "score": len(shared)})
        similar.sort(key=lambda x: x["score"], reverse=True)
        p["similar_projects"] = similar
        p["total_connections"] = len(similar)
        p["max_similarity_score"] = similar[0]["score"] if similar else 0
    return projects


def save_json(projects):
    data = [{k: v for k, v in p.items() if k not in ("start_dt", "end_dt")}
            for p in projects]
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"✔ JSON saved: {OUTPUT_JSON}")


# ═══════════════════════════════════════════════════════════════
# 4) CHARTS (UPGRADED STYLING)
# ═══════════════════════════════════════════════════════════════
def _style_ax(ax, title="", include_grid=True):
    """Applies a clean, modern aesthetic to the chart."""
    if title:
        ax.set_title(heb(title), fontsize=14, fontweight='bold',
                     fontfamily=FF, pad=18, color=C_PRIMARY)

    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_color('#CBD5E0')
    ax.tick_params(colors='#4A5568', bottom=False, left=False)

    if include_grid:
        ax.yaxis.grid(True, linestyle='-', alpha=0.5, color='#EDF2F7')
        ax.xaxis.grid(False)
        ax.set_axisbelow(True)


def chart_status_donut(projects):
    counts = Counter(p["status"] for p in projects)
    fig, ax = plt.subplots(figsize=(5.5, 4.5))
    labels = [heb(k) for k in counts.keys()]
    clrs = [STATUS_COLORS.get(k, C_GRAY) for k in counts.keys()]
    wedges, texts, autotexts = ax.pie(
        list(counts.values()), labels=labels, colors=clrs,
        autopct='%1.0f%%', startangle=90, pctdistance=0.78,
        wedgeprops=dict(width=0.45, edgecolor='white', linewidth=2),
        textprops={'fontsize': 11, 'fontfamily': FF, 'color': '#2D3748'})
    for t in autotexts:
        t.set_fontsize(10); t.set_color('white'); t.set_fontweight('bold')
    ax.text(0, 0, str(len(projects)), ha='center', va='center',
            fontsize=28, fontweight='bold', color=C_PRIMARY, fontfamily=FF)
    ax.text(0, -0.15, heb("פרוייקטים"), ha='center', va='center',
            fontsize=10, color=C_GRAY, fontfamily=FF)
    fig.tight_layout()
    return fig


def chart_timeline(projects):
    dated = [(p, p["start_dt"], p["end_dt"] or datetime.now())
             for p in projects if p["start_dt"]]
    if not dated:
        return None
    dated.sort(key=lambda x: x[1])

    fig, ax = plt.subplots(figsize=(9, min(9, max(4.5, len(dated) * 0.4))))

    for i, (p, start, end) in enumerate(dated):
        clr = STATUS_COLORS.get(p["status"], C_GRAY)
        dur = (end - start).days
        ax.barh(i, dur, left=start.toordinal(), color=clr,
                height=0.5, edgecolor='white', linewidth=0.5, alpha=0.9)
        ax.text(start.toordinal() + dur + 8, i,
                f"{dur}{heb(' ימים')}", va='center',
                fontsize=8, color='#4A5568', fontfamily=FF)

    ax.set_yticks(range(len(dated)))
    ax.set_yticklabels([heb(p["name"][:25]) for p, _, _ in dated],
                       fontsize=9, fontfamily=FF, color='#2D3748')
    ax.invert_yaxis()

    all_d = [d for _, s, e in dated for d in (s, e)]
    min_d, max_d = min(all_d), max(all_d)
    ax.set_xlim(min_d.toordinal() - 15, max_d.toordinal() + 60)

    import matplotlib.dates as mdates
    ax.xaxis_date()
    tm = (max_d.year - min_d.year) * 12 + (max_d.month - min_d.month)
    if tm > 60:   iv, fmt = 12, '%Y'
    elif tm > 30: iv, fmt = 6, '%m/%Y'
    elif tm > 12: iv, fmt = 3, '%m/%Y'
    else:         iv, fmt = 1, '%m/%Y'
    ax.xaxis.set_major_formatter(mdates.DateFormatter(fmt))
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=iv))
    plt.xticks(rotation=0, fontsize=9, ha='center', color='#4A5568')

    _style_ax(ax, "ציר זמן פרוייקטים", include_grid=False)
    ax.xaxis.grid(True, linestyle='-', alpha=0.4, color='#EDF2F7')

    legend = [mpatches.Patch(color=c, label=heb(s)) for s, c in STATUS_COLORS.items()]
    ax.legend(handles=legend, loc='upper center', bbox_to_anchor=(0.5, -0.1),
              fontsize=9, prop={'family': FF}, ncol=4, frameon=False)

    fig.tight_layout()
    return fig


def chart_monthly_starts(projects):
    dated = [p for p in projects if p["start_dt"]]
    if not dated:
        return None
    monthly = Counter(p["start_dt"].strftime("%Y-%m") for p in dated)
    months = sorted(monthly.keys())
    vals = [monthly[m] for m in months]

    fig, ax = plt.subplots(figsize=(8, 4))
    bars = ax.bar(range(len(months)), vals, color=C_ACCENT, edgecolor='white',
                  linewidth=1, width=0.6, alpha=0.9)
    ax.set_xticks(range(len(months)))
    ax.set_xticklabels(months, fontsize=9, rotation=30, ha='right', color='#4A5568')
    ax.set_ylabel(heb("מספר פרוייקטים"), fontsize=10, fontfamily=FF, color='#4A5568', labelpad=10)

    for bar, v in zip(bars, vals):
        if v > 0:
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.15,
                    str(v), ha='center', fontsize=9, fontweight='bold',
                    color=C_PRIMARY)

    _style_ax(ax, "פרוייקטים חדשים לפי חודש")
    fig.tight_layout()
    return fig


def chart_duration_distribution(projects):
    durations = []
    for p in projects:
        if p["start_dt"] and p["end_dt"]:
            durations.append((p["end_dt"] - p["start_dt"]).days)
    if not durations:
        return None

    fig, ax = plt.subplots(figsize=(8, 4))
    bins = np.linspace(0, max(durations) + 30, min(12, len(set(durations)) + 1))
    n, bins_out, patches = ax.hist(durations, bins=bins, color=C_ACCENT,
                                   edgecolor='white', linewidth=1.5, alpha=0.9)

    for patch, left in zip(patches, bins_out[:-1]):
        if left > 365:   patch.set_facecolor(C_RED)
        elif left > 180: patch.set_facecolor(C_ORANGE)
        else:            patch.set_facecolor(C_GREEN)

    avg_d, med_d = np.mean(durations), np.median(durations)
    ax.axvline(avg_d, color=C_DARK, ls='--', lw=1.5, alpha=0.8)
    ax.axvline(med_d, color=C_PURPLE, ls=':', lw=2, alpha=0.8)

    y_max = ax.get_ylim()[1]
    ax.text(avg_d + 5, y_max * 0.9, f'{heb("ממוצע")}: {avg_d:.0f}',
            fontsize=9, color=C_DARK, fontfamily=FF, fontweight='bold')
    ax.text(med_d + 5, y_max * 0.75, f'{heb("חציון")}: {med_d:.0f}',
            fontsize=9, color=C_PURPLE, fontfamily=FF, fontweight='bold')

    ax.set_xlabel(heb("ימים"), fontsize=10, fontfamily=FF, color='#4A5568', labelpad=10)
    ax.set_ylabel(heb("מספר פרוייקטים"), fontsize=10, fontfamily=FF, color='#4A5568', labelpad=10)

    _style_ax(ax, "התפלגות משך פרוייקטים (ימים)")
    fig.tight_layout()
    return fig


def chart_active_over_time(projects):
    dated = [p for p in projects if p["start_dt"]]
    if not dated:
        return None
    all_dates = [p["start_dt"] for p in dated]
    end_dates = [p["end_dt"] or datetime.now() for p in dated]
    min_d = min(all_dates).replace(day=1)
    max_d = max(end_dates).replace(day=1)
    months = []
    cur = min_d
    while cur <= max_d:
        months.append(cur)
        m, y = cur.month + 1, cur.year
        if m > 12: m, y = 1, y + 1
        cur = cur.replace(year=y, month=m)

    active_counts = [
        sum(1 for p in dated
            if p["start_dt"] <= m and (p["end_dt"] or datetime.now()) >= m)
        for m in months]
    prod = [p for p in dated if p["status"] == "בפרודקשן"]
    maint_counts = [sum(1 for p in prod if p["start_dt"] <= m) for m in months]

    fig, ax = plt.subplots(figsize=(9, 4.5))
    ax.fill_between(range(len(months)), active_counts, alpha=0.15, color=C_ACCENT)
    ax.plot(range(len(months)), active_counts, color=C_ACCENT, lw=2.5,
            marker='o', ms=4, mfc='white', mec=C_ACCENT,
            label=heb("פרוייקטים פעילים"))

    ax.fill_between(range(len(months)), maint_counts, alpha=0.1, color=C_GREEN)
    ax.plot(range(len(months)), maint_counts, color=C_GREEN, lw=2.5,
            marker='s', ms=4, mfc='white', mec=C_GREEN,
            label=heb("תחזוקה (בפרודקשן)"))

    step = max(1, len(months) // 10)
    ax.set_xticks(range(0, len(months), step))
    ax.set_xticklabels([m.strftime("%m/%Y") for m in months[::step]],
                       fontsize=9, rotation=0, ha='center', color='#4A5568')
    ax.set_ylabel(heb("מספר פרוייקטים"), fontsize=10, fontfamily=FF, color='#4A5568', labelpad=10)

    if active_counts:
        peak = max(active_counts)
        pi = active_counts.index(peak)
        ax.annotate(f'{heb("שיא")}: {peak}', xy=(pi, peak), fontsize=10,
                    fontfamily=FF, fontweight='bold', xytext=(pi + 0.5, peak + 1.5), color=C_RED,
                    arrowprops=dict(arrowstyle='->', color=C_RED, lw=1.5))

    _style_ax(ax, "פרוייקטים פעילים ותחזוקה לאורך זמן")

    ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15),
              fontsize=9, prop={'family': FF}, ncol=2, frameon=False)

    fig.tight_layout()
    return fig


def chart_manager_stacked(projects):
    mgrs = defaultdict(Counter)
    for p in projects:
        if p["manager"]:
            mgrs[p["manager"]][p["status"]] += 1
        if p["manager_b"]:
            mgrs[p["manager_b"]][p["status"]] += 1
    if not mgrs:
        return None
    names = sorted(mgrs.keys(), key=lambda m: sum(mgrs[m].values()), reverse=True)
    statuses = list(STATUS_COLORS.keys())

    fig, ax = plt.subplots(figsize=(8.5, max(4, len(names) * 0.8)))
    y = np.arange(len(names))
    left = np.zeros(len(names))

    for st in statuses:
        vals = [mgrs[m].get(st, 0) for m in names]
        ax.barh(y, vals, left=left, height=0.5, color=STATUS_COLORS[st],
                edgecolor='white', linewidth=1, label=heb(st))
        for i, v in enumerate(vals):
            if v > 0:
                ax.text(left[i] + v / 2, y[i], str(v), ha='center',
                        va='center', fontsize=9, fontweight='bold',
                        color='white', fontfamily=FF)
        left += vals

    for i, m in enumerate(names):
        ax.text(left[i] + 0.3, y[i], str(int(left[i])), ha='left',
                va='center', fontsize=10, fontweight='bold',
                color=C_PRIMARY, fontfamily=FF)

    ax.set_yticks(y)
    ax.set_yticklabels([heb(m) for m in names], fontsize=10, fontfamily=FF, color='#2D3748')
    ax.set_xlabel(heb("מספר פרוייקטים"), fontsize=10, fontfamily=FF, color='#4A5568', labelpad=10)

    _style_ax(ax, "מנהלי פרוייקטים — פילוח לפי סטטוס", include_grid=False)
    ax.xaxis.grid(True, linestyle='-', alpha=0.5, color='#EDF2F7')

    ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.12), fontsize=9,
              prop={'family': FF}, ncol=len(statuses), frameon=False)

    fig.tight_layout()
    return fig


def chart_client_managers_stacked(projects):
    """Same idea as chart_manager_stacked but for the client_manager field."""
    cms = defaultdict(Counter)
    for p in projects:
        cm = (p.get("client_manager") or "").strip()
        if cm:
            cms[cm][p["status"]] += 1
    if not cms:
        return None

    names = sorted(cms.keys(), key=lambda m: sum(cms[m].values()), reverse=True)
    statuses = list(STATUS_COLORS.keys())

    fig, ax = plt.subplots(figsize=(8.5, max(4, len(names) * 0.8)))
    y = np.arange(len(names))
    left = np.zeros(len(names))

    for st in statuses:
        vals = [cms[m].get(st, 0) for m in names]
        ax.barh(y, vals, left=left, height=0.5, color=STATUS_COLORS[st],
                edgecolor='white', linewidth=1, label=heb(st))
        for i, v in enumerate(vals):
            if v > 0:
                ax.text(left[i] + v / 2, y[i], str(v), ha='center',
                        va='center', fontsize=9, fontweight='bold',
                        color='white', fontfamily=FF)
        left += vals

    for i, m in enumerate(names):
        ax.text(left[i] + 0.3, y[i], str(int(left[i])), ha='left',
                va='center', fontsize=10, fontweight='bold',
                color=C_PRIMARY, fontfamily=FF)

    ax.set_yticks(y)
    ax.set_yticklabels([heb(m) for m in names], fontsize=10, fontfamily=FF, color='#2D3748')
    ax.set_xlabel(heb("מספר פרוייקטים"), fontsize=10, fontfamily=FF, color='#4A5568', labelpad=10)

    _style_ax(ax, "מנהלי לקוחות — פילוח לפי סטטוס", include_grid=False)
    ax.xaxis.grid(True, linestyle='-', alpha=0.5, color='#EDF2F7')

    ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.12), fontsize=9,
              prop={'family': FF}, ncol=len(statuses), frameon=False)

    fig.tight_layout()
    return fig


def chart_developers(projects):
    devs = Counter()
    dev_status = defaultdict(Counter)
    for p in projects:
        for field in ("developer",):
            d = p.get(field, "").strip()
            if d:
                devs[d] += 1
                dev_status[d][p["status"]] += 1
    if not devs:
        return None

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, max(4.5, len(devs) * 0.6)),
                                   gridspec_kw={'width_ratios': [1, 1.6]})

    # ── Donut ──
    names = [k for k, _ in devs.most_common()]
    vals = [v for _, v in devs.most_common()]
    clrs = (MANAGER_COLORS * 3)[:len(names)]
    wedges, texts, autotexts = ax1.pie(
        vals, labels=[heb(n) for n in names], colors=clrs,
        autopct='%1.0f%%', startangle=90, pctdistance=0.78,
        wedgeprops=dict(width=0.45, edgecolor='white', linewidth=2),
        textprops={'fontsize': 9, 'fontfamily': FF, 'color': '#2D3748'})
    for t in autotexts:
        t.set_fontsize(8); t.set_color('white'); t.set_fontweight('bold')
    ax1.text(0, 0, str(sum(vals)), ha='center', va='center',
             fontsize=24, fontweight='bold', color=C_PRIMARY, fontfamily=FF)
    ax1.text(0, -0.2, heb("שיבוצים"), ha='center', va='center',
             fontsize=10, color=C_GRAY, fontfamily=FF)

    # ── Stacked bar ──
    sorted_devs = [k for k, _ in devs.most_common()]
    y_pos = np.arange(len(sorted_devs))
    statuses = list(STATUS_COLORS.keys())
    left = np.zeros(len(sorted_devs))

    for st in statuses:
        v = [dev_status[d].get(st, 0) for d in sorted_devs]
        ax2.barh(y_pos, v, left=left, height=0.5,
                 color=STATUS_COLORS[st], edgecolor='white',
                 linewidth=1, label=heb(st))
        for i, val in enumerate(v):
            if val > 0:
                ax2.text(left[i] + val / 2, y_pos[i], str(val),
                         ha='center', va='center', fontsize=9,
                         fontweight='bold', color='white', fontfamily=FF)
        left += v

    for i, d in enumerate(sorted_devs):
        ax2.text(left[i] + 0.3, y_pos[i], str(int(left[i])),
                 ha='left', va='center', fontsize=10,
                 fontweight='bold', color=C_PRIMARY, fontfamily=FF)

    ax2.set_yticks(y_pos)
    ax2.set_yticklabels([heb(d) for d in sorted_devs], fontsize=10, fontfamily=FF, color='#2D3748')
    ax2.set_xlabel(heb("מספר פרוייקטים"), fontsize=10, fontfamily=FF, color='#4A5568', labelpad=10)

    _style_ax(ax2, include_grid=False)
    ax2.xaxis.grid(True, linestyle='-', alpha=0.5, color='#EDF2F7')

    ax2.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15),
               fontsize=9, prop={'family': FF}, ncol=4, frameon=False)

    fig.suptitle(heb("פרוייקטים לפי מפתח"), fontsize=16,
                 fontweight='bold', fontfamily=FF, color=C_PRIMARY, y=1.02)
    fig.tight_layout()
    return fig


def chart_keywords(projects):
    all_kw = [k for p in projects for k in p["keywords"]]
    freq = Counter(all_kw).most_common(10)
    if not freq:
        return None

    fig, ax = plt.subplots(figsize=(8, 4))
    clrs = [C_PRIMARY, C_ACCENT, C_GREEN, C_ORANGE, C_PURPLE,
            C_RED, C_YELLOW, C_GRAY, C_PRIMARY, C_ACCENT]
    bars = ax.bar(range(len(freq)), [v for _, v in freq],
                  color=clrs[:len(freq)], edgecolor='white',
                  linewidth=1, width=0.6, alpha=0.9)

    ax.set_xticks(range(len(freq)))
    ax.set_xticklabels([heb(k) for k, _ in freq], fontsize=10,
                       fontfamily=FF, rotation=30, ha='right', color='#2D3748')
    ax.set_ylabel(heb("מספר פרוייקטים"), fontsize=10, fontfamily=FF, color='#4A5568', labelpad=10)

    for bar, (_, v) in zip(bars, freq):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.15,
                str(v), ha='center', fontsize=10, fontweight='bold',
                color=C_PRIMARY)

    _style_ax(ax, "מילות מפתח נפוצות")
    fig.tight_layout()
    return fig


# ═══════════════════════════════════════════════════════════════
# 5) BUILD PDF
# ═══════════════════════════════════════════════════════════════
def build_pdf(projects):
    doc = SimpleDocTemplate(OUTPUT_PDF, pagesize=A4,
                            leftMargin=15 * mm, rightMargin=15 * mm,
                            topMargin=18 * mm, bottomMargin=18 * mm)
    W = A4[0] - 30 * mm
    story = []

    # ── Styles ────────────────────────────────────────────────
    s_title = ParagraphStyle("T", fontName="HebB", fontSize=34, leading=42,
                             alignment=TA_CENTER, textColor=HexColor(C_PRIMARY),
                             spaceAfter=4 * mm)
    s_title2 = ParagraphStyle("T2", fontName="Heb", fontSize=20, leading=26,
                              alignment=TA_CENTER, textColor=HexColor(C_ACCENT),
                              spaceAfter=8 * mm)
    s_sub = ParagraphStyle("S", fontName="Heb", fontSize=12,
                           alignment=TA_CENTER, textColor=HexColor(C_GRAY),
                           spaceAfter=8 * mm)
    s_h1 = ParagraphStyle("H1", fontName="HebB", fontSize=16,
                          alignment=TA_RIGHT, textColor=HexColor(C_PRIMARY),
                          spaceAfter=3 * mm, spaceBefore=6 * mm)
    s_body = ParagraphStyle("B", fontName="Heb", fontSize=10,
                            alignment=TA_RIGHT, leading=15, spaceAfter=2 * mm,
                            textColor=HexColor('#4A5568'))
    s_insight = ParagraphStyle("INS", fontName="Heb", fontSize=11,
                               alignment=TA_RIGHT, leading=18,
                               spaceAfter=3 * mm,
                               textColor=HexColor('#2D3748'), leftIndent=3 * mm)

    # ── Stats ─────────────────────────────────────────────────
    total  = len(projects)
    active = sum(1 for p in projects if p["status"] == "בפרודקשן")
    done   = sum(1 for p in projects if p["status"] == "הושלם")
    paused = sum(1 for p in projects if p["status"] == "מושהה")

    durations = [(p["end_dt"] - p["start_dt"]).days
                 for p in projects if p["start_dt"] and p["end_dt"]]
    avg_dur = int(np.mean(durations)) if durations else 0

    all_kw = [k for p in projects for k in p["keywords"]]
    top_kw = Counter(all_kw).most_common(1)[0] if all_kw else ("—", 0)
    all_mgrs = set()
    for p in projects:
        if p["manager"]:   all_mgrs.add(p["manager"])
        if p["manager_b"]: all_mgrs.add(p["manager_b"])
    n_mgrs = len(all_mgrs)

    n_client_mgrs = len({(p.get("client_manager") or "").strip()
                         for p in projects
                         if (p.get("client_manager") or "").strip()})

    # ══════════════════════════════════════════════════════════
    # COVER PAGE
    # ══════════════════════════════════════════════════════════
    story.append(Spacer(1, 40 * mm))
    story.append(Paragraph(heb("דוח ניהולי"), s_title))
    story.append(Paragraph(heb("סיכום פרוייקטים"), s_title2))
    story.append(ColorLine(W, C_PRIMARY, 1.5))
    story.append(Spacer(1, 5 * mm))
    story.append(Paragraph(
        heb(f"תאריך הפקה: {datetime.now().strftime('%d/%m/%Y')}"), s_sub))
    story.append(Spacer(1, 20 * mm))

    kpis = [
        (str(total),  heb("סהכ פרוייקטים"), C_PRIMARY),
        (str(active), heb("בפרודקשן"),       C_GREEN),
        (str(done),   heb("הושלמו"),         C_ACCENT),
        (str(paused), heb("מושהים"),         C_ORANGE),
    ]

    r_num = [Paragraph(f"<font color='{c}'>{v}</font>", ParagraphStyle(
                 "kv", fontName="HebB", fontSize=42, leading=50, alignment=TA_CENTER,
                 spaceAfter=0))
             for v, _, c in kpis]

    r_line = [Paragraph(f"<font color='#CBD5E0'>{'━' * 6}</font>", ParagraphStyle(
                  "kd", fontName="Heb", fontSize=10, leading=12, alignment=TA_CENTER))
              for _ in kpis]

    r_lbl = [Paragraph(lbl, ParagraphStyle(
                 "kl", fontName="HebB", fontSize=12, leading=14, alignment=TA_CENTER,
                 textColor=HexColor('#4A5568')))
             for _, lbl, _ in kpis]

    t = Table([r_num, r_line, r_lbl], colWidths=[W / 4] * 4)
    t.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, 0), 20),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 5),
        ('TOPPADDING', (0, 1), (-1, 1), 5),
        ('BOTTOMPADDING', (0, 1), (-1, 1), 5),
        ('TOPPADDING', (0, 2), (-1, 2), 5),
        ('BOTTOMPADDING', (0, 2), (-1, 2), 20),
        ('BACKGROUND', (0, 0), (-1, -1), HexColor('#F8FAFC')),
        ('LINEBELOW', (0, -1), (-1, -1), 3, HexColor(C_ACCENT)),
    ]))
    story.append(t)
    story.append(Spacer(1, 25 * mm))

    # Insights
    story.append(Paragraph(heb("תובנות מפתח"), s_h1))
    story.append(ColorLine(W, C_ACCENT, 1))
    story.append(Spacer(1, 6 * mm))

    ins = []
    if active and total:
        ins.append(heb(f"• {active} מתוך {total} פרוייקטים פעילים כרגע ({active*100//total}%)"))
    ins.append(heb(f"• {n_mgrs} מנהלי פרוייקטים פעילים"))
    if n_client_mgrs:
        ins.append(heb(f"• {n_client_mgrs} מנהלי לקוחות פעילים"))
    if durations:
        ins.append(heb(f"• משך ממוצע לפרוייקט: {avg_dur} ימים"))
    if top_kw[0] != "—":
        ins.append(heb(f"• מילת מפתח מובילה: {top_kw[0]} ({top_kw[1]} פרוייקטים)"))

    for i in ins:
        story.append(Paragraph(i, s_insight))
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════
    # STATUS + ACTIVE OVER TIME
    # ══════════════════════════════════════════════════════════
    story.append(Paragraph(heb("סטטוס פרוייקטים"), s_h1))
    story.append(ColorLine(W, C_ACCENT, 1))
    story.append(fig_to_image(chart_status_donut(projects), width=130 * mm))
    story.append(Spacer(1, 5 * mm))

    fig_a = chart_active_over_time(projects)
    if fig_a:
        story.append(Paragraph(heb("פרוייקטים פעילים ותחזוקה"), s_h1))
        story.append(ColorLine(W, C_ACCENT, 1))
        story.append(fig_to_image(fig_a, width=165 * mm))
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════
    # TIMELINE
    # ══════════════════════════════════════════════════════════
    fig_tl = chart_timeline(projects)
    if fig_tl:
        story.append(Paragraph(heb("ציר זמן פרוייקטים"), s_h1))
        story.append(ColorLine(W, C_ACCENT, 1))
        story.append(Spacer(1, 2 * mm))
        story.append(Paragraph(
            heb("כל שורה מייצגת פרוייקט. אורך הפס = משך בימים."), s_body))
        story.append(Spacer(1, 2 * mm))
        story.append(fig_to_image(fig_tl, width=175 * mm))
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════
    # MANAGER DASHBOARD
    # ══════════════════════════════════════════════════════════
    story.append(Paragraph(heb("דשבורד מנהלים"), s_h1))
    story.append(ColorLine(W, C_ACCENT, 1))
    story.append(Spacer(1, 5 * mm))

    mgr_data = defaultdict(list)
    for p in projects:
        if p["manager"]:
            mgr_data[p["manager"]].append(p)
        if p["manager_b"]:
            mgr_data[p["manager_b"]].append(p)
    mgr_sorted = sorted(mgr_data.keys(),
                        key=lambda m: len(mgr_data[m]), reverse=True)

    s_mn = ParagraphStyle("MN", fontName="HebB", fontSize=12,
                          alignment=TA_CENTER, textColor=white)
    s_ms = ParagraphStyle("MS", fontName="HebB", fontSize=14,
                          alignment=TA_CENTER, textColor=HexColor(C_DARK))
    s_ml = ParagraphStyle("ML", fontName="Heb", fontSize=8,
                          alignment=TA_CENTER, textColor=HexColor(C_GRAY))

    def render_people_cards(people_dict, sorted_names):
        """Renders 2-column manager-style cards for any people dict."""
        for i in range(0, len(sorted_names), 2):
            cells = []
            for j in range(2):
                if i + j >= len(sorted_names):
                    cells.append(Paragraph("", s_body))
                    continue
                mgr = sorted_names[i + j]
                ps = people_dict[mgr]
                ma = sum(1 for p in ps if p["status"] == "בפרודקשן")
                md = sum(1 for p in ps if p["status"] == "הושלם")
                mp = sum(1 for p in ps if p["status"] == "מושהה")
                clr = MANAGER_COLORS[(i + j) % len(MANAGER_COLORS)]
                cw = W / 2 - 4 * mm

                name_r = [Paragraph(heb(mgr), s_mn)]
                stat_r = [
                    Paragraph(f"<b>{len(ps)}</b>", s_ms),
                    Paragraph(f"<b>{ma}</b>", ParagraphStyle(
                        "x", fontName="HebB", fontSize=14, alignment=TA_CENTER,
                        textColor=HexColor(C_GREEN))),
                    Paragraph(f"<b>{md}</b>", ParagraphStyle(
                        "x", fontName="HebB", fontSize=14, alignment=TA_CENTER,
                        textColor=HexColor(C_ACCENT))),
                    Paragraph(f"<b>{mp}</b>", ParagraphStyle(
                        "x", fontName="HebB", fontSize=14, alignment=TA_CENTER,
                        textColor=HexColor(C_ORANGE))),
                ]
                lbl_r = [
                    Paragraph(heb("סה״כ"), s_ml),
                    Paragraph(heb("פעיל"), s_ml),
                    Paragraph(heb("הושלם"), s_ml),
                    Paragraph(heb("מושהה"), s_ml),
                ]
                inner = Table([name_r, stat_r, lbl_r], colWidths=[cw / 4] * 4)
                inner.setStyle(TableStyle([
                    ('SPAN', (0, 0), (3, 0)),
                    ('BACKGROUND', (0, 0), (3, 0), HexColor(clr)),
                    ('TEXTCOLOR', (0, 0), (3, 0), white),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('TOPPADDING', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
                    ('TOPPADDING', (0, 1), (-1, 1), 10),
                    ('BOTTOMPADDING', (0, 2), (-1, 2), 8),
                    ('BOX', (0, 0), (-1, -1), 0.5, HexColor('#E2E8F0')),
                    ('BACKGROUND', (0, 1), (-1, -1), HexColor('#F7FAFC')),
                ]))
                cells.append(inner)

            outer = Table([cells], colWidths=[W / 2] * 2)
            outer.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('LEFTPADDING', (0, 0), (-1, -1), 2 * mm),
                ('RIGHTPADDING', (0, 0), (-1, -1), 2 * mm),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4 * mm),
            ]))
            story.append(outer)

    render_people_cards(mgr_data, mgr_sorted)

    story.append(Spacer(1, 4 * mm))
    fig_ms = chart_manager_stacked(projects)
    if fig_ms:
        story.append(fig_to_image(fig_ms, width=165 * mm))
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════
    # CLIENT MANAGERS DASHBOARD
    # ══════════════════════════════════════════════════════════
    cm_data = defaultdict(list)
    for p in projects:
        cm = (p.get("client_manager") or "").strip()
        if cm:
            cm_data[cm].append(p)

    if cm_data:
        story.append(Paragraph(heb("דשבורד מנהלי לקוחות"), s_h1))
        story.append(ColorLine(W, C_ACCENT, 1))
        story.append(Spacer(1, 5 * mm))

        cm_sorted = sorted(cm_data.keys(),
                           key=lambda m: len(cm_data[m]), reverse=True)
        render_people_cards(cm_data, cm_sorted)

        story.append(Spacer(1, 4 * mm))
        fig_cm = chart_client_managers_stacked(projects)
        if fig_cm:
            story.append(fig_to_image(fig_cm, width=165 * mm))
        story.append(PageBreak())

    # ══════════════════════════════════════════════════════════
    # DEVELOPERS
    # ══════════════════════════════════════════════════════════
    fig_dev = chart_developers(projects)
    if fig_dev:
        story.append(Paragraph(heb("פרוייקטים לפי מפתח"), s_h1))
        story.append(ColorLine(W, C_ACCENT, 1))
        story.append(Spacer(1, 5 * mm))
        story.append(fig_to_image(fig_dev, width=175 * mm))
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════
    # MONTHLY + DURATION
    # ══════════════════════════════════════════════════════════
    fig_mo = chart_monthly_starts(projects)
    if fig_mo:
        story.append(Paragraph(heb("קצב פתיחת פרוייקטים חדשים"), s_h1))
        story.append(ColorLine(W, C_ACCENT, 1))
        story.append(fig_to_image(fig_mo, width=160 * mm))
        story.append(Spacer(1, 8 * mm))

    fig_du = chart_duration_distribution(projects)
    if fig_du:
        story.append(Paragraph(heb("התפלגות משך פרוייקטים"), s_h1))
        story.append(ColorLine(W, C_ACCENT, 1))
        story.append(Spacer(1, 2 * mm))
        story.append(Paragraph(
            heb("ירוק = עד חצי שנה, כתום = חצי שנה עד שנה, אדום = מעל שנה"),
            s_body))
        story.append(fig_to_image(fig_du, width=160 * mm))
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════
    # KEYWORDS
    # ══════════════════════════════════════════════════════════
    fig_kw = chart_keywords(projects)
    if fig_kw:
        story.append(Paragraph(heb("מילות מפתח נפוצות"), s_h1))
        story.append(ColorLine(W, C_ACCENT, 1))
        story.append(Spacer(1, 5 * mm))
        story.append(fig_to_image(fig_kw, width=160 * mm))

    doc.build(story, canvasmaker=NumberedCanvas)
    print(f"✅ PDF saved: {OUTPUT_PDF}")


# ═══════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════
if __name__ == "__main__":
    projects = read_data()
    projects = build_similarity(projects)
    save_json(projects)
    build_pdf(projects)