#!/usr/bin/env python3
"""
gh_slides — Transforme l'export JSON de gh_bench.py en présentation PowerPoint (.pptx).

Usage :
    python gh_bench.py "http client" -l python --json data.json
    python gh_slides.py data.json -o rapport.pptx

Dépendance : python-pptx  (pip install python-pptx)
"""
from __future__ import annotations

import argparse
import json
import re
from datetime import datetime

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

# --------------------------------------------------------------------------- #
# Thème
# --------------------------------------------------------------------------- #
BG = RGBColor(0x0D, 0x11, 0x17)
CARD = RGBColor(0x16, 0x1B, 0x22)
BORDER = RGBColor(0x30, 0x36, 0x3D)
TEXT = RGBColor(0xE6, 0xED, 0xF3)
MUTED = RGBColor(0x8B, 0x94, 0x9E)
ACCENT = RGBColor(0x58, 0xA6, 0xFF)
GREEN = RGBColor(0x3F, 0xB9, 0x50)
YELLOW = RGBColor(0xD2, 0x99, 0x22)
RED = RGBColor(0xF8, 0x51, 0x49)
FONT = "Calibri"

W, H = Inches(13.333), Inches(7.5)
MARGIN = Inches(0.6)

CRITERIA = {"popularite": "Popularité", "activite": "Activité", "communaute": "Communauté",
            "maintenance": "Maintenance", "licence": "Licence"}


def verdict_info(v: str) -> tuple[str, RGBColor]:
    if "Recommandé" in v:
        return "Recommandé", GREEN
    if "Acceptable" in v:
        return "Acceptable", YELLOW
    return "À éviter", RED


def fmt_int(n) -> str:
    return "n/d" if n is None else f"{n:,}".replace(",", " ")


def short(name: str) -> str:
    return name.split("/")[-1]


# --------------------------------------------------------------------------- #
# Primitives
# --------------------------------------------------------------------------- #
def new_slide(prs: Presentation):
    s = prs.slides.add_slide(prs.slide_layouts[6])  # vierge
    bg = s.background.fill
    bg.solid()
    bg.fore_color.rgb = BG
    return s


def text(slide, x, y, w, h, value, size=14, color=TEXT, bold=False, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, wrap=True):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    for side in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, side, 0)
    lines = value if isinstance(value, list) else [value]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        runs = line if isinstance(line, list) else [(line, {})]
        for chunk, style in runs:
            r = p.add_run()
            r.text = chunk
            f = r.font
            f.name = FONT
            f.size = Pt(style.get("size", size))
            f.bold = style.get("bold", bold)
            f.color.rgb = style.get("color", color)
        if i:
            p.space_before = Pt(6)
    return tb


def box(slide, x, y, w, h, fill=CARD, line=BORDER, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06):
    sh = slide.shapes.add_shape(shape, x, y, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(0.75)
    sh.shadow.inherit = False
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        sh.adjustments[0] = radius
    return sh


def pill(slide, x, y, label, color, size=11):
    w = Inches(0.12 * len(label) + 0.35)
    sh = box(slide, x, y, w, Inches(0.32), fill=color, line=None, radius=0.5)
    tf = sh.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    r.font.name, r.font.size, r.font.bold, r.font.color.rgb = FONT, Pt(size), True, BG
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    return w


def header(slide, kicker, title):
    text(slide, MARGIN, Inches(0.45), Inches(10), Inches(0.3), kicker.upper(), 12, ACCENT, True)
    text(slide, MARGIN, Inches(0.75), W - 2 * MARGIN, Inches(0.7), title, 30, TEXT, True)


def footer(slide, data, n):
    q = f"Benchmark « {data['query']} »" + (f" · {data['language']}" if data.get("language") else "")
    text(slide, MARGIN, H - Inches(0.45), Inches(8), Inches(0.25), q, 10, MUTED)
    text(slide, W - MARGIN - Inches(1), H - Inches(0.45), Inches(1), Inches(0.25), str(n), 10,
         MUTED, align=PP_ALIGN.RIGHT)


def stat(slide, x, y, w, value, label):
    box(slide, x, y, w, Inches(1.05))
    text(slide, x + Inches(0.2), y + Inches(0.15), w - Inches(0.4), Inches(0.5), value, 24, TEXT, True)
    text(slide, x + Inches(0.2), y + Inches(0.65), w - Inches(0.4), Inches(0.3), label, 11, MUTED)


# --------------------------------------------------------------------------- #
# Slides
# --------------------------------------------------------------------------- #
def slide_title(prs, data):
    s = new_slide(prs)
    box(s, 0, 0, Inches(0.18), H, fill=ACCENT, line=None, shape=MSO_SHAPE.RECTANGLE)
    gen = datetime.fromisoformat(data["generated"])
    text(s, Inches(1.1), Inches(2.0), Inches(10), Inches(0.4), "BENCHMARK DE LIBRAIRIES GITHUB",
         14, ACCENT, True)
    text(s, Inches(1.1), Inches(2.5), Inches(11), Inches(1.4), f"« {data['query']} »", 54, TEXT, True)
    sub = f"{len(data['repos'])} dépôts analysés"
    if data.get("language"):
        sub += f" · langage {data['language']}"
    text(s, Inches(1.1), Inches(3.9), Inches(11), Inches(0.5), sub, 20, MUTED)
    text(s, Inches(1.1), Inches(6.3), Inches(11), Inches(0.4),
         f"Généré le {gen.strftime('%d/%m/%Y')} par gh-bench-bot", 12, MUTED)


PROP_RE = re.compile(r"\*\*(.+?) : \[(.+?)\]\((.+?)\)\*\*\s*[—-]?\s*(.*)")


def slide_proposals(prs, data, n):
    s = new_slide(prs)
    header(s, "Synthèse", "Nos recommandations")
    props = []
    for p in data["proposals"]:
        m = PROP_RE.match(p)
        if m:
            label, name, _, rest = m.groups()
            props.append((label, name, rest.strip(" ,:").replace("**", "")))
        else:
            props.append(("Remarque", "", re.sub(r"[*`]", "", p)))
    props = props[:4]
    cols = len(props)
    gap = Inches(0.3)
    cw = int((W - 2 * MARGIN - gap * (cols - 1)) / cols)
    y, ch = Inches(1.9), Inches(4.6)
    for i, (label, name, rest) in enumerate(props):
        x = MARGIN + i * (cw + gap)
        color = RED if label.startswith("Piège") else GREEN if i == 0 else ACCENT
        box(s, x, y, cw, ch)
        box(s, x, y, cw, Inches(0.08), fill=color, line=None, shape=MSO_SHAPE.RECTANGLE)
        text(s, x + Inches(0.3), y + Inches(0.4), cw - Inches(0.6), Inches(0.4), label.upper(), 12,
             color, True)
        text(s, x + Inches(0.3), y + Inches(0.85), cw - Inches(0.6), Inches(1.2),
             short(name) if name else "", 26, TEXT, True)
        if name:
            text(s, x + Inches(0.3), y + Inches(1.55), cw - Inches(0.6), Inches(0.3), name, 11, MUTED)
        rest = rest[0].upper() + rest[1:] if rest else ""
        text(s, x + Inches(0.3), y + Inches(2.1), cw - Inches(0.6), Inches(2.3), rest, 14, TEXT)
    footer(s, data, n)


def slide_ranking(prs, data, n):
    s = new_slide(prs)
    header(s, "Classement", "Score global sur 100")
    repos = data["repos"][:10]
    cd = CategoryChartData()
    cd.categories = [short(r["name"]) for r in reversed(repos)]
    cd.add_series("Score", [r["total"] for r in reversed(repos)])
    gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, MARGIN, Inches(1.7), Inches(8.6),
                            Inches(5.2), cd)
    ch = gf.chart
    ch.has_legend = False
    ch.has_title = False
    ch.font.name, ch.font.size, ch.font.color.rgb = FONT, Pt(12), TEXT
    plot = ch.plots[0]
    plot.gap_width = 45
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.font.size, dl.font.bold, dl.font.color.rgb = Pt(12), True, TEXT
    dl.number_format, dl.number_format_is_linked = "0", False
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    va = ch.value_axis
    va.minimum_scale, va.maximum_scale = 0, 100
    va.visible = False
    va.has_major_gridlines = False
    ca = ch.category_axis
    ca.format.line.fill.background()
    ca.tick_labels.font.size, ca.tick_labels.font.color.rgb = Pt(13), TEXT
    ser = plot.series[0]
    for i, r in enumerate(reversed(repos)):
        pt = ser.points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = verdict_info(r["verdict"])[1]

    # Légende + lecture
    x = Inches(9.6)
    y = Inches(2.0)
    for label, color in (("Recommandé (≥ 75)", GREEN), ("Acceptable (≥ 55)", YELLOW),
                         ("À éviter", RED)):
        box(s, x, y + Inches(0.05), Inches(0.22), Inches(0.22), fill=color, line=None, radius=0.2)
        text(s, x + Inches(0.35), y, Inches(3), Inches(0.3), label, 13)
        y += Inches(0.45)
    counts = {k: sum(1 for r in data["repos"] if verdict_info(r["verdict"])[0] == k)
              for k in ("Recommandé", "Acceptable", "À éviter")}
    box(s, x, Inches(3.75), Inches(3.1), Inches(2.6))
    text(s, x + Inches(0.25), Inches(3.95), Inches(2.6), Inches(2.3), [
        [(str(counts["Recommandé"]), {"size": 40, "bold": True, "color": GREEN})],
        [("dépôt(s) recommandé(s) sur " + str(len(data["repos"])), {"size": 13, "color": MUTED})],
        [(f"{counts['À éviter']} à éviter", {"size": 16, "bold": True, "color": RED})],
    ])
    footer(s, data, n)


def slide_table(prs, data, n):
    s = new_slide(prs)
    header(s, "Détail", "Tableau comparatif")
    repos = data["repos"][:10]
    cols = ["#", "Dépôt", "Score", "Verdict", "Stars", "Contrib.", "Commits 90 j", "Dernier push",
            "Licence"]
    widths = [0.5, 3.0, 0.9, 1.5, 1.2, 1.1, 1.4, 1.3, 1.2]
    rows = len(repos) + 1
    tbl = s.shapes.add_table(rows, len(cols), MARGIN, Inches(1.7), Inches(sum(widths)),
                             Inches(0.42 * rows)).table
    for j, w in enumerate(widths):
        tbl.columns[j].width = Inches(w)

    def cell(i, j, value, color=TEXT, bold=False, fill=CARD, align=PP_ALIGN.LEFT):
        c = tbl.cell(i, j)
        c.fill.solid()
        c.fill.fore_color.rgb = fill
        c.margin_left = c.margin_right = Inches(0.08)
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = c.text_frame
        tf.paragraphs[0].text = ""
        p = tf.paragraphs[0]
        p.alignment = align
        r = p.add_run()
        r.text = str(value)
        r.font.name, r.font.size, r.font.bold, r.font.color.rgb = FONT, Pt(12), bold, color

    for j, c in enumerate(cols):
        cell(0, j, c, MUTED, True, BG, PP_ALIGN.RIGHT if j in (2, 4, 5, 6) else PP_ALIGN.LEFT)
    for i, r in enumerate(repos, 1):
        label, color = verdict_info(r["verdict"])
        fill = CARD if i % 2 else BG
        R = PP_ALIGN.RIGHT
        cell(i, 0, i, MUTED, fill=fill)
        cell(i, 1, r["name"], TEXT, True, fill)
        cell(i, 2, f"{r['total']:.0f}", color, True, fill, R)
        cell(i, 3, label, color, True, fill)
        cell(i, 4, fmt_int(r["stars"]), fill=fill, align=R)
        cell(i, 5, fmt_int(r["contributors"]), fill=fill, align=R)
        cell(i, 6, fmt_int(r["commits_90d"]), fill=fill, align=R)
        push = r["days_since_push"]
        cell(i, 7, f"il y a {push} j", RED if push > 365 else TEXT, fill=fill)
        cell(i, 8, r["license"] or "—", fill=fill)
    footer(s, data, n)


def slide_repo(prs, data, r, rank, n):
    s = new_slide(prs)
    label, color = verdict_info(r["verdict"])
    text(s, MARGIN, Inches(0.45), Inches(10), Inches(0.3), f"N° {rank} DU CLASSEMENT", 12, ACCENT, True)
    text(s, MARGIN, Inches(0.75), Inches(9), Inches(0.7), r["name"], 30, TEXT, True)
    if r.get("description"):
        text(s, MARGIN, Inches(1.45), Inches(9.5), Inches(0.5), r["description"][:140], 14, MUTED)
    # Score
    text(s, W - MARGIN - Inches(2.5), Inches(0.55), Inches(2.5), Inches(0.9),
         [[(f"{r['total']:.0f}", {"size": 54, "bold": True, "color": color}),
           (" /100", {"size": 18, "color": MUTED})]], align=PP_ALIGN.RIGHT)
    pw = Inches(0.12 * len(label) + 0.35)
    pill(s, W - MARGIN - pw, Inches(1.55), label, color)

    # Stats
    y = Inches(2.2)
    sw = Inches(2.85)
    gap = Inches(0.2)
    for i, (v, l) in enumerate([(fmt_int(r["stars"]), "stars"),
                                (fmt_int(r["contributors"]), "contributeurs"),
                                (fmt_int(r["commits_90d"]), "commits sur 90 jours"),
                                (f"{r['days_since_push']} j", "depuis le dernier push")]):
        stat(s, MARGIN + i * (sw + gap), y, sw, v, l)

    # Barres de critères
    y0 = Inches(3.6)
    box(s, MARGIN, y0, Inches(5.9), Inches(3.3))
    text(s, MARGIN + Inches(0.3), y0 + Inches(0.2), Inches(5), Inches(0.3), "DÉTAIL DU SCORE", 11,
         MUTED, True)
    bx, bw = MARGIN + Inches(2.0), Inches(3.0)
    for i, (k, lab) in enumerate(CRITERIA.items()):
        v = r["scores"].get(k, 0)
        yy = y0 + Inches(0.65) + i * Inches(0.5)
        text(s, MARGIN + Inches(0.3), yy, Inches(1.7), Inches(0.3), lab, 13)
        box(s, bx, yy + Inches(0.05), bw, Inches(0.2), fill=BORDER, line=None, radius=0.5)
        if v > 0.01:
            box(s, bx, yy + Inches(0.05), Emu(int(bw * v)), Inches(0.2), fill=ACCENT, line=None,
                radius=0.5)
        text(s, bx + bw + Inches(0.15), yy, Inches(0.5), Inches(0.3), f"{v * 100:.0f}", 13, TEXT,
             True)

    # Forces / vigilance
    x1 = MARGIN + Inches(6.1)
    w1 = W - MARGIN - x1
    box(s, x1, y0, w1, Inches(3.3))
    lines = [[("POINTS FORTS", {"size": 11, "bold": True, "color": GREEN})]]
    lines += [[("+  " + t, {"size": 13})] for t in r["strengths"][:4]] or [[("—", {"color": MUTED})]]
    lines += [[("POINTS DE VIGILANCE", {"size": 11, "bold": True, "color": YELLOW})]]
    lines += [[("!  " + t, {"size": 13})] for t in r["alerts"][:3]] or [[("Aucun", {"color": MUTED})]]
    text(s, x1 + Inches(0.3), y0 + Inches(0.2), w1 - Inches(0.6), Inches(3.0), lines)
    footer(s, data, n)


def slide_avoid(prs, data, n):
    bad = [r for r in data["repos"] if verdict_info(r["verdict"])[0] == "À éviter"]
    if not bad:
        return False
    s = new_slide(prs)
    header(s, "Attention", "Dépôts à éviter")
    y = Inches(1.8)
    for r in bad[:5]:
        box(s, MARGIN, y, W - 2 * MARGIN, Inches(0.9))
        box(s, MARGIN, y, Inches(0.08), Inches(0.9), fill=RED, line=None, shape=MSO_SHAPE.RECTANGLE)
        text(s, MARGIN + Inches(0.35), y + Inches(0.15), Inches(4), Inches(0.35), r["name"], 16,
             TEXT, True)
        text(s, MARGIN + Inches(0.35), y + Inches(0.5), Inches(4), Inches(0.3),
             f"{fmt_int(r['stars'])} stars · score {r['total']:.0f}/100", 12, MUTED)
        text(s, MARGIN + Inches(4.6), y, Inches(7.3), Inches(0.9),
             r["alerts"][0] if r["alerts"] else "Score trop faible.", 14, TEXT,
             anchor=MSO_ANCHOR.MIDDLE)
        y += Inches(1.05)
    footer(s, data, n)
    return True


def slide_method(prs, data, n):
    s = new_slide(prs)
    header(s, "Méthode", "Comment le score est calculé")
    desc = {
        "popularite": "Nombre de stars, sur une échelle logarithmique.",
        "activite": "Fraîcheur du dernier push et commits des 90 derniers jours.",
        "communaute": "Nombre de contributeurs.",
        "maintenance": "Fraîcheur de la dernière release et issues ouvertes par rapport aux stars.",
        "licence": "Permissive > copyleft > autre > aucune.",
    }
    y = Inches(1.9)
    for k, lab in CRITERIA.items():
        w = data["weights"][k]
        box(s, MARGIN, y, W - 2 * MARGIN, Inches(0.8))
        text(s, MARGIN + Inches(0.3), y, Inches(1.2), Inches(0.8), f"{w}", 28, ACCENT, True,
             anchor=MSO_ANCHOR.MIDDLE)
        text(s, MARGIN + Inches(1.2), y, Inches(2.5), Inches(0.8), lab, 18, TEXT, True,
             anchor=MSO_ANCHOR.MIDDLE)
        text(s, MARGIN + Inches(3.8), y, Inches(8), Inches(0.8), desc[k], 14, MUTED,
             anchor=MSO_ANCHOR.MIDDLE)
        y += Inches(0.92)
    text(s, MARGIN, y + Inches(0.1), W - 2 * MARGIN, Inches(0.4),
         "Verdict : 75 et plus = recommandé · 55 et plus = acceptable · en dessous = à éviter.", 13,
         MUTED)
    footer(s, data, n)


# --------------------------------------------------------------------------- #
def build(data: dict, out: str, details: int = 3) -> None:
    prs = Presentation()
    prs.slide_width, prs.slide_height = W, H
    slide_title(prs, data)
    n = 2
    slide_proposals(prs, data, n); n += 1
    slide_ranking(prs, data, n); n += 1
    slide_table(prs, data, n); n += 1
    for i, r in enumerate(data["repos"][:details], 1):
        slide_repo(prs, data, r, i, n); n += 1
    if slide_avoid(prs, data, n):
        n += 1
    slide_method(prs, data, n)
    prs.save(out)


def main() -> None:
    p = argparse.ArgumentParser(description="Génère une présentation PowerPoint depuis gh_bench.py")
    p.add_argument("json", help="Fichier JSON produit par gh_bench.py --json")
    p.add_argument("-o", "--out", default="rapport.pptx")
    p.add_argument("--details", type=int, default=3, help="Nombre de fiches détaillées (défaut 3)")
    a = p.parse_args()
    with open(a.json, encoding="utf-8") as f:
        build(json.load(f), a.out, a.details)
    print(f"✅ Présentation écrite dans {a.out}")


if __name__ == "__main__":
    main()
