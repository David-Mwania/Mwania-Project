#!/usr/bin/env python3
"""
Conference-grade scientific presentation for:
Zhang et al. (2026) "Functional dissection of conserved RNAi pathways in
antiviral immunity of several hemipteran insects" PLoS Pathogens 22(10): e1014753

Design: 16:9, consistent grid, section dividers, key-takeaway callouts,
uncropped figures fitted to bounding boxes with full source citations.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image as PILImage

BASE = "/home/davidmwania/Mwania Project"
FIGS = os.path.join(BASE, "paper_figures")
OUT = os.path.join(BASE, "RNAi_Hemiptera_Presentation.pptx")

# ---------------- Design system ----------------
NAVY   = RGBColor(15, 40, 71)
BLUE   = RGBColor(23, 84, 138)
STEEL  = RGBColor(66, 129, 181)
SKY    = RGBColor(219, 233, 244)
BG     = RGBColor(252, 253, 254)
WHITE  = RGBColor(255, 255, 255)
INK    = RGBColor(38, 42, 48)
GREY   = RGBColor(110, 118, 128)
AMBER  = RGBColor(224, 138, 43)
GREEN  = RGBColor(34, 139, 84)
RED    = RGBColor(192, 57, 43)

SW, SH = Inches(13.333), Inches(7.5)
MARGIN = Inches(0.55)
FONT = "Calibri"

ARTICLE = ("Zhang et al., Functional dissection of conserved RNAi pathways in "
           "antiviral immunity of several hemipteran insects, PLoS Pathogens, 2026")

def slide_bg(s, color=BG):
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    r.fill.solid(); r.fill.fore_color.rgb = color
    r.line.fill.background(); r.shadow.inherit = False
    s.shapes._spTree.remove(r._element); s.shapes._spTree.insert(2, r._element)

def txt(s, text, l, t, w, h, size=18, bold=False, color=INK, align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.TOP, spacing=1.12, italic=False):
    box = s.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    lines = text.split("\n") if isinstance(text, str) else text
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = ln; p.alignment = align; p.line_spacing = spacing
        f = p.font; f.name = FONT; f.size = Pt(size); f.bold = bold
        f.color.rgb = color; f.italic = italic
    return box

def bullets(s, items, l, t, w, h, size=17, gap=8, color=INK):
    box = s.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        indent = item.startswith("  ")
        p.text = ("– " + item.strip()) if indent else ("•  " + item if not item.startswith("•") else item)
        p.line_spacing = 1.1; p.space_after = Pt(gap)
        if indent:
            p.level = 1
        f = p.font; f.name = FONT; f.size = Pt(size)
        f.color.rgb = GREY if indent else color
        f.bold = not indent and item.endswith(":")
    return box

def header(s, title, kicker=None):
    band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, Inches(1.06))
    band.fill.solid(); band.fill.fore_color.rgb = NAVY
    band.line.fill.background(); band.shadow.inherit = False
    acc = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(1.06), SW, Inches(0.045))
    acc.fill.solid(); acc.fill.fore_color.rgb = AMBER
    acc.line.fill.background(); acc.shadow.inherit = False
    if kicker:
        txt(s, kicker.upper(), MARGIN, Inches(0.12), Inches(12.2), Inches(0.3),
            size=11, bold=True, color=RGBColor(158, 192, 219))
        txt(s, title, MARGIN, Inches(0.38), Inches(12.2), Inches(0.62),
            size=27, bold=True, color=WHITE)
    else:
        txt(s, title, MARGIN, Inches(0.22), Inches(12.2), Inches(0.72),
            size=28, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)

def footer(s, n):
    txt(s, ARTICLE, MARGIN, Inches(7.13), Inches(10.5), Inches(0.28),
        size=9.5, color=GREY)
    txt(s, str(n), Inches(12.35), Inches(7.13), Inches(0.45), Inches(0.28),
        size=10, color=GREY, align=PP_ALIGN.RIGHT)

def takeaway(s, message, l=MARGIN, t=Inches(6.28), w=Inches(12.23), h=Inches(0.72)):
    card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    card.fill.solid(); card.fill.fore_color.rgb = SKY
    card.line.color.rgb = STEEL; card.line.width = Pt(1.2); card.shadow.inherit = False
    tf = card.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.2); tf.margin_right = Inches(0.2)
    p = tf.paragraphs[0]
    r1 = p.add_run(); r1.text = "KEY TAKEAWAY   "
    r1.font.name = FONT; r1.font.size = Pt(13); r1.font.bold = True; r1.font.color.rgb = BLUE
    r2 = p.add_run(); r2.text = message
    r2.font.name = FONT; r2.font.size = Pt(14.5); r2.font.color.rgb = INK

def fit_picture(s, path, box_l, box_t, box_w, box_h):
    """Place image inside box, aspect preserved, centered. No cropping/stretching."""
    if not os.path.exists(path):
        return None
    w_px, h_px = PILImage.open(path).size
    ar = h_px / w_px
    bw, bh = box_w, box_h
    if bw * ar <= bh:
        w = bw; h = bw * ar
    else:
        h = bh; w = h / ar
    l = box_l + int((bw - w) / 2)
    t = box_t + int((bh - h) / 2)
    return s.shapes.add_picture(path, l, t, width=w, height=h)

def fig_panel(s, path, box_l, box_t, box_w, box_h, caption):
    """Framed figure panel + caption strip below."""
    frame = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, box_l, box_t, box_w, box_h)
    frame.fill.solid(); frame.fill.fore_color.rgb = WHITE
    frame.line.color.rgb = RGBColor(203, 213, 222); frame.line.width = Pt(1)
    frame.shadow.inherit = False
    fit_picture(s, path, box_l + Inches(0.08), box_t + Inches(0.08),
                box_w - Inches(0.16), box_h - Inches(0.5))
    cap_h = Inches(0.38)
    txt(s, caption, box_l + Inches(0.08), box_t + box_h - cap_h,
        box_w - Inches(0.16), cap_h, size=9.5, color=GREY, italic=True,
        align=PP_ALIGN.CENTER, spacing=1.0)

def divider(s, number, title, subtitle):
    slide_bg(s, NAVY)
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.1), Inches(3.32), Inches(0.09), Inches(1.15))
    bar.fill.solid(); bar.fill.fore_color.rgb = AMBER
    bar.line.fill.background(); bar.shadow.inherit = False
    txt(s, f"{number}", Inches(1.45), Inches(2.55), Inches(2), Inches(0.8),
        size=22, bold=True, color=RGBColor(158, 192, 219))
    txt(s, title, Inches(1.45), Inches(3.15), Inches(10.8), Inches(0.95),
        size=40, bold=True, color=WHITE)
    txt(s, subtitle, Inches(1.45), Inches(4.15), Inches(10.5), Inches(0.7),
        size=17, color=RGBColor(190, 207, 224))

def card(s, l, t, w, h, fill=WHITE, line=RGBColor(203, 213, 222)):
    c = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    c.fill.solid(); c.fill.fore_color.rgb = fill
    c.line.color.rgb = line; c.line.width = Pt(1); c.shadow.inherit = False
    return c

# ---------------- Build ----------------
prs = Presentation()
prs.slide_width, prs.slide_height = SW, SH
BLANK = prs.slide_layouts[6]
S = lambda: prs.slides.add_slide(BLANK)
n = 0

# ---- 1. TITLE ----
s = S(); slide_bg(s, NAVY)
band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(5.05), SW, Inches(0.05))
band.fill.solid(); band.fill.fore_color.rgb = AMBER; band.line.fill.background(); band.shadow.inherit = False
txt(s, "RESEARCH ARTICLE  |  RNA INTERFERENCE  |  HEMIPTERA", Inches(1.2), Inches(1.35),
    Inches(11), Inches(0.4), size=13, bold=True, color=AMBER, align=PP_ALIGN.CENTER)
txt(s, "Functional Dissection of Conserved RNAi Pathways in\nAntiviral Immunity of Several Hemipteran Insects",
    Inches(1.0), Inches(1.95), Inches(11.33), Inches(1.9), size=36, bold=True,
    color=WHITE, align=PP_ALIGN.CENTER, spacing=1.1)
txt(s, "Yan Zhang† · Zhuang-Xin Ye† · Qing-Lu Ma† · Bi-Yun Wang · Ji-Chong Zhuo · Hai-Jian Huang · "
       "Chuan-Xi Zhang · Jian-Ping Chen* · Jun-Min Li*",
    Inches(1.0), Inches(4.05), Inches(11.33), Inches(0.75), size=15, color=RGBColor(205, 220, 235),
    align=PP_ALIGN.CENTER, spacing=1.2)
txt(s, "State Key Laboratory for Quality and Safety of Agro-Products · Institute of Plant Virology · Ningbo University, China",
    Inches(1.0), Inches(4.6), Inches(11.33), Inches(0.35), size=13, color=RGBColor(160, 182, 204),
    align=PP_ALIGN.CENTER)
txt(s, "PLoS Pathogens 2026; 22(10): e1014753   ·   DOI: 10.1371/journal.ppat.1014753",
    Inches(1.0), Inches(5.35), Inches(11.33), Inches(0.4), size=14, bold=True,
    color=RGBColor(205, 220, 235), align=PP_ALIGN.CENTER)
txt(s, "Presented by  [Your Name]", Inches(1.0), Inches(6.15), Inches(11.33), Inches(0.4),
    size=16, color=WHITE, align=PP_ALIGN.CENTER)
n += 1

# ---- 2. OUTLINE ----
s = S(); slide_bg(s); header(s, "Presentation Outline")
items = [
    ("01", "Background & Rationale", "Small RNA pathways; the knowledge gap in Hemiptera"),
    ("02", "Study Design & Methods", "RNAi screen, sRNA-seq, CRISPR/Cas9 in four species"),
    ("03", "Results", "Five sequential observations linking genes to phenotypes"),
    ("04", "Discussion & Conclusions", "Evidence integration, limitations, implications"),
    ("05", "Take-Home Message", "What was asked, done, found, and why it matters"),
]
y = Inches(1.55)
for num, t, d in items:
    c = card(s, Inches(1.0), y, Inches(11.3), Inches(0.92))
    chip = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.25), y + Inches(0.16), Inches(0.6), Inches(0.6))
    chip.fill.solid(); chip.fill.fore_color.rgb = NAVY; chip.line.fill.background(); chip.shadow.inherit = False
    tfc = chip.text_frame; tfc.vertical_anchor = MSO_ANCHOR.MIDDLE
    pp = tfc.paragraphs[0]; pp.text = num; pp.alignment = PP_ALIGN.CENTER
    pp.font.size = Pt(15); pp.font.bold = True; pp.font.color.rgb = WHITE; pp.font.name = FONT
    txt(s, t, Inches(2.15), y + Inches(0.12), Inches(4.3), Inches(0.4), size=19, bold=True, color=NAVY)
    txt(s, d, Inches(6.5), y + Inches(0.15), Inches(5.6), Inches(0.65), size=13.5, color=GREY)
    y += Inches(1.08)
footer(s, 2)

# ---- 3. DIVIDER: BACKGROUND ----
s = S(); divider(s, "01", "Background & Rationale", "What is known — and what remained unexplored")
n += 1

# ---- 4. BACKGROUND ----
s = S(); slide_bg(s); header(s, "Background: Small Silencing RNA Pathways", "SECTION 01 · BACKGROUND")
cols = [
    ("miRNA pathway", "~22 nt · AGO1 / DCR1", "Gene regulation,\ndevelopment, metamorphosis", STEEL),
    ("piRNA pathway", "26–30 nt · Piwi / AGO3", "Transposon silencing,\ngermline integrity", BLUE),
    ("siRNA pathway", "21–22 nt · AGO2 / DCR2 / R2D2", "Primary antiviral\ndefense (canonical)", NAVY),
]
x = Inches(0.55)
for title, comp, role, col in cols:
    c = card(s, x, Inches(1.5), Inches(3.95), Inches(2.5))
    hdr = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, Inches(1.5), Inches(3.95), Inches(0.62))
    hdr.fill.solid(); hdr.fill.fore_color.rgb = col; hdr.line.fill.background(); hdr.shadow.inherit = False
    txt(s, title, x + Inches(0.15), Inches(1.6), Inches(3.65), Inches(0.42),
        size=18, bold=True, color=WHITE)
    txt(s, comp, x + Inches(0.2), Inches(2.3), Inches(3.55), Inches(0.4), size=13.5, bold=True, color=BLUE)
    txt(s, role, x + Inches(0.2), Inches(2.8), Inches(3.55), Inches(1.0), size=15, color=INK)
    x += Inches(4.14)
bullets(s, [
    "Well characterized in Diptera: Drosophila and mosquitoes",
    "AGO2, DCR2, or R2D2 loss increases viral susceptibility in flies",
    "Hemiptera: major rice-pest vectors — RNAi roles unresolved",
], Inches(0.55), Inches(4.35), Inches(12.2), Inches(1.7), size=17, gap=8)
takeaway(s, "The three pathways are conserved, but their antiviral roles in non-model Hemiptera were unknown.")
footer(s, 4)

# ---- 5. RESEARCH QUESTION ----
s = S(); slide_bg(s); header(s, "Research Question & Objectives", "SECTION 01 · BACKGROUND")
c = card(s, Inches(0.55), Inches(1.5), Inches(12.23), Inches(1.55), fill=SKY, line=STEEL)
txt(s, "RESEARCH QUESTION", Inches(0.85), Inches(1.7), Inches(4), Inches(0.3),
    size=12, bold=True, color=BLUE)
txt(s, "Are the conserved miRNA, piRNA, and siRNA pathways functionally active in hemipteran insects,\n"
       "and which of them mediate antiviral immunity?",
    Inches(0.85), Inches(2.05), Inches(11.6), Inches(0.9), size=19, bold=True, color=NAVY, spacing=1.15)
objs = [
    ("Objective 1", "Identify and annotate core small-RNA pathway genes in three rice planthoppers"),
    ("Objective 2", "Functionally dissect each pathway by dsRNA knockdown in L. striatellus"),
    ("Objective 3", "Test antiviral roles against four insect-specific viruses (ISVs) and RSV"),
    ("Objective 4", "Validate key findings genetically (CRISPR) and across three more hemipteran species"),
]
y = Inches(3.4)
for t, d in objs:
    chip = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), y, Inches(1.7), Inches(0.55))
    chip.fill.solid(); chip.fill.fore_color.rgb = BLUE; chip.line.fill.background(); chip.shadow.inherit = False
    tf = chip.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.text = t; p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = WHITE; p.font.name = FONT
    txt(s, d, Inches(2.5), y + Inches(0.08), Inches(10.2), Inches(0.45), size=16, color=INK)
    y += Inches(0.78)
footer(s, 5)

# ---- 6. DIVIDER: METHODS ----
s = S(); divider(s, "02", "Study Design & Methods", "A complementary toolkit: knockdown, sequencing, and gene editing")
n += 1

# ---- 7. STUDY DESIGN ----
s = S(); slide_bg(s); header(s, "Study Design", "SECTION 02 · METHODS")
steps = [
    ("STEP 1", "Gene discovery", "Homology search of 7 core RNAi genes in 3 planthoppers; phylogeny + domain analysis"),
    ("STEP 2", "Functional screen", "dsRNA knockdown (AGO1/2, DCR1/2, R2D2, Piwi, AGO3) vs ds-GFP control"),
    ("STEP 3", "Phenotyping", "Survival, fecundity, hatch rate, developmental duration; tissue expression"),
    ("STEP 4", "Antiviral assay", "Viral RNA (RT-qPCR) and vsiRNAs (sRNA-seq) for 4 ISVs + RSV"),
    ("STEP 5", "Genetic validation", "CRISPR/Cas9 DCR2 knockout; RNAi-of-RNAi reporter assay"),
    ("STEP 6", "Cross-species test", "AGO2/DCR2 knockdown in S. furcifera, N. lugens, B. tabaci"),
]
y = Inches(1.5)
for t1, t2, d in steps:
    chip = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), y, Inches(1.25), Inches(0.72))
    chip.fill.solid(); chip.fill.fore_color.rgb = NAVY; chip.line.fill.background(); chip.shadow.inherit = False
    tf = chip.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.text = t1; p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(11.5); p.font.bold = True; p.font.color.rgb = WHITE; p.font.name = FONT
    txt(s, t2, Inches(2.0), y + Inches(0.03), Inches(2.6), Inches(0.7),
        size=15.5, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, d, Inches(4.7), y + Inches(0.03), Inches(8.05), Inches(0.7),
        size=14, color=INK, anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(0.9)
footer(s, 7)

# ---- 8. METHODS ----
s = S(); slide_bg(s); header(s, "Methods Overview", "SECTION 02 · METHODS")
mcols = [
    ("Functional genomics", [
        "dsRNA microinjection: 500 ng/µL into hemocoel",
        "Knockdown efficiency confirmed by RT-qPCR",
        "RNAi-of-RNAi: ds-Tubulin reporter re-injection",
        "CRISPR/Cas9: 2 sgRNAs, 2 DCR2 mutant alleles",
    ]),
    ("Virology & sequencing", [
        "4 ISVs (>60% prevalence) + RSV quantified",
        "Absolute qPCR to quantify viral copy numbers",
        "sRNA-seq: Illumina HiSeq; vsiRNA mapping (Bowtie)",
        "vsiRNAs normalized to viral copy number (reads/copies)",
    ]),
    ("Statistics", [
        "Student's t-test for pairwise comparisons",
        "One-way ANOVA + Tukey for multiple groups",
        "Log-rank (Mantel-Cox) for survival curves",
        "3 biological replicates (qPCR); 2 (sRNA-seq)",
    ]),
]
x = Inches(0.55)
for title, pts in mcols:
    c = card(s, x, Inches(1.5), Inches(3.95), Inches(4.5))
    hdr = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, Inches(1.5), Inches(3.95), Inches(0.58))
    hdr.fill.solid(); hdr.fill.fore_color.rgb = BLUE; hdr.line.fill.background(); hdr.shadow.inherit = False
    txt(s, title, x + Inches(0.15), Inches(1.58), Inches(3.65), Inches(0.42),
        size=16, bold=True, color=WHITE)
    bl = ["  " + p for p in pts]
    bullets(s, bl, x + Inches(0.2), Inches(2.3), Inches(3.6), Inches(3.5), size=13.5, gap=12)
    x += Inches(4.14)
takeaway(s, " Knockdown + sequencing + genetics provide orthogonal evidence for each pathway's role.")
footer(s, 8)

# ---- 9. DIVIDER: RESULTS ----
s = S(); divider(s, "03", "Results", "Five sequential observations — each one shaping the next question")
n += 1

# ---- 10. RESULT 1: Fig 1 ----
s = S(); slide_bg(s); header(s, "Result 1 — Core RNAi Genes Are Conserved in Planthoppers",
                             "SECTION 03 · RESULTS")
txt(s, "Central question of this slide", MARGIN, Inches(1.28), Inches(6), Inches(0.3),
    size=11, bold=True, color=AMBER)
txt(s, "Do hemipteran genomes encode intact, conserved RNAi machinery?",
    MARGIN, Inches(1.55), Inches(6.4), Inches(0.75), size=17, bold=True, color=NAVY, spacing=1.1)
bullets(s, [
    "21 core small-RNA pathway genes identified in 3 planthoppers",
    "11 genes newly identified (Ls and Sf orthologs)",
    "Planthopper genes form distinct clades with conserved domains",
    "Ls_DCR1 lacks the DEAD/DEAH helicase_C domain",
    "Ls_R2D2 carries an extra dsRNA-binding motif",
], Inches(0.55), Inches(2.55), Inches(6.4), Inches(3.4), size=15.5, gap=12)
fig_panel(s, os.path.join(FIGS, "page_05_img_01.png"), Inches(7.15), Inches(1.28), Inches(5.6), Inches(4.85),
          "Source: Figure 1 from " + ARTICLE)
takeaway(s, "Conserved, intact RNAi machinery — setting the stage for functional testing.")
footer(s, 10)

# ---- 11. RESULT 2: Fig 2 ----
s = S(); slide_bg(s); header(s, "Result 2 — miRNA & piRNA Genes Are Vital for Host Biology",
                             "SECTION 03 · RESULTS")
txt(s, "What happens when each pathway is knocked down?",
    MARGIN, Inches(1.28), Inches(6.4), Inches(0.45), size=17, bold=True, color=NAVY)
bullets(s, [
    "AGO1/DCR1 knockdown: lower survival, molting defects",
    "Fecundity collapses; eggs fail to hatch",
    "Piwi/AGO3 knockdown: wing curling, fewer eggs",
    "Developmental duration significantly extended",
    "No effect of AGO2/DCR2/R2D2 on survival or fecundity",
], Inches(0.55), Inches(1.95), Inches(6.4), Inches(4.1), size=15.5, gap=13)
fig_panel(s, os.path.join(FIGS, "page_06_img_01.png"), Inches(7.15), Inches(1.28), Inches(5.6), Inches(4.85),
          "Source: Figure 2 from " + ARTICLE)
takeaway(s, "miRNA and piRNA pathways are essential for development and reproduction — not obviously antiviral.")
footer(s, 11)

# ---- 12. RESULT 3: Fig 3 ----
s = S(); slide_bg(s); header(s, "Result 3 — Knockdowns Functionally Deplete Their Small RNAs",
                             "SECTION 03 · RESULTS")
bullets(s, [
    "AGO1/DCR1 knockdown: 22-nt miRNA abundance drops",
    "Knockdown efficiency: 82% (AGO1), 70% (DCR1)",
    "Piwi knockdown: 26–30 nt piRNA peak strongly reduced",
    "EVE-derived piRNAs (LsETLVEs) fall with Piwi loss",
    "Confirms knockdowns truly impair pathway activity",
], Inches(0.55), Inches(1.4), Inches(6.4), Inches(4.6), size=15.5, gap=13)
fig_panel(s, os.path.join(FIGS, "page_08_img_01.png"), Inches(7.15), Inches(1.28), Inches(5.6), Inches(4.85),
          "Source: Figure 3 from " + ARTICLE)
takeaway(s, "The knockdown system works: pathway output (small RNAs) falls as expected.")
footer(s, 12)

# ---- 13. BRIDGE ----
s = S(); slide_bg(s); header(s, "The Logical Pivot", "FROM OBSERVATION TO THE NEXT EXPERIMENT")
box_defs = [
    ("OBSERVATION", "miRNA & piRNA knockdowns disrupt host biology\nbut do not increase viral accumulation", SKY, NAVY),
    ("QUESTION RAISED", "If miRNA/piRNA are not antiviral here,\ndoes the canonical siRNA pathway protect these insects?", RGBColor(255, 243, 224), RGBColor(150, 90, 20)),
    ("ANALYSIS PERFORMED", "Knockdown AGO2, DCR2, R2D2 in L. striatellus\n→ measure viral RNA + vsiRNAs for 4 ISVs and RSV", RGBColor(226, 239, 226), GREEN),
]
y = Inches(1.55)
for title, body, fill, tcol in box_defs:
    c = card(s, Inches(1.6), y, Inches(10.1), Inches(1.35), fill=fill,
             line=STEEL if fill == SKY else (AMBER if tcol != GREEN else GREEN))
    txt(s, title, Inches(1.9), y + Inches(0.14), Inches(9.5), Inches(0.32),
        size=12.5, bold=True, color=tcol)
    txt(s, body, Inches(1.9), y + Inches(0.48), Inches(9.5), Inches(0.8),
        size=16, color=INK, spacing=1.12)
    if title != "ANALYSIS PERFORMED":
        ar = s.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(6.42), y + Inches(1.4), Inches(0.45), Inches(0.28))
        ar.fill.solid(); ar.fill.fore_color.rgb = AMBER; ar.line.fill.background(); ar.shadow.inherit = False
    y += Inches(1.78)
footer(s, 13)

# ---- 14. RESULT 4: Fig 4 ----
s = S(); slide_bg(s); header(s, "Result 4 — Only AGO2 Depletion Raises Viral Load",
                             "SECTION 03 · RESULTS")
bullets(s, [
    "AGO2 knockdown: all 4 ISVs accumulate significantly",
    "Same phenotype for RSV (plant virus transmitted by vector)",
    "DCR2 and R2D2 knockdown: no viral increase",
    "21–22 nt vsiRNA size profile preserved after DCR2 loss",
    "Double/triple knockdowns add no effect beyond AGO2",
    "miRNA and piRNA genes show no antiviral role either",
], Inches(0.55), Inches(1.4), Inches(6.4), Inches(4.7), size=15, gap=11)
fig_panel(s, os.path.join(FIGS, "page_10_img_01.png"), Inches(7.15), Inches(1.28), Inches(5.6), Inches(4.85),
          "Source: Figure 4 from " + ARTICLE)
takeaway(s, "AGO2 is the antiviral effector — yet DCR2 loss leaves antiviral silencing seemingly intact.")
footer(s, 14)

# ---- 15. RESULT 4b: controls ----
s = S(); slide_bg(s); header(s, "Ruling Out Residual Silencing", "SECTION 03 · RESULTS · SUPPORTING CONTROLS")
lcards = [
    ("Normalization control", "vsiRNA reads divided by viral copy number:\nratio falls only after AGO2 knockdown (S4 Fig)"),
    ("RNAi-of-RNAi reporter", "AGO2 pre-knockdown cuts ds-Tubulin silencing by 76.5% (p < 0.01);\nDCR1/DCR2 pre-knockdown: no significant effect (S9 Fig)"),
    ("Compensation check", "DCR1+DCR2 double and DCR2+AGO2+R2D2 triple knockdowns:\nno viral increase unless AGO2 is targeted (S7 Fig)"),
]
y = Inches(1.5)
for t, d in lcards:
    c = card(s, Inches(0.55), y, Inches(12.23), Inches(1.45))
    tag = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.55), y, Inches(0.14), Inches(1.45))
    tag.fill.solid(); tag.fill.fore_color.rgb = STEEL; tag.line.fill.background(); tag.shadow.inherit = False
    txt(s, t, Inches(0.95), y + Inches(0.14), Inches(3.3), Inches(1.15),
        size=16, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, d, Inches(4.45), y + Inches(0.12), Inches(8.1), Inches(1.2),
        size=14.5, color=INK, anchor=MSO_ANCHOR.MIDDLE, spacing=1.15)
    y += Inches(1.63)
takeaway(s, "Multiple orthogonal controls converge: DCR2 depletion does not measurably weaken antiviral RNAi.")
footer(s, 15)

# ---- 16. RESULT 5: Fig 5 CRISPR ----
s = S(); slide_bg(s); header(s, "Result 5 — CRISPR Knockout: DCR2 Is Essential for Embryos",
                             "SECTION 03 · RESULTS")
bullets(s, [
    "CRISPR bypasses limits of RNAi-based knockdown",
    "Two independent DCR2 mutant alleles generated",
    "No homozygous adults recovered in G2 generation",
    "Unhatched eggs genotyped: homozygous mutants",
    "17.9–29.4% of eggs = homozygous → embryonic lethality",
    "Heterozygotes: less DCR2, delayed development, no viral change",
], Inches(0.55), Inches(1.4), Inches(6.4), Inches(4.7), size=15, gap=11)
fig_panel(s, os.path.join(FIGS, "page_12_img_01.png"), Inches(7.15), Inches(1.28), Inches(5.6), Inches(4.85),
          "Source: Figure 5 from " + ARTICLE)
takeaway(s, "DCR2 is essential for embryonic viability — explaining why antiviral roles are hard to prove genetically.")
footer(s, 16)

# ---- 17. RESULT 6: Fig 6 cross-species ----
s = S(); slide_bg(s); header(s, "Result 6 — The AGO2/DCR2 Contrast Is Conserved Across Hemiptera",
                             "SECTION 03 · RESULTS")
bullets(s, [
    "Sogatella furcifera: 2 ISVs — only AGO2 knockdown ↑ viral load",
    "Nilaparvata lugens: 4 ISVs — same AGO2-only phenotype",
    "Bemisia tabaci (different family): 3 ISVs — same result",
    "DCR2 knockdown: no detectable effect in any species",
    "Two families, four species — one consistent pattern",
], Inches(0.55), Inches(1.4), Inches(6.4), Inches(4.7), size=15, gap=12)
fig_panel(s, os.path.join(FIGS, "page_13_img_01.png"), Inches(7.15), Inches(1.28), Inches(5.6), Inches(4.85),
          "Source: Figure 6 from " + ARTICLE)
takeaway(s, "AGO2-dependent, DCR2-independent antiviral silencing is a shared hemipteran feature.")
footer(s, 17)

# ---- 18. DIVIDER: DISCUSSION ----
s = S(); divider(s, "04", "Discussion & Conclusions", "Integrating the evidence — and what remains unresolved")
n += 1

# ---- 19. FLOW DIAGRAM ----
s = S(); slide_bg(s); header(s, "How the Findings Were Connected", "SECTION 04 · DISCUSSION")
flow = [
    ("Research question", "Which RNAi pathways defend hemipterans against viruses?", NAVY),
    ("Data collected", "Knockdowns + sRNA-seq + CRISPR across 4 species", STEEL),
    ("Observation 1", "miRNA/piRNA govern development — no antiviral effect", STEEL),
    ("Analysis 1", "Target the siRNA pathway: AGO2, DCR2, R2D2", STEEL),
    ("Observation 2", "Only AGO2 depletion increases viral accumulation", STEEL),
    ("Analysis 2", "CRISPR DCR2 knockout + RNAi-of-RNAi + normalization", STEEL),
    ("Conclusion", "AGO2 is antiviral; DCR2 is essential for embryonic development", NAVY),
]
y = Inches(1.35); bh = Inches(0.62); gap = 0.815
for i, (t, d, col) in enumerate(flow):
    b = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), y, Inches(10.9), bh)
    b.fill.solid(); b.fill.fore_color.rgb = col; b.line.fill.background(); b.shadow.inherit = False
    tfb = b.text_frame; tfb.vertical_anchor = MSO_ANCHOR.MIDDLE; tfb.word_wrap = True
    p = tfb.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r1 = p.add_run(); r1.text = t.upper() + "   "
    r1.font.name = FONT; r1.font.size = Pt(13); r1.font.bold = True; r1.font.color.rgb = AMBER
    r2 = p.add_run(); r2.text = d
    r2.font.name = FONT; r2.font.size = Pt(15); r2.font.bold = True; r2.font.color.rgb = WHITE
    if i < len(flow) - 1:
        ar = s.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(6.44), y + bh + Inches(0.015),
                                Inches(0.42), Inches(0.15))
        ar.fill.solid(); ar.fill.fore_color.rgb = AMBER; ar.line.fill.background(); ar.shadow.inherit = False
    y += Inches(gap)
footer(s, 19)

# ---- 20. Q/EVIDENCE TABLE ----
s = S(); slide_bg(s); header(s, "Evidence Mapping: Question → Data → Conclusion", "SECTION 04 · DISCUSSION")
rows, cols_n = 5, 3
tbl = s.shapes.add_table(rows, cols_n, Inches(0.55), Inches(1.4), Inches(12.23), Inches(4.6)).table
tbl.columns[0].width = Inches(3.35); tbl.columns[1].width = Inches(5.1); tbl.columns[2].width = Inches(3.78)
data = [
    ["Research Question", "Evidence Found", "Conclusion"],
    ["Do miRNA/piRNA act antivirally?",
     "Knockdown → developmental defects only; viral load unchanged",
     "Developmental roles; not antiviral"],
    ["Is the siRNA pathway antiviral?",
     "AGO2 knockdown ↑ viral load across ISVs and RSV",
     "Yes — AGO2 is the key effector"],
    ["Is DCR2 required for antiviral RNAi?",
     "vsiRNA profiles intact after DCR2 loss; heterozygous mutants show no viral increase",
     "Not required under tested conditions"],
    ["Conserved across Hemiptera?",
     "Identical AGO2/DCR2 contrast in S. furcifera, N. lugens, B. tabaci",
     "Conserved in all species examined"],
]
for i, row in enumerate(data):
    for j, val in enumerate(row):
        cell = tbl.cell(i, j)
        cell.text = val
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = Inches(0.12); cell.margin_right = Inches(0.12)
        cell.margin_top = Inches(0.04); cell.margin_bottom = Inches(0.04)
        for p in cell.text_frame.paragraphs:
            for r in p.runs:
                r.font.name = FONT
                r.font.size = Pt(13 if i else 14)
                r.font.bold = (i == 0) or (j == 2 and i > 0)
                r.font.color.rgb = WHITE if i == 0 else (NAVY if j == 2 else INK)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY if i == 0 else (WHITE if i % 2 else SKY)
takeaway(s, "Every observation fed the next analysis — the evidence chain fully answers the study question.")
footer(s, 20)

# ---- 21. STRENGTHS & LIMITATIONS ----
s = S(); slide_bg(s); header(s, "Strengths & Limitations", "SECTION 04 · DISCUSSION")
c = card(s, Inches(0.55), Inches(1.45), Inches(5.95), Inches(4.55))
hdr = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.55), Inches(1.45), Inches(5.95), Inches(0.58))
hdr.fill.solid(); hdr.fill.fore_color.rgb = GREEN; hdr.line.fill.background(); hdr.shadow.inherit = False
txt(s, "Strengths", Inches(0.75), Inches(1.53), Inches(5.5), Inches(0.42), size=17, bold=True, color=WHITE)
bullets(s, ["  Four hemipteran species, two families",
            "  RNAi + CRISPR + sequencing triangulation",
            "  vsiRNAs normalized to viral copy number",
            "  Public data, scripts, and accessions"], 
        Inches(0.8), Inches(2.25), Inches(5.45), Inches(3.5), size=14.5, gap=13)
c = card(s, Inches(6.83), Inches(1.45), Inches(5.95), Inches(4.55))
hdr = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.83), Inches(1.45), Inches(5.95), Inches(0.58))
hdr.fill.solid(); hdr.fill.fore_color.rgb = RED; hdr.line.fill.background(); hdr.shadow.inherit = False
txt(s, "Limitations", Inches(7.03), Inches(1.53), Inches(5.5), Inches(0.42), size=17, bold=True, color=WHITE)
bullets(s, ["  Residual DCR2 activity cannot be excluded",
            "  Homozygous lethality blocks adult antiviral tests",
            "  Alternative vsiRNA nuclease not identified",
            "  Laboratory conditions; field relevance untested"],
        Inches(7.08), Inches(2.25), Inches(5.45), Inches(3.5), size=14.5, gap=13)
takeaway(s, "Strong convergent evidence — tempered by unavoidable limits of DCR2 genetics in this system.")
footer(s, 21)

# ---- 22. CONCLUSIONS & FUTURE ----
s = S(); slide_bg(s); header(s, "Conclusions & Future Directions", "SECTION 04 · DISCUSSION")
txt(s, "What the study demonstrated", MARGIN, Inches(1.35), Inches(6), Inches(0.4),
    size=17, bold=True, color=NAVY)
bullets(s, [
    "Dipteran DCR2-centered model does not transfer to Hemiptera",
    "AGO2 is the conserved antiviral effector",
    "DCR2 is developmentally essential (embryonic lethality)",
    "Antiviral vsiRNAs arise despite DCR2 knockdown",
], Inches(0.55), Inches(1.8), Inches(6.1), Inches(3.2), size=15, gap=12)
txt(s, "Future directions", Inches(6.95), Inches(1.35), Inches(6), Inches(0.4),
    size=17, bold=True, color=NAVY)
bullets(s, [
    "Identify the nuclease generating antiviral vsiRNAs",
    "Conditional DCR2 knockout to bypass embryonic lethality",
    "Test broader hemipteran taxonomy and field conditions",
    "Explore RNAi-based vector and crop-virus control",
], Inches(6.95), Inches(1.8), Inches(5.8), Inches(3.2), size=15, gap=12)
c = card(s, Inches(0.55), Inches(5.35), Inches(12.23), Inches(1.15), fill=SKY, line=STEEL)
txt(s, "PUBLIC HEALTH & AGRICULTURAL IMPLICATION", Inches(0.85), Inches(5.5), Inches(6), Inches(0.3),
    size=11.5, bold=True, color=BLUE)
txt(s, "AGO2-centered antiviral defense could be exploited to limit persistent viral loads in rice planthoppers "
       "and whiteflies — reducing transmission of economically devastating plant viruses such as RSV.",
    Inches(0.85), Inches(5.82), Inches(11.6), Inches(0.6), size=14.5, color=INK, spacing=1.15)
footer(s, 22)

# ---- 23. TAKE-HOME ----
s = S(); slide_bg(s, NAVY)
txt(s, "TAKE-HOME MESSAGE", Inches(1.0), Inches(0.75), Inches(11.33), Inches(0.5),
    size=20, bold=True, color=AMBER, align=PP_ALIGN.CENTER)
qa = [
    ("What did they ask?", "Which conserved RNAi pathways defend hemipteran insects against viruses?"),
    ("What did they do?", "Functional knockdown, small-RNA sequencing, and CRISPR knockout across four hemipteran species."),
    ("What did they find?", "AGO2 is the antiviral effector; DCR2 is essential for embryonic development but dispensable for the measured antiviral response."),
    ("Why does it matter?", "It revises the canonical DCR2-centered model of insect antiviral RNAi and points to lineage-specific pathway organization."),
]
y = Inches(1.55)
for q, a in qa:
    c = card(s, Inches(0.9), y, Inches(11.53), Inches(1.22), fill=WHITE, line=STEEL)
    txt(s, q, Inches(1.2), y + Inches(0.12), Inches(3.2), Inches(1.0),
        size=16, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, a, Inches(4.6), y + Inches(0.12), Inches(7.6), Inches(1.0),
        size=14.5, color=INK, anchor=MSO_ANCHOR.MIDDLE, spacing=1.15)
    y += Inches(1.38)
footer(s, 23)

# ---- 24. REFERENCES ----
s = S(); slide_bg(s); header(s, "References & Data Availability")
txt(s, "Primary article", MARGIN, Inches(1.4), Inches(4), Inches(0.35), size=15, bold=True, color=NAVY)
txt(s, "Zhang Y, Ye Z-X, Ma Q-L, Wang B-Y, Zhuo J-C, Huang H-J, Zhang C-X, Chen J-P, Li J-M. "
       "Functional dissection of conserved RNAi pathways in antiviral immunity of several hemipteran insects. "
       "PLoS Pathogens. 2026;22(10):e1014753. doi:10.1371/journal.ppat.1014753",
    MARGIN, Inches(1.75), Inches(12.23), Inches(0.95), size=14, color=INK, spacing=1.2)
txt(s, "Data & code", MARGIN, Inches(2.95), Inches(4), Inches(0.35), size=15, bold=True, color=NAVY)
txt(s, "Sequencing data: NCBI BioProject PRJNA1405795   ·   Numerical data: S1 Data (article supplement)\n"
       "Small-RNA analysis scripts: github.com/Zhuangxinye/vsirna-pipeline",
    MARGIN, Inches(3.3), Inches(12.23), Inches(0.7), size=13.5, color=INK, spacing=1.25)
txt(s, "Key referenced findings cited within this presentation", MARGIN, Inches(4.15), Inches(9), Inches(0.35),
    size=15, bold=True, color=NAVY)
txt(s, "Figures 1–6 and Supplementary Figures S1–S11 of the primary article, reproduced and adapted under the "
       "Creative Commons Attribution License (CC BY 4.0). Source of every visual: Zhang et al., PLoS Pathogens, 2026.",
    MARGIN, Inches(4.5), Inches(12.23), Inches(0.8), size=13.5, color=INK, spacing=1.25)
c = card(s, MARGIN, Inches(5.55), Inches(12.23), Inches(1.1), fill=SKY, line=STEEL)
txt(s, "All figures in this deck:  \"Source: Figure X from Zhang et al., Functional dissection of conserved RNAi "
       "pathways in antiviral immunity of several hemipteran insects, PLoS Pathogens, 2026.\"",
    Inches(0.85), Inches(5.72), Inches(11.6), Inches(0.8), size=13, italic=True, color=NAVY,
    anchor=MSO_ANCHOR.MIDDLE, spacing=1.15)
footer(s, 24)

prs.save(OUT)
print(f"Saved {OUT} with {len(prs.slides.__iter__.__self__._sldIdLst)} slides")
