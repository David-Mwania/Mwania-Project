#!/usr/bin/env python3
"""
Professional PowerPoint presentation for:
"Functional dissection of conserved RNAi pathways in antiviral immunity of several hemipteran insects"
Clean layout, properly cropped figures, consistent design.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image as PILImage

BASE_DIR = "/home/davidmwania/Mwania Project"
FIGURES_DIR = os.path.join(BASE_DIR, "paper_figures")
OUTPUT = os.path.join(BASE_DIR, "RNAi_Hemiptera_Presentation.pptx")

# Professional color palette
NAVY = RGBColor(16, 42, 67)        # Deep navy for headers
BLUE = RGBColor(31, 78, 121)       # Medium blue
STEEL = RGBColor(70, 130, 180)     # Steel blue accent
LIGHT = RGBColor(245, 247, 250)    # Light background
WHITE = RGBColor(255, 255, 255)
DARK_TEXT = RGBColor(33, 37, 41)
MUTED = RGBColor(108, 117, 125)
ORANGE = RGBColor(230, 126, 34)
GREEN = RGBColor(40, 167, 69)
RED = RGBColor(220, 53, 69)

def add_bg(slide, color=LIGHT):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    shape.fill.solid(); shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    slide.shapes._spTree.remove(shape._element)
    slide.shapes._spTree.insert(2, shape._element)

def add_header(slide, title, subtitle=None):
    # Top band
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(1.15))
    band.fill.solid(); band.fill.fore_color.rgb = NAVY
    band.line.fill.background()
    band.shadow.inherit = False
    
    # Accent line under band
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(1.15), Inches(13.333), Inches(0.06))
    line.fill.solid(); line.fill.fore_color.rgb = ORANGE
    line.line.fill.background()
    
    # Title text
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.18), Inches(12.1), Inches(0.85))
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = title
    p.font.size = Pt(28); p.font.bold = True
    p.font.color.rgb = WHITE; p.font.name = "Calibri"
    
    if subtitle:
        stb = slide.shapes.add_textbox(Inches(0.6), Inches(0.75), Inches(12.1), Inches(0.35))
        stf = stb.text_frame; stf.word_wrap = True
        sp = stf.paragraphs[0]; sp.text = subtitle
        sp.font.size = Pt(13); sp.font.color.rgb = RGBColor(200, 215, 230); sp.font.name = "Calibri"

def add_footer(slide, page_num):
    ftb = slide.shapes.add_textbox(Inches(0.6), Inches(7.08), Inches(6), Inches(0.3))
    ftf = ftb.text_frame; fp = ftf.paragraphs[0]
    fp.text = "Zhang et al., 2026 | PLoS Pathogens 22(10): e1014753"
    fp.font.size = Pt(10); fp.font.color.rgb = MUTED; fp.font.name = "Calibri"
    
    ptb = slide.shapes.add_textbox(Inches(11.5), Inches(7.08), Inches(1.2), Inches(0.3))
    ptf = ptb.text_frame; pp = ptf.paragraphs[0]
    pp.text = f"{page_num}"; pp.font.size = Pt(10); pp.font.color.rgb = MUTED
    pp.alignment = PP_ALIGN.RIGHT; pp.font.name = "Calibri"

def add_bullets(slide, items, left, top, width, height, font_size=18, gap=8, title=None, title_color=NAVY):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame; tf.word_wrap = True
    tf.margin_left = Inches(0.05); tf.margin_right = Inches(0.05)
    first = True
    if title:
        p = tf.paragraphs[0]; first = False
        p.text = title; p.font.size = Pt(font_size + 2); p.font.bold = True
        p.font.color.rgb = title_color; p.space_after = Pt(gap + 4); p.font.name = "Calibri"
    for item in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.text = "• " + item if not item.startswith(("•", "–", "–")) else item
        p.font.size = Pt(font_size); p.font.color.rgb = DARK_TEXT
        p.font.name = "Calibri"; p.space_after = Pt(gap)
        p.line_spacing = 1.15
    return box

def crop_image_for_slide(img_path, max_w_in, max_h_in):
    """Crop and resize an image to fit within max_w_in x max_h_in, removing page margins."""
    if not os.path.exists(img_path):
        return None, None
    img = PILImage.open(img_path)
    w, h = img.size
    
    # Crop: remove page header/footer margins (approx top 12%, bottom 8%)
    # and side margins (approx 6% each side) to isolate the figure
    crop_top = int(h * 0.12)
    crop_bottom = int(h * 0.92)
    crop_left = int(w * 0.06)
    crop_right = int(w * 0.94)
    
    if crop_bottom > crop_top and crop_right > crop_left:
        img = img.crop((crop_left, crop_top, crop_right, crop_bottom))
        w, h = img.size
    
    # Save cropped version
    cropped_path = img_path.replace(".png", "_crop.png")
    img.save(cropped_path)
    
    aspect = h / w
    max_h_px = max_h_in * 96  # approximate
    max_w_px = max_w_in * 96
    
    if w > max_w_px or h > max_h_px:
        if w / h > max_w_px / max_h_px:
            new_w = max_w_px
            new_h = int(new_w * aspect)
        else:
            new_h = max_h_px
            new_w = int(new_h / aspect)
        img = img.resize((new_w, new_h), PILImage.LANCZOS)
        img.save(cropped_path)
    
    return cropped_path, (w / h)

def add_fig(slide, img_path, left, top, width=None, height=None):
    """Add an image centered in the given area."""
    if not os.path.exists(img_path):
        return None
    cropped, aspect = crop_image_for_slide(img_path, width or 5, height or 5)
    if cropped is None:
        return None
    if width and not height:
        h = width * aspect
        return slide.shapes.add_picture(cropped, left, top, width=width, height=h)
    elif height and not width:
        w = height / aspect
        return slide.shapes.add_picture(cropped, left, top, width=w, height=height)
    else:
        return slide.shapes.add_picture(cropped, left, top, width=width, height=height)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

# ============================================================
# SLIDE 1: TITLE
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s, NAVY)
# Decorative bands
band1 = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(2.9), Inches(13.333), Inches(0.08))
band1.fill.solid(); band1.fill.fore_color.rgb = ORANGE; band1.line.fill.background()

tb = s.shapes.add_textbox(Inches(0.8), Inches(1.7), Inches(11.7), Inches(1.6))
tf = tb.text_frame; tf.word_wrap = True
p = tf.paragraphs[0]; p.text = "Functional dissection of conserved RNAi pathways in antiviral immunity of several hemipteran insects"
p.font.size = Pt(34); p.font.bold = True; p.font.color.rgb = WHITE
p.alignment = PP_ALIGN.CENTER; p.font.name = "Calibri"

sub = s.shapes.add_textbox(Inches(0.8), Inches(3.4), Inches(11.7), Inches(2.2))
tf = sub.text_frame; tf.word_wrap = True
lines = [
    "Yan Zhang, Zhuang-Xin Ye, Qing-Lu Ma, Bi-Yun Wang, Ji-Chong Zhuo, Hai-Jian Huang,",
    "Chuan-Xi Zhang, Jian-Ping Chen, Jun-Min Li",
    "",
    "PLoS Pathogens, 2026, 22(10): e1014753",
    "DOI: 10.1371/journal.ppat.1014753",
    "",
    "Presented by: [Your Name]",
]
for i, ln in enumerate(lines):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    p.text = ln
    p.font.size = Pt(16 if i in (0,1) else 14)
    p.font.bold = (i == 0)
    p.font.color.rgb = RGBColor(210, 220, 230) if i not in (0,1) else WHITE
    p.alignment = PP_ALIGN.CENTER; p.font.name = "Calibri"
    p.space_after = Pt(6)

# ============================================================
# SLIDE 2: BACKGROUND
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s)
add_header(s, "Background: RNA Silencing Pathways in Insects",
           "Why study RNAi in hemipteran insects?")
add_bullets(s, [
    "RNA interference (RNAi) is a sequence-specific gene-silencing mechanism found in insects.",
    "Three conserved small silencing RNA pathways:",
    "   • miRNA pathway — gene regulation, development, reproduction",
    "   • piRNA pathway — transposon suppression, germline maintenance",
    "   • siRNA pathway — canonical antiviral defense",
    "Most functional data come from Dipteran models (Drosophila, mosquitoes).",
    "Hemipteran pests (planthoppers, whiteflies) transmit economically important plant viruses,",
    "but their RNAi biology in antiviral immunity is poorly understood.",
], Inches(0.7), Inches(1.5), Inches(12.0), Inches(5.2), font_size=19)
add_footer(s, 2)

# ============================================================
# SLIDE 3: STUDY QUESTION
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s)
add_header(s, "Study Question & Objectives",
           "What did the researchers set out to answer?")
add_bullets(s, [
    "Main research question:",
    "   What are the functional roles of the conserved miRNA, piRNA, and siRNA pathways",
    "   in hemipteran antiviral immunity?",
    "",
    "Specific objectives:",
    "   1. Functionally screen core RNAi genes in Laodelphax striatellus (small brown planthopper).",
    "   2. Determine which pathway(s) restrict viral accumulation.",
    "   3. Test whether the findings are conserved across multiple hemipteran species.",
], Inches(0.7), Inches(1.5), Inches(12.0), Inches(5.2), font_size=19)
add_footer(s, 3)

# ============================================================
# SLIDE 4: STUDY DESIGN
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s)
add_header(s, "Study Design",
           "Experimental system and species tested")
add_bullets(s, [
    "Type: Laboratory functional genomics study (RNAi knockdown + CRISPR/Cas9 knockout).",
    "Primary model: Laodelphax striatellus (small brown planthopper).",
    "Cross-species validation in:",
    "   • Sogatella furcifera (rice planthopper)",
    "   • Nilaparvata lugens (brown planthopper)",
    "   • Bemisia tabaci (whitefly)",
    "Treatment: dsRNA injection targeting RNAi genes vs. ds-GFP negative control.",
    "Readouts: survival, fecundity, development, viral RNA load, small RNA profiles.",
], Inches(0.7), Inches(1.5), Inches(12.0), Inches(5.2), font_size=19)
add_footer(s, 4)

# ============================================================
# SLIDE 5: METHODS
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s)
add_header(s, "Methods Overview",
           "Core experimental approaches")
add_bullets(s, [
    "RNAi knockdown: dsRNA microinjection of AGO1/2, DCR1/2, R2D2, Piwi, AGO3.",
    "RT-qPCR: gene expression and viral RNA quantification (normalized to ACTIN).",
    "Small RNA sequencing: profile miRNA, piRNA, and virus-derived siRNAs (vsiRNAs).",
    "RNAi-of-RNAi assay: test whether pre-knockdown impairs reporter silencing.",
    "CRISPR/Cas9: generate DCR2 mutant alleles to test developmental requirement.",
    "Statistics: Student's t-test, one-way ANOVA (Tukey), Log-rank (Mantel-Cox).",
], Inches(0.7), Inches(1.5), Inches(12.0), Inches(5.2), font_size=19)
add_footer(s, 5)

# ============================================================
# SLIDE 6: KEY OBSERVATION 1 + FIGURE 2
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s)
add_header(s, "Key Observation 1: miRNA & piRNA Pathways Control Development",
           "Knockdown phenotypes in L. striatellus (Figure 2)")
# Left: bullets
add_bullets(s, [
    "miRNA pathway (AGO1, DCR1):",
    "   • Reduced survival and molting defects",
    "   • Severely reduced fecundity; eggs failed to hatch",
    "   • 22-nt miRNA production dropped",
    "piRNA pathway (Piwi, AGO3):",
    "   • Abnormal wing curls and reduced egg production",
    "   • Extended developmental duration",
    "Interpretation: essential for development/reproduction, not major antiviral players.",
], Inches(0.6), Inches(1.45), Inches(6.2), Inches(5.4), font_size=17, gap=10)
# Right: Figure 2
fig2 = os.path.join(FIGURES_DIR, "page_06_img_01.png")
if os.path.exists(fig2):
    add_fig(s, fig2, left=Inches(7.1), top=Inches(1.5), width=Inches(5.7))
add_footer(s, 6)

# ============================================================
# SLIDE 7: LOGICAL BRIDGE
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s)
add_header(s, "From Observation 1 to the Next Question",
           "The logical chain that guided the study")
add_bullets(s, [
    "Observation 1: miRNA and piRNA knockdowns affected development and reproduction",
    "   but did not increase viral accumulation.",
    "",
    "Question raised:",
    "   If miRNA/piRNA are not the main antiviral pathways, what is the role of",
    "   the canonical siRNA pathway?",
    "",
    "Next analysis: Knockdown the three core siRNA genes:",
    "   • AGO2 — antiviral effector nuclease",
    "   • DCR2 — cleaves viral dsRNA into siRNAs",
    "   • R2D2 — loads siRNAs onto AGO2",
    "   Then measure viral RNA levels and vsiRNA production.",
], Inches(0.7), Inches(1.5), Inches(12.0), Inches(5.2), font_size=19, gap=12)
add_footer(s, 7)

# ============================================================
# SLIDE 8: KEY OBSERVATION 2 + FIGURE 4
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s)
add_header(s, "Key Observation 2: AGO2, Not DCR2, Restricts Viral Accumulation",
           "siRNA pathway knockdown effects (Figure 4)")
add_bullets(s, [
    "AGO2 knockdown:",
    "   • Significantly increased accumulation of four ISVs and RSV",
    "   • vsiRNA reads increased with viral RNA",
    "   • Normalized vsiRNA per viral copy fell — impaired antiviral RNAi",
    "DCR2 and R2D2 knockdown:",
    "   • No detectable increase in viral accumulation",
    "   • Canonical 21–22 nt vsiRNA profiles preserved",
], Inches(0.6), Inches(1.45), Inches(6.2), Inches(5.4), font_size=17, gap=10)
fig4 = os.path.join(FIGURES_DIR, "page_10_img_01.png")
if os.path.exists(fig4):
    add_fig(s, fig4, left=Inches(7.1), top=Inches(1.5), width=Inches(5.7))
add_footer(s, 8)

# ============================================================
# SLIDE 9: CRISPR VALIDATION + FIGURE 5
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s)
add_header(s, "CRISPR/Cas9 Validation: DCR2 Is Essential for Embryos",
           "Genetic knockout confirms developmental role (Figure 5)")
add_bullets(s, [
    "Why CRISPR? RNAi knockdown may leave residual DCR2 activity.",
    "Two DCR2 mutant alleles generated (5-bp deletion + 9-bp insertion; 8-bp deletion).",
    "Result: No homozygous mutants recovered in G2; unhatched eggs carried",
    "homozygous DCR2 mutations → embryonic lethality.",
    "Heterozygous mutants: reduced DCR2, no viral increase, delayed nymph development.",
    "Conclusion: DCR2 is essential for development, not required for measured",
    "antiviral RNAi response.",
], Inches(0.6), Inches(1.45), Inches(6.2), Inches(5.4), font_size=17, gap=10)
fig5 = os.path.join(FIGURES_DIR, "page_12_img_01.png")
if os.path.exists(fig5):
    add_fig(s, fig5, left=Inches(7.1), top=Inches(1.5), width=Inches(5.7))
add_footer(s, 9)

# ============================================================
# SLIDE 10: CROSS-SPECIES + FIGURE 6
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s)
add_header(s, "Cross-Species Validation: Conserved AGO2-Dependent Antiviral Role",
           "Same AGO2 / DCR2 contrast in three more species (Figure 6)")
add_bullets(s, [
    "Tested: S. furcifera, N. lugens, B. tabaci",
    "In all three species:",
    "   • AGO2 knockdown increased viral accumulation",
    "   • DCR2 knockdown had no detectable effect",
    "Supports a conserved AGO2-centric antiviral phenotype across Hemiptera.",
], Inches(0.6), Inches(1.45), Inches(6.2), Inches(5.4), font_size=17, gap=10)
fig6 = os.path.join(FIGURES_DIR, "page_13_img_01.png")
if os.path.exists(fig6):
    add_fig(s, fig6, left=Inches(7.1), top=Inches(1.5), width=Inches(5.7))
add_footer(s, 10)

# ============================================================
# SLIDE 11: FLOW DIAGRAM
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s)
add_header(s, "How the Findings Were Connected",
           "Logical flow from question to conclusion")
flow_items = [
    ("Research Question", "Roles of RNAi in hemipteran antiviral immunity?"),
    ("Data Collected", "RNAi + sRNA-seq + CRISPR in 4 species"),
    ("Observation 1", "miRNA/piRNA: development/reproduction, not antiviral"),
    ("Analysis 1", "Focus on siRNA: knockdown AGO2, DCR2, R2D2"),
    ("Observation 2", "Only AGO2 depletion increases viral load"),
    ("Analysis 2", "CRISPR/Cas9 DCR2 knockout + RNAi-of-RNAi"),
    ("Conclusion", "AGO2 is antiviral; DCR2 is essential for embryos"),
]
y = 1.45
bx_h = 0.68
bx_w = 11.6
gap = 0.86
for i, (hdr, txt) in enumerate(flow_items):
    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                             Inches(0.87), Inches(y), Inches(bx_w), Inches(bx_h))
    box.fill.solid()
    box.fill.fore_color.rgb = NAVY if (i == 0 or i == len(flow_items)-1) else STEEL
    box.line.color.rgb = WHITE
    box.shadow.inherit = False
    tbf = s.shapes.add_textbox(Inches(1.05), Inches(y + 0.04), Inches(bx_w - 0.4), Inches(bx_h - 0.08))
    tff = tbf.text_frame; tff.word_wrap = True
    pp = tff.paragraphs[0]
    pp.text = f"{hdr}: {txt}"
    pp.font.size = Pt(15); pp.font.bold = True; pp.font.color.rgb = WHITE
    pp.font.name = "Calibri"
    pp.alignment = PP_ALIGN.CENTER
    if i < len(flow_items) - 1:
        ar = s.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(6.5), Inches(y + bx_h + 0.03), Inches(0.35), Inches(0.14))
        ar.fill.solid(); ar.fill.fore_color.rgb = ORANGE
        ar.line.fill.background(); ar.shadow.inherit = False
    y += gap
add_footer(s, 11)

# ============================================================
# SLIDE 12: MAIN FINDINGS
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s)
add_header(s, "Main Findings", "Five key results from the study")
add_bullets(s, [
    "1. miRNA and piRNA pathways are essential for survival, development, and reproduction in L. striatellus.",
    "2. AGO2 is required for antiviral siRNA-mediated immunity.",
    "3. DCR2 and R2D2 knockdown did not impair measured antiviral responses or vsiRNA profiles.",
    "4. DCR2 is essential for embryonic development — homozygous CRISPR mutants are lethal.",
    "5. The AGO2-essential / DCR2-dispensable antiviral pattern is conserved across three additional",
    "   hemipteran species.",
], Inches(0.7), Inches(1.5), Inches(12.0), Inches(5.2), font_size=19, gap=14)
add_footer(s, 12)

# ============================================================
# SLIDE 13: TABLE - Q&A
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s)
add_header(s, "How the Findings Answered the Research Question",
           "Mapping evidence to conclusions")
rows, cols = 5, 3
tbl_shape = s.shapes.add_table(rows, cols, Inches(0.5), Inches(1.4), Inches(12.3), Inches(5.1))
table = tbl_shape.table
# Column widths
table.columns[0].width = Inches(3.6)
table.columns[1].width = Inches(4.8)
table.columns[2].width = Inches(3.9)
data = [
    ["Research Question", "Evidence Found", "Conclusion"],
    ["Roles of miRNA/piRNA?", "Knockdown → developmental defects, lower survival and fecundity; no antiviral effect",
     "Essential for development/reproduction, not antiviral defense"],
    ["Is siRNA pathway antiviral?", "AGO2 knockdown ↑ viral load; DCR2/R2D2 knockdown did not",
     "siRNA pathway is antiviral; AGO2 is the key effector"],
    ["Is DCR2 required for antiviral RNAi?", "vsiRNA profiles unchanged after DCR2 knockdown; CRISPR heterozygotes show no viral increase",
     "DCR2 not required for measured antiviral RNAi under tested conditions"],
    ["Conserved across Hemiptera?", "Same AGO2/DCR2 contrast in S. furcifera, N. lugens, B. tabaci",
     "Pattern conserved among examined hemipteran species"],
]
for i, row in enumerate(data):
    for j, val in enumerate(row):
        cell = table.cell(i, j)
        cell.text = val
        cell.margin_left = Inches(0.08); cell.margin_right = Inches(0.08)
        cell.margin_top = Inches(0.05); cell.margin_bottom = Inches(0.05)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        for p in cell.text_frame.paragraphs:
            for r in p.runs:
                r.font.name = "Calibri"
                r.font.size = Pt(12.5 if i > 0 else 14)
                r.font.bold = (i == 0)
                r.font.color.rgb = WHITE if i == 0 else DARK_TEXT
        if i == 0:
            cell.fill.solid(); cell.fill.fore_color.rgb = NAVY
        else:
            cell.fill.solid(); cell.fill.fore_color.rgb = WHITE if i % 2 else LIGHT
add_footer(s, 13)

# ============================================================
# SLIDE 14: STRENGTHS & LIMITATIONS
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s)
add_header(s, "Strengths and Limitations", "Critical appraisal of the study")
# Two columns
lb = s.shapes.add_textbox(Inches(0.7), Inches(1.5), Inches(5.9), Inches(5.2))
ltf = lb.text_frame; ltf.word_wrap = True
p = ltf.paragraphs[0]; p.text = "Strengths"; p.font.size = Pt(19); p.font.bold = True; p.font.color.rgb = GREEN
p.font.name = "Calibri"; p.space_after = Pt(10)
for item in ["Multi-species validation across two hemipteran families",
             "RNAi knockdown combined with CRISPR/Cas9 genetics",
             "Comprehensive small RNA sequencing",
             "RNAi-of-RNAi assay as independent functional readout"]:
    p = ltf.add_paragraph(); p.text = "• " + item
    p.font.size = Pt(16); p.font.color.rgb = DARK_TEXT; p.font.name = "Calibri"
    p.space_after = Pt(9); p.line_spacing = 1.12

rb = s.shapes.add_textbox(Inches(6.8), Inches(1.5), Inches(5.9), Inches(5.2))
rtf = rb.text_frame; rtf.word_wrap = True
p = rtf.paragraphs[0]; p.text = "Limitations"; p.font.size = Pt(19); p.font.bold = True; p.font.color.rgb = RED
p.font.name = "Calibri"; p.space_after = Pt(10)
for item in ["RNAi knockdown may leave residual DCR2 activity",
             "Homozygous DCR2 mutants died as embryos; adult antiviral phenotypes not testable",
             "Alternative vsiRNA biogenesis mechanisms not identified",
             "Lab conditions only; field relevance needs further study"]:
    p = rtf.add_paragraph(); p.text = "• " + item
    p.font.size = Pt(16); p.font.color.rgb = DARK_TEXT; p.font.name = "Calibri"
    p.space_after = Pt(9); p.line_spacing = 1.12
add_footer(s, 14)

# ============================================================
# SLIDE 15: CONCLUSIONS
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s)
add_header(s, "Conclusions", "What the study demonstrated")
add_bullets(s, [
    "The canonical DCR2-centered model of antiviral RNAi from Diptera does not fully explain hemipteran antiviral immunity.",
    "AGO2 is the conserved antiviral effector across the tested hemipteran species.",
    "DCR2 has an essential developmental role, particularly during embryogenesis.",
    "DCR2 depletion does not detectably impair the measured antiviral siRNA response in L. striatellus.",
    "Future work should identify alternative nucleases or mechanisms for vsiRNA production in hemipterans.",
], Inches(0.7), Inches(1.5), Inches(12.0), Inches(5.2), font_size=19, gap=14)
add_footer(s, 15)

# ============================================================
# SLIDE 16: TAKE-HOME MESSAGE
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s, NAVY)
# Center card
card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.5), Inches(11.33), Inches(4.6))
card.fill.solid(); card.fill.fore_color.rgb = WHITE
card.line.color.rgb = ORANGE; card.line.width = Pt(2)
card.shadow.inherit = False

tb = s.shapes.add_textbox(Inches(1.4), Inches(1.9), Inches(10.5), Inches(3.9))
tf = tb.text_frame; tf.word_wrap = True
sections = [
    ("What did they ask?", "What are the roles of conserved RNAi pathways in hemipteran antiviral immunity?", False),
    ("What did they do?", "RNAi knockdown, CRISPR/Cas9 knockout, and small RNA sequencing in four hemipteran species.", False),
    ("What did they find?", "AGO2 is essential for antiviral defense; DCR2 is essential for development but not for the measured antiviral response.", True),
    ("Why does it matter?", "It challenges the prevailing DCR2-centered model of insect antiviral RNAi and suggests lineage-specific organization of RNAi pathways.", False),
]
first = True
for h, body, bold in sections:
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    first = False
    p.text = h
    p.font.size = Pt(20); p.font.bold = True; p.font.color.rgb = NAVY
    p.font.name = "Calibri"; p.space_after = Pt(4)
    p2 = tf.add_paragraph(); p2.text = body
    p2.font.size = Pt(16); p2.font.color.rgb = DARK_TEXT
    p2.font.name = "Calibri"; p2.space_after = Pt(14); p2.line_spacing = 1.15

prs.save(OUTPUT)
print(f"Saved: {OUTPUT}")
print(f"Slides: {len(prs.slides)}")
