#!/usr/bin/env python3
"""
Create an improved, visually appealing PowerPoint presentation for the paper:
"Functional dissection of conserved RNAi pathways in antiviral immunity of several hemipteran insects"
Figures extracted from the PDF are inserted on the relevant slides.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml

# Paths
BASE_DIR = "/home/davidmwania/Mwania Project"
FIGURES_DIR = os.path.join(BASE_DIR, "paper_figures")

# Color palette
DARK_BLUE = RGBColor(0, 51, 102)
MEDIUM_BLUE = RGBColor(0, 102, 153)
LIGHT_BLUE = RGBColor(173, 216, 230)
ACCENT_ORANGE = RGBColor(230, 126, 34)
TEXT_DARK = RGBColor(40, 40, 40)
TEXT_LIGHT = RGBColor(255, 255, 255)
BG_LIGHT = RGBColor(250, 250, 250)

def add_background(slide, color=BG_LIGHT):
    """Add a full-slide background rectangle at the back."""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    # Move to back
    slide.shapes._spTree.remove(shape._element)
    slide.shapes._spTree.insert(2, shape._element)

def add_header_bar(slide):
    """Add a top header bar."""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.15))
    shape.fill.solid()
    shape.fill.fore_color.rgb = DARK_BLUE
    shape.line.fill.background()

def add_title_text(slide, title, top=Inches(0.35), left=Inches(0.5), width=Inches(12.333), 
                   font_size=Pt(32), bold=True, color=DARK_BLUE, align=PP_ALIGN.LEFT):
    title_box = slide.shapes.add_textbox(left, top, width, Inches(0.8))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = font_size
    p.font.bold = bold
    p.font.color.rgb = color
    p.alignment = align
    return title_box

def add_bullets(slide, bullets, left=Inches(0.6), top=Inches(1.3), width=Inches(12.333), 
                height=Inches(5.8), font_size=Pt(20), line_spacing=1.3):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, bullet in enumerate(bullets):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = bullet
        p.font.size = font_size
        p.font.color.rgb = TEXT_DARK
        p.space_after = Pt(10)
        p.level = 0
    return box

def add_image(slide, image_path, left, top, width=None, height=None):
    """Add an image with optional width or height, preserving aspect ratio."""
    if not os.path.exists(image_path):
        print(f"Warning: image not found: {image_path}")
        return None
    from PIL import Image as PILImage
    img = PILImage.open(image_path)
    w_px, h_px = img.size
    aspect = h_px / w_px
    if width is not None and height is None:
        height = width * aspect
    elif height is not None and width is None:
        width = height / aspect
    elif width is None and height is None:
        width = Inches(4.0)
        height = width * aspect
    pic = slide.shapes.add_picture(image_path, left, top, width=width, height=height)
    return pic

def make_title_slide(prs, title, subtitle):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide, DARK_BLUE)
    
    # Decorative accent bar
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(4.7), Inches(13.333), Inches(0.05))
    bar.fill.solid()
    bar.fill.fore_color.rgb = ACCENT_ORANGE
    bar.line.fill.background()
    
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(2.0), Inches(11.733), Inches(1.8))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = TEXT_LIGHT
    p.alignment = PP_ALIGN.CENTER
    
    sub_box = slide.shapes.add_textbox(Inches(0.8), Inches(4.0), Inches(11.733), Inches(2.0))
    tf = sub_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = subtitle
    p.font.size = Pt(22)
    p.font.color.rgb = RGBColor(220, 220, 220)
    p.alignment = PP_ALIGN.CENTER
    return slide

def make_section_slide(prs, title):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide, DARK_BLUE)
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.8), Inches(12.333), Inches(1.5))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(48)
    p.font.bold = True
    p.font.color.rgb = TEXT_LIGHT
    p.alignment = PP_ALIGN.CENTER
    return slide

def make_content_slide(prs, title, bullets, image_path=None, image_left=None, image_top=None, image_width=None, 
                       image_height=None, two_column=False, right_bullets=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide)
    add_header_bar(slide)
    add_title_text(slide, title)
    
    if image_path and two_column:
        # Left: bullets, Right: image
        add_bullets(slide, bullets, left=Inches(0.5), top=Inches(1.25), width=Inches(6.5), height=Inches(5.8), font_size=Pt(18))
        add_image(slide, image_path, left=Inches(7.2), top=image_top or Inches(1.4), width=image_width or Inches(5.5), height=image_height)
    elif image_path:
        # Bullets on top/bottom, image centered below
        add_bullets(slide, bullets, left=Inches(0.6), top=Inches(1.25), width=Inches(12.0), height=Inches(2.2), font_size=Pt(18))
        add_image(slide, image_path, left=image_left or Inches(2.5), top=image_top or Inches(3.6), width=image_width or Inches(8.0), height=image_height)
    else:
        add_bullets(slide, bullets, left=Inches(0.6), top=Inches(1.25), width=Inches(12.0), height=Inches(5.8), font_size=Pt(20))
    return slide

# Initialize presentation
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Slide 1: Title
slide1 = make_title_slide(prs,
    "Functional dissection of conserved RNAi pathways\nin antiviral immunity of several hemipteran insects",
    "Zhang Y, Ye Z-X, Ma Q-L, Wang B-Y, Zhuo J-C, Huang H-J, Zhang C-X, Chen J-P, Li J-M\n"
    "PLoS Pathogens, 2026, 22(10): e1014753\n"
    "DOI: https://doi.org/10.1371/journal.ppat.1014753\n\n"
    "Presented by: [Your Name]"
)

# Slide 2: Background
s2 = make_content_slide(prs, "Background: RNAi Pathways in Insects", [
    "RNA interference (RNAi) is a sequence-specific gene-silencing mechanism.",
    "Three main small silencing RNA pathways:",
    "   • miRNA pathway → gene regulation & development",
    "   • piRNA pathway → transposon suppression & germline maintenance",
    "   • siRNA pathway → antiviral defense (canonical view)",
    "Most knowledge comes from Dipteran model insects (fruit fly, mosquito).",
    "Hemipteran pests (planthoppers, whiteflies) transmit rice viruses, but their RNAi biology is largely unexplored."
])

# Slide 3: Study Question
s3 = make_content_slide(prs, "Study Question & Objectives", [
    "Main research question:",
    "  What are the functional roles of the conserved miRNA, piRNA, and siRNA pathways in hemipteran insects, particularly in antiviral immunity?",
    "",
    "Specific objectives:",
    "  1. Screen core RNAi genes in Laodelphax striatellus (small brown planthopper).",
    "  2. Determine which pathway(s) restrict viral accumulation.",
    "  3. Test whether the findings are conserved across multiple hemipteran species."
])

# Slide 4: Study Design
s4 = make_content_slide(prs, "Study Design", [
    "Type: Laboratory functional genomics study (RNAi + CRISPR/Cas9).",
    "Primary model: Laodelphax striatellus (small brown planthopper).",
    "Cross-species validation:",
    "   • Sogatella furcifera (rice planthopper)",
    "   • Nilaparvata lugens (brown planthopper)",
    "   • Bemisia tabaci (whitefly)",
    "Treatment: dsRNA targeting core RNAi genes vs. ds-GFP control.",
    "Readouts: survival, fecundity, development, viral RNA levels, small RNA profiles."
])

# Slide 5: Methods overview
s5 = make_content_slide(prs, "Methods Overview", [
    "RNAi knockdown: dsRNA microinjection of AGO1/2, DCR1/2, R2D2, Piwi, AGO3.",
    "RT-qPCR: gene-expression and viral quantification (normalized to ACTIN).",
    "Small RNA sequencing: profile miRNA, piRNA, and virus-derived siRNAs (vsiRNAs).",
    "RNAi-of-RNAi assay: test RNAi efficiency after pre-knockdown of key genes.",
    "CRISPR/Cas9: generate DCR2 mutant alleles to assess developmental requirements.",
    "Statistics: Student’s t-test, one-way ANOVA with Tukey, and Log-rank tests."
])

# Slide 6: Key Observation 1 - miRNA/piRNA + image Fig 2
s6 = make_content_slide(prs, "Key Observation 1: miRNA & piRNA Pathways Govern Development", [
    "miRNA pathway (AGO1, DCR1) knockdown:",
    "   • Reduced survival and molting defects.",
    "   • Severely reduced fecundity; eggs failed to hatch.",
    "   • 22-nt miRNA production dropped.",
    "piRNA pathway (Piwi, AGO3) knockdown:",
    "   • Abnormal wing curls and reduced egg production.",
    "   • Extended developmental duration.",
    "Interpretation: These pathways are essential for development and reproduction, but not major antiviral players."
], image_path=os.path.join(FIGURES_DIR, "page_06_img_01.png"), image_left=Inches(1.8), image_top=Inches(3.4), image_width=Inches(10.0))

# Slide 7: From Observation 1 to next analysis
s7 = make_content_slide(prs, "From Observation 1 to the Next Question", [
    "Observation 1: miRNA and piRNA gene knockdowns affected development/reproduction but did not increase viral load.",
    "",
    "Question raised:",
    "  If miRNA/piRNA are not the main antiviral pathways, what is the role of the canonical siRNA pathway?",
    "",
    "Next analysis:",
    "  Knockdown the three core siRNA genes:",
    "   • AGO2 — antiviral effector",
    "   • DCR2 — cleaves viral dsRNA into siRNAs",
    "   • R2D2 — loads siRNAs onto AGO2",
    "  Then measure viral RNA and vsiRNA production."
])

# Slide 8: Key Observation 2 - AGO2 vs DCR2 + image Fig 4
s8 = make_content_slide(prs, "Key Observation 2: AGO2, Not DCR2, Restricts Viral Accumulation", [
    "AGO2 knockdown:",
    "   • Significantly increased accumulation of four ISVs and RSV.",
    "   • vsiRNA reads rose with viral RNA.",
    "   • Normalized vsiRNA per viral copy fell, indicating impaired antiviral RNAi.",
    "DCR2 and R2D2 knockdown:",
    "   • No detectable increase in viral accumulation.",
    "   • Canonical 21–22 nt vsiRNA profiles were preserved."
], image_path=os.path.join(FIGURES_DIR, "page_10_img_01.png"), image_left=Inches(3.5), image_top=Inches(3.4), image_width=Inches(7.5))

# Slide 9: CRISPR DCR2 knockout + image Fig 5
s9 = make_content_slide(prs, "Deeper Validation: CRISPR/Cas9 Reveals DCR2 Is Essential for Embryos", [
    "Why CRISPR? RNAi knockdown may leave residual DCR2 activity.",
    "Two DCR2 mutant alleles generated (5-bp deletion + 9-bp insertion; 8-bp deletion).",
    "Result: No homozygous mutants recovered in G2; unhatched eggs carried homozygous mutations.",
    "Heterozygous mutants: reduced DCR2 expression but no viral increase; delayed nymph development.",
    "Conclusion: DCR2 is essential for embryonic viability, not detectably required for the measured antiviral response."
], image_path=os.path.join(FIGURES_DIR, "page_12_img_01.png"), image_left=Inches(2.5), image_top=Inches(3.8), image_width=Inches(8.5))

# Slide 10: Cross-species validation + image Fig 6
s10 = make_content_slide(prs, "Cross-Species Validation: Conserved AGO2-Dependent Antiviral Role", [
    "Tested three additional hemipteran species:",
    "   • Sogatella furcifera (Delphacidae)",
    "   • Nilaparvata lugens (Delphacidae)",
    "   • Bemisia tabaci (Aleyrodidae)",
    "In all three: AGO2 knockdown increased viral accumulation, while DCR2 knockdown did not.",
    "This supports a conserved AGO2-centric antiviral phenotype across Hemiptera."
], image_path=os.path.join(FIGURES_DIR, "page_13_img_01.png"), image_left=Inches(0.8), image_top=Inches(3.5), image_width=Inches(11.8))

# Slide 11: How Findings Were Connected (flow diagram)
s11 = prs.slides.add_slide(prs.slide_layouts[6])
add_background(s11)
add_header_bar(s11)
add_title_text(s11, "How the Findings Were Connected")

flow_items = [
    ("Research Question", "Roles of RNAi pathways in hemipteran antiviral immunity?"),
    ("Data Collection", "RNAi + sRNA-seq + CRISPR in 4 hemipteran species"),
    ("Observation 1", "miRNA/piRNA genes control development/reproduction"),
    ("Analysis 1", "Focus on siRNA pathway: knockdown AGO2, DCR2, R2D2"),
    ("Observation 2", "Only AGO2 depletion increases viral load"),
    ("Analysis 2", "CRISPR/Cas9 DCR2 knockout + RNAi-of-RNAi"),
    ("Final Conclusion", "AGO2 is antiviral; DCR2 is essential for embryonic development")
]

y_start = 1.35
box_height = 0.62
box_width = 11.5
spacing = 0.82

for i, (header, text) in enumerate(flow_items):
    y = y_start + i * spacing
    fill_color = DARK_BLUE if (i == 0 or i == len(flow_items) - 1) else MEDIUM_BLUE
    shape = s11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9), Inches(y), Inches(box_width), Inches(box_height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.color.rgb = TEXT_LIGHT
    tb = s11.shapes.add_textbox(Inches(1.0), Inches(y + 0.05), Inches(box_width - 0.2), Inches(box_height - 0.1))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.text = f"{header}: {text}"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = TEXT_LIGHT
    if i < len(flow_items) - 1:
        arrow = s11.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(6.6), Inches(y + box_height + 0.04), Inches(0.25), Inches(0.12))
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = ACCENT_ORANGE
        arrow.line.fill.background()

# Slide 12: Main Findings
s12 = make_content_slide(prs, "Main Findings", [
    "1. miRNA and piRNA pathways are essential for survival, development, and reproduction in L. striatellus.",
    "2. AGO2 is required for antiviral siRNA-mediated immunity in L. striatellus.",
    "3. DCR2 and R2D2 knockdown did not impair measured antiviral responses or vsiRNA profiles.",
    "4. DCR2 is essential for embryonic development—homozygous CRISPR mutants are lethal.",
    "5. The AGO2-essential / DCR2-dispensable antiviral pattern is conserved across three additional hemipteran species."
])

# Slide 13: Research question table
s13 = prs.slides.add_slide(prs.slide_layouts[6])
add_background(s13)
add_header_bar(s13)
add_title_text(s13, "How the Findings Answered the Research Question")
rows = 5
cols = 3
left = Inches(0.5)
top = Inches(1.35)
width = Inches(12.333)
height = Inches(5.2)
table = s13.shapes.add_table(rows, cols, left, top, width, height).table
table_data = [
    ["Research Question", "Evidence Found", "Conclusion"],
    ["What roles do miRNA/piRNA pathways play?", "Knockdown caused developmental defects, reduced survival and fecundity; no antiviral effect", "miRNA/piRNA are essential for development/reproduction, not antiviral defense"],
    ["Is the siRNA pathway antiviral?", "AGO2 knockdown increased viral load; DCR2/R2D2 knockdown did not", "siRNA pathway is antiviral, with AGO2 as the key effector"],
    ["Is DCR2 required for antiviral RNAi?", "vsiRNA profiles unchanged after DCR2 knockdown; CRISPR heterozygotes showed no viral increase", "DCR2 is not required for the measured antiviral RNAi response"],
    ["Are findings conserved across Hemiptera?", "Same AGO2/DCR2 contrast in S. furcifera, N. lugens, and B. tabaci", "Pattern is conserved among the hemipteran species examined"]
]
for i, row in enumerate(table_data):
    for j, val in enumerate(row):
        cell = table.cell(i, j)
        cell.text = val
        cell.text_frame.paragraphs[0].font.size = Pt(13)
        cell.text_frame.word_wrap = True
        if i == 0:
            cell.fill.solid()
            cell.fill.fore_color.rgb = DARK_BLUE
            cell.text_frame.paragraphs[0].font.color.rgb = TEXT_LIGHT
            cell.text_frame.paragraphs[0].font.bold = True
        else:
            cell.fill.solid()
            cell.fill.fore_color.rgb = LIGHT_BLUE if i % 2 == 0 else BG_LIGHT

# Slide 14: Strengths and Limitations
s14 = make_content_slide(prs, "Strengths and Limitations", [
    "Strengths:",
    "   • Multi-species validation across two hemipteran families.",
    "   • Combination of RNAi knockdown and CRISPR/Cas9 genetics.",
    "   • Comprehensive small RNA sequencing to assess pathway activity.",
    "   • RNAi-of-RNAi assay provided an independent functional readout.",
    "",
    "Limitations:",
    "   • RNAi knockdown may leave residual DCR2 activity.",
    "   • Homozygous DCR2 mutants died as embryos, so adult antiviral phenotypes could not be tested.",
    "   • Alternative vsiRNA biogenesis mechanisms were not identified.",
    "   • Experiments performed under specific lab conditions; field relevance needs further study."
])

# Slide 15: Conclusions
s15 = make_content_slide(prs, "Conclusions", [
    "The canonical DCR2-centered model of antiviral RNAi from Diptera does not fully explain hemipteran antiviral immunity.",
    "AGO2 is the conserved antiviral effector across the tested hemipteran species.",
    "DCR2 has an essential developmental role, particularly during embryogenesis.",
    "DCR2 depletion does not detectably impair the measured antiviral siRNA response in L. striatellus.",
    "Future work should identify alternative nucleases or mechanisms responsible for vsiRNA production in hemipterans."
])

# Slide 16: Take-Home Message
s16 = make_content_slide(prs, "Take-Home Message", [
    "What did they ask?",
    "   What are the roles of conserved RNAi pathways in hemipteran antiviral immunity?",
    "",
    "What did they do?",
    "   Used RNAi knockdown, CRISPR/Cas9, and small RNA sequencing in four hemipteran species.",
    "",
    "What did they find?",
    "   AGO2 is essential for antiviral defense; DCR2 is essential for development but not for the measured antiviral response.",
    "",
    "Why does it matter?",
    "   It challenges the prevailing DCR2-centered model of insect antiviral RNAi and suggests lineage-specific organization of RNAi pathways."
])

# Save
output_path = os.path.join(BASE_DIR, "RNAi_Hemiptera_Presentation.pptx")
prs.save(output_path)
print(f"Presentation saved to: {output_path}")
