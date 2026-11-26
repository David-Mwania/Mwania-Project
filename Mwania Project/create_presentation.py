#!/usr/bin/env python3
"""
Create a PowerPoint presentation for the paper:
"Functional dissection of conserved RNAi pathways in antiviral immunity of several hemipteran insects"
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# Create presentation object
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Helper functions
def add_title_slide(prs, title, subtitle):
    slide_layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(slide_layout)
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.0), Inches(12.333), Inches(1.5))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 51, 102)
    p.alignment = PP_ALIGN.CENTER
    
    # Subtitle
    sub_box = slide.shapes.add_textbox(Inches(0.5), Inches(3.6), Inches(12.333), Inches(2.0))
    tf = sub_box.text_frame
    p = tf.paragraphs[0]
    p.text = subtitle
    p.font.size = Pt(22)
    p.font.color.rgb = RGBColor(80, 80, 80)
    p.alignment = PP_ALIGN.CENTER
    
    return slide

def add_section_slide(prs, section_title):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    
    # Section title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.8), Inches(12.333), Inches(1.5))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = section_title
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER
    
    # Background rectangle
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0, 51, 102)
    shape.line.fill.background()
    # Send to back
    slide.shapes._spTree.remove(shape._element)
    slide.shapes._spTree.insert(2, shape._element)
    
    # Re-add title on top
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.8), Inches(12.333), Inches(1.5))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = section_title
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER
    
    return slide

def add_content_slide(prs, title, bullets, notes=None):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.333), Inches(0.8))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 51, 102)
    
    # Content
    content_box = slide.shapes.add_textbox(Inches(0.7), Inches(1.3), Inches(12.333), Inches(5.8))
    tf = content_box.text_frame
    tf.word_wrap = True
    
    for i, bullet in enumerate(bullets):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = bullet
        p.font.size = Pt(20)
        p.space_after = Pt(10)
        p.level = 0
    
    return slide

# Slide 1: Title
slide1 = add_title_slide(prs,
    "Functional dissection of conserved RNAi pathways\nin antiviral immunity of several hemipteran insects",
    "Yan Zhang, Zhuang-Xin Ye, Qing-Lu Ma, et al.\n"
    "PLoS Pathogens, 2026, 22(10): e1014753\n"
    "DOI: https://doi.org/10.1371/journal.ppat.1014753\n\n"
    "Presented by: [Your Name]"
)

# Slide 2: Background
slide2 = add_content_slide(prs, "Background: What We Already Know", [
    "RNA interference (RNAi) is a key post-transcriptional gene silencing mechanism in insects.",
    "Three main small silencing RNA pathways: miRNA, piRNA, and siRNA.",
    "siRNA pathway is considered the primary antiviral defense in insects.",
    "Most knowledge comes from model Dipteran insects (fruit flies and mosquitoes).",
    "Hemipteran insects (e.g., planthoppers, whiteflies) are major agricultural pests and virus vectors, but their RNAi pathways are poorly understood."
])

# Slide 3: Study Question / Objective
slide3 = add_content_slide(prs, "Study Question & Objectives", [
    "Main Research Question:",
    "  What are the functional roles of the conserved small silencing RNA pathways (miRNA, piRNA, siRNA) in hemipteran insects, particularly in antiviral immunity?",
    "",
    "Specific Objectives:",
    "  1. Functionally characterize core RNAi pathway genes in Laodelphax striatellus.",
    "  2. Determine which pathway(s) mediate antiviral defense.",
    "  3. Assess whether findings are conserved across multiple hemipteran species."
])

# Slide 4: Study Design
slide4 = add_content_slide(prs, "Study Design", [
    "Study type: Laboratory-based functional genomics study (RNAi knockdown + CRISPR/Cas9 knockout).",
    "Model organism: Small brown planthopper (Laodelphax striatellus).",
    "Additional species for cross-species validation:",
    "  • Sogatella furcifera (rice planthopper)",
    "  • Nilaparvata lugens (brown planthopper)",
    "  • Bemisia tabaci (whitefly)",
    "Experimental groups: dsRNA-injected (target genes) vs. ds-GFP control.",
    "Phenotypes measured: survival, fecundity, developmental duration, viral RNA accumulation, and small RNA profiles."
])

# Slide 5: Methods
slide5 = add_content_slide(prs, "Methods", [
    "RNAi knockdown: dsRNA microinjection of core RNAi genes (AGO1, DCR1, Piwi, AGO3, AGO2, DCR2, R2D2).",
    "Gene expression: RT-qPCR normalized to ACTIN using the 2–ΔΔCt method.",
    "Small RNA sequencing (sRNA-seq): profiled miRNA, piRNA, and virus-derived siRNAs (vsiRNAs).",
    "Viral quantification: RT-qPCR and absolute qPCR for insect-specific RNA viruses (ISVs) and RSV.",
    "RNAi-of-RNAi assay: tested RNAi efficiency after pre-knockdown of DCR1, DCR2, or AGO2.",
    "CRISPR/Cas9: generated DCR2 mutant alleles to assess developmental requirements.",
    "Statistics: Student’s t-test, one-way ANOVA with Tukey’s test, and Log-rank (Mantel-Cox) test."
])

# Slide 6: Key Observation 1
slide6 = add_content_slide(prs, "Key Observation 1: miRNA & piRNA Pathways Regulate Development", [
    "miRNA pathway genes (AGO1 and DCR1):",
    "  • Knockdown reduced survival rate and caused molting defects.",
    "  • Female fecundity was severely reduced; eggs failed to hatch.",
    "  • miRNA (22-nt) production was strongly reduced.",
    "",
    "piRNA pathway genes (Piwi and AGO3):",
    "  • Knockdown caused abnormal wing curls and reduced egg production.",
    "  • Developmental duration was significantly extended.",
    "  • Piwi was required for endogenous viral element (EVE)-derived piRNAs.",
    "",
    "Interpretation: miRNA and piRNA pathways are essential for development, survival, and reproduction—but not major antiviral players."
])

# Slide 7: How Observation 1 Led to the Next Analysis
slide7 = add_content_slide(prs, "From Observation 1 to the Next Question", [
    "Observation 1: miRNA and piRNA pathway knockdowns affected development/reproduction but did not increase viral accumulation.",
    "",
    "Question Raised:",
    "  If miRNA/piRNA pathways are not driving antiviral defense, what is the role of the canonical siRNA pathway?",
    "",
    "Next Analysis:",
    "  Knockdown the three core siRNA pathway genes:",
    "    • AGO2 (effector nuclease)",
    "    • DCR2 (dsRNA cleavage into siRNAs)",
    "    • R2D2 (dsRNA-binding partner for loading)",
    "  Then measure viral RNA levels and vsiRNA production."
])

# Slide 8: Key Observation 2
slide8 = add_content_slide(prs, "Key Observation 2: AGO2, Not DCR2, Restricts Viral Accumulation", [
    "AGO2 knockdown (Ls_AGO2):",
    "  • Significantly increased accumulation of four ISVs and RSV.",
    "  • Increased total vsiRNA reads (consistent with higher viral RNA).",
    "  • Normalized vsiRNA per viral copy was reduced, indicating impaired antiviral RNAi.",
    "",
    "DCR2 and R2D2 knockdowns:",
    "  • No detectable increase in viral accumulation.",
    "  • Canonical vsiRNA size profiles (21–22 nt) were maintained.",
    "",
    "Conclusion: AGO2 is the key antiviral effector; DCR2 and R2D2 are not required for the measured antiviral response under these conditions."
])

# Slide 9: How Findings Were Connected (flow diagram)
slide9 = prs.slides.add_slide(prs.slide_layouts[6])
title_box = slide9.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.333), Inches(0.8))
tf = title_box.text_frame
p = tf.paragraphs[0]
p.text = "How the Findings Were Connected"
p.font.size = Pt(32)
p.font.bold = True
p.font.color.rgb = RGBColor(0, 51, 102)

flow_items = [
    ("Research Question", "Roles of RNAi pathways in hemipteran antiviral immunity"),
    ("Data Collected", "RNAi knockdown + sRNA-seq + CRISPR in 4 hemipteran species"),
    ("Observation 1", "miRNA/piRNA genes control development/reproduction, not antiviral response"),
    ("Analysis 1", "Focus on siRNA pathway: knockdown AGO2, DCR2, R2D2"),
    ("Observation 2", "Only AGO2 depletion increases viral load; DCR2/R2D2 do not"),
    ("Analysis 2", "CRISPR/Cas9 DCR2 knockout + RNAi-of-RNAi assay"),
    ("Final Conclusion", "AGO2 is antiviral; DCR2 is essential for embryonic development")
]

y_start = 1.3
box_height = 0.7
box_width = 12.0
spacing = 0.85

for i, (header, text) in enumerate(flow_items):
    y = y_start + i * spacing
    # box
    shape = slide9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.7), Inches(y), Inches(box_width), Inches(box_height))
    shape.fill.solid()
    if i == 0 or i == len(flow_items) - 1:
        shape.fill.fore_color.rgb = RGBColor(0, 51, 102)
    else:
        shape.fill.fore_color.rgb = RGBColor(70, 130, 180)
    shape.line.color.rgb = RGBColor(255, 255, 255)
    # text
    tb = slide9.shapes.add_textbox(Inches(0.9), Inches(y + 0.05), Inches(box_width - 0.4), Inches(box_height - 0.1))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.text = f"{header}: {text}"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    # arrow
    if i < len(flow_items) - 1:
        arrow = slide9.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(6.5), Inches(y + box_height + 0.02), Inches(0.3), Inches(0.1))
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = RGBColor(100, 100, 100)

# Slide 10: Main Findings
slide10 = add_content_slide(prs, "Main Findings", [
    "1. miRNA and piRNA pathways are essential for survival, development, and reproduction in L. striatellus.",
    "2. AGO2 is required for antiviral siRNA-mediated immunity.",
    "3. DCR2 and R2D2 knockdown did not impair measured antiviral responses or vsiRNA profiles.",
    "4. DCR2 is essential for embryonic development—homozygous CRISPR mutants are lethal.",
    "5. The AGO2-essential / DCR2-dispensable pattern is conserved across three additional hemipteran species."
])

# Slide 11: How Findings Answered the Research Question
slide11 = prs.slides.add_slide(prs.slide_layouts[6])
title_box = slide11.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.333), Inches(0.8))
tf = title_box.text_frame
p = tf.paragraphs[0]
p.text = "How the Findings Answered the Research Question"
p.font.size = Pt(32)
p.font.bold = True
p.font.color.rgb = RGBColor(0, 51, 102)

# Table
rows = 5
cols = 3
left = Inches(0.7)
top = Inches(1.4)
width = Inches(12.0)
height = Inches(5.0)
table = slide11.shapes.add_table(rows, cols, left, top, width, height).table

table_data = [
    ["Research Question", "Evidence Found", "Conclusion"],
    ["What roles do miRNA/piRNA pathways play?", "Knockdown caused developmental defects, reduced survival and fecundity; no antiviral effect", "miRNA/piRNA pathways are essential for development/reproduction, not antiviral defense"],
    ["Is the siRNA pathway antiviral?", "AGO2 knockdown increased viral load; DCR2/R2D2 knockdown did not", "siRNA pathway is antiviral, but AGO2 is the key effector"],
    ["Is DCR2 required for antiviral RNAi?", "vsiRNA profiles unchanged after DCR2 knockdown; CRISPR heterozygotes showed no viral increase", "DCR2 is not required for measured antiviral RNAi under tested conditions"],
    ["Are findings conserved across Hemiptera?", "Same AGO2/DCR2 contrast in S. furcifera, N. lugens, and B. tabaci", "Pattern is conserved among the hemipteran species examined"]
]

for i, row in enumerate(table_data):
    for j, val in enumerate(row):
        cell = table.cell(i, j)
        cell.text = val
        cell.text_frame.paragraphs[0].font.size = Pt(14)
        if i == 0:
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(0, 51, 102)
            cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
            cell.text_frame.paragraphs[0].font.bold = True
        else:
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(240, 248, 255) if i % 2 == 0 else RGBColor(255, 255, 255)

# Slide 12: Strengths and Limitations
slide12 = add_content_slide(prs, "Strengths and Limitations", [
    "Strengths:",
    "  • Multi-species validation across two families of Hemiptera.",
    "  • Combined RNAi knockdown and CRISPR/Cas9 genetic analysis.",
    "  • Comprehensive small RNA sequencing to assess pathway activity.",
    "  • RNAi-of-RNAi assay provided independent functional readout.",
    "",
    "Limitations:",
    "  • RNAi knockdown may leave residual DCR2 activity.",
    "  • Homozygous DCR2 mutants died as embryos, so adult antiviral phenotypes could not be tested in mutants.",
    "  • Alternative vsiRNA biogenesis mechanisms were not identified.",
    "  • Experiments were performed under specific lab conditions; field relevance needs further study."
])

# Slide 13: Conclusions
slide13 = add_content_slide(prs, "Conclusions", [
    "The canonical DCR2-centered model of antiviral RNAi from Diptera does not fully explain hemipteran antiviral immunity.",
    "AGO2 is the conserved antiviral effector across the tested hemipteran species.",
    "DCR2 has an essential developmental role, particularly during embryogenesis.",
    "DCR2 depletion does not detectably impair the measured antiviral siRNA response in L. striatellus.",
    "Future work should identify alternative nucleases or mechanisms responsible for vsiRNA production in hemipterans."
])

# Slide 14: Take-Home Message
slide14 = add_content_slide(prs, "Take-Home Message", [
    "What did they ask?",
    "  What are the roles of conserved RNAi pathways in hemipteran antiviral immunity?",
    "",
    "What did they do?",
    "  Used RNAi knockdown, CRISPR/Cas9 knockout, and small RNA sequencing in four hemipteran species.",
    "",
    "What did they find?",
    "  AGO2 is essential for antiviral defense; DCR2 is essential for development but not required for the measured antiviral response.",
    "",
    "Why does it matter?",
    "  It challenges the prevailing DCR2-centered model of insect antiviral RNAi and suggests lineage-specific organization of RNAi pathways."
])

# Save presentation
output_path = "/home/davidmwania/Mwania Project/RNAi_Hemiptera_Presentation.pptx"
prs.save(output_path)
print(f"Presentation saved to: {output_path}")
