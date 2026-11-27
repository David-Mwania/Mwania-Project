#!/usr/bin/env python3
"""
Builder script for the research presentation study document.
Run this script (or extend APPENDICES) to regenerate/update the Word document.
"""

import os
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

BASE = "/home/davidmwania/Mwania Project"
FOLDER = os.path.join(BASE, "research presentation")
OUT = os.path.join(FOLDER, "RNAi_Paper_Study_Notes.docx")

NAVY = RGBColor(15, 40, 71)
BLUE = RGBColor(23, 84, 138)
GREY = RGBColor(90, 98, 108)

def heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for r in h.runs:
        r.font.color.rgb = NAVY if level == 1 else BLUE
        r.font.name = "Calibri"
    return h

def para(doc, text, size=11, italic=False, bold=False, color=None, align=None):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.size = Pt(size); r.font.italic = italic; r.font.bold = bold
    r.font.name = "Calibri"
    if color:
        r.font.color.rgb = color
    if align:
        p.alignment = align
    return p

def qa_block(doc, date, question, answer_paras):
    """Add a Q&A block: question heading + dated + answer paragraphs."""
    heading(doc, question, level=2)
    para(doc, f"Added: {date}", size=9, italic=True, color=GREY)
    for block in answer_paras:
        if isinstance(block, tuple):  # (text, style) where style in {"bullet","para"}
            text, style = block
            if style == "bullet":
                p = doc.add_paragraph(text, style="List Bullet")
            else:
                p = doc.add_paragraph(text)
        else:
            p = doc.add_paragraph(block)
        for r in p.runs:
            r.font.size = Pt(11); r.font.name = "Calibri"
    doc.add_paragraph()

APPENDICES = [
    {
        "date": "Session 3",
        "question": "Result 3: Did the switch-offs actually work? (Checking the knockdowns at the small-RNA level)",
        "paras": [
            "Before trusting any 'switching off this gene had no effect' claim, there is a fair challenge the researchers had to answer first: how do we know the genes were truly switched off? If the knockdowns failed, the 'no effect' results would mean nothing. So they checked the knockdowns at two levels: did the target gene's message (mRNA) drop, and — more importantly — did the pathway's actual product (its small RNAs) disappear?",
            ("The gene-message level was convincing. Injection of the knockout RNA cut AGO1's message by 82% and DCR1's by 70%, with similar strong reductions for the other target genes. But the real test was whether the pathway's output changed, because a pathway is defined by what it produces, not just by whether its gene is quiet.", "para"),
            ("For the miRNA pathway, the answer was clear. In normal insects, small RNAs of about 22 nucleotides — the miRNAs — are abundant. After switching off either AGO1 or DCR1, the 22-nucleotide miRNA pool dropped sharply (roughly halved or worse). This mirrors what happens in fruit flies, showing the miRNA pathway in the planthopper works the same way — and that the knockdown truly disabled it.", "para"),
            ("For the piRNA pathway, the test was even more specific. Planthoppers carry ancient virus-like DNA sequences in their own genome (called endogenous viral elements, or LsETLVEs), and these produce piRNAs. When Piwi was switched off, both the general 26–30 nucleotide piRNA pool and the piRNAs made from these embedded viral elements dropped dramatically — direct proof that Piwi is the engine making them. AGO3 knockdown did not clearly reduce these piRNAs, which the authors interpret cautiously: either AGO3's role is smaller, or leftover protein and weak amplification masked the effect.", "para"),
            ("Why does all this matter for the paper's story? Because it validates the whole experiment. Every 'no effect' result that comes later can now be trusted — not as a failed knockdown, but as a genuine biological finding. The miRNA and piRNA pathways were truly silenced and truly disrupted (Result 2 showed the damage). The stage is now set for the central test: switch off the siRNA pathway genes and see whether viruses are released from control.", "para"),
        ],
    },
    {
        "date": "Session 2",
        "question": "Result 2: What happens when each RNAi pathway is switched off?",
        "paras": [
            "With the genes identified, the researchers moved to the obvious next question: what do these genes actually do in a living insect? They used a simple trick called RNAi knockdown — injecting a tailor-made double-stranded RNA that tells the insect's own machinery to destroy the message of one specific target gene, effectively 'switching it off'. They did this separately for the seven key genes and then watched three things: whether the insects survived, whether females laid eggs, and how long young insects took to develop into adults. A harmless control RNA (against GFP, a jellyfish gene that planthoppers don't have) was injected into a comparison group.",
            ("The results split the three pathways into two camps. Switching off the miRNA pathway genes (AGO1 or DCR1) was devastating: many insects died, nymphs got stuck trying to molt (the old skin would not come off), and females produced almost no eggs — DCR1 knockdown stopped egg-laying completely, while AGO1 knockdown eggs never hatched. The piRNA pathway genes (Piwi or AGO3) mattered mainly for reproduction and growth: adults developed warped, curled wings, females laid far fewer eggs that mostly failed to hatch, and development took noticeably longer. But survival was unaffected.", "para"),
            ("The surprise came from the siRNA pathway — the one supposedly dedicated to antiviral defense. Switching off AGO2, DCR2, or R2D2 caused no increase in death and no drop in fertility or hatching. Only development timing shifted mildly (slower with AGO2 loss, faster with R2D2 loss). The expression patterns made sense of this: miRNA and piRNA genes were most active in reproductive organs and early embryos — exactly where their defects showed up — while AGO2 and DCR2 were most active in the gut, the entry point for viruses, hinting their real job might be antiviral rather than developmental.", "para"),
            ("This created the paper's first big clue. The miRNA and piRNA pathways are clearly essential for the insect's daily life — development, molting, and reproduction — but nothing so far suggested they fight viruses. Meanwhile the siRNA pathway looked unimportant for ordinary biology yet perfectly placed (in the gut) for antiviral duty. That observation set up the next question: if these siRNA genes don't matter for survival, are they the ones actually holding the viruses in check?", "para"),
        ],
    },
    {
        "date": "Session 1",
        "question": "Result 1: Identifying the RNAi genes — what did the researchers do and find?",
        "paras": [
            "Before the researchers could test how insects fight viruses, they first had to check something very basic: does the insect even have the right genes for RNA interference? Think of RNAi genes as the insect's built-in antiviral toolkit. Using well-studied fruit flies and mosquitoes as a reference, they searched the genomes of three rice planthoppers (Laodelphax striatellus, Sogatella furcifera, and Nilaparvata lugens) to see whether similar genes exist there.",
            ("They found all seven key RNAi genes in each species — 21 genes in total — and 11 of them had never been described before. This matters because these planthoppers are major agricultural pests that spread plant viruses, yet scientists knew almost nothing about their RNAi machinery.", "para"),
            ("To be sure they had found the right genes, the researchers did two checks. First, they built family trees (phylogenies) of the proteins, which showed the planthopper genes grouping together separately from flies, beetles, and worms — exactly as expected if they are true counterparts of the same genes. Second, they inspected the protein 'parts' (domains) and confirmed each gene carries the equipment needed to do its job, like the cutting machinery in Dicer enzymes or the RNA-gripping clamp in Argonaute proteins.", "para"),
            ("Two findings stood out. The Ls_DCR1 gene is missing one part (a helicase domain) that its fly counterpart has, and Ls_R2D2 has an extra RNA-binding piece that others lack. The researchers did not claim these differences change how the genes work — they simply flagged them as hints that these genes might behave differently from what we see in fruit flies. That hint turned out to be important later in the paper.", "para"),
            ("In short: the genes are present, they look complete, and they look like the real thing. This gave the researchers the green light to move to the next step — switching these genes off one by one to see what happens (Result 2).", "para"),
        ],
    },
]

def build():
    doc = Document()
    # Title page content
    t = doc.add_heading("RNAi Pathways in Hemipteran Insects — Study Notes", level=0)
    for r in t.runs:
        r.font.color.rgb = NAVY; r.font.name = "Calibri"
    para(doc, "Companion notes for the presentation of: Zhang Y, Ye Z-X, Ma Q-L, et al. "
              "\"Functional dissection of conserved RNAi pathways in antiviral immunity of several "
              "hemipteran insects.\" PLoS Pathogens 2026;22(10):e1014753.",
         size=10, italic=True, color=GREY)
    para(doc, "This document grows over time: every question discussed in our sessions is added as a new "
              "section below, written in plain language for presentation prep.",
         size=10, italic=True, color=GREY)
    doc.add_paragraph()

    for entry in APPENDICES:
        qa_block(doc, entry["date"], entry["question"], entry["paras"])

    doc.save(OUT)
    print(f"Saved: {OUT}")

if __name__ == "__main__":
    build()
