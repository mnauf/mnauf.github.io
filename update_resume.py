from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

path = '/Users/naufil/Documents/work/mnauf.github.io/Resume - Muhammad Naufil - HiWi.docx'
doc = Document(path)

def has_numpR(para):
    return para._element.find('.//' + qn('w:numPr')) is not None

def clone_with_text(source_para, new_text):
    """Clone paragraph keeping pPr/formatting, replace text content."""
    new_elem = copy.deepcopy(source_para._element)
    to_remove = [c for c in list(new_elem) if c.tag in [
        qn('w:r'), qn('w:hyperlink'), qn('w:ins'), qn('w:del'),
        qn('w:bookmarkStart'), qn('w:bookmarkEnd')
    ]]
    for c in to_remove:
        new_elem.remove(c)
    if new_text:
        r_elem = OxmlElement('w:r')
        src_runs = [c for c in source_para._element if c.tag == qn('w:r')]
        if src_runs:
            rpr = src_runs[0].find(qn('w:rPr'))
            if rpr is not None:
                r_elem.append(copy.deepcopy(rpr))
        t_elem = OxmlElement('w:t')
        t_elem.text = new_text
        t_elem.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        r_elem.append(t_elem)
        new_elem.append(r_elem)
    return new_elem

# ── Step 1: Delete unwanted paragraphs (reverse order) ──────────────────────
original = list(doc.paragraphs)
for idx in sorted([17, 27, 28, 29, 30, 35, 36], reverse=True):
    p = original[idx]
    p._element.getparent().remove(p._element)

# ── Step 2: Identify template paragraphs ────────────────────────────────────
p_list = list(doc.paragraphs)

title_para     = p_list[0]
blank_tmpl     = next(p for p in p_list if p.text.strip() == '')
h1_tmpl        = next(p for p in p_list if p.style.name == 'Heading 1')
h2_tmpl        = next(p for p in p_list if p.style.name == 'Heading 2')
bullet_tmpl    = next(p for p in p_list if p.style.name == 'normal' and has_numpR(p))
# org-line: normal, long, NOT a bullet
org_tmpl       = next(p for p in p_list if
                      p.style.name == 'normal' and not has_numpR(p) and
                      len(p.text) > 50 and p.text.strip())

print("Templates found:")
print(f"  h1  : {h1_tmpl.text[:40]!r}")
print(f"  h2  : {h2_tmpl.text[:40]!r}")
print(f"  bullet: {bullet_tmpl.text[:40]!r}")
print(f"  org : {org_tmpl.text[:40]!r}")

# ── Step 3: Add headline subtitle after title ────────────────────────────────
headline_elem = clone_with_text(blank_tmpl, 'Computer Vision  ·  LLMs  ·  Agentic AI')
pPr = headline_elem.find(qn('w:pPr'))
if pPr is None:
    pPr = OxmlElement('w:pPr')
    headline_elem.insert(0, pPr)
jc = OxmlElement('w:jc'); jc.set(qn('w:val'), 'center')
pPr.append(jc)
r = headline_elem.find(qn('w:r'))
if r is not None:
    rPr = r.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr'); r.insert(0, rPr)
    sz = OxmlElement('w:sz'); sz.set(qn('w:val'), '20'); rPr.append(sz)
    szCs = OxmlElement('w:szCs'); szCs.set(qn('w:val'), '20'); rPr.append(szCs)
    col = OxmlElement('w:color'); col.set(qn('w:val'), '555555'); rPr.append(col)
title_para._element.addnext(headline_elem)

# ── Step 4: Insert MPI Informatics after Zema blank ─────────────────────────
p_list = list(doc.paragraphs)
zema_blank = None
for i, p in enumerate(p_list):
    if 'Zema' in p.text:
        for j in range(i+1, min(i+6, len(p_list))):
            if p_list[j].text.strip() == '':
                zema_blank = p_list[j]
                break
        break

if zema_blank:
    mpi_elems = [
        clone_with_text(blank_tmpl, ''),
        clone_with_text(org_tmpl,
            'MPI Informatics | Saarbrücken                                                    '
            '  Jun 2024 – Sep 2024'),
        clone_with_text(h2_tmpl, 'HiWi Research Assistant'),
        clone_with_text(bullet_tmpl,
            'Generative Models for Monadic Co-Speech Gesture Synthesis'),
        clone_with_text(bullet_tmpl,
            'Progressed from unconditional diffusion models (image generation) '
            'to class-conditioned, then latent diffusion models'),
        clone_with_text(bullet_tmpl,
            'Extended to latent diffusion models for gesture synthesis — '
            'conditioned on speech transcription'),
    ]
    current = zema_blank._element
    for elem in mpi_elems:
        current.addnext(elem)
        current = elem

# ── Step 5: Insert Personal Projects before Education ───────────────────────
p_list = list(doc.paragraphs)
edu_para = next((p for p in p_list if 'Education' in p.text and '\t' in p.text), None)

if edu_para:
    proj_section = [
        clone_with_text(h1_tmpl, 'Personal Projects\t'),
        clone_with_text(org_tmpl, 'linkinsertion.ai'),
        clone_with_text(bullet_tmpl,
            'Agentic AI tool for blogger outreach built with Openclaw (formerly ClawdCode)'),
        clone_with_text(bullet_tmpl,
            'Automates end-to-end link insertion outreach — prospect discovery, '
            'personalised pitches, and negotiation'),
        clone_with_text(blank_tmpl, ''),
        clone_with_text(org_tmpl, 'Personal AI Companion'),
        clone_with_text(bullet_tmpl,
            'RAG-based LLM with vision, function calling, and long-term '
            'conversational memory (Pinecone)'),
        clone_with_text(bullet_tmpl,
            'Token-efficient: processes images only when contextually appropriate; '
            'recalls past conversations like a true friend'),
        clone_with_text(blank_tmpl, ''),
    ]
    # Insert in forward order using addprevious on edu_para
    for elem in proj_section:
        edu_para._element.addprevious(elem)

# ── Save ─────────────────────────────────────────────────────────────────────
doc.save(path)
print("\nSaved successfully.")

doc2 = Document(path)
for i, p in enumerate(doc2.paragraphs):
    print(f'[{i:02d}] {p.style.name:<12} | {repr(p.text[:72])}')
