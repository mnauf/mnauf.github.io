"""
Fixes the corrupted resume doc by:
1. Removing duplicates created by running the update script twice
2. Restoring accidentally deleted structural elements
3. Adding missing content
"""
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

path = '/Users/naufil/Documents/work/mnauf.github.io/Resume - Muhammad Naufil - HiWi.docx'
doc = Document(path)

def has_numpr(para):
    return para._element.find('.//' + qn('w:numPr')) is not None

def clone_with_text(source_para, new_text):
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

# ── Step 1: Delete bad paragraphs ───────────────────────────────────────────
# Current indices (from inspection):
# [02] duplicate headline
# [20] extra blank before MPI
# [27-31] duplicate MPI section
# [39-49] reversed/garbage Personal Projects from run 1

p_list = list(doc.paragraphs)
to_delete_idx = sorted([2, 20, 27, 28, 29, 30, 31, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49],
                       reverse=True)
for idx in to_delete_idx:
    p = p_list[idx]
    p._element.getparent().remove(p._element)

print("After cleanup deletions:")
for i, p in enumerate(doc.paragraphs):
    print(f'  [{i:02d}] {p.style.name:<12} | {repr(p.text[:60])}')

# ── Step 2: Find template paragraphs ────────────────────────────────────────
p_list = list(doc.paragraphs)
bullet_tmpl = next(p for p in p_list if p.style.name == 'normal' and has_numpr(p))
h2_tmpl     = next(p for p in p_list if p.style.name == 'Heading 2')
org_tmpl    = next(p for p in p_list if p.style.name == 'normal' and not has_numpr(p)
                   and len(p.text) > 50 and p.text.strip())
blank_tmpl  = next(p for p in p_list if p.text.strip() == '')

# ── Step 3: Add back "Tested instruction-following..." after LLMs/AI Agents ─
p_list = list(doc.paragraphs)
llms_para = next((p for p in p_list if 'LLMs/AI Agents' in p.text), None)
if llms_para:
    tested_elem = clone_with_text(bullet_tmpl,
        'Tested instruction-following and few-shot learning capabilities of '
        'the latest open-source and closed-source LLMs')
    llms_para._element.addnext(tested_elem)
    print("\nAdded 'Tested instruction-following...' after LLMs/AI Agents")

# ── Step 4: Insert Retrocausal structure back ────────────────────────────────
# After Retrocausal org: add H2 "Computer Vision Engineer", then H2 "3D Human Pose"
p_list = list(doc.paragraphs)
retro_org = next((p for p in p_list if 'Retrocausal' in p.text), None)
if retro_org:
    # Insert H2 "3D Human Pose Estimation" first (addnext order)
    # Then H2 "Computer Vision Engineer"
    # Result: Retrocausal → H2 CV Eng → blank → H2 3D Pose → bullets
    h2_3dpose_elem = clone_with_text(h2_tmpl, '3D Human Pose Estimation')
    h2_cveng_elem  = clone_with_text(h2_tmpl, 'Computer Vision Engineer')
    blank_elem     = clone_with_text(blank_tmpl, '')

    # Insert in reverse using addnext on retro_org to maintain order:
    # retro_org → h2_cveng → blank → h2_3dpose → (existing bullets)
    p_list = list(doc.paragraphs)
    retro_org = next((p for p in p_list if 'Retrocausal' in p.text), None)
    # Find the first bullet after Retrocausal to insert before
    retro_idx = p_list.index(retro_org)
    first_bullet_after = p_list[retro_idx + 1]

    # Insert before first bullet (in order): h2_cveng, blank, h2_3dpose
    first_bullet_after._element.addprevious(h2_3dpose_elem)
    first_bullet_after._element.addprevious(blank_elem)
    first_bullet_after._element.addprevious(h2_cveng_elem)
    print("Added H2 Computer Vision Engineer + H2 3D Human Pose Estimation")

# ── Step 5: Insert H2 "Human Action Segmentation" + demo bullet before "Reduced..." ─
p_list = list(doc.paragraphs)
reduced_para = next((p for p in p_list if 'Reduced purchasing price' in p.text), None)
if reduced_para:
    h2_actionseg_elem = clone_with_text(h2_tmpl, 'Human Action Segmentation')
    demo_bullet_elem = clone_with_text(bullet_tmpl,
        'Created the human activity segmentation demo for the SBIR grant '
        'from NASA\'s Human Research Program (HRP)')

    # Insert before "Reduced..." in order: blank, H2, demo_bullet
    blank2 = clone_with_text(blank_tmpl, '')
    reduced_para._element.addprevious(demo_bullet_elem)
    reduced_para._element.addprevious(h2_actionseg_elem)
    reduced_para._element.addprevious(blank2)
    print("Added H2 Human Action Segmentation + demo bullet")

# ── Save and verify ──────────────────────────────────────────────────────────
doc.save(path)
print("\nSaved. Final structure:")
doc2 = Document(path)
for i, p in enumerate(doc2.paragraphs):
    nump = p._element.find('.//' + qn('w:numPr')) is not None
    print(f'[{i:02d}] {p.style.name:<12} num={nump} | {repr(p.text[:65])}')
