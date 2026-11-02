"""Fix the ordering issues in Retrocausal section from the previous fix."""
from docx import Document
from docx.oxml.ns import qn
import copy

path = '/Users/naufil/Documents/work/mnauf.github.io/Resume - Muhammad Naufil - HiWi.docx'
doc = Document(path)

# Current state (from last run):
# [26] Retrocausal org
# [27] H2 3D Human Pose Estimation  ← wrong order
# [28] blank
# [29] H2 Computer Vision Engineer  ← should come first
# [30] Created MoCap bullet
# [31] Experimenting monocular
# [32] Addressed monocular ambiguity
# [33] blank
# [34] Created activity seg demo bullet ← before its H2!
# [35] H2 Human Action Segmentation    ← should be before [34]
# [36] blank
# [37] Reduced price 150x
# [38] Wrote algorithm

# Fix 1: Remove [27], [28], [29] and re-insert in right order before [30]
# Fix 2: Remove [34], [35], [36] and re-insert in right order before [37]

p_list = list(doc.paragraphs)

# Grab and detach the misplaced elements (indices 27, 28, 29)
h2_3dpose = p_list[27]
blank28    = p_list[28]
h2_cveng   = p_list[29]
first_mocap = p_list[30]  # "Created this markerless MoCap..."

# Also grab action seg misplaced elements
demo_bullet = p_list[34]
h2_actionseg = p_list[35]
blank36      = p_list[36]
reduced      = p_list[37]

# Detach [27],[28],[29] from document
for p in [h2_3dpose, blank28, h2_cveng]:
    p._element.getparent().remove(p._element)

# Re-insert before first_mocap in desired order:
# Retrocausal org → H2 CV Eng → H2 3D Pose → MoCap bullets
# addprevious(x) inserts x immediately before the reference.
# To get: h2_cveng, h2_3dpose, first_mocap
# We do: addprevious(h2_cveng) → h2_cveng, first_mocap
#        addprevious(h2_3dpose) → h2_cveng, h2_3dpose, first_mocap ✓
first_mocap._element.addprevious(h2_cveng._element)
first_mocap._element.addprevious(h2_3dpose._element)

# Now fix action seg section.
# After the above insertions, indices have shifted.
# Detach demo_bullet, h2_actionseg, blank36 from current positions
for p in [demo_bullet, h2_actionseg, blank36]:
    p._element.getparent().remove(p._element)

# Re-insert before "Reduced price" in desired order:
# blank, H2 Human Action Seg, demo_bullet, Reduced price
# addprevious(blank36) → blank36, reduced
# addprevious(h2_actionseg) → blank36, h2_actionseg, reduced
# addprevious(demo_bullet) → blank36, h2_actionseg, demo_bullet, reduced ✓
reduced._element.addprevious(blank36._element)
reduced._element.addprevious(h2_actionseg._element)
reduced._element.addprevious(demo_bullet._element)

doc.save(path)
print("Saved. Final structure:")

doc2 = Document(path)
for i, p in enumerate(doc2.paragraphs):
    nump = p._element.find('.//' + qn('w:numPr')) is not None
    print(f'[{i:02d}] {p.style.name:<12} num={nump} | {repr(p.text[:65])}')
