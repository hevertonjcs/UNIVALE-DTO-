from pathlib import Path
from copy import deepcopy
from docx import Document
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

root = Path('AULAS')
doc = Document(root / 'Artigo estelionato UNIVALI - fontes revisadas.docx')
paras = doc.paragraphs
idx = next(i for i, p in enumerate(paras) if p.text.strip() == 'REFERÊNCIAS DAS FONTES CITADAS')
refs = [p for p in paras[idx+1:] if p.text.strip()]
def key(p):
    t = p.text.strip()
    order = {'BRANDÃO': 0, 'BRASIL': 1, 'COSTA': 2, 'DIVAN': 3, 'NUCCI': 4, 'OLIVEIRA': 5, 'SANTOS': 6, 'STJ': 7, 'VOLPI': 8}
    author = t.split('.',1)[0].split(',',1)[0]
    if author == 'BRASIL':
        if 'Constituição' in t: sub = 0
        elif '2.848' in t: sub = 1
        elif '3.689' in t: sub = 2
        elif '13.964' in t: sub = 3
        elif '15.229' in t: sub = 4
        elif '15.397' in t: sub = 5
        else: sub = 6
    elif author == 'STJ':
        if '411.987' in t: sub = 0
        elif '143.621' in t: sub = 1
        elif '610.201' in t: sub = 2
        elif '2.468.274' in t: sub = 3
        elif '990.448' in t: sub = 4
        else: sub = 5
    else: sub = 0
    return (order.get(author, 99), sub, t)
ordered = sorted(refs, key=key)
anchor = paras[idx]._p
for p in ordered:
    anchor.addnext(p._p)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.line_spacing = 1.0
    anchor = p._p
doc.save(root / 'Artigo estelionato UNIVALI - referências organizadas.docx')
