from pathlib import Path
from copy import deepcopy
from docx import Document
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement

root = Path('AULAS')
d = Document(root / 'Artigo estelionato UNIVALI - versão preliminar completa.docx')
text = (root / 'Capítulo 1 - estelionato e representação - versão ampliada.md').read_text(encoding='utf-8')
body, refs = text.split('## Fontes usadas nesta ampliação — referências provisórias', 1)
paragraphs = d.paragraphs
start = next(i for i,p in enumerate(paragraphs) if p.text.startswith('1. O CRIME DE ESTELIONATO'))
end = next(i for i,p in enumerate(paragraphs) if p.text.startswith('2. A NATUREZA JURÍDICA'))
anchor = paragraphs[start]._p
template = paragraphs[start + 1]
for p in paragraphs[start + 1:end]:
    p._element.getparent().remove(p._element)
for block in body.split('\n\n')[1:]:
    block = block.strip()
    if not block:
        continue
    if block.startswith('## '):
        new = OxmlElement('w:p')
        anchor.addnext(new)
        p = Paragraph(new, template._parent)
        p.style = template.style
        run = p.add_run(block[3:])
        run.bold = True
    else:
        new = deepcopy(template._p)
        anchor.addnext(new)
        p = Paragraph(new, template._parent)
        for run in p.runs:
            run.text = ''
        if p.runs:
            p.runs[0].text = block.replace('**', '')
        else:
            p.add_run(block.replace('**', ''))
    anchor = new
ps = d.paragraphs
idx = next(i for i,p in enumerate(ps) if p.text.strip() == 'REFERÊNCIAS DAS FONTES CITADAS')
last = ps[-1]
for block in refs.strip().split('\n\n'):
    if '2013/0343206-7' in block or '2024/0312514-9' in block or '15.438' in block:
        new = deepcopy(last._p)
        last._p.addnext(new)
        last = Paragraph(new, last._parent)
        for run in last.runs:
            run.text = ''
        if last.runs:
            last.runs[0].text = block.replace('**', '')
        else:
            last.add_run(block.replace('**', ''))
d.save(root / 'Artigo estelionato UNIVALI - capítulo 1 ampliado.docx')
