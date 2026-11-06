from pathlib import Path
from copy import deepcopy
from docx import Document
from docx.text.paragraph import Paragraph

base = Path('AULAS')
doc = Document(base / 'Artigo estelionato UNIVALI - em desenvolvimento.docx')
text = (base / 'Artigo estelionato - introdução e capítulo 1.md').read_text(encoding='utf-8')
chapter = text.split('## 3. OS EFEITOS', 1)[1].split('\n\n## REFERÊNCIAS', 1)[0].split('\n\n', 1)[1]
refs = text.split('## REFERÊNCIAS PROVISÓRIAS DAS FONTES CITADAS\n\n', 1)[1]
ps = doc.paragraphs
heading = next(i for i, p in enumerate(ps) if p.text.startswith('3. OS EFEITOS INTERTEMPORAIS'))
target = ps[heading + 1]
def replace(p, s):
    for r in p.runs:
        r.text = ''
    if p.runs:
        p.runs[0].text = s
    else:
        p.add_run(s)
def insert_paragraphs(p, chunks):
    replace(p, chunks[0].replace('*', ''))
    current = p._p
    for s in chunks[1:]:
        new = deepcopy(p._p)
        current.addnext(new)
        q = Paragraph(new, p._parent)
        replace(q, s.replace('*', ''))
        current = new
insert_paragraphs(target, chapter.strip().split('\n\n'))
ps = doc.paragraphs
ref = next(i for i, p in enumerate(ps) if p.text.strip() == 'REFERÊNCIAS DAS FONTES CITADAS')
existing = '\n'.join(p.text for p in ps[ref+1:])
last = ps[-1]
for block in refs.strip().split('\n\n'):
    if block.startswith('SANTOS, Erinaldo') or 'nº 3.303.220/SP' in block:
        if block[:35] not in existing:
            new = deepcopy(last._p)
            last._p.addnext(new)
            last = Paragraph(new, last._parent)
            replace(last, block.replace('*', ''))
doc.save(base / 'Artigo estelionato UNIVALI - capítulo 3.docx')
