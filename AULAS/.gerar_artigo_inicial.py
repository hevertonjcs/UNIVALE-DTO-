from pathlib import Path
from copy import deepcopy
from docx import Document

root = Path('AULAS')
d = Document(root / 'Modelo UNIVALI - cópia de trabalho.docx')
source = (root / 'Artigo estelionato - introdução e capítulo 1.md').read_text(encoding='utf-8')
intro = source.split('## INTRODUÇÃO\n\n', 1)[1].split('\n\n## 1.', 1)[0]
chapter = source.split('## 1. O CRIME', 1)[1].split('\n\n## 2.', 1)[0].split('\n\n', 1)[1]
chapter2 = source.split('## 2. A NATUREZA', 1)[1].split('\n\n## REFERÊNCIAS', 1)[0].split('\n\n', 1)[1]
refs = source.split('## REFERÊNCIAS PROVISÓRIAS DAS FONTES CITADAS\n\n', 1)[1]
ps = d.paragraphs
def replace(p, value):
    for run in p.runs:
        run.text = ''
    if p.runs:
        p.runs[0].text = value
    else:
        p.add_run(value)
def paragraphs(p, value):
    parts = value.strip().split('\n\n')
    replace(p, parts[0].replace('*', ''))
    current = p._p
    for s in parts[1:]:
        new = deepcopy(p._p)
        current.addnext(new)
        from docx.text.paragraph import Paragraph
        q = Paragraph(new, p._parent)
        replace(q, s.replace('*', ''))
        current = new
replace(ps[0], 'A REVOGAÇÃO DO § 5º DO ARTIGO 171 DO CÓDIGO PENAL: A NATUREZA JURÍDICA DA REPRESENTAÇÃO E OS EFEITOS INTERTEMPORAIS DA LEI Nº 15.397/2026')
replace(ps[2], 'THE REPEAL OF ARTICLE 171, PARAGRAPH 5, OF THE BRAZILIAN PENAL CODE: THE LEGAL NATURE OF THE VICTIM’S REPRESENTATION AND THE TEMPORAL EFFECTS OF LAW NO. 15,397/2026')
replace(ps[4], 'Heverton Junior Cristaldo da Silva')
replace(ps[5], '')
replace(ps[7], '')
replace(ps[9], 'Palavras-chave:')
replace(ps[11], '')
replace(ps[13], 'Keywords:')
paragraphs(ps[16], intro)
replace(ps[18], '1. O CRIME DE ESTELIONATO E A EXIGÊNCIA DE REPRESENTAÇÃO DA VÍTIMA NO DIREITO PENAL BRASILEIRO')
paragraphs(ps[19], chapter)
replace(ps[21], '2. A NATUREZA JURÍDICA DA REPRESENTAÇÃO E A APLICAÇÃO DA LEI PENAL E PROCESSUAL PENAL NO TEMPO')
paragraphs(ps[22], chapter2)
replace(ps[24], '3. OS EFEITOS INTERTEMPORAIS DA LEI Nº 15.397/2026: A REVOGAÇÃO DA REPRESENTAÇÃO NO ESTELIONATO E SUA APLICAÇÃO A FATOS ANTERIORES')
replace(ps[25], '')
replace(ps[28], '')
paragraphs(ps[31], refs)
d.save(root / 'Artigo estelionato UNIVALI - em desenvolvimento.docx')
