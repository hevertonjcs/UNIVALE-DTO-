from pathlib import Path
from copy import deepcopy
from docx import Document
from docx.text.paragraph import Paragraph

root = Path('AULAS')
doc = Document(root / 'Artigo estelionato UNIVALI - capítulo 3.docx')
text = (root / 'Artigo estelionato - introdução e capítulo 1.md').read_text(encoding='utf-8')
def section(start, end):
    return text.split(start, 1)[1].split(end, 1)[0].strip()
abstract = section('## RESUMO\n', '\n## ABSTRACT')
english = section('## ABSTRACT\n', '\n## INTRODUÇÃO')
conclusion = section('## CONSIDERAÇÕES FINAIS\n', '\n## REFERÊNCIAS')
def replace(p, s):
    for run in p.runs:
        run.text = ''
    if p.runs:
        p.runs[0].text = s
    else:
        p.add_run(s)
ps = doc.paragraphs
def index(label):
    return next(i for i,p in enumerate(ps) if p.text.strip() == label)
summary, keywords = abstract.split('\n\n', 1)
english_summary, english_keywords = english.split('\n\n', 1)
replace(ps[index('RESUMO')+1], summary)
replace(next(p for p in ps if p.text.strip().startswith('Palavras-chave:')), keywords.replace('**', ''))
replace(ps[index('RESUMEN ou ABSTRACT')], 'ABSTRACT')
replace(ps[index('RESUMEN ou ABSTRACT')+1] if False else ps[index('ABSTRACT')+1], english_summary)
replace(next(p for p in ps if p.text.strip().startswith('Keywords:')), english_keywords.replace('**', ''))
ps = doc.paragraphs
idx = next(i for i,p in enumerate(ps) if p.text.strip() == 'CONSIDERAÇÕES FINAIS')
p = ps[idx+1]
parts = conclusion.split('\n\n')
replace(p, parts[0].replace('**', ''))
current = p._p
for part in parts[1:]:
    new = deepcopy(p._p)
    current.addnext(new)
    q = Paragraph(new, p._parent)
    replace(q, part.replace('**', ''))
    current = new
doc.save(root / 'Artigo estelionato UNIVALI - versão preliminar completa.docx')
