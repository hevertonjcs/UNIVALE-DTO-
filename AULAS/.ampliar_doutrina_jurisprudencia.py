from pathlib import Path
from copy import deepcopy
from docx import Document
from docx.text.paragraph import Paragraph

root = Path('AULAS')
doc = Document(root / 'Artigo estelionato UNIVALI - princípios e doutrina.docx')
text = (root / 'Aprofundamento - doutrina e jurisprudência do estelionato.md').read_text(encoding='utf-8')
first = text.split('## Representação e decadência na sucessão de leis', 1)[1].split('## Fatos anteriores, regra intermediária e exceções', 1)[0]
second = text.split('## Fatos anteriores, regra intermediária e exceções', 1)[1].split('## Fonte adicional', 1)[0]
def add_before(anchor, text):
    parts = text.strip().split('\n\n')
    template = next(p for p in doc.paragraphs if p.text.startswith('A sucessão de leis exige distinguir'))
    for part in parts:
        node = deepcopy(template._p)
        anchor._p.addprevious(node)
        p = Paragraph(node, anchor._parent)
        for run in p.runs:
            run.text = ''
        value = part.strip()
        heading = value.startswith('## ')
        if heading:
            value = value[3:]
        if p.runs:
            p.runs[0].text = value
            p.runs[0].bold = heading
        else:
            p.add_run(value).bold = heading
anchor2 = next(p for p in doc.paragraphs if p.text.startswith('3. OS EFEITOS INTERTEMPORAIS'))
add_before(anchor2, first)
anchor3 = next(p for p in doc.paragraphs if p.text.strip() == 'CONSIDERAÇÕES FINAIS')
add_before(anchor3, second)
last = doc.paragraphs[-1]
node = deepcopy(last._p)
last._p.addnext(node)
p = Paragraph(node, last._parent)
for run in p.runs:
    run.text = ''
p.runs[0].text = 'VOLPI, Murilo Alan; VOLPI, Matheus Tauan. Retroatividade da lei que alterou a natureza da ação penal nos crimes de estelionato. Boletim IBCCRIM, São Paulo, v. 28, n. 331, p. 27–30, jun. 2020. Disponível em: https://publicacoes.ibccrim.org.br/index.php/boletim_1993/article/view/548.'
doc.save(root / 'Artigo estelionato UNIVALI - aprofundamento doutrinário.docx')
