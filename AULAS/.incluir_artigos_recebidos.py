from pathlib import Path
from copy import deepcopy
from docx import Document
from docx.text.paragraph import Paragraph

root = Path('AULAS')
d = Document(root / 'Artigo estelionato UNIVALI - referências organizadas.docx')
def insert_after(p, value):
    node = deepcopy(p._p)
    p._p.addnext(node)
    q = Paragraph(node, p._parent)
    for r in q.runs:
        r.text = ''
    if q.runs:
        q.runs[0].text = value
    else:
        q.add_run(value)
    return q
def insert_before(p, value):
    node = deepcopy(p._p)
    p._p.addprevious(node)
    q = Paragraph(node, p._parent)
    for r in q.runs:
        r.text = ''
    if q.runs:
        q.runs[0].text = value
    else:
        q.add_run(value)
    return q
anchor = next(p for p in d.paragraphs if p.text.startswith('Segundo Oliveira e Silva (2026)'))
insert_after(anchor, 'Ao examinar o estelionato em ambiente digital, Meireles e Pasitto (2024, p. 6311) relacionam a prática à obtenção de vantagem mediante fraude e ao erro da vítima. O estudo se volta às dificuldades de prevenção e enfrentamento dos golpes virtuais. Por ser anterior à Lei nº 15.397/2026 e possuir objeto distinto, sua contribuição aqui se limita à contextualização da modalidade digital, sem tratar da revogação da representação.')
anchor = next(p for p in d.paragraphs if p.text.startswith('Divan e Santiago (2024) estudam a retroatividade'))
insert_after(anchor, 'Em estudo sobre o acordo de não persecução penal, Faraco Neto e Lopes (2020, p. 24–25) distinguem a regra de aplicação imediata das normas processuais da discussão sobre a retroatividade da norma mista mais favorável. Os autores relacionam o conteúdo material do acordo à possibilidade de extinção da punibilidade e defendem sua aplicação benéfica a processos em andamento. O argumento fornece uma comparação para o debate sobre normas mistas, mas se refere ao ANPP e não decide, por si, os efeitos da revogação da representação no estelionato.')
before_nucci = next(p for p in d.paragraphs if p.text.startswith('NUCCI, Guilherme'))
insert_before(before_nucci, 'FARACO NETO, Pedro; LOPES, Vinicius Basso. Acordo de não persecução penal: a retroatividade da lei penal mista e a possibilidade dos acordos após a instrução processual. Boletim IBCCRIM, ano 28, n. 331, p. 22–25, jun. 2020. Arquivo do boletim fornecido para a pesquisa.')
insert_before(before_nucci, 'MEIRELES, Isys Gonzaga; PASITTO, Fernando Teles. Estelionato e suas implicações: o constante crescimento dos golpes virtuais. Revista Ibero-Americana de Humanidades, Ciências e Educação, v. 10, n. 11, p. 6303–6316, nov. 2024. DOI: https://doi.org/10.51891/rease.v10i11.17063.')
d.save(root / 'Artigo estelionato UNIVALI - artigos analisados.docx')
