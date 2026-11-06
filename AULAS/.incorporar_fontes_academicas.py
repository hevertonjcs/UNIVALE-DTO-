from pathlib import Path
from copy import deepcopy
from docx import Document
from docx.text.paragraph import Paragraph

root = Path('AULAS')
d = Document(root / 'Artigo estelionato UNIVALI - capítulo 1 ampliado.docx')
def after(p, value):
    new = deepcopy(p._p)
    p._p.addnext(new)
    q = Paragraph(new, p._parent)
    for r in q.runs:
        r.text = ''
    if q.runs:
        q.runs[0].text = value
    else:
        q.add_run(value)
    return q
ps = d.paragraphs
anchor = next(p for p in ps if p.text.startswith('A Lei nº 15.397/2026 alterou também'))
after(anchor, 'Segundo Oliveira e Silva (2026), a Lei nº 14.155/2021 integra o quadro normativo de enfrentamento ao phishing e ao estelionato eletrônico, embora persistam desafios de investigação e responsabilização penal. As autoras estudam esse contexto mediante pesquisa bibliográfica e documental. Essa abordagem contribui para situar uma modalidade de fraude, sem resolver, por si, a controvérsia sobre a revogação da representação em 2026.')
ps = d.paragraphs
anchor = next(p for p in ps if p.text.startswith('Brandão (2023, p. 30 e 34) sustenta'))
new = after(anchor, 'Segundo Nucci (2020), a exigência de representação introduzida em 2019 possui reflexos sobre a punibilidade em razão da decadência. Para o autor, essa dimensão material justifica a aplicação retroativa benéfica da nova condição de procedibilidade a processos em andamento. Nucci distingue a condição de procedibilidade da condição de prosseguibilidade e admite a verificação da vontade da vítima quando ela não estiver clara nos autos. Sua análise se refere à introdução da exigência, e não à revogação de 2026.')
after(new, 'Divan e Santiago (2024) estudam a retroatividade do acordo de não persecução penal e recorrem, em sua argumentação, ao debate sobre a exigência de representação no estelionato introduzida em 2019. Os autores identificam aproximações na discussão da lei mais favorável, mas também assinalam diferenças entre os institutos. A comparação contribui para compreender o debate sobre normas mistas; não corresponde a uma conclusão dos autores acerca da revogação de 2026.')
last = d.paragraphs[-1]
for ref in [
'DIVAN, Gabriel Antinolfi; SANTIAGO, Nestor Eduardo Araruna. Acordo de Não Persecução Penal como instrumento político-criminal: possibilidades, reconfigurações jurisprudenciais vinculantes e os novos rumos do processo penal brasileiro. Revista Brasileira de Direito Processual Penal, v. 10, n. 1, 2024. DOI: https://doi.org/10.22197/rbdpp.v10i1.920.',
'NUCCI, Guilherme de Souza. Ação pública condicionada no estelionato e a retroatividade da lei processual benéfica. Migalhas, 23 out. 2020. Disponível em: https://www.migalhas.com.br/depeso/335360/acao-publica-condicionada-no-estelionato-e-a-retroatividade-da-lei-processual-benefica.',
'OLIVEIRA, Rafaele Louise de Souza Silva; SILVA, Rosana Reis de Melo. Phishing e estelionato eletrônico: desafios da investigação e da responsabilidade penal após a Lei 14.155/2021. Revista Ibero-Americana de Humanidades, Ciências e Educação, v. 12, n. 6, p. 1–21, 2026. DOI: https://doi.org/10.51891/rease.v12i6.27553.'
]:
    last = after(last, ref)
d.save(root / 'Artigo estelionato UNIVALI - fontes acadêmicas.docx')
