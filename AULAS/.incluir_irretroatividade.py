from pathlib import Path
from copy import deepcopy
from docx import Document
from docx.text.paragraph import Paragraph

root = Path('AULAS')
d = Document(root / 'Artigo estelionato UNIVALI - fontes acadêmicas.docx')
def after(p, text):
    new = deepcopy(p._p)
    p._p.addnext(new)
    q = Paragraph(new, p._parent)
    for r in q.runs:
        r.text = ''
    if q.runs:
        q.runs[0].text = text
    else:
        q.add_run(text)
    return q
anchor = next(p for p in d.paragraphs if p.text.startswith('Para fatos praticados durante a vigência do § 5º'))
a = after(anchor, 'A vedação à retroatividade da lei penal mais gravosa decorre do art. 5º, XL, da Constituição Federal, enquanto o art. 2º, parágrafo único, do Código Penal prevê a aplicação da lei posterior mais favorável. Sob a premissa de que a exigência de representação possui dimensão material e de que sua retirada prejudica o acusado no caso examinado, esses dispositivos fundamentam a preservação da regra anterior mais benéfica para fatos pretéritos aos quais ela era aplicável. Isso não impede a vigência da lei posterior para situações que ela possa reger legitimamente, nem dispensa a análise das exceções e da sucessão de normas (BRASIL, 1940; BRASIL, 1988).')
after(a, 'Ao examinar especificamente a revogação de 2026, Costa (2026) sustenta que a retirada da condição de representação configura alteração prejudicial e, por isso, não deve retroagir. Trata-se de posicionamento do autor, consonante com a hipótese investigada, e não de tese judicial cuja consolidação tenha sido demonstrada neste trabalho.')
anchor = next(p for p in d.paragraphs if p.text.startswith('Segundo Nucci (2020), a exigência de representação'))
after(anchor, 'Na apresentação feita por Brandão (2023, p. 33–34), Aury Lopes Jr. defende uma leitura constitucional das normas processuais penais, considerando a repercussão da sucessão legislativa sobre a situação do acusado. Como a obra original de Lopes Jr. não foi consultada diretamente, essa posição é aqui registrada a partir da exposição de Brandão, referente ao debate anterior à revogação de 2026.')
last = d.paragraphs[-1]
after(last, 'COSTA, Fernando José da. Direito intertemporal e a revogação da condição de procedibilidade. Veja, 4 set. 2026. Disponível em: https://veja.abril.com.br/coluna/radar-juridico/direito-intertemporal-e-a-revogacao-da-condicao-de-procedibilidade/.')
d.save(root / 'Artigo estelionato UNIVALI - princípios e doutrina.docx')
