from pathlib import Path
from docx import Document

folder = Path('AULAS')
doc = Document(folder / 'Artigo estelionato UNIVALI - artigos analisados.docx')
def replace(p, value):
    for run in p.runs:
        run.text = ''
    if p.runs: p.runs[0].text = value
    else: p.add_run(value)

changes = {
    'Na ação penal pública, a iniciativa pertence': None,
    'A ação penal pública é promovida pelo Ministério Público.': 'A ação penal pública é promovida pelo Ministério Público. Em alguns casos, a lei exige que a vítima, ou outra pessoa legalmente legitimada, manifeste sua vontade de que o fato seja apurado: essa manifestação é chamada de representação. Ela não se confunde com a denúncia, que é apresentada pelo Ministério Público para propor a ação penal (BRASIL, 1941, arts. 24 e 39).',
    'Quando exigida, a representação atua como condição de procedibilidade.': 'Quando exigida, a representação funciona como condição de procedibilidade: é um requisito necessário para que a ação penal pública possa ser proposta. A necessidade de manifestação da vítima não transforma a ação em privada; a iniciativa da ação continua com o Ministério Público. A existência ou ausência da representação deve ser examinada de acordo com a lei aplicável e as circunstâncias do caso.',
    'O exercício da representação está sujeito a prazo.': 'O direito de representação deve ser exercido dentro do prazo previsto em lei. Como regra, o art. 103 do Código Penal estabelece seis meses contados do conhecimento da autoria, salvo disposição expressa em contrário. A perda desse direito pelo decurso do prazo é chamada de decadência e constitui causa de extinção da punibilidade, conforme o art. 107, IV, do Código Penal. Não se confunde com a prescrição, que possui regras próprias relativas à pretensão punitiva ou executória (BRASIL, 1940).',
    'Essa distinção não resolve, por si, a classificação de uma disposição': 'Algumas normas, chamadas mistas ou híbridas, reúnem efeitos processuais e materiais. No caso da representação, o aspecto processual está em funcionar como requisito da ação penal; o aspecto material decorre de sua relação com a decadência, que pode extinguir a punibilidade. Essa classificação é relevante porque pode influenciar a escolha da regra de aplicação da lei no tempo.',
}
remove_prefixes = [
    'Segundo Nucci (2020), a representação introduzida no estelionato em 2019 não pode',
    'Brandão (2023, p. 30 e 34) também identifica efeitos materiais',
    'O art. 5º, XL, da Constituição dispõe sobre a irretroatividade da lei penal',
    'No AgRg no AgRg no AREsp 2.468.274/GO, a Quinta Turma do STJ manteve',
    'A existência de manifestação anterior é uma questão distinta da abertura',
    'Na decisão de 2024, o STJ descreveu o HC 208.817 AgR/RJ',
    'Divan e Santiago (2024), em estudo sobre o acordo de não persecução penal, utilizam',
    'Para fatos ocorridos durante a vigência do § 5º, a análise deve verificar',
    'Em relação a fatos anteriores à entrada em vigor da Lei nº 13.964/2019, há uma sequência distinta',
    'O efeito da decadência já consumada e a suficiência de manifestações da vítima também dependem',
]
for p in list(doc.paragraphs):
    text = p.text.strip()
    if any(text.startswith(x) for x in remove_prefixes):
        p._element.getparent().remove(p._element)
        continue
    for prefix, value in changes.items():
        if value is not None and text.startswith(prefix):
            replace(p, value)
            break
doc.save(folder / 'Artigo estelionato UNIVALI - revisão para leitura.docx')
