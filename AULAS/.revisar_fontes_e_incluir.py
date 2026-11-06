from pathlib import Path
from copy import deepcopy
from docx import Document
from docx.text.paragraph import Paragraph

folder = Path('AULAS')
doc = Document(folder / 'Artigo estelionato UNIVALI - aprofundamento doutrinário.docx')
def settext(p, value):
    for r in p.runs: r.text = ''
    if p.runs: p.runs[0].text = value
    else: p.add_run(value)
def after(p, value):
    element = deepcopy(p._p)
    p._p.addnext(element)
    q = Paragraph(element, p._parent)
    settext(q, value)
    return q
for p in doc.paragraphs:
    t = p.text
    if t.startswith('Em julgamento sobre o art. 171, o Superior'):
        settext(p, 'No Conflito de Competência nº 143.621/PR, o ministro Rogerio Schietti Cruz examinou os elementos do estelionato em caso de cheque clonado. Em seu voto, destacou a relação entre fraude, erro, obtenção de vantagem ilícita e prejuízo alheio. O julgamento tratava de competência territorial, circunstância que limita sua utilização para o problema intertemporal deste artigo (STJ, 2016).')
    elif t.startswith('No precedente referido, o STJ tratou'):
        settext(p, 'No voto proferido no CC 143.621/PR, o relator descreveu a modalidade fundamental de estelionato como crime material e analisou a vantagem ilícita e o prejuízo como elementos relevantes para a consumação (STJ, 2016). A distinção entre consumação e tentativa exige observar os requisitos legais e as particularidades do caso, sem presumir a consumação apenas pela presença de uma conduta enganosa.')
    elif t.startswith('Em decisão relativa a uma transação contratual, o STJ examinou'):
        settext(p, 'No AgRg no AREsp 411.987/SP, o STJ registrou que a corte de origem reconhecera o estelionato a partir das provas do processo. Ao negar o agravo, o tribunal considerou inviável reexaminar o conjunto fático-probatório em recurso especial. O caso ilustra a necessidade de distinguir os fatos reconhecidos pela instância de origem daquilo que o STJ efetivamente decidiu (STJ, 2014).')
    elif t.startswith('STJ. Superior Tribunal de Justiça. Acórdão, registro 2013/034'):
        settext(p, 'STJ. Superior Tribunal de Justiça. AgRg no AREsp nº 411.987/SP. Relatora Ministra Marilza Maynard, Sexta Turma, julgado em 17 dez. 2013, publicado em 13 fev. 2014. Disponível em: https://scon.stj.jus.br/SCON/GetInteiroTeorDoAcordao?dt_publicacao=13%2F02%2F2014&num_registro=201303432067.')
    elif t.startswith('STJ. Superior Tribunal de Justiça. Acórdão, registro 2024/031'):
        settext(p, 'STJ. Superior Tribunal de Justiça. Conflito de Competência nº 143.621/PR. Relator Ministro Rogerio Schietti Cruz, Terceira Seção, julgado em 25 maio 2016, publicado em 7 jun. 2016. Disponível em: https://scon.stj.jus.br/SCON/GetInteiroTeorDoAcordao?dt_publicacao=07%2F06%2F2016&num_registro=201502572320.')
    else:
        new = t.replace('(STJ, 2025b)', '(STJ, 2025)').replace('(BRASIL, 2026b)', '(BRASIL, 2026b)').replace('(BRASIL, 2026)', '(BRASIL, 2026a)').replace('(BRASIL, 2019; BRASIL, 2026)', '(BRASIL, 2019; BRASIL, 2026a)').replace('(BRASIL, 1940; BRASIL, 2026)', '(BRASIL, 1940; BRASIL, 2026a)').replace('(BRASIL, 2019; BRASIL, 2025; BRASIL, 2026)', '(BRASIL, 2019; BRASIL, 2025; BRASIL, 2026a)')
        if new != t: settext(p, new)
        if t.startswith('BRASIL. Lei nº 15.397, de 30 de abril de 2026.'):
            settext(p, t.replace('de 30 de abril de 2026.', 'de 30 de abril de 2026. [2026a].', 1))
        elif t.startswith('BRASIL. Lei nº 15.438, de 18 de junho de 2026.'):
            settext(p, t.replace('de 18 de junho de 2026.', 'de 18 de junho de 2026. [2026b].', 1))
anchor = next(p for p in doc.paragraphs if p.text.startswith('A sucessão de leis exige distinguir'))
a = after(anchor, 'A distinção entre retroatividade e ultratividade é útil para identificar o período de incidência de uma norma. A retroatividade refere-se à aplicação de lei posterior a fato anterior; a ultratividade, à aplicação de lei que já deixou de vigorar a fato ocorrido sob sua vigência. No caso de uma lei intermediária, vigente entre a prática do fato e o julgamento, discute-se a aplicação de uma regra que não estava em vigor em nenhum desses dois momentos. Para normas penais materiais, o art. 2º, parágrafo único, do Código Penal determina a aplicação da lei posterior que favoreça o agente (BRASIL, 1940).')
after(a, 'A distinção exige cautela neste estudo: afirmar que o antigo § 5º pode ter efeitos após sua revogação depende de demonstrar sua natureza jurídica, sua maior favorabilidade e sua aplicabilidade ao caso. O § 5º também continha exceções e sofreu alteração em 2025. Não basta identificar cronologicamente uma lei intermediária para dispensar a verificação dessas condições e do alcance dos precedentes.')
anchor = next(p for p in doc.paragraphs if p.text.startswith('A análise é diferente para fatos anteriores à entrada em vigor'))
a = after(anchor, 'Para identificar o tempo do crime, o art. 4º do Código Penal adota como referência a ação ou omissão, ainda que o resultado ocorra em outro momento. Isso não deve ser confundido com a discussão sobre consumação nem com as regras de competência territorial. A eventual aplicação dessa referência a condutas continuadas ou permanentes requer análise específica e não pode ser resolvida apenas pela leitura de um exemplo isolado (BRASIL, 1940).')
after(a, 'Em uma hipótese de estudo, se a conduta ocorreu quando a regra de representação era exigível e não incidia exceção, a investigação posterior deverá considerar se houve manifestação válida e se o prazo decadencial transcorreu. Já para conduta anterior à Lei nº 13.964/2019, acrescenta-se o problema da eventual incidência da regra como lei intermediária mais favorável. São cenários hipotéticos, não decisões judiciais.')
last = doc.paragraphs[-1]
after(last, 'STJ. Superior Tribunal de Justiça. AgRg no AREsp nº 3.303.220/SP. Relator Ministro Ribeiro Dantas, Quinta Turma, julgado em 18 ago. 2026, publicado em 25 ago. 2026. Disponível em: https://scon.stj.jus.br/SCON/GetInteiroTeorDoAcordao?dt_publicacao=25%2F08%2F2026&num_registro=202602547159.')
doc.save(folder / 'Artigo estelionato UNIVALI - fontes revisadas.docx')
