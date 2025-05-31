from docx import Document
import pandas as pd
from datetime import datetime
tabela = pd.read_excel('Informações.xlsx')
for linha in tabela.index:
    nome = tabela.loc[linha,'Nome']
    item1 = tabela.loc[linha,'Item1']
    item2 = tabela.loc[linha,'Item2']
    item3 = tabela.loc[linha,'Item3']
    print(nome)
    parametros = {
        'XXXX':nome,
        'YYYY':item1,
        'ZZZZ':item2,
        'WWWW':item3,
        'DD':datetime.today().day,
        'MM':datetime.today().month,
        'AAAA':datetime.today().year

    }
    #documento
    documento = Document('Contrato.docx')
    for paragrafo in documento.paragraphs:
        for codigo in parametros:
            paragrafo.text = paragrafo.text.replace(str(codigo),str(parametros[codigo]))
        documento.save(f'Contrato-{nome}.docx')

