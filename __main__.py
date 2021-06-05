from openpyxl import Workbook, load_workbook

receber = load_workbook('receber.xlsx', read_only=False)

#LIMPA O CONTEÚDO DAS COLUNAS A-O (DADOS) DE UMA PLANILHA(ARG0)
def purge_data(ws):
    for r in range (7,150):
        for c in range (1, 16):
            print(ws , 'has been purged')

#LIMPEZA DAS PLANILHAS À VENCER E VENCIDO DE AMBAS AS CARTEIRAS(09/28)
purge_data(receber['VE09BRAD'])
purge_data(receber['AV09BRAD'])
purge_data(receber['VE28BRAD'])
purge_data(receber['AV28BRAD'])
receber.save(filename='receber.xlsx')


