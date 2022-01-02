import pandas as pd

#Passo a passo da solucao

#Abrir o 6 arquivos em Excel

listaMeses = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho']

for mes in listaMeses:
    tabelaVendas = pd.read_excel(f'{mes}.xlsx')
    if (tabelaVendas['Vendas'] > 55000).any():
        vendedor = tabelaVendas.loc[tabelaVendas['Vendas'] > 55000, 'Vendedor'].values[0]
        vendas = tabelaVendas.loc[tabelaVendas['Vendas'] > 55000, 'Vendas'].values[0]
        print(f'No mes {mes}, o vendedor: {vendedor}, vendeu: {vendas}.')

#Para cada arquivo:

#Verificar se algum valor na coluna VENDAS daquele arquivo e maior que 55.000

#Se for maior do que 55.000 -> Mostrar o Nome, o mes e as vendas do vendedor

#Caso nao seja maior do que 55.000 nao fazer nada